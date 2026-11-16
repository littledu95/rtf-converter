# -*- coding: utf-8 -*-
""" RTF 批量转换服务（纯 API 形态，多进程版）
================================
引擎：Microsoft Word COM，格式保真度最高
架构：Flask 主进程只管 HTTP/队列/监管；每个转换 worker 是独立子进程
      （spawn），各自独享 COM STA + 独立 Word 实例
并发：WORKERS 个子进程并行消费；队列有上限，超限直接 503
进度：POST /api/convert -> GET /api/task/<id> 轮询
下载：GET /api/task/<id>/download 打包 zip

为什么用子进程而不是线程：
    - Python 无法 kill 卡死的线程；卡死的 Word 会把宿主线程永久挂起，
      线程方案只能"放弃线程再建一个"，旧线程和 Word 实例会不断累积
    - 子进程可以被 terminate() 强杀回收（连同它的 WINWORD.EXE），
      HTTP 服务始终存活，多人并发互不影响

Word 生命周期（每个子进程独立管理自己的实例）：
    - 懒加载：worker 启动不拉 Word，真有任务才启动
    - 空闲释放：队列空 WORD_IDLE_TIMEOUT 秒 -> Quit()
    - 主动回收：每 WORD_RECYCLE_FILES 个文件 / WORD_RECYCLE_SECONDS 秒换新实例
    - RPC 自愈：-2147023174(0x800706BA) 等通道断开时自动重建并重试单文件
    - 可选每任务一实例：WORD_PER_TASK=1，任务结束立刻释放，全程零常驻

主进程监管（supervisor + watchdog）：
    - 子进程通过消息队列上报：心跳 / 进度 / Word 状态 / 结果
    - 健康、释放等 HTTP 接口只读子进程上报的普通状态，绝不跨进程碰 Word 对象
    - 某 worker 的当前任务超过 WORD_STUCK_TIMEOUT 无心跳 -> terminate 子进程
      + 按 PID 强杀其 WINWORD.EXE -> 只标记该任务失败 -> 自动拉起替补子进程

清理策略（延迟删除）：
    - uploads_* 目录：上传完成后 UPLOAD_EXPIRE_DAYS 天过期
    - tmp_* 临时输出目录（未下载/失败）：TMP_TASK_EXPIRE_DAYS 天过期
    - tmp_* 临时输出目录（下载完成后）：DOWNLOAD_CLEAN_DELAY 秒后延迟删除
    - 所有清理由单一后台线程 _auto_clean_all 统一扫描执行
"""
import gc
import os
import queue
import tempfile
import zipfile
import logging
import threading
import shutil
import time
import uuid
import ctypes
import multiprocessing
import subprocess
from contextlib import contextmanager
from datetime import datetime, timedelta

import pythoncom
import win32com.client
from flask import Flask, Response, jsonify, request, send_file, json

# ===================== 配置（环境变量可覆盖） =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8000"))
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", os.path.join(BASE_DIR, "output"))
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", os.path.join(OUTPUT_DIR, "uploads"))
MAX_FILES = int(os.environ.get("MAX_FILES", "500"))
API_TOKEN = os.environ.get("API_TOKEN", "")
DEFAULT_FORMATS = ["doc", "docx", "pdf"]
FORMAT_MAP = {"doc": 0, "docx": 12, "pdf": 17}
LARGE_RTF_BYTES = int(os.environ.get("LARGE_RTF_BYTES", str(50 * 1024 * 1024)))
# 每个 worker 是一个独立子进程（独立 COM STA + 独立 Word 实例）。
# 多 worker 提升多人并发吞吐；Word 实例间会争抢配置/加载项，数值按机器实测调整。
WORKERS = max(1, int(os.environ.get("WORKERS", "2")))
# 任务队列上限：待处理任务超过该数量直接 503，防止并发提交打爆内存/磁盘
QUEUE_MAX = int(os.environ.get("QUEUE_MAX", "100"))
# 单次上传请求总大小上限（MB），超过返回 413
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "2048"))
INSTANCE_MUTEX_NAME = "Global\\RTFConvertServer_%s" % PORT

# ---- Word 实例生命周期配置 ----
# 空闲多少秒后释放 Word 实例（0=不自动释放，常驻）
WORD_IDLE_TIMEOUT = int(os.environ.get("WORD_IDLE_TIMEOUT", "300"))
# 队列空转检测间隔（秒）
WORD_POLL_INTERVAL = int(os.environ.get("WORD_POLL_INTERVAL", "5"))
# 每任务一实例：任务结束立刻 Quit，全程零常驻（冷启动开销换绝对干净）
WORD_PER_TASK = os.environ.get("WORD_PER_TASK", "0").strip().lower() in ("1", "true", "on", "yes")
# 保险栓：释放后仍残留 WINWORD.EXE 时强杀（会误杀用户自己开的 Word，默认关）
WORD_KILL_STALE = os.environ.get("WORD_KILL_STALE", "0").strip().lower() in ("1", "true", "on", "yes")

# PDF 书签来源：Word 的显式书签（2）或标题/大纲级别（1）。
# RTF 中的 \bkmkstart/\bkmkend 书签属于 Word 书签，默认使用 2 才能保留。
# 设为 1 可改为根据标题样式生成 PDF 书签，设为 0 禁用。
PDF_CREATE_BOOKMARKS = int(os.environ.get("PDF_CREATE_BOOKMARKS", "1"))

# ---- RPC 自愈配置 ----
# 单文件 RPC 失败后重建 Word 并重试的次数
RPC_RETRY = int(os.environ.get("RPC_RETRY", "2"))
# 每次重建后等待秒数
RPC_RETRY_SLEEP = float(os.environ.get("RPC_RETRY_SLEEP", "1.5"))
# 重建 / 启动 Word 的最大尝试次数
RPC_WORD_RELAUNCH_TRIES = int(os.environ.get("RPC_WORD_RELAUNCH_TRIES", "3"))
# 看门狗：worker 正在处理任务(busy>0)且超过这么久没有心跳 -> 判定 Word 卡死
# （Word 弹出模态对话框时 COM 调用会永久挂起，唯一 worker 一挂全站 pending）
WORD_STUCK_TIMEOUT = int(os.environ.get("WORD_STUCK_TIMEOUT", "600"))
# 卡死 worker 被放弃后，隔多少秒拉起替补 worker
WORD_WORKER_RESTART_DELAY = int(os.environ.get("WORD_WORKER_RESTART_DELAY", "5"))
# 每转换多少个文件主动重启 Word（0=不启用）
WORD_RECYCLE_FILES = int(os.environ.get("WORD_RECYCLE_FILES", "50"))
# 或每多少秒主动重启 Word（0=不启用）
WORD_RECYCLE_SECONDS = int(os.environ.get("WORD_RECYCLE_SECONDS", "3600"))

# ---- 清理策略配置 ----
UPLOAD_EXPIRE_DAYS = int(os.environ.get("UPLOAD_EXPIRE_DAYS", "3"))       # uploads 保留天数
TMP_TASK_EXPIRE_DAYS = int(os.environ.get("TMP_TASK_EXPIRE_DAYS", "20"))  # 未下载/失败 tmp_ 与 rtf_wps_ 保留天数
DOWNLOAD_CLEAN_DELAY = int(os.environ.get("DOWNLOAD_CLEAN_DELAY", str(3 * 86400)))  # 下载后延迟删除（秒，默认3天）
CLEAN_SCAN_INTERVAL = int(os.environ.get("CLEAN_SCAN_INTERVAL", "600"))   # 清理线程扫描间隔（秒，默认10分钟）
CLEANUP_HOUR = int(os.environ.get("CLEANUP_HOUR", "2"))                    # 每天统一清理时间（本机时间）

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(BASE_DIR, "server.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("rtf-server")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


def _is_child_process():
    """是否运行在 multiprocessing 拉起的转换子进程里。"""
    return multiprocessing.current_process().name != "MainProcess"


def _acquire_instance_mutex():
    """Allow only one app.py process to own a port on this Windows host."""
    if os.name != "nt" or _is_child_process():
        return None
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, INSTANCE_MUTEX_NAME)
    if not mutex:
        raise RuntimeError("无法创建服务实例锁")
    if ctypes.windll.kernel32.GetLastError() == 183:
        print("服务已在运行：%s" % INSTANCE_MUTEX_NAME)
        raise SystemExit(2)
    return mutex


_INSTANCE_MUTEX = _acquire_instance_mutex()


# ===================== RPC 故障识别 =====================
# -2147023174 = 0x800706BA = RPC_S_SERVER_UNAVAILABLE
# 含义：COM 对象背后的 WINWORD.EXE 已不在（崩溃/被杀/自行退出），
# 但进程内仍持有僵尸引用，后续每次调用都会抛"RPC 服务器不可用"。
RPC_HRESULTS = {
    -2147023174,   # 0x800706BA  RPC 服务器不可用
    -2147023170,   # 0x800706BE  远程过程调用失败
    -2147023169,   # 0x800706BF  RPC 调用失败
    -2147417848,   # 0x80010108  被调用的对象已与其客户端断开连接
    -2147417842,   # 0x8001010E  RPC_E_WRONG_THREAD
}
RPC_TEXT_MARKERS = (
    "-2147023174", "-2147023170", "-2147023169", "-2147417848", "-2147417842",
    "0x800706ba", "0x800706be", "0x800706bf", "0x80010108", "0x8001010e",
    "rpc 服务器不可用", "rpc server is unavailable",
    "远程过程调用失败", "已与其客户端断开连接",
)


def _is_rpc_error(exc):
    """判断异常/错误串是否属于 COM 通道断开类故障。"""
    if exc is None:
        return False
    hr = getattr(exc, "hresult", None)
    if hr is None:
        args = getattr(exc, "args", None)
        if args and isinstance(args[0], int):
            hr = args[0]
    try:
        if hr is not None and int(hr) in RPC_HRESULTS:
            return True
    except Exception:
        pass
    text = ("%s" % exc).lower()
    return any(m in text for m in RPC_TEXT_MARKERS)


def _results_have_rpc(results):
    """_save_document 会把异常吞进 results，这里补一层检查。"""
    if not results:
        return False
    for r in results.values():
        if not r.get("ok") and _is_rpc_error(r.get("error")):
            return True
    return False


def _safe_remove(path):
    for attempt in range(3):
        try:
            if not os.path.exists(path):
                return True
            os.remove(path)
            return True
        except (PermissionError, OSError) as e:
            if attempt == 2:
                log.warning("文件删除失败（可能仍被占用）：%s: %s", path, e)
                return False
            time.sleep(0.25 * (attempt + 1))
        except Exception as e:
            log.warning("文件删除失败：%s: %s", path, e)
            return False
    return False


# ===================== 任务管理器（主进程侧） =====================
class QueueFullError(Exception):
    """任务队列已满"""


class TaskManager:
    _used_global = {}
    _used_global_lock = threading.Lock()

    def __init__(self):
        self._tasks = {}
        # 可重入锁：清理线程会在持有锁的情况下调用 is_path_active()（它也要这把锁）
        self._lock = threading.RLock()
        self._request_cache = {}
        self._request_locks = {}
        self._request_cache_ttl = 10 * 60
        # ---- 多进程 IPC（start() 里创建，只能在主进程调用）----
        self._queue = None          # 主进程 -> 子进程：任务分发
        self._result_queue = None   # 子进程 -> 主进程：心跳/进度/Word状态/结果
        self._large_lock = None     # 跨进程大文件互斥
        self._workers = {}          # wid -> {"proc","current_tid","last_beat","word_alive","word_pid","respawn_at"}
        self._started = False
        self._stopping = False

    def start(self):
        """在主进程里拉起 WORKERS 个转换子进程 + 监管线程。
        子进程各自独立初始化 COM（STA）、各持一个 Word 实例；
        卡死时整个子进程可被强杀回收，HTTP 服务不受影响。"""
        if self._started or _is_child_process():
            return
        ctx = multiprocessing.get_context("spawn")
        self._queue = ctx.Queue(maxsize=QUEUE_MAX)
        self._result_queue = ctx.Queue()
        self._large_lock = ctx.Lock()
        for wid in range(WORKERS):
            self._spawn_worker(wid)
        threading.Thread(target=self._supervisor_loop, daemon=True,
                         name="rtf-supervisor").start()
        self._started = True
        log.info("多进程架构已启动：%d 个转换子进程（各持独立 Word 实例），队列上限=%d，卡死超时=%d 秒",
                 WORKERS, QUEUE_MAX, WORD_STUCK_TIMEOUT)

    def _spawn_worker(self, wid):
        ctx = multiprocessing.get_context("spawn")
        p = ctx.Process(target=_worker_main,
                        args=(wid, self._queue, self._result_queue, self._large_lock),
                        daemon=True)
        p.start()
        self._workers[wid] = {
            "proc": p, "current_tid": None,
            "last_beat": time.time(),
            "word_alive": False, "word_pid": None,
            "respawn_at": 0.0,
        }
        log.info("[worker-%d] 转换子进程已启动 (pid=%s)", wid, p.pid)

    def submit(self, files, formats, out_dir, split_dirs=False):
        return self._submit([(src, base, list(formats)) for src, base in files],
                            formats, out_dir, split_dirs)

    def submit_map(self, items, out_dir, split_dirs=False):
        groups = {}
        for src, fmt in items:
            groups.setdefault(src, []).append(fmt)
        files3 = [(src, os.path.splitext(os.path.basename(src))[0], sorted(set(fmts)))
                  for src, fmts in groups.items()]
        formats = sorted({f for _, _, fmts in files3 for f in fmts})
        return self._submit(files3, formats, out_dir, split_dirs)

    def _submit(self, files3, formats, out_dir, split_dirs=False):
        tid = uuid.uuid4().hex[:12]
        is_temp = False
        if not out_dir:
            out_dir = os.path.join(OUTPUT_DIR, "tmp_" + tid)
            is_temp = True
        # 输出名改在主进程一次性分配（原来在 worker 线程里分配，
        # 多进程后各进程计数独立会撞名 -> 多进程必须由主进程统一分配）
        files4 = []
        with TaskManager._used_global_lock:
            d = TaskManager._used_global.setdefault(out_dir, {})
            for src, base, fmts in files3:
                n = d.get(base, 0)
                d[base] = n + 1
                files4.append((src, base, base if n == 0 else "%s_%d" % (base, n), fmts))
        with self._lock:
            self._tasks[tid] = {
                "id": tid, "status": "pending",
                "total": len(files3), "done": 0,
                "progress": 0.0,
                "current": None,
                "phase": "queued",
                "heartbeat": time.time(),
                "formats": list(formats), "out_dir": out_dir,
                "is_temp": is_temp,
                "split_dirs": bool(split_dirs),
                "source_paths": [os.path.abspath(src) for src, _, _, _ in files4],
                "results": [], "log": [],
                "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "started": None, "finished": None,
                "clean_after": None,          # 下载后延迟删除时间戳
            }
        item = (tid, files4, out_dir, bool(split_dirs))
        try:
            self._queue.put_nowait(item)
        except queue.Full:
            with self._lock:
                self._tasks.pop(tid, None)
            raise QueueFullError("转换队列已满（上限 %d 个任务），请稍后再试" % QUEUE_MAX)
        log.info("任务 %s 入队：%d 个文件 -> %s，临时目录=%s", tid, len(files3), formats, is_temp)
        return tid

    def get_cached_request(self, request_id):
        now = time.time()
        with self._lock:
            for key, value in list(self._request_cache.items()):
                if now - value["created"] > self._request_cache_ttl:
                    self._request_cache.pop(key, None)
            value = self._request_cache.get(request_id)
            return dict(value["response"]) if value else None

    def cache_request(self, request_id, response):
        with self._lock:
            self._request_cache[request_id] = {
                "created": time.time(),
                "response": dict(response),
            }

    @contextmanager
    def request_guard(self, request_id):
        with self._lock:
            lock = self._request_locks.setdefault(request_id, threading.Lock())
        with lock:
            yield

    def is_path_active(self, path):
        target = os.path.abspath(path).lower()
        with self._lock:
            for task in self._tasks.values():
                if task["status"] not in ("pending", "running"):
                    continue
                paths = [task["out_dir"]] + task.get("source_paths", [])
                for active_path in paths:
                    active = os.path.abspath(active_path).lower()
                    if target == active or target.startswith(active + os.sep) or active.startswith(target + os.sep):
                        return True
        return False

    def get(self, tid):
        with self._lock:
            t = self._tasks.get(tid)
            return dict(t) if t else None

    # ------------------------------------------------------------------
    # 供路由/清理线程使用的查询与控制。
    # 只读子进程上报的普通状态，绝不在 HTTP 请求线程里访问 Word COM 对象
    # ------------------------------------------------------------------
    def set_clean_after(self, tid, ts):
        """真正把 clean_after 写到原任务上（原来 download 里改的是 get() 返回的副本）。"""
        with self._lock:
            t = self._tasks.get(tid)
            if t is not None:
                t["clean_after"] = ts

    def busy_count(self):
        with self._lock:
            return sum(1 for w in self._workers.values() if w["current_tid"])

    def request_release_all(self):
        """向每个转换子进程投递释放指令，由子进程自己 Quit 自己的 Word。"""
        notified = 0
        if self._queue is None:
            return 0
        for wid in self._workers:
            try:
                self._queue.put_nowait(("__release__", wid))
                notified += 1
            except queue.Full:
                log.warning("[release] 队列已满，worker-%d 暂时收不到释放指令", wid)
        return notified

    def health(self):
        with self._lock:
            workers = []
            word_alive = False
            busy = 0
            for wid, w in sorted(self._workers.items()):
                proc_alive = bool(w["proc"] and w["proc"].is_alive())
                if w["word_alive"]:
                    word_alive = True
                if w["current_tid"]:
                    busy += 1
                workers.append({
                    "id": wid,
                    "pid": w["proc"].pid if w["proc"] else None,
                    "proc_alive": proc_alive,
                    "busy": bool(w["current_tid"]),
                    "task": w["current_tid"],
                    "word_alive": w["word_alive"],
                    "word_pid": w["word_pid"],
                    "last_beat_age": round(time.time() - w["last_beat"], 1),
                })
            queued = None
            if self._queue is not None:
                try:
                    queued = self._queue.qsize()
                except Exception:
                    queued = None
        return {
            "status": "ok",
            "pid": os.getpid(),
            "mode": "multiprocessing",
            "workers": WORKERS,
            "workers_detail": workers,
            "word_alive": word_alive,
            "word_idle_timeout": WORD_IDLE_TIMEOUT,
            "word_per_task": WORD_PER_TASK,
            "queue_max": QUEUE_MAX,
            "queued": queued,
            "busy": busy,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    # ------------------------------------------------------------------
    # 监管线程：消费子进程消息 + 看门狗
    # 与旧线程方案的本质区别：卡死的 worker 是子进程，可以被 terminate() 真正杀死，
    # 只影响它自己的任务，HTTP 服务和其它子进程照常运行
    # ------------------------------------------------------------------
    def _supervisor_loop(self):
        while True:
            msg = None
            try:
                msg = self._result_queue.get(timeout=2)
            except queue.Empty:
                pass
            except Exception as e:
                log.warning("[监管] 读取子进程消息失败：%s", e)
            if msg:
                try:
                    self._handle_msg(msg)
                except Exception as e:
                    log.warning("[监管] 消息处理异常：%s (%r)", e, msg)
            try:
                self._watchdog()
            except Exception as e:
                log.warning("[监管] 看门狗异常：%s", e)

    def _handle_msg(self, m):
        typ = m.get("type")
        wid = m.get("wid")
        tid = m.get("tid")
        now = time.time()
        with self._lock:
            w = self._workers.get(wid)
            t = self._tasks.get(tid) if tid else None
            if t is not None:
                t["heartbeat"] = now
            if typ == "hello":
                log.info("[worker-%s] 子进程就绪 (pid=%s)", wid, m.get("pid"))
            elif typ == "word":
                if w is not None:
                    w["word_alive"] = bool(m.get("alive"))
                    w["word_pid"] = m.get("pid")
                log.info("[worker-%s] Word 状态：alive=%s pid=%s（%s）",
                         wid, m.get("alive"), m.get("pid"), m.get("reason") or "")
            elif typ == "claimed":
                if w is not None:
                    w["current_tid"] = tid
                    w["last_beat"] = now
                if t is not None and t["status"] == "pending":
                    t["status"] = "running"
                    t["started"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    t["current"] = "正在启动 Word 转换引擎"
                    t["phase"] = "starting_engine"
            elif typ == "beat":
                if w is not None:
                    w["last_beat"] = now
            elif typ == "status":
                if t is not None:
                    t["current"] = m.get("current")
                    t["phase"] = m.get("phase")
                if w is not None:
                    w["last_beat"] = now
            elif typ == "file":
                if t is not None:
                    t["results"].append(m["entry"])
                    t["done"] = m["done"]
                    t["progress"] = round(m["done"] * 100.0 / max(1, m["total"]), 1)
                    t["current"] = None
                    t["phase"] = "file_done"
                    t["log"].append(m["log"])
                if w is not None:
                    w["last_beat"] = now
                log.info(m["log"])
            elif typ == "log":
                log.info("[worker-%s] %s", wid, m.get("line"))
            elif typ == "done":
                if w is not None:
                    w["current_tid"] = None
                    w["last_beat"] = now
                if t is not None and t["status"] not in ("done", "failed"):
                    self._finish_locked(t, m.get("status") or "done", m.get("error"))

    def _finish_locked(self, t, status, error):
        """完成任务（必须在持有 self._lock 时调用）。"""
        t["status"] = status
        t["finished"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        t["phase"] = "done" if status == "done" else "failed"
        t["heartbeat"] = time.time()
        if status == "done":
            t["progress"] = 100.0
            t["current"] = None
        if error:
            t["error"] = error
            t["log"].append("任务失败: " + error)
            log.error("[%s] %s", t["id"], error)
        else:
            ok = sum(1 for r in t["results"] if r["ok"])
            t["log"].append("任务完成：%d/%d 成功" % (ok, t["total"]))
            log.info("[%s] 完成 %d/%d", t["id"], ok, t["total"])

    @staticmethod
    def _kill_worker_word(w):
        """强杀该 worker 持有的 WINWORD.EXE（子进程被 terminate 时 Word 会变孤儿）。"""
        pid = w.get("word_pid")
        if not pid:
            return
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, text=True, timeout=10)
            log.warning("[看门狗] 已强制结束该 worker 的 WINWORD.EXE (pid=%s)", pid)
        except Exception as e:
            log.warning("[看门狗] 结束 WINWORD pid=%s 失败：%s", pid, e)

    def _watchdog(self):
        now = time.time()
        with self._lock:
            for wid, w in list(self._workers.items()):
                proc = w["proc"]
                # 1) 子进程意外退出：只失败它手里的任务，再拉替补
                if proc is not None and not proc.is_alive():
                    tid = w["current_tid"]
                    log.error("[worker-%d] 转换子进程退出 (exitcode=%s)", wid, proc.exitcode)
                    if tid:
                        t = self._tasks.get(tid)
                        if t and t["status"] in ("pending", "running"):
                            self._finish_locked(
                                t, "failed",
                                "转换子进程异常退出，该任务已标记失败（请重新提交）")
                    w["proc"] = None
                    w["current_tid"] = None
                    w["word_alive"] = False
                    w["word_pid"] = None
                    w["respawn_at"] = now + WORD_WORKER_RESTART_DELAY
                    continue
                # 2) 卡死检测：该 worker 有任务且超过 WORD_STUCK_TIMEOUT 无心跳
                #    （心跳按转换里程碑上报：领任务/每个文件开始与结束。
                #     Word 弹模态框时 COM 调用永久挂起 -> 不再有里程碑 -> 在此被杀）
                if proc is not None and w["current_tid"] and WORD_STUCK_TIMEOUT \
                        and (now - w["last_beat"]) > WORD_STUCK_TIMEOUT:
                    tid = w["current_tid"]
                    log.error("[worker-%d] 任务 %s 超过 %d 秒无心跳，判定 Word 卡死："
                              "终止转换子进程并回收其 Word 实例", wid, tid, WORD_STUCK_TIMEOUT)
                    self._kill_worker_word(w)
                    try:
                        proc.terminate()
                        proc.join(5)
                    except Exception as e:
                        log.warning("[worker-%d] 终止子进程失败：%s", wid, e)
                    t = self._tasks.get(tid)
                    if t and t["status"] in ("pending", "running"):
                        self._finish_locked(
                            t, "failed",
                            "转换引擎无响应（Word 卡死或弹出对话框），已终止对应转换子进程")
                    w["proc"] = None
                    w["current_tid"] = None
                    w["word_alive"] = False
                    w["word_pid"] = None
                    w["respawn_at"] = now + WORD_WORKER_RESTART_DELAY
                    continue
                # 3) 拉起替补子进程
                if proc is None and not self._stopping and now >= w["respawn_at"]:
                    self._spawn_worker(wid)

# ===================== 转换子进程（worker） =====================
# 以下函数全部运行在独立子进程里：每个子进程 CoInitialize 自己的 STA，
# 自己启动/持有/释放 Word，与主进程通过消息队列通信
_WPS_LOCK = threading.Lock()


def _send_msg(result_queue, msg):
    try:
        result_queue.put(msg)
    except Exception:
        pass


def _list_winword_pids():
    """当前所有 WINWORD.EXE 的 PID（用于定位本子进程启动的那个 Word）。"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
        )
    except Exception:
        return set()
    pids = set()
    for line in (out.stdout or "").splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip('"') for p in line.split('","')]
        if parts and parts[0].upper() == "WINWORD.EXE":
            try:
                pids.add(int(parts[1]))
            except (IndexError, ValueError):
                pass
    return pids


def _kill_stale_winword():
    """保险栓：干掉 Quit 后仍赖着不走的 WINWORD.EXE。
    默认关闭（WORD_KILL_STALE=1 才启用），因为会误杀用户自己开的 Word。"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
            capture_output=True, text=True, timeout=10,
        )
        if "WINWORD.EXE" not in (out.stdout or ""):
            return 0
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"],
                       capture_output=True, text=True, timeout=15)
        log.warning("[Word] 已强制结束残留 WINWORD.EXE")
        return 1
    except Exception as e:
        log.warning("[Word] 清理残留 Word 进程失败：%s", e)
        return 0


def _start_word():
    """启动一个新的 Word COM 实例（带重试 + 握手校验）。
    用启动前后的 WINWORD 快照差集定位新实例 PID，供看门狗强杀时精确回收。"""
    before = _list_winword_pids()
    for attempt in range(RPC_WORD_RELAUNCH_TRIES):
        try:
            w = win32com.client.DispatchEx("Word.Application")
            w.Visible = False
            w.DisplayAlerts = 0
            w.ScreenUpdating = False
            try:
                w.FeatureInstall = 0   # msoFeatureInstallNone：避免"正在安装功能"弹窗把 Word 卡死
            except Exception:
                pass
            try:
                w.Options.CheckSpellingAsYouType = False
                w.Options.CheckGrammarAsYouType = False
            except Exception:
                pass
            try:
                w.AutomationSecurity = 3
            except Exception:
                pass
            w.Documents.Count          # 握手：确认 COM 通道真的通
            pid = None
            new_pids = _list_winword_pids() - before
            if new_pids:
                pid = sorted(new_pids)[0]
            log.info("Word COM 转换引擎已启动 (pid=%s)", pid)
            return w, pid
        except Exception as e:
            log.warning("Word COM 启动失败（%d/%d）：%s",
                        attempt + 1, RPC_WORD_RELAUNCH_TRIES, e)
            time.sleep(1.0 * (attempt + 1))
    log.error("Word COM 连续 %d 次启动失败", RPC_WORD_RELAUNCH_TRIES)
    return None, None


def _word_alive(word):
    """轻量体检：能读到 Documents.Count 说明通道还活着（只在子进程线程里调用）。"""
    if word is None:
        return False
    try:
        word.Documents.Count
        return True
    except Exception as e:
        return not _is_rpc_error(e)


def _release_word(word_state, wid, result_queue, reason=""):
    """子进程主动释放自己的 Word 实例：Quit + 置空 + gc，WINWORD.EXE 随之退出。"""
    word = word_state.get("app")
    had_word = word is not None    # 先记下结果：del word 会解除绑定，之后不能再引用
    word_state["app"] = None
    word_state["pid"] = None
    if word is not None:
        try:
            word.Quit()
        except Exception as e:
            log.warning("[Word] Quit 失败（%s）：%s", reason or "主动释放", e)
        try:
            del word
        except Exception:
            pass
        gc.collect()
        if WORD_KILL_STALE:
            _kill_stale_winword()
        log.info("[Word] 已释放实例（%s）", reason or "主动释放")
    _send_msg(result_queue, {"type": "word", "wid": wid,
                             "alive": False, "pid": None,
                             "reason": reason or "释放"})
    return had_word


def _relaunch_word_into(word_state, wid, result_queue, reason=""):
    """彻底丢弃旧实例，重新拉起一个干净的 Word 并写回 word_state。"""
    old = word_state.get("app")
    if old is not None:
        try:
            old.Quit()
        except Exception:
            pass
        try:
            del old
        except Exception:
            pass
    word_state["app"] = None
    word_state["pid"] = None
    gc.collect()
    for i in range(RPC_WORD_RELAUNCH_TRIES):
        time.sleep(0.5 + i)
        w, pid = _start_word()
        if w is not None and _word_alive(w):
            word_state["app"] = w
            word_state["pid"] = pid
            _send_msg(result_queue, {"type": "word", "wid": wid,
                                     "alive": True, "pid": pid,
                                     "reason": reason or "重建"})
            log.warning("[RPC] Word COM 实例已重建（第 %d 次尝试）", i + 1)
            return True
    _send_msg(result_queue, {"type": "word", "wid": wid,
                             "alive": False, "pid": None, "reason": "重建失败"})
    log.error("[RPC] Word COM 重建失败，本批后续文件将标记为失败")
    return False


def _ensure_word(word_state, wid, result_queue, reason=""):
    """每次转换前体检，发现 RPC 断开就地重建。"""
    if _word_alive(word_state.get("app")):
        return True
    log.warning("[RPC] 检测到 Word COM 通道断开（%s），开始重建实例", reason or "健康检查")
    return _relaunch_word_into(word_state, wid, result_queue, reason)


def _process_task(tid, files4, out_dir, split_dirs, word_state, wid, result_queue):
    """转换一个任务的全部文件。返回 (large_seen, error)。
    进度/心跳/结果通过消息发回主进程（子进程不直接改主进程内存里的任务表）。"""
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception as e:
        return False, "output directory error: %s" % e

    wps_state = {"app": None}
    large_seen = False
    files_since_recycle = 0
    recycle_at = time.time() + WORD_RECYCLE_SECONDS if WORD_RECYCLE_SECONDS else 0

    try:
        for i, (src, base, name, fmts) in enumerate(files4, 1):
            _send_msg(result_queue, {"type": "beat", "wid": wid, "tid": tid})
            entry = {"src": os.path.basename(src), "ok": False,
                     "outputs": [], "formats": {}, "error": None}
            try:
                # ---- 0) 主动回收 Word：句柄/内存累积是崩溃的头号诱因 ----
                need_recycle = (
                    (WORD_RECYCLE_FILES and files_since_recycle >= WORD_RECYCLE_FILES)
                    or (WORD_RECYCLE_SECONDS and time.time() > recycle_at)
                )
                if need_recycle:
                    if _relaunch_word_into(word_state, wid, result_queue, "主动回收"):
                        log.info("[%s] 主动回收 Word 实例（已连续转换 %d 个文件）",
                                 tid, files_since_recycle)
                    files_since_recycle = 0
                    recycle_at = time.time() + WORD_RECYCLE_SECONDS

                # ---- 1) 源文件体检：别拿空文件/不存在的文件去喂崩 Word ----
                if not os.path.isfile(src):
                    raise RuntimeError("源文件不存在，跳过")
                if os.path.getsize(src) == 0:
                    raise RuntimeError("源文件为空（0 字节），跳过")
                size_mb = os.path.getsize(src) / (1024.0 * 1024.0)
                is_large = os.path.getsize(src) >= LARGE_RTF_BYTES
                large_seen = large_seen or is_large

                _send_msg(result_queue, {
                    "type": "status", "wid": wid, "tid": tid,
                    "current": ("正在打开 %s（%.1f MB）" % (os.path.basename(src), size_mb)
                                if is_large else "正在打开 " + os.path.basename(src)),
                    "phase": "opening",
                })

                # ---- 2) 单次转换 + RPC 自愈重试（输出名已在主进程分配好）----
                results = None
                last_err = None
                for attempt in range(RPC_RETRY + 1):
                    if not _ensure_word(word_state, wid, result_queue,
                                        "转换 %s 前" % os.path.basename(src)):
                        raise RuntimeError("Word COM 不可用且无法重建")
                    try:
                        results = _convert_one(
                            word_state["app"], wps_state, src, fmts,
                            out_dir, name, bool(split_dirs))
                    except Exception as e:
                        last_err = e
                        results = None
                        if _is_rpc_error(e) and attempt < RPC_RETRY:
                            log.warning("[%s] %s 打开/保存时 RPC 断开（第 %d 次），重建 Word 后重试：%s",
                                        tid, os.path.basename(src), attempt + 1, e)
                            _relaunch_word_into(word_state, wid, result_queue, "RPC 重试")
                            time.sleep(RPC_RETRY_SLEEP)
                            continue
                        raise
                    # _save_document 会把异常吞进 results，补一次检查
                    if _results_have_rpc(results) and attempt < RPC_RETRY:
                        log.warning("[%s] %s 保存时 RPC 断开（第 %d 次），重建 Word 后重试",
                                    tid, os.path.basename(src), attempt + 1)
                        _relaunch_word_into(word_state, wid, result_queue, "RPC 重试")
                        time.sleep(RPC_RETRY_SLEEP)
                        continue
                    break

                if results is None:
                    raise last_err or RuntimeError("转换未返回结果")

                # ---- 3) 结果汇总 ----
                entry["ok"] = any(r["ok"] for r in results.values())
                entry["outputs"] = [r["path"] for r in results.values() if r["ok"]]
                entry["formats"] = {
                    k: ("OK" if v["ok"] else ("FAIL: " + (v["error"] or "")))
                    for k, v in results.items()}
                detail = " ".join(
                    "%s:%s" % (k, ("WPS" if v.get("via_wps") else "OK")
                               if v["ok"] else ("FAIL:" + (v["error"] or "")))
                    for k, v in results.items())
            except Exception as e:
                entry["error"] = str(e)[:300]
                detail = "FAIL: " + entry["error"]
                log.warning("[%s] %s 转换失败：%s", tid, os.path.basename(src), e)

            files_since_recycle += 1
            _send_msg(result_queue, {
                "type": "file", "wid": wid, "tid": tid,
                "entry": entry, "done": i, "total": len(files4),
                "log": "[%d/%d] %s -> %s" % (i, len(files4), entry["src"], detail),
            })
    finally:
        _close_app(wps_state.get("app"))
        wps_state["app"] = None

    return large_seen, None


def _worker_main(wid, task_queue, result_queue, large_lock):
    """转换子进程入口：独享 COM STA 与 Word 实例，可被主进程整体回收。"""
    pythoncom.CoInitialize()
    _send_msg(result_queue, {"type": "hello", "wid": wid, "pid": os.getpid()})
    word_state = {"app": None, "pid": None}
    idle_since = None
    try:
        while True:
            # 主服务已死 -> 自行退出，避免留下孤儿子进程
            parent = multiprocessing.parent_process()
            if parent is not None and not parent.is_alive():
                log.info("[worker-%d] 主服务已退出，转换子进程自行结束", wid)
                return
            # ---- 取任务：带超时，超时即判定为空闲 ----
            try:
                item = task_queue.get(timeout=WORD_POLL_INTERVAL)
            except queue.Empty:
                if word_state["app"] is not None and WORD_IDLE_TIMEOUT:
                    now = time.time()
                    if idle_since is None:
                        idle_since = now
                    elif now - idle_since >= WORD_IDLE_TIMEOUT:
                        _release_word(word_state, wid, result_queue,
                                      "空闲 %d 秒" % WORD_IDLE_TIMEOUT)
                        idle_since = None
                continue

            idle_since = None

            # ---- 控制指令：手动释放 Word ----
            if item[0] == "__release__":
                _release_word(word_state, wid, result_queue, "手动调用 /api/word/release")
                continue

            tid, files4, out_dir, split_dirs = item
            _send_msg(result_queue, {"type": "claimed", "wid": wid, "tid": tid})
            status = "done"
            error = None
            try:
                large_job = any(os.path.getsize(src) >= LARGE_RTF_BYTES
                                for src, _, _, _ in files4)
                if large_job:
                    # 大文件任务跨进程互斥：避免多个 Word 同时装载超大文档。
                    # 用带超时的循环拿锁，等待期间照常上报心跳，防止看门狗误杀
                    acquired = False
                    while not acquired:
                        _send_msg(result_queue, {"type": "beat", "wid": wid, "tid": tid})
                        acquired = large_lock.acquire(timeout=10)
                    try:
                        large_seen, err = _process_task(
                            tid, files4, out_dir, split_dirs, word_state, wid, result_queue)
                    finally:
                        large_lock.release()
                else:
                    large_seen, err = _process_task(
                        tid, files4, out_dir, split_dirs, word_state, wid, result_queue)

                if err:
                    status = "failed"
                    error = err
                    _release_word(word_state, wid, result_queue, "任务出错，回收后重建")
                elif large_seen:
                    _release_word(word_state, wid, result_queue, "大文件任务结束")
                elif WORD_PER_TASK:
                    _release_word(word_state, wid, result_queue,
                                  "任务 %s 结束（每任务一实例）" % tid)
            except Exception as e:
                log.error("[%s] 任务处理异常，回收 Word 实例：%s", tid, e)
                status = "failed"
                error = "任务处理异常：%s" % str(e)[:300]
                _release_word(word_state, wid, result_queue, "任务异常，回收后下次重建")
            _send_msg(result_queue, {"type": "done", "wid": wid, "tid": tid,
                                     "status": status, "error": error})
    finally:
        _release_word(word_state, wid, result_queue, "worker 退出")
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def _close_app(app):
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            del app
        except Exception:
            pass
    gc.collect()


def _convert_one(word, wps_state, src, formats, out_dir, name, split_dirs=False):
    """name 由主进程预先分配（RPC 重试时复用），不再在这里自增。"""
    doc = None
    try:
        open_args = {
            "FileName": os.path.abspath(src),
            "ConfirmConversions": False,
            "ReadOnly": True,
            "AddToRecentFiles": False,
            "Visible": False,
            "OpenAndRepair": False,
        }
        try:
            doc = word.Documents.Open(**open_args)
        except Exception as e_first:
            if _is_rpc_error(e_first):
                raise                      # RPC 断开：不要硬撑，交给上层重建 Word
            open_args["OpenAndRepair"] = True
            doc = word.Documents.Open(**open_args)
        return _save_document(
            doc, wps_state, src, formats, out_dir, name, split_dirs)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception as e:
                if _is_rpc_error(e):
                    log.warning("关闭文档时 RPC 已断开（实例需重建）：%s", src)
                else:
                    log.warning("关闭 Word 文档失败：%s: %s", src, e)
            doc = None
            gc.collect()


def _save_document(doc, wps_state, src, formats, out_dir, name, split_dirs=False):
    results = {}
    for fmt in formats:
        if split_dirs:
            fmt_dir = os.path.join(out_dir, fmt)
            os.makedirs(fmt_dir, exist_ok=True)
            out_path = os.path.join(fmt_dir, "%s.%s" % (name, fmt))
        else:
            out_path = os.path.join(out_dir, "%s.%s" % (name, fmt))
        results[fmt] = {"ok": False, "path": out_path, "error": None, "via_wps": False}
        try:
            if fmt == "pdf":
                # 明确指定书签来源。Word 默认值并不稳定，且 RTF 的显式
                # 书签不会自动写入 PDF；CreateBookmarks=2 保留原文书签。
                doc.ExportAsFixedFormat(
                    out_path,
                    ExportFormat=17,
                    CreateBookmarks=PDF_CREATE_BOOKMARKS,
                )
            elif fmt == "docx":
                try:
                    doc.SaveAs2(out_path, FileFormat=12)
                except Exception as e1:
                    results[fmt]["error"] = str(e1)[:200]
                    wps = _get_wps(wps_state)
                    if wps is None:
                        raise
                    fallback_src = src
                    fallback_rtf = None
                    try:
                        fd, fallback_rtf = tempfile.mkstemp(
                            prefix="rtf_wps_", suffix=".rtf", dir=OUTPUT_DIR)
                        os.close(fd)
                        doc.SaveAs2(fallback_rtf, FileFormat=6)
                        fallback_src = fallback_rtf
                    except Exception:
                        fallback_src = src
                    try:
                        with _WPS_LOCK:
                            _wps_save_docx(wps, fallback_src, out_path)
                    finally:
                        # 不做立即删除：WPS/Word 可能仍持有句柄，马上删必报
                        # WinError 32。留在原地交给 _auto_clean_all 定时清理
                        pass
                    results[fmt]["via_wps"] = True
                    results[fmt]["error"] = None
            else:
                doc.SaveAs2(out_path, FileFormat=FORMAT_MAP[fmt])
            results[fmt]["ok"] = True
        except Exception as e:
            results[fmt]["error"] = str(e)[:300]
    return results


def _get_wps(wps_state):
    if wps_state["app"] is None:
        try:
            wps = win32com.client.DispatchEx("Kwps.Application")
            wps.Visible = False
            try:
                wps.DisplayAlerts = 0
            except Exception:
                pass
            wps_state["app"] = wps
        except Exception:
            return None
    return wps_state["app"]


def _wps_save_docx(wps, src, out_path):
    wdoc = None
    try:
        wdoc = wps.Documents.Open(os.path.abspath(src), ReadOnly=True)
        wdoc.SaveAs2(out_path, FileFormat=12)
    finally:
        if wdoc is not None:
            try:
                wdoc.Close(False)
            except Exception as e:
                log.warning("关闭 WPS 文档失败：%s: %s", src, e)
            wdoc = None
            gc.collect()


# ===================== 统一自动清理（延迟删除） =====================
def _safe_rmtree(path):
    """安全删除目录，记录日志"""
    if not os.path.exists(path):
        return True
    if manager.is_path_active(path):
        log.info("[自动清理] 跳过活动任务目录：%s", path)
        return False
    for attempt in range(3):
        try:
            shutil.rmtree(path)
            log.info("[自动清理] 删除目录：%s", path)
            return True
        except (PermissionError, OSError) as e:
            if attempt == 2:
                log.warning("[自动清理] 删除失败（可能仍被占用）%s: %s", path, e)
                return False
            time.sleep(0.5 * (attempt + 1))
        except Exception as e:
            log.warning("[自动清理] 删除失败 %s: %s", path, e)
            return False
    return False


def _next_cleanup_time(now=None):
    now = now or datetime.now()
    target = now.replace(hour=CLEANUP_HOUR, minute=0, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return target.timestamp()


def _auto_clean_all():
    """后台统一清理线程：
    1) uploads_* 目录：按 UPLOAD_EXPIRE_DAYS 天过期
    2) tmp_* 目录（未下载/失败）：按 TMP_TASK_EXPIRE_DAYS 天过期
    3) tmp_* 目录（已下载）：按 clean_after 时间戳延迟删除
    """
    upload_cutoff = UPLOAD_EXPIRE_DAYS * 86400
    tmp_task_cutoff = TMP_TASK_EXPIRE_DAYS * 86400

    log.info(
        "[自动清理] 启动：uploads=%d天, tmp任务=%d天, 下载后延迟=%d秒, 扫描间隔=%d秒",
        UPLOAD_EXPIRE_DAYS, TMP_TASK_EXPIRE_DAYS, DOWNLOAD_CLEAN_DELAY, CLEAN_SCAN_INTERVAL
    )

    def _clean_by_prefix(parent_dir, prefix, cutoff_ts, desc):
        if not os.path.isdir(parent_dir):
            return 0
        removed = 0
        now = time.time()
        for entry in os.listdir(parent_dir):
            if not entry.startswith(prefix):
                continue
            full = os.path.join(parent_dir, entry)
            if not os.path.isdir(full):
                continue
            try:
                age = now - os.path.getmtime(full)
            except Exception:
                continue
            if age > cutoff_ts:
                if _safe_rmtree(full):
                    removed += 1
                    log.info("[自动清理] [%s] 删除 %s (年龄=%.2f天)", desc, full, age / 86400)
        return removed

    def _clean_old_files(prefix, cutoff_ts):
        if not os.path.isdir(OUTPUT_DIR):
            return
        now = time.time()
        for entry in os.listdir(OUTPUT_DIR):
            if not entry.startswith(prefix):
                continue
            full = os.path.join(OUTPUT_DIR, entry)
            if os.path.isfile(full):
                try:
                    if now - os.path.getmtime(full) > cutoff_ts:
                        _safe_remove(full)
                        log.info("[自动清理] 删除临时文件：%s", full)
                except Exception:
                    pass

    while True:
        target = _next_cleanup_time()
        log.info("[自动清理] 下次清理时间：%s",
                 datetime.fromtimestamp(target).strftime("%Y-%m-%d %H:%M:%S"))
        time.sleep(max(1, target - time.time()))
        try:
            _clean_by_prefix(UPLOAD_DIR, "uploads_", upload_cutoff, "uploads")
            _clean_by_prefix(OUTPUT_DIR, "tmp_", tmp_task_cutoff, "tmp任务")
            _clean_old_files("rtf_download_", tmp_task_cutoff)
            # WPS 兜底转换的中间 rtf_wps_*.rtf（docx 失败时生成），异常中断会泄漏
            _clean_old_files("rtf_wps_", tmp_task_cutoff)

            # 已下载 tmp_* 按 clean_after 延迟删除
            with manager._lock:
                for tid, t in list(manager._tasks.items()):
                    if (t.get("is_temp") and t.get("clean_after")
                            and isinstance(t["clean_after"], (int, float))
                            and time.time() > t["clean_after"]):
                        out_dir = t["out_dir"]
                        t["clean_after"] = None   # 标记已处理，避免重复删
                        if os.path.exists(out_dir):
                            log.info("[自动清理] 下载后延迟到期，删除 tmp 目录：%s (task=%s)", out_dir, tid)
                            _safe_rmtree(out_dir)

        except Exception as e:
            log.warning("[自动清理] 扫描异常: %s", e)


manager = TaskManager()
_dir_pick_lock = threading.Lock()


# ===================== 前端页面 =====================
INDEX_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RTF 批量转换</title>
<style>
*{box-sizing:border-box;font-family:"Segoe UI","Microsoft YaHei",sans-serif}
body{max-width:1100px;margin:30px auto;padding:0 20px;background:#f5f6f8;color:#1a1b1c}
h1{font-size:22px;margin:0 0 4px}
.sub{color:#6b7280;font-size:13px;margin-bottom:20px}
.card{background:#fff;border:1px solid #e4e3dd;border-radius:12px;padding:18px;margin-bottom:16px}
.row{display:flex;gap:16px;flex-wrap:wrap;align-items:center}
.upload-area{border:2px dashed #9bbbf4;background:#fafbfe;border-radius:10px;padding:30px 12px;text-align:center;cursor:pointer;color:#6b7280}
.upload-area.over{background:#eef3fd;border-color:#6d96e8}
#file-input{display:none}
input[type=text]{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #d0d3da;border-radius:8px;font-size:14px}
.filelist{font-size:13px;color:#374151;margin-top:10px;max-height:160px;overflow:auto}
.chk{display:inline-flex;align-items:center;gap:4px;margin-right:14px;font-size:14px}
.btn{background:#2563eb;color:#fff;border:0;border-radius:6px;padding:9px 16px;font-size:14px;cursor:pointer}
.btn:disabled{background:#9ca3af;cursor:not-allowed}
.btn2{background:#fff;color:#1a1b1c;border:1px solid #d0d3da;border-radius:6px;padding:8px 14px;font-size:14px;cursor:pointer}
.btn2:hover{background:#f3f4f6}

table{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px}
th,td{border-bottom:1px solid #eee;padding:10px;text-align:left}
th{background:#fafafa;font-weight:600}
.st-queue{color:#6b7280}
.st-processing{color:#2563eb;font-weight:600}
.st-done{color:#16a34a;font-weight:600}
.st-fail{color:#dc2626;font-weight:600}

.progress-track{height:12px;background:#e9edf3;border-radius:999px;overflow:hidden;margin-top:8px}
.progress-fill{height:100%;width:0;background:#2563eb;border-radius:inherit;transition:width .35s ease}
.log-label{margin-top:16px;color:#6b7280;font-size:12px;font-weight:600}
.log{font-size:12px;color:#4b5563;background:#f8fafc;border:1px solid #edf0f4;border-radius:8px;padding:10px;margin-top:6px;max-height:220px;overflow:auto;white-space:pre-wrap;font-family:Consolas,monospace}
.hide{display:none}
</style>
</head>
<body>
<h1>RTF 批量转换</h1>
<div class="sub">转换引擎：Word · 多进程 worker 并发 · 支持 doc / docx / pdf · 空闲自动释放 · RPC 断开自动重建重试</div>

<div class="card">
    <div style="font-size:14px;font-weight:600;margin-bottom:8px">输入（二选一：拖拽上传 或 输入服务端文件夹路径）</div>
    <div class="upload-area" id="uploadBox">
        <p style="font-size:16px;color:#374151;margin:0 0 4px"><b>点击选择 RTF 文件</b></p>
        <p style="margin:0">或将 .rtf 文件/整个文件夹拖拽到此处（自动收集全部rtf，支持多选）</p>
        <input type="file" id="file-input" multiple accept=".rtf">
    </div>
    <div id="filebar" style="margin-top:10px;display:flex;justify-content:space-between;align-items:center">
        <span id="filecount" style="font-size:12px;color:#6b7280"></span>
        <button class="btn2" onclick="clearFiles()">清空全部</button>
    </div>
    <div class="filelist" id="filelist"></div>
    <div style="margin-top:10px">
        <input type="text" id="pathInput" placeholder="或直接输入服务端文件夹路径（自动扫描全部 .rtf），例如 D:\报告\rtf">
    </div>
</div>

<div class="card">
    <div style="font-size:14px;font-weight:600;margin-bottom:8px">输出设置</div>
    <div class="row" style="margin-bottom:10px">
        <span style="font-size:14px">格式：</span>
        <label class="chk"><input type="checkbox" class="fmt" value="doc" checked> doc</label>
        <label class="chk"><input type="checkbox" class="fmt" value="docx" checked> docx</label>
        <label class="chk"><input type="checkbox" class="fmt" value="pdf" checked> pdf</label>
    </div>
    <label class="chk" style="margin-bottom:10px"><input type="checkbox" id="splitDirs" checked> 按格式分开存放（doc/、docx/、pdf/ 各一个文件夹）</label>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
        <input type="text" id="outDir" placeholder="输出目录（留空 = 默认临时output/任务ID，下载后延迟清理）" style="flex:1 1 300px;min-width:0">
        <button class="btn2" onclick="pickDir(this)">选择输出文件夹</button>
    </div>
    <div style="margin-top:14px">
        <button class="btn" id="goBtn" onclick="submit()">开始转换</button>
    </div>
</div>

<div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:8px">
        <h3 style="margin:0">任务列表</h3>
        <div style="display:flex;align-items:center;gap:10px">
            <span id="svcStatus" style="font-size:12px;color:#6b7280">服务状态加载中…</span>
            <button class="btn2" id="dlAllBtn" style="padding:4px 10px;font-size:12px" onclick="downloadAll()" disabled>下载全部</button>
            <button class="btn2" style="padding:4px 10px;font-size:12px" onclick="releaseWord()">释放 Word</button>
        </div>
    </div>
    <table>
        <thead>
            <tr>
                <th>文件名</th>
                <th>大小</th>
                <th>状态</th>
                <th>进度</th>
                <th>处理信息</th>
                <th>操作</th>
            </tr>
        </thead>
        <tbody id="taskList"></tbody>
    </table>
    <div class="log-label">全局处理记录（多任务并行时按任务号区分）</div>
    <div class="log" id="log"></div>
</div>

<script>
var files = [];
var tasks = new Map();          // tid -> {name, finished, lastLog}，支持多任务并行跟踪
var submitInFlight = false;
var submissionRequestId = null;
var pollTimer = null;
const uploadBox = document.getElementById("uploadBox");
const fileInput = document.getElementById("file-input");
const taskListEl = document.getElementById("taskList");

function pickDir(btn) {
  var old = btn.textContent;
  btn.disabled = true;
  btn.textContent = '请在弹出的对话框中选择…';
  fetch('/api/pick-dir')
    .then(function(r){ return r.json(); })
    .then(function(j){
      btn.disabled = false;
      btn.textContent = old;
      if (j.path) document.getElementById('outDir').value = j.path;
      else if (j.error) alert(j.error);
    })
    .catch(function(e){
      btn.disabled = false;
      btn.textContent = old;
      alert('弹出文件夹选择对话框失败：' + e);
    });
}

function releaseWord() {
  if (!confirm('通知所有转换 worker 立即退出 Word 实例？\\n（空闲超过 5 分钟本来会自动释放，一般不需要手动点）')) return;
  fetch('/api/word/release', { method: 'POST' })
    .then(function(r){ return r.json(); })
    .then(function(j){
      setLog(j.error ? '释放 Word 失败：' + j.error
                     : '已通知 ' + (j.notified || 0) + ' 个转换 worker 释放 Word 实例');
      refreshSvc();
    })
    .catch(function(e){ setLog('释放请求失败：' + e); });
}

uploadBox.onclick = () => fileInput.click();
uploadBox.ondragover = e => { e.preventDefault(); uploadBox.classList.add("over"); };
uploadBox.ondragleave = () => uploadBox.classList.remove("over");
uploadBox.ondrop = handleDrop;
fileInput.onchange = function(){ addFiles(fileInput.files); fileInput.value = ""; };

function entryToFiles(entry) {
  return new Promise(function(resolve){
    if (entry.isFile) {
      entry.file(function(f){ resolve([f]); }, function(){ resolve([]); });
    } else if (entry.isDirectory) {
      var reader = entry.createReader();
      var all = [];
      (function readBatch(){
        reader.readEntries(function(batch){
          if (!batch.length) {
            Promise.all(all.map(entryToFiles)).then(function(arrs){
              var flat = [];
              arrs.forEach(function(a){ flat = flat.concat(a); });
              resolve(flat);
            });
          } else {
            all = all.concat(batch);
            readBatch();
          }
        }, function(){ resolve([]); });
      })();
    } else resolve([]);
  });
}

function handleDrop(e) {
  e.preventDefault();
  uploadBox.classList.remove('over');
  var items = e.dataTransfer.items;
  if (items && items.length && items[0] && items[0].webkitGetAsEntry) {
    var entries = [];
    for (var i = 0; i < items.length; i++) {
      var en = items[i].webkitGetAsEntry();
      if (en) entries.push(en);
    }
    Promise.all(entries.map(entryToFiles)).then(function(arrs){
      var flat = [];
      arrs.forEach(function(a){ flat = flat.concat(a); });
      addFiles(flat);
    });
  } else {
    addFiles(e.dataTransfer.files);
  }
}

function addFiles(list) {
  for (var i = 0; i < list.length; i++) {
    var f = list[i];
    if (!f.name.toLowerCase().endsWith('.rtf')) { alert('仅支持 .rtf：' + f.name); continue; }
    files.push(f);
  }
  renderFiles();
}
function renderFiles() {
  filelist.innerHTML = files.length
    ? files.map(function(f, i){
        return '<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #f0f1f3">'
          + '<span style="word-break:break-all;flex:1">' + f.name + '</span>'
          + '<span data-i="' + i + '" style="cursor:pointer;color:#ef4444;font-weight:bold;margin-left:10px;padding:0 4px" title="移除此文件">&times;</span>'
          + '</div>';
      }).join('')
    : '';
  var bar = document.getElementById('filebar');
  bar.style.display = files.length ? 'flex' : 'none';
  document.getElementById('filecount').textContent = '已选 ' + files.length + ' 个 .rtf 文件';
}
function clearFiles() {
  if (!files.length) return;
  if (!confirm('清空已选的 ' + files.length + ' 个文件？')) return;
  files = [];
  renderFiles();
}
filelist.addEventListener('click', function(e){
  if (e.target.tagName === 'SPAN' && e.target.dataset.i !== undefined) {
    files.splice(+e.target.dataset.i, 1);
    renderFiles();
  }
});

function submit() {
  if (submitInFlight) return;
  var fmt = [].slice.call(document.querySelectorAll('.fmt:checked')).map(function(x){ return x.value; });
  if (!fmt.length) { alert('请至少选择一种输出格式'); return; }
  var p = document.getElementById('pathInput').value.trim();
  if (!p && !files.length) { alert('请拖拽上传 .rtf 文件/文件夹，或输入文件夹路径'); return; }
  if (p && files.length) { alert('拖拽上传与输入文件夹路径请二选一'); return; }
  var fd = new FormData();
  if (p) {
    fd.append('path', p);
  } else {
    files.forEach(function(f){ fd.append('files', f); });
  }
  fd.append('formats', fmt.join(','));
  var od = document.getElementById('outDir').value.trim();
  if (od) fd.append('out_dir', od);
  if (document.getElementById('splitDirs').checked) fd.append('split_dirs', '1');

  submitInFlight = true;
  setLog('正在上传文件…（文件越多越久，请等待服务器返回任务号）');
  if (!submissionRequestId) {
    submissionRequestId = (window.crypto && crypto.randomUUID)
      ? crypto.randomUUID()
      : 'rtf-' + Date.now() + '-' + Math.random().toString(16).slice(2);
  }
  document.getElementById('goBtn').disabled = true;

  fetch('/api/convert', {
    method: 'POST',
    body: fd,
    headers: { 'X-Request-ID': submissionRequestId }
  })
    .then(function(r){
      return r.json().catch(function(){
        throw new Error('HTTP ' + r.status + '（服务器返回的不是 JSON）');
      }).then(function(j){
        if (!r.ok) throw new Error('HTTP ' + r.status + '：' + (j.error || JSON.stringify(j)));
        return j;
      });
    })
    .then(function(j){
      submitInFlight = false;
      submissionRequestId = null;
      if (!j.task_id) {
        setLog('错误：' + JSON.stringify(j));
        alert('提交失败：' + (j.error || JSON.stringify(j)));
        document.getElementById('goBtn').disabled = false;
        return;
      }
      var tid = j.task_id;
      var nFiles = files.length;
      tasks.set(tid, { name: '批量任务（' + nFiles + ' 个文件）', finished: false, lastLog: [] });
      addRow(tid, '批量任务（' + nFiles + ' 个文件）', '-');
      files = [];
      renderFiles();
      document.getElementById('pathInput').value = '';
      document.getElementById('goBtn').disabled = false;
      startPolling();
      setLog('任务已提交：' + tid.slice(0,8) + '（可继续提交下一批，无需等待）');
    })
    .catch(function(e){
      submitInFlight = false;
      submissionRequestId = null;   // 不重置会命中幂等缓存，重试拿到旧结果
      setLog('请求失败：' + e.message);
      alert('提交失败，请修正后重试：\n' + e.message);
      document.getElementById('goBtn').disabled = false;
    });
}

function addRow(tid, name, size) {
  const tr = document.createElement("tr");
  tr.id = "row_" + tid;
  tr.innerHTML = `
    <td>${name}</td>
    <td class="size">${size}</td>
    <td class="st-queue">queue</td>
    <td><div class="progress-track"><div class="progress-fill" id="bar_${tid}"></div></div></td>
    <td>排队中…</td>
    <td></td>`;
  taskListEl.appendChild(tr);
}

// ---- 多任务轮询：所有未完结任务并行拉状态，互不干扰 ----
function startPolling() {
  if (pollTimer) return;
  pollTimer = setTimeout(pollOnce, 800);
}

function pollOnce() {
  pollTimer = null;
  var active = [];
  tasks.forEach(function(st, tid){ if (!st.finished) active.push(tid); });
  if (!active.length) { renderLog(); return; }
  var jobs = active.map(function(tid){
    return fetch('/api/task/' + tid)
      .then(function(r){
        if (!r.ok) throw new Error('gone');
        return r.json();
      })
      .then(function(t){ updateRow(tid, t); })
      .catch(function(){ markRowDead(tid); });
  });
  Promise.all(jobs).catch(function(){}).then(function(){
    renderLog();
    pollTimer = setTimeout(pollOnce, 1500);
  });
}

function updateRow(tid, t) {
  const st = tasks.get(tid);
  if (!st) return;
  st.lastLog = t.log || [];
  st.status = t.status;
  var pct = typeof t.progress === 'number'
    ? t.progress
    : (t.total ? (t.done * 100.0 / t.total) : 0);
  pct = Math.max(0, Math.min(100, pct));
  var statusText = t.status === 'done' ? 'done'
    : t.status === 'failed' ? 'fail'
    : t.status === 'pending' ? 'queue'
    : 'processing';
  const tr = document.getElementById('row_' + tid);
  if (!tr) return;
  const cls = "st-" + statusText;
  let action = "";
  let info;
  if (t.status === "done") {
    st.finished = true;
    action = `<a class="btn" href="/api/task/${tid}/download">下载 ZIP</a>`;
    info = "全部完成";
  } else if (t.status === "failed") {
    st.finished = true;
    info = "失败：" + (t.error || "详见处理记录");
  } else if (t.status === "pending") {
    info = "排队中…（前面的任务转换完就轮到）";
  } else {
    info = t.current ? "正在处理：" + t.current : "等待转换引擎响应";
  }
  updateDlAll();
  tr.innerHTML = `
    <td>${st.name} <span style="color:#9ca3af">(${tid.slice(0,8)})</span></td>
    <td class="size">${t.total}个文件</td>
    <td class="${cls}">${statusText}</td>
    <td><div class="progress-track"><div class="progress-fill" id="bar_${tid}" style="width:${pct}%"></div></div></td>
    <td>${info}</td>
    <td>${action}</td>`;
}

function markRowDead(tid) {
  const st = tasks.get(tid);
  if (!st) return;
  st.finished = true;
  st.status = "failed";
  updateDlAll();
  const tr = document.getElementById('row_' + tid);
  if (tr) tr.innerHTML = `
    <td>${st.name} <span style="color:#9ca3af">(${tid.slice(0,8)})</span></td>
    <td class="size">-</td>
    <td class="st-fail">fail</td>
    <td><div class="progress-track"><div class="progress-fill" id="bar_${tid}" style="width:0%"></div></div></td>
    <td>任务不存在或服务已重启，请重新提交</td>
    <td></td>`;
}

// ---- 日志：系统提示 + 各任务逐文件记录，按任务号区分 ----
var sysMsgs = [];
function setLog(s) {
  if (s) sysMsgs.push('[' + new Date().toLocaleTimeString() + '] ' + s);
  if (sysMsgs.length > 50) sysMsgs = sysMsgs.slice(-50);
  renderLog();
}

function renderLog() {
  var lines = sysMsgs.slice();
  tasks.forEach(function(st, tid){
    (st.lastLog || []).forEach(function(l){ lines.push('[' + tid.slice(0,8) + '] ' + l); });
  });
  document.getElementById('log').textContent = lines.join('\n') || '';
}

// ---- 合并下载：所有已完成任务的输出打进一个 ZIP（服务端按任务号分目录） ----
function doneTaskIds() {
  var ids = [];
  tasks.forEach(function(st, tid){ if (st.status === 'done') ids.push(tid); });
  return ids;
}

function updateDlAll() {
  var btn = document.getElementById('dlAllBtn');
  if (!btn) return;
  var n = doneTaskIds().length;
  btn.textContent = n > 0 ? ('下载全部（' + n + '）') : '下载全部';
  btn.disabled = n === 0;
}

function downloadAll() {
  var ids = doneTaskIds();
  if (!ids.length) { alert('还没有已完成的任务'); return; }
  setLog('合并下载 ' + ids.length + ' 个任务的输出…');
  window.location.href = '/api/tasks/download?ids=' + ids.join(',');
}

// ---- 服务状态栏：只读 /api/health，5 秒刷新 ----
function refreshSvc() {
  fetch('/api/health')
    .then(function(r){ return r.json(); })
    .then(function(h){
      var detail = h.workers_detail || [];
      var alive = detail.filter(function(w){ return w.proc_alive; }).length;
      var wordCount = detail.filter(function(w){ return w.word_alive; }).length;
      document.getElementById('svcStatus').textContent =
        '转换进程 ' + alive + '/' + h.workers +
        '（忙 ' + h.busy + '）· 排队 ' + (h.queued == null ? '-' : h.queued) +
        ' · Word 实例 ' + wordCount;
    })
    .catch(function(){
      document.getElementById('svcStatus').textContent = '服务状态获取失败';
    });
}
refreshSvc();
setInterval(refreshSvc, 5000);
</script>
</body>
</html>
"""


# ===================== API =====================
@app.before_request
def auth():
    if API_TOKEN and request.path != "/api/health":
        if request.headers.get("X-API-Key") != API_TOKEN:
            return jsonify({"error": "unauthorized"}), 401


@app.errorhandler(413)
def upload_too_large(e):
    return jsonify({"error": "上传内容超过上限 %d MB，请减少文件数量或体积" % MAX_UPLOAD_MB}), 413


@app.get("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html; charset=utf-8")


@app.get("/api/health")
def health():
    # 只读子进程上报的普通状态；绝不在请求线程里碰 Word COM 对象
    return jsonify(manager.health())


@app.post("/api/word/release")
def word_release():
    """通知所有转换子进程释放各自的 Word 实例。有任务在跑时拒绝（409）。"""
    with manager._lock:
        running = any(t["status"] in ("pending", "running")
                      for t in manager._tasks.values())
    if running or manager.busy_count() > 0:
        return jsonify({"error": "有任务正在处理，拒绝释放"}), 409
    notified = manager.request_release_all()
    if notified <= 0:
        return jsonify({"error": "没有可通知的转换子进程（服务未启动或队列已满）"}), 500
    return jsonify({
        "released": True,
        "notified": notified,
        # 释放是异步的：子进程收到指令后自己 Quit，几秒后再查 /api/health 可看到 word_alive=false
        "word_alive": any(w["word_alive"] for w in manager._workers.values()),
    })


def _save_upload(f):
    """统一的上传保存逻辑：UPLOAD_DIR/uploads_<8位>/文件名"""
    sub_dir = os.path.join(UPLOAD_DIR, "uploads_" + uuid.uuid4().hex[:8])
    os.makedirs(sub_dir, exist_ok=True)
    path = os.path.join(sub_dir, os.path.basename(f.filename))
    f.save(path)
    return path


def _bad(msg, status=400):
    """400/4xx 统一出口：把拒绝原因写进 server.log。
    否则日志里只有一行 "POST /api/convert 400"，根本无从排查。"""
    try:
        log.error("[convert] 拒绝请求(%d)：%s", status, msg)
    except Exception:
        pass
    return jsonify({"error": msg}), status


@app.post("/api/convert")
def convert():
    request_id = (request.headers.get("X-Request-ID") or "").strip()
    if len(request_id) > 128:
        return _bad("X-Request-ID 过长")
    if not request_id:
        return _convert_impl()
    with manager.request_guard(request_id):
        cached = manager.get_cached_request(request_id)
        if cached:
            return jsonify(cached)
        result = _convert_impl()
        response = result[0] if isinstance(result, tuple) else result
        status = result[1] if isinstance(result, tuple) and len(result) > 1 else 200
        if 200 <= status < 300 and hasattr(response, "get_json"):
            body = response.get_json(silent=True)
            if isinstance(body, dict) and body.get("task_id"):
                manager.cache_request(request_id, body)
        return result


def _convert_impl():
    form = request.form

    # ---- 映射模式 ----
    fmt_inputs = {}
    for fmt in FORMAT_MAP:
        ups = request.files.getlist("files_" + fmt)
        if ups:
            fmt_inputs[fmt] = {"files": ups}
    map_raw = (form.get("map") or "").strip()
    if map_raw:
        try:
            mapping = json.loads(map_raw)
        except Exception as e:
            log.error("map parse fail: %r -> %s", map_raw, e)
            return _bad("map 参数不是合法 JSON")
        if not isinstance(mapping, dict):
            return _bad("map 参数格式错误")
        for fmt, spec in mapping.items():
            if fmt not in FORMAT_MAP:
                return _bad("不支持的格式: %s，可选 %s" % (fmt, list(FORMAT_MAP)))
            if fmt in fmt_inputs:
                return _bad("格式 %s 输入重复" % fmt)
            if not isinstance(spec, dict) or not (spec.get("path") or "").strip():
                return _bad("map.%s 需提供 path 目录" % fmt)
            fmt_inputs[fmt] = {"path": spec["path"].strip()}
    if fmt_inputs:
        items = []
        for fmt, src in fmt_inputs.items():
            if "files" in src:
                for f in src["files"]:
                    if not f or not f.filename:
                        continue
                    if not f.filename.lower().endswith(".rtf"):
                        return _bad("仅支持 .rtf 文件: %s" % f.filename)
                    p = _save_upload(f)
                    items.append((p, fmt))
            else:
                p = src["path"]
                if not os.path.isdir(p):
                    return _bad("目录不存在(注意：path 必须是运行本服务的那台机器上的路径，"
                                "客户端填自己电脑的路径一律找不到): %s" % p)
                rtf_list = sorted(x for x in os.listdir(p) if x.lower().endswith(".rtf"))
                if not rtf_list:
                    return _bad("目录下没有 .rtf 文件: %s" % p)
                items.extend((os.path.join(p, x), fmt) for x in rtf_list)
        if not items:
            return _bad("没有可转换的文件")
        if len({s for s, _ in items}) > MAX_FILES:
            return _bad("文件数超过上限 %d" % MAX_FILES)
        out_dir = (form.get("out_dir") or "").strip() or None
        split_dirs = (form.get("split_dirs") or "").strip().lower() in ("1", "true", "on", "yes")
        try:
            task_id = manager.submit_map(items, out_dir, split_dirs)
        except QueueFullError as e:
            return _bad(str(e), 503)
        return jsonify({
            "task_id": task_id,
            "total": len({s for s, _ in items}),
            "formats": sorted({f for _, f in items}),
            "out_dir": manager.get(task_id)["out_dir"],
            "poll": "/api/task/<task_id>", "download": "/api/task/<task_id>/download",
        })

    # ---- 统一模式 ----
    raw = (form.get("formats") or "").strip().lower()
    formats = [f for f in raw.split(",") if f] if raw else list(DEFAULT_FORMATS)
    bad = [f for f in formats if f not in FORMAT_MAP]
    if bad:
        return _bad("不支持的格式: %s，可选 %s" % (bad, list(FORMAT_MAP)))

    has_files = bool(request.files)
    src_path = (form.get("path") or "").strip()
    # 提交来源诊断：客户端 IP / 上传文件数 / 是否走服务端路径模式
    log.info("[convert] 来源=%s files=%d path=%r out_dir=%r formats=%s",
             request.remote_addr, len(request.files.getlist("files")),
             src_path, (form.get("out_dir") or "").strip(), formats)
    if has_files and src_path:
        return _bad("files 上传与 path 目录模式请二选一")
    if has_files:
        files = []
        for f in request.files.getlist("files"):
            if not f or not f.filename:
                continue
            if not f.filename.lower().endswith(".rtf"):
                return _bad("仅支持 .rtf 文件: %s" % f.filename)
            path = _save_upload(f)
            files.append((path, os.path.splitext(f.filename)[0]))
    elif src_path:
        if not os.path.isdir(src_path):
            # 说明：path 是"服务端本机"路径，不是访问者电脑上的路径
            return _bad("目录不存在(注意：path 必须是运行本服务的那台机器上的路径，"
                        "客户端填自己电脑的路径一律找不到): %s" % src_path)
        rtf_list = sorted(x for x in os.listdir(src_path) if x.lower().endswith(".rtf"))
        if not rtf_list:
            return _bad("目录下没有 .rtf 文件: %s" % src_path)
        files = [(os.path.join(src_path, x), os.path.splitext(x)[0]) for x in rtf_list]
    else:
        return _bad("请上传 files 或提供服务端目录 path")

    if len(files) > MAX_FILES:
        return _bad("文件数 %d 超过上限 %d" % (len(files), MAX_FILES))

    out_dir = (form.get("out_dir") or "").strip() or None
    split_dirs = (form.get("split_dirs") or "").strip().lower() in ("1", "true", "on", "yes")
    try:
        task_id = manager.submit(files, formats, out_dir, split_dirs)
    except QueueFullError as e:
        return _bad(str(e), 503)
    return jsonify({
        "task_id": task_id,
        "total": len(files), "formats": formats,
        "split_dirs": split_dirs,
        "out_dir": manager.get(task_id)["out_dir"],
        "poll": "/api/task/<task_id>", "download": "/api/task/<task_id>/download",
    })


@app.get("/api/pick-dir")
def pick_dir():
    if not _dir_pick_lock.acquire(blocking=False):
        return jsonify({"error": "已有文件夹选择对话框打开中"}), 409
    try:
        import tkinter
        from tkinter import filedialog
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.lift()
        try:
            path = filedialog.askdirectory(title="选择输出文件夹")
        finally:
            root.destroy()
        if not path:
            return jsonify({"path": None, "canceled": True})
        return jsonify({"path": os.path.normpath(path)})
    except Exception as e:
        log.error("pick-dir 失败: %s", e)
        return jsonify({"error": "无法弹出文件夹选择对话框: %s" % e}), 500
    finally:
        _dir_pick_lock.release()


@app.get("/api/task/<tid>")
def task_status(tid):
    t = manager.get(tid)
    if not t:
        return jsonify({"error": "任务不存在"}), 404
    if "progress" not in t:
        t["progress"] = round(t["done"] * 100.0 / t["total"], 1) if t["total"] else 0
    return jsonify(t)


@app.get("/api/tasks/download")
def tasks_download():
    """合并下载多个已完成任务的输出：ids=逗号分隔的任务号。
    ZIP 内按任务号前缀分目录，避免不同任务的同名输出互相覆盖。"""
    ids_raw = (request.args.get("ids") or "").strip()
    tids = [x.strip() for x in ids_raw.split(",") if x.strip()]
    if not tids:
        return jsonify({"error": "缺少 ids 参数（逗号分隔的任务号）"}), 400
    if len(tids) > 50:
        return jsonify({"error": "一次最多合并下载 50 个任务"}), 400
    tasks_data = []
    unknown = []
    for tid in tids:
        t = manager.get(tid)
        if t and t["status"] == "done":
            tasks_data.append(t)
        else:
            unknown.append(tid[:8])
    if not tasks_data:
        return jsonify({"error": "没有已完成的任务可下载"}), 400
    fd, zip_path = tempfile.mkstemp(prefix="rtf_download_", suffix=".zip", dir=OUTPUT_DIR)
    os.close(fd)
    try:
        added = 0
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for t in tasks_data:
                prefix = t["id"][:8]
                for entry in t["results"]:
                    for p in entry["outputs"]:
                        if os.path.exists(p):
                            rel = os.path.relpath(p, t["out_dir"])
                            if rel.startswith(".."):
                                rel = os.path.basename(p)
                            z.write(p, os.path.join(prefix, rel))
                            added += 1
        if added == 0:
            _safe_remove(zip_path)
            return jsonify({"error": "所有输出文件都已不存在（可能已被清理）"}), 400
    except Exception as e:
        _safe_remove(zip_path)
        return jsonify({"error": "打包失败：%s" % e}), 500
    resp = send_file(zip_path, mimetype="application/zip", as_attachment=True,
                     download_name="rtf_tasks_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S"))
    resp.call_on_close(lambda: _safe_remove(zip_path))
    if unknown:
        log.info("[合并下载] 跳过 %d 个未完成/不存在的任务：%s", len(unknown), unknown)
    return resp


@app.get("/api/task/<tid>/download")
def download(tid):
    t = manager.get(tid)
    if not t:
        return jsonify({"error": "任务不存在"}), 404
    if t["status"] != "done":
        return jsonify({"error": "任务未完成，当前状态: %s" % t["status"]}), 400
    fd, zip_path = tempfile.mkstemp(prefix="rtf_download_", suffix=".zip", dir=OUTPUT_DIR)
    os.close(fd)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for entry in t["results"]:
            for p in entry["outputs"]:
                if os.path.exists(p):
                    rel = os.path.relpath(p, t["out_dir"])
                    if rel.startswith(".."):
                        rel = os.path.basename(p)
                    z.write(p, rel)
    resp = send_file(zip_path, mimetype="application/zip", as_attachment=True,
                     download_name="%s_output.zip" % tid)
    resp.call_on_close(lambda: _safe_remove(zip_path))

    # 下载完成：设置延迟删除时间（本次修正：manager.get() 返回的是副本，
    # 原来改副本等于没改；现在用 set_clean_after 真正写到原任务上）
    if t.get("is_temp"):
        manager.set_clean_after(tid, time.time() + DOWNLOAD_CLEAN_DELAY)
        log.info("[下载] 任务 %s 下载完成，将在 %d 秒后清理: %s",
                 tid, DOWNLOAD_CLEAN_DELAY, t["out_dir"])

    return resp


if __name__ == "__main__":
    # 启动诊断
    print("=" * 60)
    print(">>> BASE_DIR     = %s" % BASE_DIR)
    print(">>> OUTPUT_DIR   = %s (存在=%s)" % (OUTPUT_DIR, os.path.isdir(OUTPUT_DIR)))
    print(">>> UPLOAD_DIR   = %s (存在=%s)" % (UPLOAD_DIR, os.path.isdir(UPLOAD_DIR)))
    print(">>> 清理策略：uploads=%d天, tmp任务=%d天, 下载延迟=%d秒"
          % (UPLOAD_EXPIRE_DAYS, TMP_TASK_EXPIRE_DAYS, DOWNLOAD_CLEAN_DELAY))
    print(">>> 架构：Flask 主进程 + %d 个转换子进程（各持独立 Word 实例）" % WORKERS)
    print(">>> 并发保护：队列上限=%d, 上传上限=%dMB, 大文件跨进程互斥=%.1fMB"
          % (QUEUE_MAX, MAX_UPLOAD_MB, LARGE_RTF_BYTES / (1024.0 * 1024.0)))
    print(">>> Word 生命周期：空闲释放=%s秒, 每任务一实例=%s, 回收阈值=%d文件/%d秒, RPC重试=%d次"
          % (WORD_IDLE_TIMEOUT or "永不", WORD_PER_TASK,
             WORD_RECYCLE_FILES, WORD_RECYCLE_SECONDS, RPC_RETRY))
    print(">>> 看门狗：worker %d 秒无心跳 -> 终止该转换子进程 + 强杀其 WINWORD + 只失败该任务"
          % WORD_STUCK_TIMEOUT)
    if os.path.isdir(OUTPUT_DIR):
        print(">>> OUTPUT_DIR 内容: %s" % os.listdir(OUTPUT_DIR))
    if os.path.isdir(UPLOAD_DIR):
        print(">>> UPLOAD_DIR 内容: %s" % os.listdir(UPLOAD_DIR))

    # 拉起转换子进程 + 监管线程（必须在主进程里做，spawn 子进程会重新 import 本模块）
    manager.start()
    print(">>> 转换子进程已启动（懒加载 Word），监管线程运行中")

    # 启动统一清理线程
    cleaner = threading.Thread(target=_auto_clean_all, daemon=True)
    cleaner.start()
    print(">>> 清理线程已启动 (alive=%s)" % cleaner.is_alive())
    print("=" * 60)

    log.info("RTF 转换服务启动 http://%s:%s（多进程模式：%d 个转换子进程）",
             HOST, PORT, WORKERS)
    app.run(host=HOST, port=PORT, threaded=True, use_reloader=False)
