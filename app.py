# -*- coding: utf-8 -*-
""" RTF 批量转换服务（纯 API 形态）
================================
引擎：Microsoft Word COM，格式保真度最高
队列：多 worker 并行消费（WORKERS 个独立 Word 实例）
进度：POST /api/convert -> GET /api/task/<id> 轮询
下载：GET /api/task/<id>/download 打包 zip

Word 生命周期（本次改造重点）：
    - 懒加载：worker 启动不再拉起 Word，真有任务才启动
    - 空闲释放：队列空 WORD_IDLE_TIMEOUT 秒 -> Quit()，WINWORD.EXE 退出
    - 主动回收：每 WORD_RECYCLE_FILES 个文件 / WORD_RECYCLE_SECONDS 秒换新实例
    - RPC 自愈：-2147023174(0x800706BA) 等通道断开时自动重建并重试单文件
    - 可选每任务一实例：WORD_PER_TASK=1，任务结束立刻释放，全程零常驻

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
# Word COM is not a normal CPU-bound worker pool. Multiple Word instances
# often contend over profiles, add-ins, and file locks; one stable worker is
# the safer default. Increase explicitly only after measuring this machine.
WORKERS = max(1, int(os.environ.get("WORKERS", "1")))
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
TMP_TASK_EXPIRE_DAYS = int(os.environ.get("TMP_TASK_EXPIRE_DAYS", "1"))   # 未下载/失败 tmp_ 保留天数
DOWNLOAD_CLEAN_DELAY = int(os.environ.get("DOWNLOAD_CLEAN_DELAY", str(1 * 3600)))  # 下载后延迟删除（秒，默认1小时）
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
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


def _acquire_instance_mutex():
    """Allow only one app.py process to own a port on this Windows host."""
    if os.name != "nt":
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


# ===================== 任务管理器 =====================
class TaskManager:
    _used_global = {}
    _used_global_lock = threading.Lock()
    _wps_global_lock = threading.Lock()
    _large_job_lock = threading.Lock()

    def __init__(self):
        self._queue = queue.Queue()
        self._tasks = {}
        # 可重入锁：清理线程会在持有锁的情况下调用 is_path_active()（它也要这把锁）
        self._lock = threading.RLock()
        self._request_cache = {}
        self._request_locks = {}
        self._request_cache_ttl = 10 * 60
        # ---- Word 生命周期相关状态 ----
        self._busy = 0                      # 正在处理中的任务数（防止空闲误杀）
        self._beat = time.time()            # 看门狗心跳：卡死检测用
        self._word_state = {"app": None}    # 全局持有，供 health/release 查询
        for _ in range(WORKERS):
            threading.Thread(target=self._worker, daemon=True).start()
        log.info("任务队列启动：%d 个并行 worker（Word 懒加载，空闲 %s 秒后释放）",
                 WORKERS, WORD_IDLE_TIMEOUT or "永不")

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
                "source_paths": [os.path.abspath(src) for src, _, _ in files3],
                "results": [], "log": [],
                "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "started": None, "finished": None,
                "clean_after": None,          # 下载后延迟删除时间戳
            }
        self._queue.put((tid, files3, out_dir))
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
    # Word 实例生命周期：懒加载 / 空闲释放 / 主动回收 / RPC 自愈
    # ------------------------------------------------------------------
    @staticmethod
    def _start_word():
        """启动 Word COM 实例（带重试 + 握手校验）。"""
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
                log.info("Word COM 转换引擎已启动")
                return w
            except Exception as e:
                log.warning("Word COM 启动失败（%d/%d）：%s",
                            attempt + 1, RPC_WORD_RELAUNCH_TRIES, e)
                time.sleep(1.0 * (attempt + 1))
        log.error("Word COM 连续 %d 次启动失败", RPC_WORD_RELAUNCH_TRIES)
        return None

    @staticmethod
    def _word_alive(word):
        """轻量体检：能读到 Documents.Count 说明通道还活着。"""
        if word is None:
            return False
        try:
            word.Documents.Count
            return True
        except Exception as e:
            return not _is_rpc_error(e)

    def word_alive(self):
        """给 /api/health 用：当前 Word 实例是否活着。"""
        return self._word_alive(self._word_state.get("app"))

    def _kill_stale_winword(self):
        """保险栓：干掉 Quit 后仍赖着不走的 WINWORD.EXE。
        默认关闭（WORD_KILL_STALE=1 才启用），因为会误杀用户自己开的 Word。"""
        try:
            import subprocess
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

    def _relaunch_word(self, old_word, tries=None):
        """彻底丢弃旧实例，重新拉起一个干净的 Word。"""
        if old_word is not None:
            try:
                old_word.Quit()
            except Exception:
                pass
            try:
                del old_word
            except Exception:
                pass
        gc.collect()
        for i in range(tries or RPC_WORD_RELAUNCH_TRIES):
            time.sleep(0.5 + i)
            w = self._start_word()
            if w is not None and self._word_alive(w):
                log.warning("[RPC] Word COM 实例已重建（第 %d 次尝试）", i + 1)
                return w
        return None

    def _relaunch_word_into(self, word_state):
        """重建并把新实例写回 word_state，供 worker 主循环复用。"""
        w = self._relaunch_word(word_state.get("app"))
        if w is None:
            return False
        word_state["app"] = w
        return True

    def _ensure_word(self, word_state, reason=""):
        """每次转换前体检，发现 RPC 断开就地重建。"""
        if self._word_alive(word_state.get("app")):
            return True
        log.warning("[RPC] 检测到 Word COM 通道断开（%s），开始重建实例", reason or "健康检查")
        new_word = self._relaunch_word(word_state.get("app"))
        if new_word is None:
            log.error("[RPC] Word COM 重建失败，本批后续文件将标记为失败")
            return False
        word_state["app"] = new_word
        return True

    def _release_word(self, word_state=None, reason=""):
        """主动释放 Word 实例：Quit + 置空 + gc，WINWORD.EXE 随之退出。"""
        state = word_state if word_state is not None else self._word_state
        word = state.get("app")
        if word is None:
            return False
        try:
            word.Quit()
        except Exception as e:
            log.warning("[Word] Quit 失败（%s）：%s", reason or "主动释放", e)
        try:
            del word
        except Exception:
            pass
        state["app"] = None
        gc.collect()
        if WORD_KILL_STALE:
            self._kill_stale_winword()
        log.info("[Word] 已释放实例（%s），空闲态不再占用进程/内存", reason or "主动释放")
        return True

    @staticmethod
    def _alloc_name(out_dir, base):
        """输出名只在这里分配一次，RPC 重试时复用同一个名字（不会退化成 xxx_1）。"""
        with TaskManager._used_global_lock:
            d = TaskManager._used_global.setdefault(out_dir, {})
            n = d.get(base, 0)
            d[base] = n + 1
        return base if n == 0 else "%s_%d" % (base, n)

    # ------------------------------------------------------------------
    # worker 看门狗：Python 无法 kill 卡死的线程，只能"放弃 + 换一个"
    # ------------------------------------------------------------------
    def _fail_stuck_task(self):
        """把卡住的任务就地判失败，别让前端永远停在 processing。"""
        with self._lock:
            stuck = [tid for tid, t in self._tasks.items()
                     if t["status"] in ("pending", "running")]
        for tid in stuck:
            try:
                self._finish(tid, "failed",
                             "转换引擎无响应（Word 卡死或弹出对话框），已放弃该任务并重启转换线程")
            except Exception as e:
                log.warning("[看门狗] 标记卡死任务失败出错 %s: %s", tid, e)
        return len(stuck)

    def _worker(self):
        """监管壳：内部真正干活的 _worker_loop 卡死后，拉起替补继续服务。"""
        while True:
            self._beat = time.time()
            th = threading.Thread(target=self._worker_loop, daemon=True)
            th.start()
            while th.is_alive():
                th.join(5)
                if not th.is_alive():
                    break
                if WORD_STUCK_TIMEOUT and self._busy > 0 and \
                        (time.time() - self._beat) > WORD_STUCK_TIMEOUT:
                    log.error("[看门狗] 转换线程 %d 秒无心跳（busy=%d），判定 Word 卡死，"
                              "放弃该线程并拉起替补", WORD_STUCK_TIMEOUT, self._busy)
                    self._fail_stuck_task()
                    self._busy = 0
                    if WORD_KILL_STALE:
                        self._kill_stale_winword()
                    # 旧线程的 app 引用已是僵尸（连 Quit 都会挂起），弃用旧槽位
                    self._word_state = {"app": None}
                    break
            log.warning("[看门狗] 转换线程已退出，%d 秒后拉起替补", WORD_WORKER_RESTART_DELAY)
            time.sleep(WORD_WORKER_RESTART_DELAY)

    # ------------------------------------------------------------------
    # worker：懒加载 + 空闲释放
    # ------------------------------------------------------------------
    def _worker_loop(self):
        pythoncom.CoInitialize()
        word_state = {"app": None}
        # WORKERS=1 时本线程即唯一转换线程，用 self._word_state 暴露给外部接口
        if self._word_state.get("app") is None:
            self._word_state = word_state
        idle_since = None
        try:
            while True:
                self._beat = time.time()
                # ---- 取任务：带超时，超时即判定为空闲 ----
                try:
                    tid, files3, out_dir = self._queue.get(timeout=WORD_POLL_INTERVAL)
                except queue.Empty:
                    if word_state["app"] is None:
                        idle_since = None
                        continue
                    if self._busy > 0:
                        idle_since = None
                        continue
                    if not WORD_IDLE_TIMEOUT:
                        continue
                    now = time.time()
                    if idle_since is None:
                        idle_since = now
                    elif now - idle_since >= WORD_IDLE_TIMEOUT:
                        self._release_word(word_state, reason="空闲 %d 秒" % WORD_IDLE_TIMEOUT)
                        idle_since = None
                    continue

                idle_since = None
                self._busy += 1
                try:
                    with self._lock:
                        task = self._tasks.get(tid)
                        if task and task["status"] == "pending":
                            task["status"] = "running"
                            task["started"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            task["current"] = "正在启动 Word 转换引擎"
                            task["phase"] = "starting_engine"
                            task["heartbeat"] = time.time()

                    # ---- 懒加载：只有真要转了才启动 Word ----
                    if word_state["app"] is None:
                        word_state["app"] = self._start_word()
                        if word_state["app"] is None:
                            self._finish(tid, "failed", "Word COM 初始化失败")
                            continue
                    try:
                        is_large_job = any(
                            os.path.getsize(src) >= LARGE_RTF_BYTES
                            for src, _, _ in files3
                        )
                        if is_large_job:
                            with TaskManager._large_job_lock:
                                restart_word = self._process_optimized(
                                    tid, files3, out_dir, word_state)
                        else:
                            restart_word = self._process_optimized(
                                tid, files3, out_dir, word_state)

                        # ---- 每任务一实例：任务收尾立刻释放 ----
                        if WORD_PER_TASK and not restart_word:
                            self._release_word(word_state, reason="任务 %s 结束（每任务一实例）" % tid)

                        if restart_word:
                            self._release_word(word_state, reason="大文件任务结束")
                    except Exception as e:
                        log.error("[%s] 任务处理异常，重建 Word 实例: %s", tid, e)
                        with self._lock:
                            task = self._tasks.get(tid)
                            unfinished = task and task["status"] not in ("done", "failed")
                        if unfinished:
                            self._finish(tid, "failed", "任务处理异常：%s" % str(e)[:300])
                        self._release_word(word_state, reason="任务异常，回收后下次重建")
                finally:
                    self._busy = max(0, self._busy - 1)
        finally:
            self._release_word(word_state, reason="worker 退出")
            pythoncom.CoUninitialize()

    # ------------------------------------------------------------------
    # 批量转换主流程：体检 + RPC 重试 + 主动回收
    # ------------------------------------------------------------------
    def _process_optimized(self, tid, files3, out_dir, word_state):
        """word_state: {"app": word} —— 字典传递，重建后能回传 worker。"""
        t = self._tasks[tid]
        with self._lock:
            t["status"] = "running"
            t["started"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            self._finish(tid, "failed", "output directory error: %s" % e)
            return False

        wps_state = {"app": None}
        large_seen = False
        files_since_recycle = 0
        recycle_at = time.time() + WORD_RECYCLE_SECONDS if WORD_RECYCLE_SECONDS else 0

        try:
            with self._lock:
                t["total"] = len(files3)

            for i, (src, base, fmts) in enumerate(files3, 1):
                self._beat = time.time()   # 看门狗心跳：单文件卡死也能被检测到
                entry = {"src": os.path.basename(src), "ok": False,
                         "outputs": [], "formats": {}, "error": None}
                try:
                    # ---- 0) 主动回收 Word：句柄/内存累积是崩溃的头号诱因 ----
                    need_recycle = (
                        (WORD_RECYCLE_FILES and files_since_recycle >= WORD_RECYCLE_FILES)
                        or (WORD_RECYCLE_SECONDS and time.time() > recycle_at)
                    )
                    if need_recycle:
                        if self._relaunch_word_into(word_state):
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

                    with self._lock:
                        t["current"] = ("正在打开 %s（%.1f MB）" % (os.path.basename(src), size_mb)
                                        if is_large else "正在打开 " + os.path.basename(src))
                        t["phase"] = "opening"
                        t["heartbeat"] = time.time()

                    # ---- 2) 输出名只分配一次 ----
                    name = self._alloc_name(out_dir, base)

                    # ---- 3) 单次转换 + RPC 自愈重试 ----
                    results = None
                    last_err = None
                    for attempt in range(RPC_RETRY + 1):
                        if not self._ensure_word(word_state, "转换 %s 前" % os.path.basename(src)):
                            raise RuntimeError("Word COM 不可用且无法重建")
                        try:
                            results = self._convert_one(
                                word_state["app"], wps_state, src, fmts,
                                out_dir, name, bool(t.get("split_dirs")))
                        except Exception as e:
                            last_err = e
                            results = None
                            if _is_rpc_error(e) and attempt < RPC_RETRY:
                                log.warning("[%s] %s 打开/保存时 RPC 断开（第 %d 次），重建 Word 后重试：%s",
                                            tid, os.path.basename(src), attempt + 1, e)
                                self._relaunch_word_into(word_state)
                                time.sleep(RPC_RETRY_SLEEP)
                                continue
                            raise
                        # _save_document 会把异常吞进 results，补一次检查
                        if _results_have_rpc(results) and attempt < RPC_RETRY:
                            log.warning("[%s] %s 保存时 RPC 断开（第 %d 次），重建 Word 后重试",
                                        tid, os.path.basename(src), attempt + 1)
                            self._relaunch_word_into(word_state)
                            time.sleep(RPC_RETRY_SLEEP)
                            continue
                        break

                    if results is None:
                        raise last_err or RuntimeError("转换未返回结果")

                    # ---- 4) 结果汇总 ----
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
                with self._lock:
                    t["results"].append(entry)
                    t["done"] = i
                    t["progress"] = round(i * 100.0 / len(files3), 1)
                    t["current"] = None
                    t["phase"] = "file_done"
                    t["heartbeat"] = time.time()
                    t["log"].append("[%d/%d] %s -> %s" % (i, len(files3), entry["src"], detail))
                log.info("[%s] %d/%d %s -> %s", tid, i, len(files3), entry["src"], detail)
        finally:
            self._close_app(wps_state.get("app"))
            wps_state["app"] = None

        self._finish(tid, "done", None)
        return large_seen

    @staticmethod
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

    @staticmethod
    def _convert_one(word, wps_state, src, formats, out_dir, name, split_dirs=False):
        """name 由调用方预先分配（重试时复用），不再在这里自增。"""
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
            return TaskManager._save_document(
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

    @staticmethod
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
                        wps = TaskManager._get_wps(wps_state)
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
                            with TaskManager._wps_global_lock:
                                TaskManager._wps_save_docx(wps, fallback_src, out_path)
                        finally:
                            if fallback_rtf:
                                _safe_remove(fallback_rtf)
                        results[fmt]["via_wps"] = True
                        results[fmt]["error"] = None
                else:
                    doc.SaveAs2(out_path, FileFormat=FORMAT_MAP[fmt])
                results[fmt]["ok"] = True
            except Exception as e:
                results[fmt]["error"] = str(e)[:300]
        return results

    @staticmethod
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

    @staticmethod
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

    def _finish(self, tid, status, error):
        with self._lock:
            t = self._tasks[tid]
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
                log.error("[%s] %s", tid, error)
            else:
                ok = sum(1 for r in t["results"] if r["ok"])
                t["log"].append("任务完成：%d/%d 成功" % (ok, t["total"]))
                log.info("[%s] 完成 %d/%d", tid, ok, t["total"])


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
.log{font-size:12px;color:#4b5563;background:#f8fafc;border:1px solid #edf0f4;border-radius:8px;padding:10px;margin-top:6px;max-height:180px;overflow:auto;white-space:pre-wrap;font-family:Consolas,monospace}
.hide{display:none}
#dlBtn[href="#"]{visibility:hidden;pointer-events:none}
</style>
</head>
<body>
<h1>RTF 批量转换</h1>
<div class="sub">转换引擎：Word · 支持 doc / docx / pdf · 空闲自动释放 · RPC 断开自动重建重试</div>

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
    <h3 style="margin:0 0 8px">任务列表</h3>
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
    <div class="log-label">全局处理记录</div>
    <div class="log" id="log"></div>
    <div style="margin-top:12px">
        <a class="btn" id="dlBtn" style="text-decoration:none;display:inline-block" href="#">下载结果 zip</a>
        <button class="btn2" style="margin-left:8px" onclick="releaseWord()">释放 Word 实例</button>
    </div>
</div>

<script>
var files = [];
var taskId = null;
var submitInFlight = false;
var submissionRequestId = null;
var pollDelay = 1200;
var pollTimer = null;
const uploadBox = document.getElementById("uploadBox");
const fileInput = document.getElementById("file-input");
const taskListEl = document.getElementById("taskList");
const taskMap = new Map();

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
  fetch('/api/word/release', { method: 'POST' })
    .then(function(r){ return r.json(); })
    .then(function(j){
      alert(j.error ? ('释放失败：' + j.error) : ('已释放，Word 当前存活=' + j.word_alive));
    })
    .catch(function(e){ alert('请求失败：' + e); });
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
      taskId = j.task_id;
      addRow(taskId, "批量任务", "-");
      poll();
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
  taskMap.set(tid, tr);
}

function poll() {
  fetch('/api/task/' + taskId)
    .then(function(r){
      if (!r.ok) throw new Error('任务已不存在或服务已重启，请重新提交');
      return r.json();
    })
    .then(function(t){
      var pct = typeof t.progress === 'number'
        ? t.progress
        : (t.total ? (t.done * 100.0 / t.total) : 0);
      pct = Math.max(0, Math.min(100, pct));
      var statusText = t.status === 'done' ? 'done'
        : t.status === 'failed' ? 'fail'
        : t.status === 'pending' ? 'queue'
        : 'processing';
      const tr = taskMap.get(taskId);
      if(!tr) return;
      const cls = "st-" + statusText;
      const barEl = document.getElementById(`bar_${taskId}`);
      if(barEl) barEl.style.width = pct.toFixed(1) + '%';

      let action = "";
      if (t.status === "done") {
        action = `<a class="btn" href="/api/task/${taskId}/download">下载 ZIP</a>`;
        document.getElementById('dlBtn').href = `/api/task/${taskId}/download`;
        document.getElementById('goBtn').disabled = false;
      }else if(t.status === "failed"){
        document.getElementById('goBtn').disabled = false;
      }

      tr.innerHTML = `
        <td>批量任务 ${taskId.slice(0,8)}</td>
        <td class="size">${t.total}个文件</td>
        <td class="${cls}">${statusText}</td>
        <td><div class="progress-track"><div class="progress-fill" id="bar_${taskId}" style="width:${pct}%"></div></div></td>
        <td>${t.current? "正在处理："+t.current : (t.status==="done"?"全部完成":"等待转换引擎响应")}</td>
        <td>${action}</td>`;

      setLog(t.log.join('\n'));
      if (t.status === 'done' || t.status === 'failed') {
        return;
      } else {
        pollDelay = t.status === 'pending'
          ? Math.min(5000, pollDelay + 400)
          : 1200;
        pollTimer = setTimeout(poll, pollDelay);
      }
    })
    .catch(function(e){
      if (pollTimer) clearTimeout(pollTimer);
      setLog('轮询异常: ' + e.message);
    });
}

function setLog(s) { document.getElementById('log').textContent = s || ''; }
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


@app.get("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html; charset=utf-8")


@app.get("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "pid": os.getpid(),
        "workers": WORKERS,
        "word_alive": manager.word_alive(),          # 一眼看出现在占不占
        "word_idle_timeout": WORD_IDLE_TIMEOUT,
        "word_per_task": WORD_PER_TASK,
        "queued": manager._queue.qsize(),
        "busy": getattr(manager, "_busy", 0),
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


@app.post("/api/word/release")
def word_release():
    """手动释放 Word 实例。有任务在跑时拒绝（409）。"""
    with manager._lock:
        running = any(t["status"] in ("pending", "running")
                      for t in manager._tasks.values())
    if running or getattr(manager, "_busy", 0) > 0:
        return jsonify({"error": "有任务正在处理，拒绝释放"}), 409
    released = manager._release_word(None, reason="手动调用 /api/word/release")
    return jsonify({"released": released, "word_alive": manager.word_alive()})


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
        task_id = manager.submit_map(items, out_dir, split_dirs)
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
    task_id = manager.submit(files, formats, out_dir, split_dirs)
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

    # 下载完成：设置延迟删除时间（本次修正：用 DOWNLOAD_CLEAN_DELAY，而非下一个凌晨2点）
    if t.get("is_temp"):
        with manager._lock:
            t["clean_after"] = time.time() + DOWNLOAD_CLEAN_DELAY
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
    print(">>> Word 生命周期：空闲释放=%s秒, 每任务一实例=%s, 回收阈值=%d文件/%d秒, RPC重试=%d次"
          % (WORD_IDLE_TIMEOUT or "永不", WORD_PER_TASK,
             WORD_RECYCLE_FILES, WORD_RECYCLE_SECONDS, RPC_RETRY))
    print(">>> 看门狗：单文件/单任务无心跳超过 %d 秒判定卡死并换线程（残留 Word 强杀=%s）"
          % (WORD_STUCK_TIMEOUT, WORD_KILL_STALE))
    if os.path.isdir(OUTPUT_DIR):
        print(">>> OUTPUT_DIR 内容: %s" % os.listdir(OUTPUT_DIR))
    if os.path.isdir(UPLOAD_DIR):
        print(">>> UPLOAD_DIR 内容: %s" % os.listdir(UPLOAD_DIR))

    # 启动统一清理线程
    cleaner = threading.Thread(target=_auto_clean_all, daemon=True)
    cleaner.start()
    print(">>> 清理线程已启动 (alive=%s)" % cleaner.is_alive())
    print(">>> Word 未启动（懒加载），首次提交任务时才拉起")
    print("=" * 60)

    log.info("RTF 转换服务启动 http://%s:%s", HOST, PORT)
    app.run(host=HOST, port=PORT, threaded=True, use_reloader=False)
