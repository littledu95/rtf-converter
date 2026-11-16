# RTF 批量转换服务（纯 API）

把 RTF 批量转换为 doc / docx / pdf。转换引擎：
- **主引擎：Microsoft Word COM**（格式保真度最高），doc / docx / pdf 全格式
- **docx 兜底引擎：WPS COM**（Kwps.Application）——当 Word 对含大量图片/域的复杂报告（如 SAS 输出的 7MB 大 RTF）转 docx 报"文件发生错误"时，自动切换 WPS 完成 docx 输出
- 每格式独立处理：doc / docx / pdf 互不影响，哪个能出就出哪个

多人同时提交自动排队串行执行（Word 单实例安全），每个任务独立进度与日志。

## 目录



```
rtf-convert-server/

├── app.py            # Flask API + Word COM 转换核心 + 任务队列

├── start.bat         # 启动：后台静默 + PID 防重复 + 端口检测清理 + 依赖自动安装

├── stop.bat          # 停止

├── requirements.txt  # 依赖：flask、pywin32

├── server.log        # 服务日志（运行时生成）

├── server.pid        # 进程 PID 文件（运行时生成）

└── output/           # 转换结果（默认输出目录）
```

## 快速开始



1. 改 `start.bat` 顶部 `PYTHON` 路径（当前已指向 `D:\app\python3.10.10\python.exe`）

2. 双击 `start.bat`（首次会自动安装 flask、pywin32；需要装 Microsoft Word）

3. 浏览器打开 [http://127.0.0.1:8000/](http://127.0.0.1:8000/) 使用网页界面；或直接调用 API（见下）

4. 不用时双击 `stop.bat`

## 网页界面

打开 `http://127.0.0.1:8000/` 即用：



* 输入方式：**拖拽 .rtf 文件或整个文件夹**（自动收集文件夹内全部 .rtf），或**直接输入服务端文件夹路径**
* 勾选输出格式（doc / docx / pdf），可多选
* 输出目录可手填或选择文件夹，留空默认 `output/任务ID`
* **按格式分开存放**（默认开）：doc / docx / pdf 各输出到独立子文件夹，不勾选则全部混在一个目录
* 实时进度条 + 逐文件日志，完成后一键下载 zip（zip 内同样按格式分目录）

## API 说明



| 方法   | 路径                             | 说明                      |
| ---- | ------------------------------ | ----------------------- |
| GET  | `/api/health`                  | 健康检查                    |
| POST | `/api/convert`                 | 创建转换任务，返回 `task_id`     |
| GET  | `/api/task/{task_id}`          | 轮询：状态 / 进度 / 每文件结果 / 日志 |
| GET  | `/api/task/{task_id}/download` | 下载结果 zip（任务完成后）         |

### POST /api/convert 参数（表单 multipart）



| 参数        | 必填  | 说明                           |
| --------- | --- | ---------------------------- |
| `files`   | 二选一 | 上传的 .rtf 文件（可多个，字段名均为 files） |
| `path`    | 二选一 | 服务端目录路径，自动扫描全部 .rtf          |
| `formats` | 否   | 输出格式，逗号分隔，默认 `doc,docx,pdf`  |
| `out_dir` | 否   | 输出目录，默认 `output/<task_id>`   |
| `split_dirs` | 否 | `1` 时按格式分目录输出（`out_dir/doc/`、`out_dir/docx/`、`out_dir/pdf/`） |
| `map`     | 否   | 高级映射模式：JSON，每格式独立指定目录，如 `{"pdf":{"path":"D:\\a"},"doc":{"path":"D:\\b"}}` |
| `files_<fmt>` | 否 | 高级映射模式上传文件，如 `files_pdf`（可多文件），与 `map` 中同名格式互斥 |

### 调用示例

上传模式（curl）：



```
curl.exe -s -F "files=@C:\报告\a.rtf" -F "files=@C:\报告\b.rtf" -F "formats=docx,pdf" http://127.0.0.1:8000/api/convert
```

目录模式（curl）：



```
curl.exe -s -F "path=C:\报告\rtf目录" -F "formats=doc,docx,pdf" http://127.0.0.1:8000/api/convert
```

Python 调用（并入平台）：



```
import requests, time

r = requests.post("http://127.0.0.1:8000/api/convert",

&#x20;                 files=\[("files", open("a.rtf", "rb")), ("files", open("b.rtf", "rb"))],

&#x20;                 data={"formats": "doc,docx,pdf"})

tid = r.json()\["task\_id"]

while True:                       # 轮询进度

&#x20;   t = requests.get(f"http://127.0.0.1:8000/api/task/{tid}").json()

&#x20;   print(t\["status"], t\["progress"], "%")

&#x20;   if t\["status"] in ("done", "failed"):

&#x20;       break

&#x20;   time.sleep(2)

open("out.zip", "wb").write(requests.get(f"http://127.0.0.1:8000/api/task/{tid}/download").content)
```

任务详情返回字段：`status`(pending/running/done/failed)、`progress`(0-100)、`total`、`done`、

`results`(每个文件：src/ok/outputs/error)、`log`(逐文件处理日志)、`out_dir`。

## 配置项（环境变量）



| 变量           | 默认        | 说明                              |
| ------------ | --------- | ------------------------------- |
| `PORT`       | 8000      | 服务端口                            |
| `MAX_FILES`  | 100       | 单批文件数上限                         |
| `OUTPUT_DIR` | `output/` | 默认输出根目录                         |
| `API_TOKEN`  | 空         | 设置后请求需带请求头 `X-API-Key: <token>` |

## 并发与性能



* 所有任务进入内存队列，由单 worker 串行执行（Word COM 必须单实例），多人提交自动排队，互不干扰。

* 每个文件约 2\~5 秒，100 个文件约 5\~8 分钟。任务详情页可实时看进度和日志。

* 服务重启后未完成任务丢失（内存队列）；如需持久化队列请告知。

## 故障排查



| 现象                       | 处理                                           |
| ------------------------ | -------------------------------------------- |
| 任务 failed：Word COM 初始化失败 | 确认运行机安装了 Microsoft Word（WPS 不支持）             |
| docx 输出为 FAIL：文件发生错误 | 已自动切换 WPS 兜底；若仍未装 WPS，任务会明确标记 docx 失败，doc/pdf 不受影响 |
| 端口被占用                    | start.bat 会自动清理同机 python 进程；非 python 占用需手动处理 |
| 服务重启后任务消失                | 内存队列特性，属预期                                   |