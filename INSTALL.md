# WPS Skills 安装与实机准备

当前根目录 Skill 使用纯 Python 本地 HTTP 桥，不需要 MCP Server、Node.js、WPS JS 加载项、外网或 pip 安装。

## 1. 平台要求

### Windows

- Python 3.8 或更高版本。
- Windows PowerShell 5.1（系统自带即可）。
- 已安装 WPS Office，且三个 COM ProgID 至少注册了需要使用的应用：
  - Excel：`Ket.Application`
  - PPT：`Kwpp.Application`
  - Word：`Kwps.Application`

不要求提前手动打开 WPS。首个目标 Action 会复用已有实例；没有时会尝试通过 COM 创建应用。

### Linux

- Python 3.8 或更高版本。
- 文件级 Excel/PPT/Word 自动化不要求安装 WPS，也不需要 pip；Excel 依赖已随 `vendor/` 分发。
- 安装 WPS 或 LibreOffice 仅用于可选的 GUI 预览/PDF 转换。

### macOS

当前 Python 执行桥不支持 macOS WPS 自动化。可以启动 HTTP 服务，但真实 WPS Action 会返回平台不支持。

## 2. 环境检查

在项目根目录运行：

```bash
python scripts/install.py --check
```

Windows 上检查器会查询 WPS COM 注册；Linux 上会检查文件后端及可选的 WPS 命令行。

## 3. 注册 Skill

将本项目根目录的 `SKILL.md` 注册到所用智能体宿主，或把整个根目录作为本地 Skill 目录加载。不同宿主的注册位置不同；注册后应确认智能体实际读取的是根目录这份 `SKILL.md`，不是旧的 MCP/加载项说明。

即使尚未注册自然语言 Skill，也可以显式调用执行入口：

```bash
python scripts/call.py getContext '{}'
```

## 4. 首次调用

`call.py` 会检查 `http://127.0.0.1:58891/health`，未运行时自动后台启动 `bridge/server.py`，不需要单独起服务。它只复用 checkout 路径和代码指纹完全匹配的新版实例；如果端口上是旧版服务、另一 checkout 或旧代码，会拒绝派发并给出明确状态，不会把 Action 发错进程。

如果返回 `BRIDGE_UNAVAILABLE`，端口仍被占用但健康检查暂时没有完成，常见原因是 bridge 正在执行长 Action。不要重复启动或强杀，先等待当前 Action 结束后重试 `python scripts/service.py status`。

Windows PowerShell 5.1 下建议使用 JSON 参数文件：

```powershell
python scripts/call.py addSlide --params-file C:\tmp\slide.json
python scripts/call.py findReplace --app word --params-file C:\tmp\replace.json
```

调用前先查询 Windows Action Contract；查询不会启动 bridge 或 WPS：

```bash
python scripts/actions.py search replace
python scripts/actions.py describe findReplace --app word
```

如需前台观察服务，也可以单独运行：

```bash
python scripts/start.py
```

## 5. 服务生命周期

bridge 用来维持一个多 Action 任务中的 WPS/文件状态。任务完成时必须先保存并验证，再显式停止：

```bash
python scripts/service.py status
python scripts/service.py stop
```

修改运行时代码后，使用下面的命令协作式关闭旧实例并启动新实例：

```bash
python scripts/service.py restart
```

正常停止会释放三个应用控制器；Windows PowerShell/COM 桥先接收 `EXIT`，超时后才会被强制回收。若调用方意外中断，服务默认在最后一个 Action 完成后空闲 15 分钟退出。可在服务启动前设置 `WPS_BRIDGE_IDLE_SECONDS`（秒）调整，设为 `0` 会关闭自动回收。

另一 checkout 或缺少实例身份的旧版 bridge 可能持有未保存状态，因此工具不会自动强杀。先确认并保存对应文档，再处理原进程；仅新版、可识别的另一 checkout 可在明确需要时用 `python scripts/service.py stop --takeover` 协作式关闭。

首次升级若 `status` 显示 `legacy`，先定位监听进程并核对启动路径，确认其文档已经保存后再人工终止：

```powershell
# Windows PowerShell
$bridgePid = (Get-NetTCPConnection -LocalPort 58891 -State Listen).OwningProcess
Get-CimInstance Win32_Process -Filter "ProcessId=$bridgePid" | Select-Object ProcessId, ExecutablePath, CommandLine
Stop-Process -Id $bridgePid
```

```bash
# macOS / Linux
lsof -nP -iTCP:58891 -sTCP:LISTEN
ps -p <PID> -o pid,ppid,etime,command
kill -TERM <PID>
```

不要跳过路径核对，也不要在可能存在未保存状态时直接强杀。

## 6. 实机 trace

Action trace 默认开启。正常响应和错误响应都会返回 `traceId` 与 `traceLog`。

默认日志位于：

```text
<skill根目录>/logs/traces/YYYY-MM-DD/<traceId>.jsonl
<skill根目录>/logs/server-YYYY-MM-DD.log
```

LLM 自动调用场景建议直接编辑 `bridge/action_trace.py` 顶部的手动开关：

```python
TRACE_LEVEL = "info"  # 常规模式；排障时把 info 改成 debug
```

修改后，后续 `call.py` 进程会自动加载；已经运行的桥接服务执行 `python scripts/service.py restart` 后加载新值，LLM 的 Action 调用命令本身无需改变。

skill 目录不可写时，Windows 降级到 `%LOCALAPPDATA%\wps-skills\logs`。也可以在启动桥接服务前用环境变量临时覆盖代码开关：

```powershell
$env:WPS_TRACE_DIR = "D:\wps-traces"
$env:WPS_TRACE = "debug"
python scripts/call.py getAppInfo --app ppt '{}'
```

`debug` 只增加脱敏摘要，不落完整正文或凭据。日志超过 24 小时后由后续 Action 惰性清理，清理范围不包含项目其他文件。

## 7. 验证

无需 WPS 的回归测试：

```bash
PYTHONPATH=bridge python -m unittest \
  bridge/test_action_trace.py \
  bridge/test_action_manifest.py \
  bridge/test_action_catalog.py \
  bridge/test_server_routing.py \
  bridge/test_controller_trace.py \
  bridge/test_service_lifecycle.py \
  bridge/test_server_lifecycle.py \
  bridge/test_service_cli.py
```

Windows/Linux 目标机上的连通和功能检查：

```bash
python scripts/test.py
python scripts/test_functional.py
```

## 8. 旧安装脚本说明

仓库内若仍保留 `wps-office-mcp/`、`wps-claude-addon/`、`wps-claude-assistant/` 或旧的 Node/JS 安装脚本，它们属于此前的 MCP/加载项方案，不是当前根目录 `call.py → server.py → controller` 执行链的依赖。不要为了运行当前 Skill 额外安装这些组件。
