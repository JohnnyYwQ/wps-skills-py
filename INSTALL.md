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

`call.py` 会检查 `http://127.0.0.1:58891/health`，未运行时自动后台启动 `bridge/server.py`，不需要单独起服务。

Windows PowerShell 5.1 下建议使用 JSON 参数文件：

```powershell
python scripts/call.py addSlide --params-file C:\tmp\slide.json
python scripts/call.py findReplace --app word --params-file C:\tmp\replace.json
```

如需前台观察服务，也可以单独运行：

```bash
python scripts/start.py
```

## 5. 实机 trace

Action trace 默认开启。正常响应和错误响应都会返回 `traceId` 与 `traceLog`。

默认日志位于：

```text
<skill根目录>/logs/traces/YYYY-MM-DD/<traceId>.jsonl
<skill根目录>/logs/server-YYYY-MM-DD.log
```

skill 目录不可写时，Windows 降级到 `%LOCALAPPDATA%\wps-skills\logs`。也可以在启动桥接服务前设置：

```powershell
$env:WPS_TRACE_DIR = "D:\wps-traces"
$env:WPS_TRACE = "debug"
python scripts/call.py getAppInfo --app ppt '{}'
```

`debug` 只增加脱敏摘要，不落完整正文或凭据。日志超过 24 小时后由后续 Action 惰性清理，清理范围不包含项目其他文件。

## 6. 验证

无需 WPS 的回归测试：

```bash
PYTHONPATH=bridge python -m unittest \
  bridge/test_action_trace.py \
  bridge/test_server_routing.py \
  bridge/test_controller_trace.py
```

Windows/Linux 目标机上的连通和功能检查：

```bash
python scripts/test.py
python scripts/test_functional.py
```

## 7. 旧安装脚本说明

仓库内若仍保留 `wps-office-mcp/`、`wps-claude-addon/`、`wps-claude-assistant/` 或旧的 Node/JS 安装脚本，它们属于此前的 MCP/加载项方案，不是当前根目录 `call.py → server.py → controller` 执行链的依赖。不要为了运行当前 Skill 额外安装这些组件。
