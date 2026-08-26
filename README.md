# WPS Skills

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

WPS Skills 让智能体通过本地 Python 桥接操控 WPS Excel、PPT 和 Word。当前运行时不依赖 MCP、Node.js、外网或 pip 安装。

## 运行架构

```text
智能体 / 用户
  → python scripts/call.py <action> ...
  → HTTP 127.0.0.1:58891/dispatch
  → bridge/server.py 路由到 Excel / PPT / Word 控制器
  → Windows: 持久 PowerShell line-RPC → WPS COM
    Linux: 进程内 openpyxl / OpenXML 文件后端
```

一次 `call.py` 只执行一个 Action。多 Action 的任务由智能体逐步编排；执行桥不创建 `task-id`。

## 平台

| 平台 | 支持情况 | 后端 |
|---|---|---|
| Windows | 支持 | PowerShell 5.1 + WPS COM：`Ket.Application` / `Kwpp.Application` / `Kwps.Application` |
| Linux | 支持文件级自动化 | vendored openpyxl（Excel）及标准库 OpenXML（PPT/Word） |
| macOS | 不支持真实 WPS 自动化 | 没有对应控制器 |

Python 需要 3.8 或更高版本。Windows 需要已安装且正确注册 COM 的 WPS Office；无需提前手动打开应用，首个目标 Action 会复用现有实例或尝试创建它。

## 快速使用

先检查环境：

```bash
python scripts/install.py --check
```

直接调用 Action；`call.py` 会自动启动本地桥接服务：

```bash
# Excel
python scripts/call.py setCellValue '{"cell":"A1","value":42}'

# PPT
python scripts/call.py createPresentation '{}'
python scripts/call.py addSlide '{"layout":"blank"}'

# Word
python scripts/call.py createDocument '{}'
python scripts/call.py insertText '{"text":"hello"}'
```

PowerShell 5.1 下推荐参数文件，避免 shell 改写 JSON 引号：

```powershell
python scripts/call.py addSlide --params-file C:\tmp\slide.json
```

唯一归属的 Action 自动路由。重名 Action 必须显式指定应用：

```bash
python scripts/call.py findReplace --app word --params-file replace.json
python scripts/call.py insertImage --app ppt --params-file image.json
```

缺少 `--app` 会返回 `AMBIGUOUS_ACTION` 和候选应用，不会猜测并操作错误的软件。直接 HTTP 调用时使用与 `action` 同级的 `app`。

完整 Action 契约和操作清单见 [SKILL.md](SKILL.md)，整条执行链说明见 [understand.md](understand.md)。

## Action trace

每个 Action 默认创建一个结构化 JSONL trace。成功和失败响应都返回：

```json
{
  "success": false,
  "error": "…",
  "traceId": "act-20260826-…",
  "traceLog": "C:\\...\\wps-skills\\logs\\traces\\2026-08-26\\act-20260826-….jsonl"
}
```

默认日志位置：

```text
logs/traces/YYYY-MM-DD/<traceId>.jsonl
logs/server-YYYY-MM-DD.log
```

- `traceId` 贯穿 `call.py → HTTP → 路由 → 控制器 → PowerShell/COM`。
- 默认 `WPS_TRACE=info` 只记录低敏元数据、耗时和错误。
- 实机复现前可设置 `WPS_TRACE=debug`，增加脱敏后的参数和响应摘要。
- `WPS_TRACE_DIR` 可覆盖日志根目录；skill 目录不可写时，Windows 降级到 `%LOCALAPPDATA%\wps-skills\logs`。
- 所有候选目录都不可写时，Action 仍执行，响应以 `traceLog:null`/`traceWarning` 明确降级。
- trace 和 server 日志只保留 24 小时，自动清理不会触碰其他项目文件。

## 路由和可靠性

- HTTP 服务单线程执行，避免 PowerShell 单行协议交错。
- 每次 PowerShell 尝试使用独立 `reqId`；同一 Action 的自动重试保持相同 `traceId`。
- stderr 会被持续排空并写入对应 Action trace，避免管道阻塞。
- 单 Action 超过 60 秒会终止桥接进程；可恢复 COM 故障会自动重连并重试一次。
- `saveAs` 等通用 Action 建议显式指定应用；缺失时可以按目标文件扩展名推断。

## 验证

不需要 Windows 或 WPS 的单元测试：

```bash
PYTHONPATH=bridge python -m unittest \
  bridge/test_action_trace.py \
  bridge/test_server_routing.py \
  bridge/test_controller_trace.py
```

实机连通和功能检查：

```bash
python scripts/test.py
python scripts/test_functional.py
```

## 重要文档

- [SKILL.md](SKILL.md)：智能体调用约定与 Action 清单
- [understand.md](understand.md)：端到端执行链、任务边界和排障说明
- [CONTEXT.md](CONTEXT.md)：项目统一术语
- [docs/adr/0001-explicit-app-for-ambiguous-actions.md](docs/adr/0001-explicit-app-for-ambiguous-actions.md)：重名 Action 路由决策
- [docs/adr/0002-action-level-tracing-boundary.md](docs/adr/0002-action-level-tracing-boundary.md)：Action trace 边界决策

## 许可证

MIT
