# 启动与持有一个 Word Action Session

## 运行位置

下面的 Python 任务脚本必须在安装了 WPS 的 Windows 主机上执行。若需要让用户看到窗口，应在该用户已登录的桌面会话中启动。SSH 通常运行在不同会话，不能承诺从 SSH 直接启动的 WPS 会出现在远程桌面；只有任务明确需要远程执行时，才使用环境中已授权的桌面启动机制。本 Skill 不内置主机名、账号、计划任务或远程凭据。

能力查询可在任意支持的 Python 环境执行；它和实际运行使用同一个正式 Contract Set。Windows 文件路径不会自动映射成调用方 Mac 的路径。

## 使用 Python Session Client

将下面示例保存为 UTF-8 编码的 Python 文件，替换 `SKILL_DIR` 为本 Skill 在 Windows 上的真实绝对目录，然后运行 `python task.py`。仅当用户要求新建并保留未保存文档时，采用这个示例；不要把已有文件任务改成新建。

这个最小任务创建一份含标题与中英文段落的文档，读回检查，再结束 Session。执行前先解析 `createDocument`、`writeContent`、`inspectDocument` 的完整契约。

```python
import json
from pathlib import Path
import sys

SKILL_DIR = Path(r"C:\path\to\wps-word")
sys.path.insert(0, str(SKILL_DIR / "scripts"))
from word import ActionFailed, SessionClientError, open_session

client = open_session()
try:
    with client:
        created = client.call({"app": "word", "action": "createDocument"}, {})
        written = client.call({"app": "word", "action": "writeContent"}, {
            "anchor": {"kind": "documentEnd"},
            "blocks": [
                {"kind": "heading", "level": 1, "runs": [
                    {"text": "项目周报", "format": {"bold": True, "fontSizePt": 20}}
                ]},
                {"kind": "paragraph", "runs": [
                    {"text": "本周已完成文档自动化验证。Word automation is ready."}
                ]},
            ],
        })
        inspected = client.call({"app": "word", "action": "inspectDocument"}, {
            "scope": {"kind": "document"},
            "limits": {"maxTextCharacters": 4096, "maxParagraphs": 64, "maxRuns": 256},
        })
        data = inspected["data"]
        if data["truncated"] or "项目周报" not in data["text"]:
            raise RuntimeError("Document verification did not establish the requested text")
        if data["structure"]["headingCount"] != 1:
            raise RuntimeError("Document verification did not establish one heading")
        if data["documentState"]["persistenceState"] != "unsaved":
            raise RuntimeError("Unexpected persistence state")
        print(json.dumps({"verified": True, "documentState": data["documentState"]}, ensure_ascii=True))
except ActionFailed as exc:
    print(json.dumps(exc.response, ensure_ascii=True), file=sys.stderr)
    raise
except SessionClientError as exc:
    print(str(exc), file=sys.stderr)
    print(client.stderr, file=sys.stderr)
    # may_have_effect=True means that an attempted Action has no usable response.
    # Do not replay it. last_response may still describe an earlier completed Action.
    raise
finally:
    print(json.dumps({
        "sessionOutcome": client.session_outcome,
        "cleanupError": str(client.cleanup_error) if client.cleanup_error else None,
    }, ensure_ascii=True))
```

`open_session()` 尚未启动进程；`with client` 启动一个 Python Session Host 并等待 `session.ready`。Host 在需要时启动一条 Session 自有的 PowerShell bridge。离开 `with` 时发送正常关闭请求并等待清理，保留文档打开。

## Client 的接口与行为

- `open_session(timeout=60)`：每次启动、调用或关闭的最长等待时间，单位为秒。任务若确实需要更久，在执行前设置合理值；超时不是重试理由。
- `client.call(address, params)`：只提交一个 Action。成功返回完整 Action Response，结果在 `response["data"]` 中；失败或结果不确定时抛出 `ActionFailed`，完整响应在 `exc.response` 中。
- `client.ready`：启动记录，包括 Session 身份及 trace 路径。
- `client.last_response`：最近一个完整 Action Response，不代表丢失响应的当前操作。
- `client.can_execute`：Client 是否仍允许提交。每次提交还会检查已收到的关闭记录；可用状态不是重试授权。
- `client.session_outcome`：结束后报告的 Session Outcome；没有可确认的结束记录时为 `None`。
- `client.cleanup_error`：清理或结束通道的问题。原始 Action 错误和清理失败并存时，优先抛出 Action 错误，并保留清理错误。
- `client.close()`：可重复调用；不会保存或关闭 WPS 文档。通常交给上下文管理器调用。

Client 不支持多个调用者并发使用，也不会保存一个可供下一次进程重新连接的 Session ID。需要按智能体工具调用逐步操作时，应保持同一个 Python 进程/REPL；不要每一步重新运行 `task.py`。

如果调用方直接持有持续的终端通道，也可启动：

```powershell
python "<skill-dir>/scripts/word.py" --session --app word
```

这是原始 Session Host，使用方必须自行遵循下面的协议。优先使用上面的 Client，减少重复实现。

## Protocol v1

Host 首先输出 `session.ready`。请求是单行 JSON：

```json
{"address":{"app":"word","action":"createDocument"},"params":{}}
```

收到完整 Action Response 后才能发送下一个请求。若先收到 `session.closing`，立即停止发送，但继续读取当前 Action Response（如果存在）以及最终的 `session.closed`。关闭通知不等于 Action Response；不能用它覆盖后面的具体错误。

结束时发送 `{"control":"close"}`。EOF 或异常丢失通道会结束 Session。空闲 Session 会超时关闭，不能把它当作后台常驻服务。

JSONL 使用 ASCII-safe JSON 转义传输 Unicode；Python 通过 `json.dumps(..., ensure_ascii=True)` 编码，再通过 JSON 解码恢复原文。不要依赖 PowerShell 的默认本地代码页。
