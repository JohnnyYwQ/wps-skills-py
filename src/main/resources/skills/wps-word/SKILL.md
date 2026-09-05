---
name: wps-word
description: 在 Windows WPS Writer 中创建、读取和编辑 Word 文档，设置文字与页面格式，插入表格和图片，保存已有文档并导出 PDF。用户要求通过 WPS 操作 Word 文档时使用。
---

# WPS Word

使用本 Skill 的 Python 入口和正式 Action Contracts 完成用户的 Word 任务。执行环境需要 Windows、Python 3.8+、已注册 `KWPS.Application` 的 WPS Writer；无需第三方 Python 包。能力查询也可在 macOS/Linux 上运行。

所有下列路径都相对于本 `SKILL.md` 所在的目录。先确定这个绝对目录，不要假设当前工作目录是仓库。入口为 `scripts/word.py`；安装时必须保留完整 Skill 目录及随包的 `runtime/`。

## 1. 确定文档与保存意图

- 用户给出已有 `.docx` 文件时，使用 Windows 主机上的绝对路径执行 `openDocument`；不能把打开失败解释为新建，也不能使用当前活动窗口或焦点猜测目标。
- 用户要求新建时使用 `createDocument`。每个 Action Session 只绑定一个文档；后续操作始终复用这个 Session。
- 修改已有文件默认需要显式 `save`，除非用户要求保留未保存状态。只读任务不保存。
- 如果目标已有未保存修改，且本次任务需要保存，在第一次修改前说明保存会包含这些已有修改，并取得确认；已有明确授权可直接沿用。
- 当前 `saveAs` 不可用，`save` 只能保存已有路径的文档。用户要求新建并保存到新 `.docx` 路径时，先说明能力限制，确定用户接受的交付方式，不能在最后才发现无法保存。新建且没有输出路径的文档可保留打开、未保存；PDF 导出不会替代 DOCX 保存。

## 2. 发现并解析能力

先读取由正式 Application Contract Set 生成的 Action Index：

```powershell
python "<skill-dir>/scripts/word.py" --app word --index
```

根据任务选择 Actions，然后一次解析所需的完整契约。例如新建、写入和检查：

```powershell
python "<skill-dir>/scripts/word.py" --app word --resolve createDocument writeContent inspectDocument
```

这两个命令不会启动 Session、PowerShell 或 WPS。解析结果包含参数与结果 schema、约束、错误、验证要求和参数示例。只有 `status: complete` 才可按当前计划执行；`partial`/`failed` 时先调整计划。查询不授予执行权限，也不会执行多个 Actions。

参数以查询到的契约为准，不从 COM 文档、旧接口或记忆猜测。当前正式能力包括文档创建/打开、内容写入/读取/查找/替换、表格/图片、页眉页脚/页面布局/分隔符、原位保存和 PDF 导出。

## 3. 在一个会话中执行

执行前读 [references/session.md](references/session.md)，其中提供可直接运行的 Python 任务脚本和 Session Client 的用法。使用 `open_session()` 与 `client.call(address, params)`；每次调用都传入显式的 `{"app":"word","action":"..."}`。

每次完整响应返回后，再根据结果决定下一步。`call` 在 `failed` 或 `unknown` 时抛出 `ActionFailed`，其 `response` 保留完整真实错误；不要捕获后无条件继续后续修改。不要将一份预写好的 JSONL 动作列表整体管道输入 Session。

需要文字、范围或格式操作时，读 [references/content.md](references/content.md)。已有文件的编辑先检查相关内容；替换优先使用有明确匹配数量约束的查找/替换，而不是猜测字符位置。

## 4. 验证、保存与结束

读 [references/verification.md](references/verification.md)，按用户要求检查内容、格式和结构。必要时分页读取 `inspectDocument` 的剩余范围，不能把截断结果当成全文。

需要保存时，内容验证通过后显式 `save`，检查保存响应的 artifact 与状态。导出 PDF 时检查导出响应。结束 Session 只释放自动化资源，不隐式保存，也不关闭用户文档。

最终向用户说明：完成了什么、验证了什么、文档是否已保存、实际输出路径以及是否仍保持打开。Task Outcome 与 Session Outcome 分开判断：清理成功不代表编辑成功，清理失败也不会抹去已经验证的文件结果。

## 失败后的边界

`unknown` 或响应丢失意味着操作可能已部分发生，不能自动重试写入、创建替代文档或切换路径。先进行可用的只读验证。Session 已终止时，只能在仍能明确定位同一文档且能安全重新绑定的情况下建立新 Session；无法定位的未保存文档需要用户协助，不能重建来掩盖原任务状态。

不要绕过 Document Lease/Quarantine 或直接写 COM 来绕过失败。`--debug-close-created-document` 只用于明确可丢弃的受控测试，不用于用户文档任务。
