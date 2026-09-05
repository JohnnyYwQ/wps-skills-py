# WPS 自动化基础

[English](README.md) | [简体中文](README.zh-CN.md)

本项目为智能体提供 WPS 自动化能力，采用独立 Application Skill、明确的 Action Contracts 和单文档 Action Session 架构。目前已完成 Windows WPS Writer 的 Word 实现，并提供可独立安装的 `wps-word` Skill。

智能体可以通过 Skill 了解可用能力，再用 Python Session Client 创建或打开文档、执行编辑、验证结果并按需保存。Python Session Host 负责请求校验和会话管理，PowerShell bridge 通过 WPS COM 操作同一个实际文档。

## 当前能力

Word 正式 Application Contract Set 包含 13 个可执行 Action：

| Action | 作用 |
| --- | --- |
| `createDocument` | 创建一个新的未保存文档 |
| `openDocument` | 打开或复用指定路径的已有 `.docx` 文档 |
| `writeContent` | 写入结构化正文、标题及文字和段落格式 |
| `inspectDocument` | 读取内容、格式、结构、页面设置和文档状态 |
| `findContent` | 在指定范围内查找文字 |
| `replaceContent` | 按范围或有匹配数量约束的查询替换内容 |
| `insertTable` | 插入表格 |
| `insertImage` | 插入图片 |
| `setHeaderFooter` | 设置页眉页脚 |
| `setPageLayout` | 设置页面布局 |
| `insertBreak` | 插入分页符或分节符 |
| `save` | 将绑定文档保存到已有路径 |
| `exportPdf` | 导出 PDF |

支持分别设置西文字体与东亚字体，并在操作后读回验证。内容范围带有 Content Revision，避免后续操作误用文档修改前的旧位置。

`saveAs` 已有目标契约，但尚未进入正式能力集：保持同一个实际文档时，目标文件的 Document Lease 迁移机制仍待实现。因此，`save` 不能用于将新建文档首次保存到一个新路径。Excel 和 PPT 尚未接入正式 CLI。

## 环境要求

- Python 3.8 或更高版本。
- 执行文档操作需要 Windows，以及已注册 `KWPS.Application` 的 WPS Writer。
- 使用系统原生 Windows PowerShell：`%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe`。
- 无需第三方 Python 包，也无需安装常驻 Host 服务。

能力查询、Skill 构建和本地测试也可在 macOS/Linux 上运行。实际文档操作必须在 WPS 所在的 Windows 主机执行。

## 构建与使用 Word Skill

在仓库根目录运行，将 `<输出目录>` 替换为实际目标目录：

```bash
python scripts/build_word_skill.py --output "<输出目录>/wps-word"
python "<输出目录>/wps-word/scripts/word.py" --app word --index
python "<输出目录>/wps-word/scripts/word.py" --app word --resolve createDocument writeContent inspectDocument
```

构建会在指定位置创建 `wps-word/` 目录，包含：

```text
wps-word/
  SKILL.md
  agents/openai.yaml
  references/
  scripts/word.py
  runtime/
    files.sha256.json
    src/main/python/wps_skills/
    src/main/resources/wps_skills/word/windows/
```

将完整的 `wps-word` 目录复制到目标智能体支持的技能目录。部署后的 Skill 使用随包 Runtime，无需仓库源码；只复制 `SKILL.md` 无法执行文档操作。

构建过程生成文件 SHA-256 清单，并拒绝覆盖已有目标目录。再次构建时，可用 `--output <新目录>/wps-word` 指定新的输出位置。源码是唯一维护位置，构建产物无需单独修改。

`--index` 返回由正式契约生成的紧凑 Action Index；`--resolve` 一次返回所需 Actions 的完整契约。这两个命令不会启动 Session、PowerShell 或 WPS。批量解析结果为 `partial` 或 `failed` 时，先调整操作计划，再执行文档修改。

阅读 [Word Skill](src/main/resources/skills/wps-word/SKILL.md) 了解完整工作流程。[会话使用说明](src/main/resources/skills/wps-word/references/session.md) 提供可运行的 Python 示例，使用 `open_session()` 和 `client.call(address, params)` 逐条执行操作，并在每次响应后决定下一步。

源码中的 `src/main/resources/skills/wps-word/scripts/word.py` 也可以直接使用。更多安装说明见 [INSTALL.md](INSTALL.md)。

## 直接启动 Word Session

在安装了 WPS 的 Windows 主机上，从仓库根目录运行：

```powershell
python scripts/call.py --session --app word
```

Host 输出 `session.ready` 后等待单行 JSON 请求。例如：

```json
{"address":{"app":"word","action":"createDocument"},"params":{}}
```

每次收到完整 Action Response 后才能发送下一个请求。一个 Session 始终绑定同一个实际文档并复用一条 bridge，不通过当前活动窗口或 UI 焦点选择后续操作目标。结束时发送：

```json
{"control":"close"}
```

日常任务优先使用 Skill 提供的 Python Session Client。它处理请求顺序、超时、终止通知和真实错误，避免调用方重复实现通信逻辑。收到 `session.closing` 后应停止提交新请求，但继续读取后续 Action Response 和最终关闭记录；操作结果不确定时不得自动重放。

需要在远程桌面观察 WPS 时，应在用户已登录的桌面会话中启动。仅通过 SSH 执行命令不能保证窗口出现在远程桌面；远程启动方式取决于执行环境，Skill 不内置主机地址、账号或计划任务。

## 文档保存与清理

普通 Session 结束时释放自动化资源，保留文档打开，不会隐式保存。修改已有文件通常需要显式执行 `save`；用户要求保持未保存状态时除外。新建且没有输出路径的文档可以保留未保存状态，但必须在结果中明确说明。

仅对本次创建的可丢弃测试文档，可显式启用测试清理：

```powershell
python scripts/call.py --session --app word --debug-close-created-document
```

该选项只丢弃并关闭当前 Session 创建的文档，不关闭通过 `openDocument` 获取的文档，也不是 Word Action。需要保留内容的用户任务不应使用它。

Session 使用 Windows Job Object 管理自有桥接进程，使用 Guard、Document Lease 和 Document Quarantine 协调跨进程文档访问。正常启动会显示 WPS 文档窗口，隐藏桥接控制台；新建的 WPS 应用窗口使用普通窗口状态。对于异常小的窗口，会调整到合理的可见范围。

## 日志与耗时

Action Response 的 `traceLog` 指向该次操作的日志，其中记录 `elapsedMs`。Session 日志末尾记录：

- `sessionElapsedMs`：会话总耗时。
- `actionExecutionElapsedMs`：各 Action 执行耗时之和。
- `cleanupElapsedMs`：清理耗时。
- `actionCount`：执行的 Action 数量。

这些是诊断字段，不会扩展 Protocol v1 的 Action Response。进程墙钟时间、调用往返时间与 Action 执行时间应分别理解。Session 清理成功也不代表用户的文档任务已经完成，仍需检查编辑、验证和保存结果。

## 运行测试

在 macOS/Linux 的仓库根目录运行：

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

Windows PowerShell：

```powershell
$env:PYTHONPATH = "src/main/python"
python -m unittest discover -s src/test/python -p "test_*.py"
```

本地测试覆盖 fake 实现、真实子进程通信以及迁移到独立目录后的 Skill 产物，不需要安装 WPS 或连接外部账号。

## 项目结构

项目采用 Java 风格的 source set，同时保留标准 Python 包：

```text
src/
  main/
    python/wps_skills/
      cli/          # 命令行入口的实际实现和程序组装
      client/       # 调用方的 Session Client
      core/         # 应用无关的契约和 Action Session 机制
      host/         # JSONL Session Host
      word/         # Word 契约、handlers、Adapter 和 Skill 组装
      windows/      # Windows 进程管理、文档协调和桥接实现
    resources/
      wps_skills/word/windows/  # PowerShell bridge 和 Word 操作实现
      skills/wps-word/         # Skill 源文件、参考文档和薄入口
  test/
    python/tests/   # 与生产模块对应的测试
    resources/      # 测试资源与能力证据
scripts/            # 仓库级薄入口
```

测试使用独立的 `tests` 命名空间，避免在测试发现时用另一个顶层 `wps_skills` 包遮蔽生产实现。

## Skill 与能力参考

- [Word Skill](src/main/resources/skills/wps-word/SKILL.md)：任务工作流程与能力发现。
- [会话使用说明](src/main/resources/skills/wps-word/references/session.md)：Python 客户端用法与可运行示例。
- [内容与格式](src/main/resources/skills/wps-word/references/content.md)：内容范围、revision、文字格式和单位。
- [验证与保存](src/main/resources/skills/wps-word/references/verification.md)：结果验证、持久化和失败处置。
- [Word 契约实现](src/main/python/wps_skills/word/contracts.py)：Action 的权威定义与正式能力集。
- [WPS Writer Type Library 快照](src/test/resources/wps_skills/word/type_library/wps_writer_api.py)：能力证据，不被 Runtime 导入，也不决定正式可用 Actions。

旧版组合式 Skill、全局 Manifest、多应用 Runtime、控制器、Linux/OpenXML 后端和旧验证脚本已完成切换。历史实现可从 Git 历史查看，当前系统不提供旧接口兼容层。
