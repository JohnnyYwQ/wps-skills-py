---
name: wps-skills-no-server
description: 使用 Catalog 与单 Action CLI 操作 WPS Excel、PPT 和 Word 的本地 Skill。
disable: false
---

# WPS Skills

当用户请求操作 Excel、PPT 或 Word 时，先查询 Action Catalog，再逐个执行 Action。你是编排者：你决定 WPS 任务边界、选择活动文档、串行等待每一步结果并根据结果决定后续 Action。

## 调用流程

1. 搜索或读取精确 Contract，绝不从本说明猜测参数：

   ```bash
   python scripts/actions.py search chart
   python scripts/actions.py describe setCellValue --app excel
   ```

2. 按 Contract 执行一个 Action：

   ```bash
   python scripts/call.py setCellValue --app excel '{"row": 1, "col": 1, "value": 42}'
   ```

3. 等待 JSON 响应，保存 `traceId` 和 `traceLog`，然后才安排下一步。

PowerShell 5.1 用 `--params-file` 或 `--stdin` 传 JSON，避免命令行引号转义。唯一归属的 Action 可省略 `--app`；重名 Action 必须显式指定 `excel`、`ppt` 或 `word`。

Catalog 只读取 Manifest，不会初始化 WPS。每次 `call.py` 调用创建一个 Action Runtime，执行一个 Action，并在返回前释放本次 controller 和子进程；不会关闭 WPS 应用、活动文档或未保存内容。

## 文档与风险

普通 Action 只操作相应 WPS 应用中的活动文档。需要改变目标时，先使用 Contract 提供的创建、打开、列举或切换 Action。

先读取 Contract 的 `risk` 与 `prerequisites`：

- `read` Action 可以在可识别的短暂 COM/RPC 故障后自动重试一次。
- `write` 和 `destructive` Action 的不确定结果会返回 `outcomeUnknown:true`；先执行合适的只读 Action 核验活动文档，不要盲目重放。
- `saveAs`、`convertToPDF` 和 `convertFormat` 的既有目标默认安全拒绝；仅当用户明确要求替换时才传 `overwrite:true`。

多个 Action 组成的 WPS 任务必须串行提交并等待响应。系统 mutex 只防止不同进程误并发，不能表达业务顺序。

## 环境与验证

运行 `python scripts/install.py --check` 只报告 Python、平台、Windows PowerShell、WPS COM 注册和仓内资源，不会改动环境。

本规格范围的自动化验证不要求 Windows 或 WPS：

```bash
python scripts/validate_action_manifest.py
python -m unittest discover -s bridge -p 'test_*.py'
```

Linux 后端及现有仓内 Linux 依赖不属于本轮的设计或验收范围。
