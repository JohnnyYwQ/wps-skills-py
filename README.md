# WPS Skills

WPS Skills 让任何能加载 Agent Skills 并执行本地 shell 的编排者，通过一个本地 Action Runtime 操作 WPS Excel、PPT 和 Word。编排者拥有 WPS 任务：先查询 Catalog，逐个执行 Action，读取每步结果后再决定下一步。

## 使用方式

运行环境只需要已有的 Python；Windows 实机运行还需要 Windows PowerShell 和已注册 COM 的 WPS Office。仓内已随附 Linux Excel 后端所需源码；本轮没有改动、重新设计或验收 Linux 后端。

先做只读环境报告：

```bash
python scripts/install.py --check
```

每次调用前，先从 Catalog 读取精确的 Action Contract：

```bash
python scripts/actions.py search chart
python scripts/actions.py describe setCellValue --app excel
python scripts/call.py setCellValue --app excel '{"row": 1, "col": 1, "value": 42}'
```

PowerShell 5.1 推荐传入 JSON 文件，避免命令行转义：

```powershell
python scripts/call.py addSlide --app ppt --params-file C:\tmp\slide.json
```

`call.py` 一次只执行一个 Action，并在输出响应前清理本次 Runtime 和 controller。Catalog 仅读取 Manifest，不会初始化 WPS。普通 Action 面向对应应用的活动文档；需要改变目标时，先使用显式的创建、打开、列举或切换 Action。

## 安全与可靠性

- 多 Action 的 WPS 任务由编排者串行提交，并等待每一步 JSON 响应；系统 mutex 只防止不同进程误并发。
- 先检查 Contract 的 `risk`：只有 `read` Action 在可识别的短暂 COM/RPC 故障后自动重试一次。`write` 或 `destructive` 的不确定结果带 `outcomeUnknown:true`，应先用只读 Action 核验活动文档。
- `saveAs`、`convertToPDF` 和 `convertFormat` 在目标已存在时默认拒绝；只有明确的 `overwrite:true` 才允许覆盖。
- 每个响应都有 `traceId` 和 `traceLog`。排障时从该 trace 的末尾向前查看；不要重复提交不确定的写入 Action。

## 验证

无需 Windows 或 WPS 实机即可验证本规格范围：

```bash
python scripts/validate_action_manifest.py
python -m unittest discover -s bridge -p 'test_*.py'
```

自动化套件覆盖 Catalog、Manifest、Action CLI、Runtime、mutex、风险策略、三个控制器的静态 Contract、trace 与子进程清理。Windows/WPS 实机验证仍是运行环境验收，非本轮交付前提。
