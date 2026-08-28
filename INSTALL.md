# WPS Skill 环境说明

此 Skill 没有仓外运行时依赖，也没有安装步骤。它面向能够加载 Agent Skills 并执行本地 shell 的编排者；不要求特定宿主、私有 API、全局工具注册或直接导入 Python 模块。

## 前置条件

- Python 3.8 或更高版本。
- Windows 实机运行：Windows PowerShell 与已安装、已注册 COM 的 WPS Office。所需 ProgID 为 `Ket.Application`、`Kwpp.Application` 与 `Kwps.Application`。
- Linux 后端及其仓内源码依赖保持原样，但不属于本轮实机验收范围。

环境入口只报告状态，不会修改系统、下载内容或创建运行环境：

```bash
python scripts/install.py --check
```

## 加载后工作流

1. 编排者加载根目录 `SKILL.md`。
2. 在 shell 中使用 `scripts/actions.py` 搜索和读取 Action Contract。
3. 使用 `scripts/call.py` 执行一个 Action，等待 JSON 响应后再提交下一步。
4. 记录每次响应的 `traceId`；发生不确定的写入结果时，先用只读 Action 检查活动文档。

示例：

```bash
python scripts/actions.py describe findReplace --app word
python scripts/call.py findReplace --app word --params-file C:\tmp\replace.json
```

每个 CLI 调用都有独立 Runtime 与清理边界，不会关闭 WPS 应用、活动文档或未保存内容。跨进程 mutex 会等待冲突调用结束；编排者仍必须自行维持 Action 的业务顺序。

## 本地验证

以下检查不要求 Windows 或 WPS：

```bash
python scripts/validate_action_manifest.py
python -m unittest discover -s bridge -p 'test_*.py'
```

真实 WPS 的功能验收只应在目标 Windows 环境进行；Linux 后端不因本轮架构收缩而被视为已验收。
