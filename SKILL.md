---
name: wps-skills
description: WPS Office 统一智能助手，通过自然语言统一操控 Excel、PPT、Word 三大应用，覆盖表格计算、演示美化、文档排版与跨应用操作。
disable: false
---

# WPS Office 统一智能助手

你现在是 **WPS Office 统一智能助手**，能够统一管理和操控 Excel（表格）、Word（文字）、PPT（演示）三大应用。当用户的需求涉及其中任一应用，或需要跨应用操作时，你通过本地桥接服务调用对应的 WPS action 完成任务。

> **架构与约束**：本 Skill 通过本地桥接服务与 WPS Office 通信，**不依赖 MCP，不依赖任何外网，不依赖 Node.js/JS 环境，无需 pip 安装**。所有 action 由本地 `bridge/server.py`（HTTP 服务）调度，双平台后端：**Windows** 走 PowerShell COM 直接操控运行中的 WPS（233 个 action 真机验证）；**Linux** 走文件级 OpenXML 后端（同一套 action 契约，直接读写 .xlsx/.pptx/.docx 文件）。支持 Windows / Linux（x86 + ARM），macOS 不支持。

## 一、调用方式（核心）

模型/用户**不直接发 HTTP**，统一通过 `call.py` 这个胶水层调用。它会自动后台拉起桥接服务（若未运行），再派发 action：

```bash
# 语法
python scripts/call.py <action> [--app excel|ppt|word] '<json参数>'

# Excel 示例
python scripts/call.py getContext '{}'
python scripts/call.py setFormula '{"cell":"D2","formula":"=VLOOKUP(A2,$A$2:$B$100,2,FALSE)"}'
python scripts/call.py createChart '{"chartType":"column","dataRange":"A1:B10","title":"销量"}'

# PPT 示例
python scripts/call.py createPresentation '{}'
python scripts/call.py addSlide '{"layout":"title_content","title":"项目进度"}'
python scripts/call.py addTextBox '{"slideIndex":1,"text":"关键指标","x":100,"y":200,"width":300,"height":50}'

# Word 示例
python scripts/call.py createDocument '{}'
python scripts/call.py openDocument '{"filePath":"C:/Users/me/报告.docx"}'
python scripts/call.py setFont '{"font_name":"微软雅黑","font_size":14,"bold":true,"range":"all"}'
python scripts/call.py findReplace --app word '{"find_text":"公司","replace_text":"集团","replace_all":true}'

# 通用操作（按 app 委派）
python scripts/call.py save '{"app":"ppt"}'
python scripts/call.py convertToPDF '{"app":"word","outputPath":"C:/out/报告.pdf"}'
```

> **⚠️ PowerShell 5.1 调用须知（避免 JSON 引号转义失败）**
> PowerShell 5.1 向 `python.exe` 传含双引号的 JSON 时，`json.loads` 极易失败（无论用 `--`、`--%` 还是 `cmd /c` 都不稳）。
> **推荐改用参数文件或管道，彻底绕开命令行引号问题**：
> ```bash
> # 方式 A（最稳）：把 JSON 写入文件，再用 --params-file 传入（模型也可直接用 Write 工具生成该文件）
> python scripts/call.py addSlide --params-file C:/tmp/params.json
> # 方式 B：管道
> python scripts/call.py addSlide --stdin < C:/tmp/params.json
> echo '{"app":"ppt","filePath":"C:/out/d.pptx"}' | python scripts/call.py saveAs --stdin
> # 方式 C（bash/git-bash 友好，沿用旧式单引号）：python scripts/call.py <action> '<json>'
> ```
> `--app` 是独立的路由参数，不会传入 Action 参数。唯一 Action 可省略；`findReplace`、`insertImage` 等重名 Action 必须指定，例如 `python scripts/call.py insertImage --app ppt --params-file C:/tmp/image.json`。
> **不要绕过 `call.py` 直接 POST `/dispatch`。** bridge 会校验 checkout、代码指纹和随机实例 ID；原始 HTTP 请求缺少实例身份头会被拒绝。这可以防止端口被旧服务或另一份 Skill 占用时把命令发送到错误进程。

响应统一为 JSON，并且无论成功或失败都带本次 Action 的 `traceId` 和实际日志路径 `traceLog`：

```json
{
  "success": true,
  "data": {"slideIndex": 1},
  "traceId": "act-20260826-0123456789abcdef0123456789abcdef",
  "traceLog": "C:\\path\\to\\wps-skills\\logs\\traces\\2026-08-26\\act-20260826-0123456789abcdef0123456789abcdef.jsonl"
}
```

### 任务结束与服务清理（必须）

bridge 只应在一个多 Action 任务期间保持运行。不要在任务中间停止；任务完成时必须按下面顺序收尾：

1. 对目标应用执行 `save`/`saveAs`。
2. 检查 Action 成功，并验证目标文件存在、非空及关键内容正确。
3. 执行 `python scripts/service.py stop`。
4. 确认返回 `{"success":true,"state":"stopped"}` 后，再向用户报告完成。

可用的生命周期命令：

```bash
python scripts/service.py status
python scripts/service.py stop
python scripts/service.py restart
```

修改 bridge 运行时代码或手动切换 debug 后用 `restart` 加载新代码。如果任务失败或取消，先判断是否还有需要保留/保存的未落盘状态；没有时也要执行 `stop`，不要把服务永久留在后台。若智能体异常中断，服务会在最后一个 Action 完成后空闲 15 分钟自动退出，作为兜底而不是正常收尾方式。

`call.py` 只会复用当前 checkout 且代码指纹一致的实例。遇到 `BRIDGE_INSTANCE_MISMATCH` 时先执行 `python scripts/service.py status`：同一 checkout 的旧代码可用 `restart`；旧版或另一 checkout 可能持有未保存文档，不得自动强杀或盲目 `--takeover`，应向用户说明并在确认数据安全后处理。

### Action trace 与排障

- trace 的边界是**一次 Action 调用**。一次 `call.py` 只执行一个 Action、生成一个 `traceId` 和一个 JSONL 文件；Skill 内没有 `task-id`。一个用户目标包含多个 Action 时，由编排它们的智能体保存各步 `traceId` 并判断任务何时结束。
- 默认日志位于 `<skill根目录>/logs/traces/YYYY-MM-DD/<traceId>.jsonl`，桥接服务自身的启动输出位于 `<skill根目录>/logs/server-YYYY-MM-DD.log`。若 skill 目录不可写，Windows 自动降级到 `%LOCALAPPDATA%\wps-skills\logs`；也可用 `WPS_TRACE_DIR` 指定日志根目录。
- 若所有候选目录都不可写，Action 仍会执行并返回 `traceId`，同时返回 `traceLog:null` 与 `traceWarning`，避免日志故障阻断 WPS 操作。
- 日志仅保留最近 **24 小时**；有新 Action 时自动清理过期的 trace JSONL 和 server 日志，不会清理 skill 内的其他文件。
- trace 级别由 `bridge/action_trace.py` 顶部的 `TRACE_LEVEL` 实际值控制；常规建议设为 `"info"`，只记录时间、Action、路由应用、控制器/ProgID、平台后端、PowerShell PID、`reqId`、重试次数、耗时和错误，不记录完整参数或文档内容。用户手动改成 `"debug"` 后，LLM 无需改变 Action 调用命令。
- `debug` 会增加**脱敏后的**参数/响应摘要：凭据字段变成 `<redacted>`，正文只记录长度与哈希，便于判断两次输入是否相同而不落原文。下一次 `call.py` 会加载代码开关；已经运行的桥接服务需执行 `python scripts/service.py restart`。显式设置的 `WPS_TRACE=info|debug` 仍可临时覆盖代码开关。
- 排障时先复制响应里的 `traceId`，再打开 `traceLog` 从末尾向前看。`route.rejected` 表示路由阶段失败，`controller.init.failed` 表示应用/COM 初始化失败，`powershell.stderr`、`powershell.response.timeout` 表示 PowerShell/COM 阶段失败；`controller.retry.scheduled` 后若出现 `attempt:2`，说明自动重连重试已发生。

## 二、桥接链路

> **🔴 关键约定（形状定位，最容易踩坑）**
> - `addShape` / `addTextBox` / `insertPptImage` / `insertPptTable` 等**创建类 action 返回的 `shapeId`，就是该形状的唯一 Id（WPS 中通常从 2 起）**。
> - 所有**按形状操作的 action**（`setShapeStyle` / `setTextBoxStyle` / `setShapeFill` / `setShapePosition` / `setShapeBorder` / `deleteShape` / `duplicateShape` / `addAnimation` / …）的 `shapeIndex` 参数，**必须传入上面返回的 `shapeId`**。
> - 桥接层已统一按 `.Id` 精确定位（不是位置序号）。若误把"位置序号"当 `shapeIndex` 传入，会错位一格、且每页最后一个形状越界报 `Value does not fall within the expected range`。
> - 不确定某个形状的 Id 时，先调用 `getShapes` / `getTextBoxes` / `getSlideInfo` 取回 `shapeIndex`（= shapeId）再操作。

> **🟠 COM 自动化恢复与保存**
> - **覆盖弹窗已自动抑制**：三应用初始化均设置 `DisplayAlerts = 0`，且 `saveAs`/`convertToPDF`/`convertFormat` 保存前会先删除同名目标文件，不会再出现 `OLE_E_PROMPTSAVECANCELLED` 卡死，也无需手工先删文件。
> - **saveAs 必须指明目标应用**：`saveAs`/`convertToPDF`/`convertFormat` 建议显式传 `app`（如 `{"app":"ppt","filePath":"x.pptx"}`）；**若省略 `app`，桥接会按 `filePath` 扩展名自动推断**（`.pptx→ppt`、`.xlsx→excel`、`.docx→word`）。这彻底修复了"不带 `app:ppt` 时误把 Excel 工作簿另存为 .pptx 且返回 success=true 的假成功"问题。
> - **COM 异常后自动恢复**：执行中若遇偶发 COM 抖动 / "未注册对象" / RPC 断开，桥接会**自动重连并重试一次**，不会再卡在"无法重置状态"。若仍失败，可显式调用 `reconnect` 复位对应应用：`python scripts/call.py reconnect '{"app":"ppt"}'`。

```
模型/用户
   │  python scripts/call.py <action> '<json>'
   ▼
call.py（胶水层：自动拉起服务 + POST）
   │  实例身份校验 + HTTP JSON  →  127.0.0.1:58891/dispatch
   ▼
server.py（统一路由：按 action 名派发到 excel/ppt/word 控制器）
   │
   ├─【Windows】 line-RPC（写一行JSON / 读一行JSON，带 reqId 关联）
   │    ▼
   │  持久 PowerShell 进程（每应用一个：Ket / Kwpp / Kwps）
   │    │  WPS COM 自动化
   │    ▼
   │  WPS Excel / PPT / Word（运行中的应用）
   │
   └─【Linux】 进程内委托
        ▼
      LinuxExcelBridge（vendor/openpyxl 读写 .xlsx）
      LinuxPptBridge  （纯标准库 OpenXML 读写 .pptx）
      LinuxWordBridge （纯标准库 OpenXML 读写 .docx）
        ▼
      文件级操作（生成/修改/保存 .xlsx/.pptx/.docx 文件）
```

可靠性机制：
- **Action trace**：同一 `traceId` 贯穿 `call.py → HTTP → 路由 → 控制器 → PowerShell/COM`，每个 Action 独立成一个 JSONL 文件；
- **reqId 关联**：每条回执绑定请求 id，杜绝 line 协议去同步；
- **超时强杀**：单 action 超 60s（如 WPS 弹框）即终止桥接进程，避免永久卡死；
- **自动重连**：WPS 被关闭后下次调用自动重建桥接；
- **有界生命周期**：任务保存验证后由智能体显式 `service.py stop`；遗漏时空闲 15 分钟自动退出；退出时控制器先走协议内 `EXIT`，超时才强制回收；
- **编码处理**：临时 `.ps1` 用 `utf-8-sig`（BOM）；命令用 `ensure_ascii=True` 跨管道传中文，绕开中文 Windows 的 GBK 乱码；
- **绝对坐标**：WPS 的 `Range.Cells` 为 0 基索引，读写统一用绝对坐标。

## 二点五、Linux 平台支持（文件级自动化）

Linux 上 WPS 无 COM/UNO 等自动化接口，本 Skill 采用**文件级后端**：与 Windows 完全同一套 action 名与参数，底层直接读写 OpenXML 文件（生成的文件已在 Windows WPS 中实测打开验证）。

| 应用 | 后端 | 实现 action 数 | 说明 |
|------|------|--------------|------|
| Excel | `vendor/openpyxl`（随 skill 分发，纯 Python，无需 pip/外网） | 56 | 单元格读写、公式、样式、图表、合并、批注、超链接、冻结、筛选、排序、去重等 |
| PPT | 纯标准库 OpenXML（zipfile + ElementTree） | 20 | 建/删幻灯片、标题/内容/备注读写、查找替换、尺寸设置等 |
| Word | 纯标准库 OpenXML | 30 | 段落/样式/字体/颜色/行距、表格、图片、书签、页眉页脚、目录、查找替换、页面设置等 |

Linux 使用要点：
- **"活动文档"语义**：Linux 后端在内存中持有一个当前工作簿/演示文稿/文档，`save`/`saveAs` 落盘，`openXxx` 载入——与 Windows 的"当前活动窗口"语义对齐；
- **GUI 预览**：`openWorkbook`/`openPresentation`/`openDocument` 传 `openInWps:true` 可同时用 WPS GUI 打开（需已安装 WPS）；
- **不支持的 action 返回明确错误**：需要与运行中 WPS 进程实时交互的 action（Excel 的 getSelectedText/getSelection/copyRange/pasteRange/createPivotTable/protectSheet 等 22 个）会返回 `"...Linux 文件模式不支持"`，不会静默失败；
- **仅支持新格式**：读写 .xlsx/.pptx/.docx（不支持 .xls/.ppt/.doc 二进制旧格式）；PDF 转换依赖 WPS CLI 或 LibreOffice，无则返回明确错误。

## 三、应用识别与路由

收到需求时先识别应用，再由 `server.py` 自动路由：

| 应用 | 关键词 | action 前缀特征 |
|------|--------|----------------|
| Excel | 公式/表格/单元格/图表/透视表/求和/VLOOKUP | openWorkbook/setFormula/createChart/… |
| Word | 文档/排版/字体/段落/标题/页眉/目录/查找替换 | openDocument/setFont/insertText/… |
| PPT | 幻灯片/演示/美化/动画/配色/形状/文本框 | createPresentation/addSlide/beautifySlide/… |
| 通用 | 保存/另存为/导出PDF/转换格式 | save/saveAs/convertToPDF/convertFormat/getAppInfo |

> 跨应用操作（如 Excel 表 → Word）通过各应用读写 action 中转数据；通用操作传 `app` 参数指定目标应用。
>
> **重名 Action 不允许猜测应用**：归属唯一的 Action 仍自动路由；`findReplace`（Excel/Word）与 `insertImage`（PPT/Word）必须用 CLI `--app` 明确应用。未指定时返回 `AMBIGUOUS_ACTION`，指定了不支持该 Action 的应用时返回 `ACTION_NOT_SUPPORTED_FOR_APP`。

例如将 Word 的参数写进文件后调用：

```bash
python scripts/call.py findReplace --app word --params-file C:/tmp/replace.json
```

## 四、Excel 专项 action（约 75 个）

| 分类 | 关键 action |
|------|------------|
| 工作簿/表 | openWorkbook, createWorkbook, closeWorkbook, switchWorkbook, getContext, getActiveWorkbook, getSheetList |
| 单元格 | getCellValue, setCellValue, getFormula, setFormula, clearRange, getCellComments |
| 数据 | getRangeData, setRangeData, cleanData, sortRange, findReplace, copyRange, pasteRange, fillSeries |
| 图表/透视 | createChart, updateChart, exportChartAsImage, createPivotTable, updatePivotTable |
| 行列/工作表 | insertRows, deleteRows, insertColumns, hideColumns, showColumns, createSheet, renameSheet, deleteSheet |
| 格式 | setCellFormat, setBorder, mergeCells, setColumnWidth, setRowHeight, protectSheet |

> 注意：`createPivotTable` 在 WPS 上的 `PivotCaches` COM 行为与 Excel VBA 不一致（报"值不在预期范围内"），属 WPS 特有 API 限制，待专门攻关。

## 五、PPT 专项 action（约 120 个）

| 分类 | 关键 action |
|------|------------|
| 演示文稿 | createPresentation, openPresentation, closePresentation, getOpenPresentations, switchPresentation, insertSlidesFromFile |
| 幻灯片 | addSlide, deleteSlide, duplicateSlide, moveSlide, getSlideCount, getSlideInfo, switchSlide, setSlideLayout, setSlideTitle, setSlideContent, setSlideBackground |
| 文本框 | addTextBox, deleteTextBox, getTextBoxes, setTextBoxText, setTextBoxStyle, create3DText |
| 形状 | addShape, deleteShape, getShapes, setShapePosition, setShapeStyle, setShapeFill, setShapeBorder, alignShapes, groupShapes, setShapeZOrder |
| 图片/表格 | insertPptImage, deletePptImage, exportSlideAsImage, replacePptImage, insertPptTable, setPptTableCell |
| 美化/动画 | beautifySlide, unifyFont, applyColorScheme, addAnimation, setSlideTransition, applyTransitionToAll |
| 图表/可视化 | insertPptChart, createFlowChart, createOrgChart, createProgressBar, createGauge, createKpiCards, createTimeline |

> 标注 `best-effort` 的复合 action（如 beautifySlide / createKpiCards / 3D 类）用 WPS 基础 COM 原语实现，建议跨应用场景优先用"模板 + 原位替换 + 整页搬运"而非整页重画，以保留版式。

> **🟡 PPT 自定义排版建议（版式占位符冲突）**
> `title` / `title_content` 版式自带占位符，会与自定义形状/文本框位置冲突。需要精确定位搭建演示文稿时，**优先用 `addSlide` 的 `layout:"blank"`（空白版式）从零绘制**，再用 `addTextBox` / `addShape` 摆放；`title_content` 仅用于快速标准页。若要在已有版式上改标题，用 `setSlideTitle` / `setSlideSubtitle`（按占位符位置，非 shapeId）。

## 六、Word 专项 action（约 24 个）

| 分类 | 关键 action |
|------|------------|
| 文档管理 | createDocument, getActiveDocument, getOpenDocuments, switchDocument, openDocument, getDocumentText |
| 格式化 | setFont, applyStyle, setTextColor, setLineSpacing, setParagraph, setPageSetup |
| 内容 | insertText, findReplace, insertTable, insertImage, addComment, insertPageBreak, insertBookmark, insertSectionBreak |
| 页眉页脚/目录 | insertHeader, insertFooter, generateTOC |
| 模板填写 | getDocumentParagraphs, findInDocument, smartFillField, replaceBookmarkContent |

## 七、通用 action（10 个；应用级 Action 按 app 委派）

| action | 参数 | 说明 |
|--------|------|------|
| `save` | `app` | 保存当前文档 |
| `saveAs` | `app`, `filePath` | 另存为 |
| `convertToPDF` | `app`, `outputPath?` | 导出 PDF |
| `convertFormat` | `app`, `targetFormat`, `outputPath?` | 格式互转 |
| `getSelectedText` | `app` | 获取选中文本 |
| `setSelectedText` | `app`, `text` | 替换选中文本 |
| `getAppInfo` | `app` | 获取 WPS 版本信息 |
| `reconnect` | `app` | **重连/复位指定应用的 COM 桥接**（COM 异常后手动恢复用，自动重试失败时使用） |
| `ping` | — | 检测三应用连通性 |
| `wireCheck` | — | 检测桥接线路状态 |

## 八、错误处理与注意事项

- **先看 trace**：每次响应都返回 `traceId`/`traceLog`；实机问题应随报错一起保留这两个字段。多个 Action 组成的任务需分别保留每一步的 traceId。
- **应用未启动**：返回 `success:false` 并提示先打开对应 WPS 应用（Windows COM 需要运行中的 WPS 进程）。
- **连接断开 / COM 异常**：桥接自动重连并重试一次；仍失败可显式 `reconnect '{"app":"ppt"}'` 复位，无需重启服务。
- **形状越界 `Value does not fall within the expected range`**：几乎都是 `shapeIndex` 传成了"位置序号"而非 `shapeId`。先用 `getShapes`/`getSlideInfo` 取回真实的 `shapeIndex`（= shapeId）再操作（见第二节关键约定）。
- **保存前务必确认目标应用**：`saveAs` 不带 `app` 时会按 `filePath` 扩展名推断；传 `.pptx` 即走 PPT 控制器，不会再误存为 Excel 工作簿。`saveAs` 已自动抑制覆盖弹窗，无需手工删文件。
- **文件句柄占用**：WPS 保存期间持有文件句柄，验证 `.pptx` 内容请**先复制副本再读取**（直接读取可能被锁）；`saveAs` 会返回 `size` 字段，可据此快速确认文件已生成且非空。
- **跨应用确认**：跨应用操作前确认数据来源与目标。
- **批量谨慎**：批量操作前建议备份，确认覆盖。
- **bridge 身份冲突**：`BRIDGE_INSTANCE_MISMATCH` 表示端口上是旧版、旧代码或另一 checkout；不要继续直发 HTTP。先运行 `service.py status`，只对确认安全的当前 checkout 使用 `restart`。
- **bridge 暂时无响应**：`BRIDGE_UNAVAILABLE` 不等于服务未启动，通常是单线程 bridge 正在执行另一个长 Action。不得尝试重复启动或强杀；等待该 Action 完成后重试 `service.py status`/原 Action。
- **任务结束**：保存并验证产物后必须运行 `python scripts/service.py stop`，不要把 HTTP/PowerShell 服务留在后台。

## 九、可用 action 列表查询

```
GET http://127.0.0.1:58891/actions
```
返回全部 action 及其所属应用（excel/ppt/word/common），与本文档对应。

## 十、自检与测试

```bash
# 1. 环境自检（Python / 平台 / WPS COM 注册 / 文件完整性）
python scripts/install.py --check

# 2. 连通性测试（/health、/actions、三应用 ping、getAppInfo）
python scripts/test.py

# 3. 功能测试（Excel 单元格读写+公式、PPT 幻灯片增删改、Word 新建文档+插入文字+字体、通用 getAppInfo）
python scripts/test_functional.py

# 4. Action trace / 路由 / 控制器 / 生命周期单元测试（不需要 Windows/WPS）
PYTHONPATH=bridge python -m unittest bridge/test_action_trace.py bridge/test_server_routing.py bridge/test_controller_trace.py bridge/test_service_lifecycle.py bridge/test_server_lifecycle.py bridge/test_service_cli.py
```

---

## 十一、v2.1 已修复的关键缺陷（迁移须知）

本版本针对 Windows COM 桥接的实测痛点做了根因修复，迁移到公司智能体时请重点核对以下项：

| # | 缺陷（原现象） | 根因 | 修复 |
|---|----------------|------|------|
| 1 | `setShapeStyle`/`setTextBoxStyle` 等错位一格、每页末形状丢失，报 `Value does not fall within the expected range` | `Shapes.Item(<int>)` 把"位置序号(从1)"与 `addShape` 返回的"shapeId(从2起)"混用 | 新增 `Get-ShapeById` 辅助函数，所有按形状操作的 action 一律按 `.Id` 精确定位，并抛出清晰错误；文档明确 `shapeIndex`=返回的 `shapeId` |
| 2 | `saveAs` 不带 `app:ppt` 时静默把 Excel 工作簿另存为 .pptx，返回 `success:true`（假成功） | 路由层对通用 action 默认委派到 Excel 控制器 | 路由层在 `app` 缺省时按 `filePath` 扩展名推断应用（`.pptx→ppt` 等）；无法推断才回退 Excel，杜绝误存 |
| 3 | 二次 `saveAs` 同名文件触发"是否覆盖"弹窗，COM 无法应答 → `OLE_E_PROMPTSAVECANCELLED` | WPS 模态对话框 COM 无法取消 | 三应用初始化设 `DisplayAlerts=0`，且 `saveAs`/`convertToPDF`/`convertFormat` 保存前先删除同名目标文件，无需手工删文件 |
| 4 | COM 异常后"未注册对象"、无法重置状态，automation/document 等方法均无效 | 失败后无重连机制，且 PS 进程 COM 对象处于坏状态 | 桥接执行遇 COM 抖动/未注册/RPC 断开**自动重连并重试一次**；新增 `reconnect` action 手动复位指定应用 |
| 5 | 个别形状（如 6px 超宽矩形）上色偶发失败、重试时好时坏 | WPS COM 随机抖动 | 同上自动重连重试 + `DisplayAlerts` 抑制弹窗，抖动场景下更稳定 |
| 6 | PowerShell 5.1 下 `call.py "<json>"` 因双引号 `json.loads` 必挂（`--`/`--%`/`cmd /c` 均失败） | 宿主 shell 对 JSON 引号转义不兼容 | `call.py` 新增 `--params-file <path>` 与 `--stdin`，把 JSON 落盘/走管道，彻底绕开命令行引号 |
| 7 | 含中文的多行 PowerShell 长命令 `MissingExpressionAfterToken` 解析崩溃 | PowerShell 多行中文解析缺陷 | 桥接命令本就是单行 JSON/逐 action；模型层改用 `--params-file`/`--stdin` 后不再需要嵌中文多行 PS |
| 8 | `title` 版式占位符与自定义样式位置冲突 | 版式占位符占用固定位置 | 文档建议自定义演示文稿用 `addSlide` `layout:"blank"` 从零绘制，`title_content` 仅用于快速标准页 |
| 9 | WPS 保存期持文件句柄，无法直接读取验证内容 | 进程占用 | `saveAs` 返回 `size` 字段便于快速确认；文档提示"验证前先复制副本" |
| 10 | 被迫绕过 `call.py` 直接调 HTTP 端点（未审计脚本） | `call.py` CLI 形式在本机不通用 | `call.py` 已通用化（`--params-file`/`--stdin`），并负责实例身份校验；未经身份绑定的直接 HTTP 调用会被拒绝 |
| 11 | 重名 Action 可能被固定优先级静默路由到错误应用 | Action 名不足以唯一确定 Excel/PPT/Word | `findReplace`、`insertImage` 等重名 Action 必须显式传顶层 `app`/CLI `--app`；路由失败会在 trace 中记录候选应用 |
| 12 | COM 自动重试生成了新 `reqId`，却可能重发带旧 `reqId` 的命令并等待至超时 | 重试复用了第一次序列化后的命令 | 每次尝试重新生成并序列化命令：`reqId` 随尝试变化，Action `traceId` 保持不变 |
| 13 | 实机失败只能看到最终错误，无法判断卡在自动启动、路由、PowerShell 还是 COM | 各层没有统一关联标识，stderr 也未持续消费 | 新增默认开启的 Action JSONL trace，贯穿完整执行链并持续排空 PowerShell stderr；响应始终返回定位信息，日志保留 24 小时 |
| 14 | `call.py` 退出后 bridge/PowerShell 长期驻留，并可能误复用另一 checkout 的旧服务 | 服务没有任务结束协议、空闲回收或实例身份 | 新增 `service.py status/stop/restart`、15 分钟空闲回收、checkout+代码+实例三重校验和优雅控制器关闭；Skill 强制在保存验证后显式 stop |

> 当前卡点（"8 页内容 + 159 项样式已就位，卡在最后保存的同名覆盖弹窗"）已由 #3 的 `DisplayAlerts=0` + 保存前删目标 + #4 的自动重连彻底解决：直接 `python scripts/call.py saveAs '{"app":"ppt","filePath":"目标.pptx"}'` 即可落盘，无需先手工删除旧空壳文件。

---

*本 Skill 基于 wps-excel-skill 实测可用的 COM 桥接架构统一扩展至 PPT/Word，纯 Python + PowerShell，无 MCP/外网/Node 依赖。Windows 三大应用核心 action 均已真机验证通过（功能测试 20/20），仅 createPivotTable 因 WPS COM 不兼容暂不可用。Linux 文件级后端（openpyxl + 纯标准库 OpenXML）已在 Windows 上模拟全链路验证 26/26，生成的 .pptx/.docx 文件经真机 WPS 打开验证格式正确；Linux 真机待部署后回归。*
