# Word 内容与格式

使用前解析相应 Action 的完整契约。下面解释选择和定位原则，不重复完整参数 schema。

## 阅读、查找与定位

`inspectDocument` 返回一个 revision 一致的正文快照、段落/runs、结构、页面和文档状态。请求必须明确 `scope` 和三个读取上限。结果 `truncated` 为真时，`remainingRange` 指向未返回的正文；仍在同一 revision 时可将它作为下一次 `scope` 的 range。

`findContent` 适用于定位已有文字。使用返回的 Content Range，不要从归一化后的 `text` 自行推算 WPS 位置。

Content Range 是主文档正文中从 0 开始、左闭右开的 UTF-16 code-unit 范围，并带有生成它的 Content Revision。它只在原 Action Session 内有效。表情等字符可能占两个 UTF-16 单位，段落标记又不同于归一化正文，因此 Python 字符串索引不能直接当成范围。

内容或结构修改后旧范围失效。下一步需要范围时，使用该次响应提供的新范围，或者重新检查/查找。`STALE_CONTENT_RANGE` 应通过重新读取并确认目标解决，不能只把旧坐标换成新 revision。

## 选择修改方式

- 追加或从文首插入：`writeContent`，使用 `documentEnd` 或 `documentStart` Body Anchor，无需读取光标位置。
- 在已定位内容前后插入：使用 `before`/`after` 和当前 Content Range。
- 替换已有内容：使用 `replaceContent` 的 range 或 query target。Query 指定 `expectedMatchCount`；先查找并评估目标，不能猜测数量或无界全局替换。
- 写正文：用 `paragraph` 和 `heading` blocks，runs 表达一段中的不同文字格式。标题使用语义级别，不传本地化的样式名称。
- 表格、图片与分隔符：分别使用 `insertTable`、`insertImage`、`insertBreak`，不要藏进正文字符串。图片文件路径必须存在于 Windows 执行主机。
- 页眉页脚和页面布局：先取得最新 Content Revision，再构造 section selector；section indexes 从 0 开始。

一个 `writeContent` 可承载一组自然段，按契约上限批量写入，避免每个字或每个 run 各执行一次 Action。仍要在每个 Action 响应后作下一步决定。

## 字体与单位

`fontFamily` 为统一字体；`westernFontFamily` 与 `eastAsiaFontFamily` 分别指定西文与东亚文字字体。同一个 run 不能同时使用统一字体字段和任一分文字类别字段。

只显式设置用户需要的格式。当前 WPS bridge 对请求字体名做严格读回校验，尚不归一化中英文别名。例如已验证的中文 WPS 环境将 `Microsoft YaHei` 读回为 `微软雅黑`，会导致 `CONTENT_VERIFICATION_FAILED`；该环境可用 `微软雅黑`、`宋体` 配合 `Times New Roman`、`Arial`。不要假定这些名称或字体在所有主机上都可用，也不要在失败后自动重复写入。

字号、段落间距与缩进中带 `Pt` 后缀的字段使用磅。页面边距和图片尺寸等空间长度使用契约定义的 `{value, unit}`，单位为 `pt`、`mm`、`cm`、`in`；不要传像素或 COM 原始单位。

如果任务涉及当前契约未提供的能力，例如另存为 DOCX、修订跟踪或批注，不要因为 WPS UI/COM 存在类似功能就声称 Skill 支持。
