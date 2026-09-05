# 验证、持久化与失败处置

## 验证用户要求

每个写 Action 已执行自己的操作级读回校验，任务结束前仍需针对用户要求做验证。用 `inspectDocument` 检查实际文字、段落/标题、runs 格式、表格/图片数量、页面设置等；只验证与本任务有关的事实。

有截断时读取剩余范围，或改用定位明确的小范围检查。若期间发生用户修改，重新检查，不能拼接不同 revision 的结果。截图或界面可用于用户要求的视觉检查，但不能替代文档内容与保存状态的验证。

## 保存已有文件

已有文件通常按以下顺序处理，每一步成功并评估后才进入下一步：

1. `openDocument`：检查实际返回的文档状态。若已有未保存修改，而本次任务需要保存，先说明会一起保存并确认授权。
2. 读取相关内容，执行请求的修改。
3. `inspectDocument` 验证本次任务要求。
4. `save`，检查成功结果中的 artifact 路径、格式、大小，以及文档状态。必要时再检查保存后的文档状态。
5. 正常结束 Session，报告保存结果。

读取失败、修改失败或结果不确定时，不能继续执行预定的保存步骤。新建未保存文档没有已有 locator，不能通过 `save` 实现首次保存到新路径；`saveAs` 尚未进入正式 Contract Set。

## PDF 导出

`exportPdf` 必须提供用户授权的 Windows 输出绝对路径和显式 `overwritePolicy`。默认可选择 `failIfExists`；只有已有明确覆盖授权时才使用契约允许的覆盖选项。检查返回的 PDF artifact；导出成功不意味着 DOCX 已保存。

## 如何处理不同结果

| 结果 | 下一步 |
| --- | --- |
| `succeeded` | 检查返回数据满足任务要求，再决定下一步；不等于整个任务已经完成 |
| `failed` | 查看 `error.code/message`；只有确定解决了原因且 Session 仍可用时，才显式调整请求 |
| `unknown` | 操作可能已有部分效果，禁止自动重放；优先只读验证 |
| 无完整响应、超时或 EOF | 不凭进程退出码推断文档未变化；检查 Client 的 `may_have_effect`、最近响应和 traces，保留不确定性 |
| `session.closing` | 停止提交，读取后续 Action Response 与 `session.closed`；使用真实 Action 错误诊断 |
| 清理失败 | 分别报告已验证的文档结果与未确认的资源清理，不重做已经完成的文档修改 |

已终止的 Session 不能继续使用或重新绑定。只有用户原来指定的同一文件仍可定位、Lease/Quarantine 检查允许时，才考虑新 Session 的只读核验。未保存且无法精确定位的文档需要用户协助；不得新建一份当作自动恢复。

## 日志与耗时

Action Response 的 `traceLog` 指向该次 Action 日志；`client.ready["traceLog"]` 指向 Session 日志。默认 Windows 日志在 `%LOCALAPPDATA%\wps-skills\logs`，可在任务启动前使用 `WPS_TRACE_DIR` 指定本次任务的诊断目录。

Action trace 的 `elapsedMs` 是 Host 执行耗时；Session trace 末尾记录 `sessionElapsedMs`、`actionExecutionElapsedMs`、`cleanupElapsedMs`。Python 调用的往返耗时和整个进程墙钟时间是不同指标。只有用户要求性能信息或正在诊断延迟时，才展开这些数据。

最终报告使用具体事实，例如“标题及三行表格已读回验证，已保存到实际返回的路径，文档保持打开”。没有保存时明确说明“已编辑，仍未保存”。不要用 `Session Outcome: succeeded` 替代任务完成证明。
