# PDF 路线子任务记录

- 子代理：pdf_route_review；只读评估，主代理记录。
- 已核对：现有 Python 3.8 / Tkinter / onedir 约束；官方包版本元数据；
  ReportLab JPEG 原始数据嵌入实现；冻结入口原先只做主窗口自检。
- 结论：ReportLab 3.6.13 + Pillow 9.5.0，新增真实 PDF 冻结自检。
- 明确边界：wheel 与 PE 依赖信息不等于真实 Win7 运行通过；未把未证实的
  “Pillow 10+ 停止 Win7 支持”写成事实。
- 用户要求控制验证数量后结束研究，不继续扩展审查。
- 可追溯决定及来源：`docs/adr/0004-automatic-image-pdf.md`。
