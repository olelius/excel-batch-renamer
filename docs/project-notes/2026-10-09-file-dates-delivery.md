# v0.5.0 文件日期更新交付

日期：2026-10-09。任务分支：`agent/file-creation-date`。

## 需求及实现

新增第五个独立标签页“更新文件日期”。用户已明确包含全部子文件夹，并补充修改日期
也必须更新，因此创建/修改日期同时更新为输入年月日，两者各自原时分秒及小数秒保留。
访问时间、文件内容、文件名和文档内部属性不变。

默认处理 JPG/JPEG、PDF、xls/xlsx/xlsm/xlsb，大小写不敏感；不修改目录和其他文件，
不跟随符号链接/目录联接。输入 `YYYY-MM-DD`；预检全部完成后执行，失败停止、不回滚。
仅当两项日期均已一致时才计为未变化。

采用标准库 `ctypes` 调用 Win7 支持的文件属性 API，不新增依赖。详见
[ADR 0006](../adr/0006-windows-file-dates.md)。没有更改原四项业务规则。

## 已验证

- `unittest discover -s tests -v`：115 项通过，包含 16 项文件日期测试及 3 项新增 UI 参数/状态检查。
- `compileall -q src tests`、`pip check` 通过。
- Windows 真实属性测试验证创建 `23:59:58`、修改 `01:02:03` 各自保持不变；
  日期统一替换、访问时间、100 纳秒小数部分、文件内容和目录创建时间不变。
- 覆盖递归类型匹配、闰年/非法日期、跨年/午夜、仅一项日期需更新、修改时间预检失败、
  预检零写入、写入失败统计、重跑未变化和目录联接不跟随。
- 默认窗口下五个标签及新页布局已进行实际截图检查，无控件/结果文字遮挡；
  本地截图为 `.artifacts/file-dates-ui.png`（不纳入版本管理，截图中的数量仅为布局检查文本）。
- 双 `onedir` 构建与便携运行通过；两目录各 994 个文件，控制台版 29,733,442 字节，
  正式版 29,734,978 字节。
- 发布 ZIP 重新解压后的测试版与正式版自检均通过；冻结自检实际生成 JPEG/PDF，
  再递归更新其创建/修改日期和嵌套 Excel 后缀文件日期，不只测试导入。

## 交付及未验证范围

- 正式便携目录：`.artifacts/portable/ExcelBatchRenamer`，必须复制整个目录。
- 正式 ZIP：`.artifacts/release/ExcelBatchRenamer-win7-x64.zip`。
- 测试 ZIP：`.artifacts/release/ExcelBatchRenamer-Test-win7-x64.zip`。
- 校验和：`.artifacts/release/SHA256SUMS.txt`。

当前构建机不是 Win7。真实离线 Windows 7 SP1 x64 兼容性、客户机文件系统时间精度及
权限策略仍须用户目标机验收。具体场景见 [验收场景第 11 项](../requirements/acceptance-scenarios.md)。
