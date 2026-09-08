# v0.3.0 自动 PDF 集成交付记录

## 完成范围

现有单文件夹和批量图片任务在 JPG 改名后自动按文件题名合并 PDF，不加页面、
开关或确认步骤。PDF 输出原图片目录，保留 JPG；重跑更新当前同名 PDF，
不清理旧题名 PDF。实现默认与技术取舍见 ADR 0004。

主代理完成 PDF 适配器、依赖固定、冻结运行自检、集成与交付；路线、服务/UI、
文档子任务记录分别位于本目录 `auto-pdf-route`、`auto-pdf-integration`、
`auto-pdf-docs` 对应的同日记录。所有修改属于 `agent/image-pdf-export` 单一任务。

## 实际验证

- 一次完整 `unittest discover -s tests -v`：93 项通过，耗时 0.763 秒；
  UI 故障用例主动注入错误并确认状态提示，输出异常日志是用例预期，不是测试失败。
- `compileall -q src tests`、`pip check`、`git diff --check` 通过。
- 已在用户要求减少验证前完成 PDF 适配器 10 个基础用例；随后只运行上述统一回归，
  不另做多轮独立审查。
- 合成示例 Excel 的 `验收文件 / 6-10` 经真实服务生成五页 PDF；
  Poppler 渲染首页核对中文内容和版面正常。演示文件留在忽略的
  `output/pdf/auto-pdf-example/001——演示/验收文件.pdf`，不提交样本二进制。
- 既有 `build-portable.ps1` 双 onedir 构建通过，日志位于
  `.artifacts/build-pdf-v0.3.0.log`。
- `verify-portable.ps1` 两个版本通过，均实际生成临时 JPEG 和两页 PDF：
  测试版 994 文件 / 29,711,784 字节，正式版 994 文件 / 29,713,308 字节。
- `package-release.ps1`、`verify-release-archives.ps1` 均通过；
  解压后两个 EXE 的真实 PDF 自检通过。

## 本地交付物

| 文件 | 字节数 | SHA-256 |
| --- | ---: | --- |
| `.artifacts/release/ExcelBatchRenamer-win7-x64.zip` | 14,301,747 | `141aafb566c64c9dfd841b0acc04f256246418613861d2e89e0e5a8d24d9d68a` |
| `.artifacts/release/ExcelBatchRenamer-Test-win7-x64.zip` | 14,309,346 | `2c6374565e20918f8c0531c66e8f59ee4681c1f3e9dd28f3fa245492477a7338` |

正式版目录：`.artifacts/portable/ExcelBatchRenamer`；
测试版目录：`.artifacts/portable/ExcelBatchRenamer-Test`。
旧 v0.2.1 本地发布 ZIP 与校验文件已复制保留到 `.artifacts/archive/v0.2.1`，
需要退回应用版本时使用该旧包；此项不撤销用户数据操作。

新增依赖使用官方 wheel 保存于 `.tools/pdf-wheelhouse` 后离线安装到项目环境。
旧 pip 代理握手失败后改用开发工具下载，缺少的 typing_extensions 已补齐；
客户机不联网、不安装任何开发环境。

## 验证边界与 Git 收尾

实际开发机验证不代表离线 Win7 SP1 x64 已通过，目标机仍由用户复制完整目录实测。
按仓库授权进行本任务提交、PR squash 合并、分支清理和 v0.3.0 版本发布；
以 GitHub PR/Release 和最终主分支核验结果为收尾依据，不绕过任何必需检查。
