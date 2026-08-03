# Excel Batch Renamer

面向离线 Windows 7 流水线的 Excel 批量文件整理工具。

## 功能

- 根据固定 `.xlsx` 模板创建 `001——` 或 `001——文件夹名称`；
- 根据后续整理完成的 `.xlsx` 补全或更新文件夹名称；
- 按图片任务表的单个工作表和页次范围重命名一个文件夹中的直属 JPG；
- 遍历所有数字工作表，自动匹配父目录下前三位序号相同的多个文件夹并批量重命名 JPG；
- 同一窗口提供四个相互独立的 Tkinter 标签页。

## 开发验证

```powershell
$env:PYTHONPATH = "$PWD\src"
& "$PWD\.venv\Scripts\python.exe" -m unittest discover -s tests -v
```

## 便携构建

项目使用 Python 3.8.10、PyInstaller 4.10 `onedir`，分别生成带控制台测试版和无控制台正式版。构建流程与兼容性边界见 [`docs/delivery/portable-build.md`](docs/delivery/portable-build.md)。

项目协作规则、Git 分支流程和自动收尾授权见 [`AGENTS.md`](AGENTS.md)。
