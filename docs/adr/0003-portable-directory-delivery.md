---
status: accepted
---

# 使用完整便携目录交付

应用使用项目虚拟环境中的 PyInstaller 4.10 `onedir` 生成完整便携目录，而不是复制虚拟环境、手工维护嵌入式 Python、使用安装程序或生成单文件 EXE。用户将整个目录原样复制到 64 位 Windows 7 SP1 客户机，并双击目录中的 EXE；这种方式由打包工具统一收集 Python、Tk/Tcl 和第三方依赖，也符合流水线小工具的直接复制使用方式。

## Consequences

- 交付目录必须包含 Python 运行时、Tk/Tcl、`openpyxl` 及全部间接依赖。
- 交付目录不得依赖客户机安装 .NET Framework、Java、Python、Visual C++ 运行库安装包、Microsoft Excel、WPS 或其他应用运行时。
- 应用不得在客户机运行期间联网下载依赖或更新组件。
- 用户不得只复制 EXE，文档和目录名称必须明确提示“复制整个文件夹”。
- 构建后必须执行依赖完整性和便携启动检查。
- 使用同一代码生成两个目录构建：测试版保留 CMD 控制台，正式版隐藏控制台；除控制台配置外不得形成不同业务逻辑。
- 开发环境验证不能替代真实 Win7 验证；用户在另一台 Win7 SP1 电脑完成最终测试前，兼容性状态保持未验证。

## Delivery Layout

```text
dist/
├─ ExcelBatchRenamer-Win7-Test/
│  ├─ ExcelBatchRenamer-Test.exe
│  ├─ 使用说明.txt
│  └─ <PyInstaller 自动生成的 Python、Tk/Tcl、DLL、PYD 和库文件>
└─ ExcelBatchRenamer-Win7/
   ├─ ExcelBatchRenamer.exe
   ├─ 使用说明.txt
   └─ <PyInstaller 自动生成的 Python、Tk/Tcl、DLL、PYD 和库文件>
```

PyInstaller 4.10 的实际运行文件通常直接位于 EXE 同级目录，例如 `python38.dll`、`base_library.zip`、`_tkinter.pyd`、`tcl86t.dll`、`tk86t.dll`、`tcl/`、`tk/` 以及其他依赖文件。具体文件清单以构建产物为准，不允许客户删除、移动或只复制其中一部分。
