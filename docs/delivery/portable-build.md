# 便携目录构建与验证

状态：开发侧已完成，真实 Windows 7 测试待用户执行

## 构建基线

- 64 位 Python 3.8.10；
- Tk 8.6；
- openpyxl 3.1.5；
- PyInstaller 4.10；
- `onedir` 完整目录，不使用 `onefile`；
- 测试版使用控制台启动器，正式版使用无控制台启动器。

PyInstaller 4.10 文档说明其正式运行基线为 Windows 8 或更高版本，Windows 7 “应可运行但不受支持”。因此本项目只能把 Windows 7 SP1 x64 作为目标兼容环境，不能在用户完成真实目标机测试前声明已经兼容。

## UCRT 本地部署

Python 3.8 的 `python38.dll` 依赖 Universal C Runtime。仅由 PyInstaller 生成的目录包含 `VCRUNTIME140.dll`，但不保证未经更新的 Windows 7 已安装 UCRT。

为满足“客户机不安装 Visual C++ 运行库”的交付要求，本项目按微软 UCRT 本地部署说明，将 Windows SDK 的 x64 UCRT 全套 DLL 放在主 EXE 同目录。包含：

- `ucrtbase.dll`；
- 全套 `api-ms-win-core-*.dll`；
- 全套 `api-ms-win-crt-*.dll`。

微软参考：

- https://learn.microsoft.com/en-us/cpp/windows/universal-crt-deployment
- https://pyinstaller.org/en/v4.10/requirements.html
- https://pyinstaller.org/en/v4.10/usage.html

## 首次准备

以下脚本仅在开发机首次准备项目本地 UCRT 时需要联网。它下载经过 Authenticode 校验的微软官方 Windows SDK 安装器，只创建项目内离线布局并以 MSI 管理提取方式解包 UCRT，不把 SDK 安装到客户机：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\packaging\prepare-ucrt.ps1
```

生成内容位于 `.tools/windows-sdk/`，由 `.gitignore` 排除。

## 构建与验证

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\packaging\build-portable.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\packaging\verify-portable.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\packaging\package-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\packaging\verify-release-archives.ps1
```

生成目录：

```text
.artifacts/portable/
├─ ExcelBatchRenamer-Test/
│  ├─ ExcelBatchRenamer-Test.exe
│  ├─ README.txt
│  └─ Python、Tk/Tcl、openpyxl、VC/UCRT 等运行文件
└─ ExcelBatchRenamer/
   ├─ ExcelBatchRenamer.exe
   ├─ README.txt
   └─ Python、Tk/Tcl、openpyxl、VC/UCRT 等运行文件
```

`verify-portable.ps1` 会：

1. 核对两个目录中的 EXE、Python、Tkinter、Tk/Tcl 数据和 VC 运行文件；
2. 逐个核对项目本地 x64 UCRT DLL；
3. 启动控制台测试版自检；
4. 启动无控制台正式版自检；
5. 输出目录文件数和总字节数。

`package-release.ps1` 生成正式版、测试版 ZIP 和 `SHA256SUMS.txt`；`verify-release-archives.ps1` 把两个 ZIP 解压到项目 `.artifacts` 下的唯一临时目录，分别运行自检后再清理该目录。

## 验证边界

当前构建机自检可以确认：

- 打包入口可以启动；
- Tkinter 主窗口可以创建；
- openpyxl 和四个业务服务已经被打入程序；
- 正式版与测试版使用同一业务代码；
- 正式版使用无控制台启动器；
- 两个目录包含声明的项目本地运行依赖。

当前构建机不是 Windows 7，不能证明 Win7 系统 API、显卡、权限策略和补丁状态下的真实行为。最终验证必须把整个目录复制到离线 Windows 7 SP1 x64 电脑执行，不能只复制 EXE。
