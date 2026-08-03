# 系统架构

状态：已确认

## 架构风格

项目采用轻量模块化单体。所有功能运行在一个本地桌面进程中，不引入数据库、网络服务、插件系统、依赖注入框架或自动回滚系统。

依赖方向固定为：

```text
Tkinter UI -> 应用服务 -> 纯业务规则
                     -> 基础设施适配
```

UI 不直接解析 Excel，也不直接操作文件系统。业务规则不依赖 Tkinter、openpyxl 或 Windows UI。

## 组件

### UI

- 创建文件夹标签页；
- 重命名文件夹标签页；
- 重命名图片标签页；
- 批量重命名图片标签页；
- 每个标签页维护自己的目录和输入状态；
- 显示完成数量、失败对象和错误原因。

### 应用服务

- 组织一次任务的读取、计划和执行；
- 文件夹创建与重命名服务；
- 图片重命名服务；
- 多工作表与多文件夹批量图片重命名服务；
- 失败时停止后续操作，不复制备份、不自动回滚。

### 纯业务规则

- 三位序号格式化；
- 标准文件夹名和占位文件夹名生成；
- 文件夹序号匹配；
- 页次区间推导；
- 图片目标名称生成；
- 工作表与文件夹绑定判断。

### 基础设施适配

- 使用 openpyxl 读取 `.xlsx`；
- 使用 Python 标准库执行本地文件系统操作；
- 隔离具体路径、工作表和文件重命名 API。

## 目录结构

```text
ExcelBatchRenamer/
├─ src/excel_batch_renamer/
│  ├─ app.py
│  ├─ ui/
│  │  ├─ main_window.py
│  │  ├─ create_folders_tab.py
│  │  ├─ rename_folders_tab.py
│  │  ├─ rename_images_tab.py
│  │  └─ batch_rename_images_tab.py
│  ├─ create_folders.py
│  ├─ rename_folders.py
│  ├─ rename_images.py
│  ├─ batch_rename_images.py
│  ├─ core/
│  │  ├─ naming.py
│  │  └─ page_ranges.py
│  └─ infrastructure/
│     └─ xlsx_reader.py
├─ tests/
├─ packaging/
├─ docs/
├─ requirements.txt
└─ pyproject.toml
```

## 环境与交付

- 开发、测试和打包依赖全部安装在项目 `.venv`。
- 源码使用 Python 3.8，UI 使用 Tkinter/ttk，Excel 读取使用 openpyxl。
- PyInstaller 4.10 `onedir` 生成测试版和正式版完整便携目录。
- 测试版显示 CMD 控制台；正式版隐藏控制台，只显示 UI。
- 交付目录自包含 Python、Tk/Tcl、openpyxl 及间接依赖。
- 客户机完全离线，不安装 .NET Framework、Java、Python、Visual C++ 运行库安装包、Excel/WPS 或其他运行环境。

## 明确不做

- 不联网；
- 不依赖 Office COM；
- 不复制开发虚拟环境作为交付物；
- 不使用单文件打包；
- 不建设数据库或配置服务；
- 不清洗规范外输入；
- 不复制备份或自动回滚。
