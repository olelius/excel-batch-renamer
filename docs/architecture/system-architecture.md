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
- 两种图片任务直接集成自动 PDF，不新增标签页或执行入口；
- 显示完成数量、PDF 数量和保存位置、失败步骤、对象与原因。

### 应用服务

- 组织一次任务的读取、计划和执行；
- 文件夹创建与重命名服务；
- 图片重命名服务，复用本次页码/题名计划在改名后按题名生成 PDF；
- 双目录全局编排服务，先按档号形成文件目录在前、图纸目录在后的映射，再统一改工作表、文件夹、JPG 并生成 PDF；
- 失败时停止后续操作，不复制备份、不自动回滚。

### 纯业务规则

- 三位序号格式化；
- 标准文件夹名和占位文件夹名生成；
- 文件夹序号匹配，兼容标准数字前缀与来源编码的末段数字；
- 页次区间推导；
- 图片目标名称生成；
- 工作表与文件夹绑定判断。

### 基础设施适配

- 使用 openpyxl 读取 `.xlsx`；
- 使用 Pillow 校验 JPEG 解码、读取尺寸/DPI 和 EXIF 方向，使用 ReportLab 嵌入原 JPEG 生成 PDF；
- 使用 Python 标准库执行本地文件系统操作；
- 隔离具体路径、工作表、文件重命名和 PDF 写入 API。

### 双目录编排

`catalog_reindex.py` 负责固定顺序：读取两个目录工作簿的档号集合、按档号自然顺序
生成全局新序号、写出映射表、按档号重命名工作表和文件夹、把每个文件夹内 JPG
从 001 重新起号，最后复用批量图片服务完成题名改名和 PDF。工作表标签的原位置
不参与映射。

## 图片任务与 PDF 执行边界

应用服务的图片计划保留页码、源路径、目标路径和本次文件题名；PDF 分组复用该计划，不从已生成 PDF 或图片旧题名反推业务数据。同一工作表内完全相同题名的所有图片按页码升序生成一个 PDF，不跨文件夹合并。

在首张图片改名前，完成绑定、页码覆盖、JPEG 解码、PDF 同名目录及题名大小写歧义校验；批量服务先构建全部文件夹计划。每个文件夹先改名，再生成 PDF，错误立即停止并汇总已完成的改名、未变化和 PDF 数量。

PDF 适配器每张图片生成一页，遵循 EXIF 旋转/镜像与图片 DPI，无有效 DPI 时使用 72；不强制 A4，也不重压缩 JPEG。每个 PDF 先写入同目录临时文件，关闭后再替换同名目标，生成失败保留旧 PDF。该文件级写入保护不提供整批回滚；重跑更新当前同名 PDF，但不清理旧题名 PDF。

路线取舍见 [ADR 0004：自动图片 PDF](../adr/0004-automatic-image-pdf.md)。

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
│  │  ├─ models.py
│  │  └─ page_ranges.py
│  └─ infrastructure/
│     ├─ xlsx_reader.py
│     └─ pdf_writer.py
├─ tests/
├─ packaging/
├─ docs/
├─ requirements.txt
└─ pyproject.toml
```

## 环境与交付

- 开发、测试和打包依赖全部安装在项目 `.venv`。
- 源码使用 Python 3.8，UI 使用 Tkinter/ttk，Excel 读取使用 openpyxl；PDF 运行依赖固定为 ReportLab 3.6.13 和 Pillow 9.5.0。
- PyInstaller 4.10 `onedir` 生成测试版和正式版完整便携目录。
- 测试版显示 CMD 控制台；正式版隐藏控制台，只显示 UI。
- 交付目录自包含 Python、Tk/Tcl、openpyxl、ReportLab、Pillow 及间接依赖。
- pypdf 4.3.1 只用于构建侧自动化 PDF 检查，不是应用运行依赖。
- 客户机完全离线，不安装 .NET Framework、Java、Python、Visual C++ 运行库安装包、Excel/WPS 或其他运行环境。

## 明确不做

- 不联网；
- 不依赖 Office COM；
- 不复制开发虚拟环境作为交付物；
- 不使用单文件打包；
- 不建设数据库或配置服务；
- 不清洗规范外输入；
- 不复制备份或自动回滚。
