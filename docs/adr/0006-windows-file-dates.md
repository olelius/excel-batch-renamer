# ADR 0006：更新 Windows 创建和修改日期

日期：2026-10-09。状态：采用。

## 需求

新增第五个独立标签页；输入年月日，递归更新所选文件夹及全部子文件夹中的
JPG、PDF、Excel 文件创建及修改日期，两者各自原时分秒不变。客户机仍为离线 Win7 SP1 x64。

## 候选路线

- `os.utime`：无新增依赖，但只能修改访问/修改时间，不能实现 Windows 创建时间（并且本任务须保留访问时间），排除。
- PowerShell 或 pywin32：可以实现，但前者引入子进程和命令/路径编码边界，后者引入额外
  二进制依赖与 Win7 打包成本，对本任务无收益。
- 标准库 `ctypes` + Win32：无额外依赖，支持 Win7，能精确指定创建和修改时间，采用。

## 决定及边界

通过 `CreateFileW` 仅申请属性读写权限、`GetFileTime` 读取原时间、`SetFileTime`
传入创建和修改时间，访问时间传 NULL。不读取或改写文件内容。
时间经 SYSTEMTIME 与系统本地时区互转，只替换年月日；以整数 FILETIME 保留全部
100 纳秒刻度的小数秒。不直接用当前 UTC 偏移加减，以免跨日期的时区转换错误。
目标本地时间转回后校验年月日时分秒，不能表示的时间预检拒绝。

本次默认支持 `.jpg/.jpeg/.pdf/.xls/.xlsx/.xlsm/.xlsb`，大小写不敏感；不修改文件夹、
其他类型及文档内部属性。不跟随符号链接或目录联接，避免越界或递归循环。
扫描及逐文件日期/属性权限预检全部完成后才写入；执行中重新读取原创建和修改时间，
首次失败立即停止、显示已修改/未变化数量、不回滚。文件内容被打开不必然阻止
属性写入，实际属性权限及文件系统返回结果为准。

小数秒精度受目标文件系统约束（例如 FAT），不声称所有磁盘均具有 NTFS 精度。
真实 Win7 仍须目标机实测。

## 验证

自动化检查闰年/无效日期、递归类型筛选、预检零写入、写入失败停止统计、UI 参数传递。
当前 Windows 实际文件验证跨年/午夜、原时分秒和小数秒、访问时间不变、原内容、
目录创建时间不变及重复执行。便携自检实际更新临时 JPG、PDF 和嵌套 Excel 后缀文件。

## 官方依据

- [SetFileTime](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-setfiletime)：
  分别指定创建和修改时间，访问时间指针为 NULL 时不修改访问时间；支持 Windows XP 及以上。
- [SystemTimeToTzSpecificLocalTime](https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-systemtimetotzspecificlocaltime)
  和 [TzSpecificLocalTimeToSystemTime](https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-tzspecificlocaltimetosystemtime)：
  系统本地时间转换及夏令时边界。
