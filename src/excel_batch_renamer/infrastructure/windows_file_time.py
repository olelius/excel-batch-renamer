"""通过 Win7 支持的 Win32 API 读写文件时间，不读取或重写文件内容。"""

import ctypes
import os
from contextlib import contextmanager
from ctypes import wintypes
from datetime import date
from pathlib import Path


class SystemTime(ctypes.Structure):
    """Win32 SYSTEMTIME，字段顺序与系统 ABI 一致。"""

    _fields_ = [(name, wintypes.WORD) for name in (
        "year", "month", "weekday", "day", "hour", "minute", "second", "milliseconds"
    )]


def filetime_ticks(value: wintypes.FILETIME) -> int:
    """把 FILETIME 合成为自 1601 年 UTC 起的 100 纳秒刻度。"""

    return (value.dwHighDateTime << 32) | value.dwLowDateTime


def make_filetime(ticks: int) -> wintypes.FILETIME:
    """从整数刻度构造 FILETIME，避免浮点数丢失小数秒精度。"""

    return wintypes.FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)


class WindowsFileTime:
    """封装文件句柄、文件时间以及系统本地时区日期转换。"""

    def __init__(self) -> None:
        if os.name != "nt":
            raise OSError("修改创建日期仅支持 Windows 系统")
        self.api = ctypes.WinDLL("kernel32", use_last_error=True)
        filetime_pointer = ctypes.POINTER(wintypes.FILETIME)
        systemtime_pointer = ctypes.POINTER(SystemTime)
        signatures = {
            "CreateFileW": (wintypes.HANDLE, [
                wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
            ]),
            "CloseHandle": (wintypes.BOOL, [wintypes.HANDLE]),
            "GetFileTime": (wintypes.BOOL, [wintypes.HANDLE] + [filetime_pointer] * 3),
            "SetFileTime": (wintypes.BOOL, [wintypes.HANDLE] + [filetime_pointer] * 3),
            "FileTimeToSystemTime": (wintypes.BOOL, [filetime_pointer, systemtime_pointer]),
            "SystemTimeToFileTime": (wintypes.BOOL, [systemtime_pointer, filetime_pointer]),
            "SystemTimeToTzSpecificLocalTime": (wintypes.BOOL, [
                ctypes.c_void_p, systemtime_pointer, systemtime_pointer,
            ]),
            "TzSpecificLocalTimeToSystemTime": (wintypes.BOOL, [
                ctypes.c_void_p, systemtime_pointer, systemtime_pointer,
            ]),
        }
        for name, (result, arguments) in signatures.items():
            function = getattr(self.api, name)
            function.restype = result
            function.argtypes = arguments

    @staticmethod
    def _check(success) -> None:
        """保留 Win32 错误码及系统错误信息。"""

        if not success:
            raise ctypes.WinError(ctypes.get_last_error())

    @contextmanager
    def open_file(self, path: Path):
        """仅申请属性读写权限；支持中文、UNC 及长路径，始终关闭句柄。"""

        name = os.path.abspath(str(path))
        if not name.startswith("\\\\?\\"):
            name = "\\\\?\\UNC\\" + name[2:] if name.startswith("\\\\") else "\\\\?\\" + name
        # FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES；允许其他程序共享读写和删除。
        handle = self.api.CreateFileW(name, 0x180, 7, None, 3, 0x00200000, None)
        if handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            yield handle
        finally:
            self._check(self.api.CloseHandle(handle))

    def read_times(self, handle):
        """返回创建、访问、修改时间的原始整数刻度，不打开文件内容。"""

        values = [wintypes.FILETIME() for _ in range(3)]
        self._check(self.api.GetFileTime(handle, *(ctypes.byref(value) for value in values)))
        return tuple(filetime_ticks(value) for value in values)

    def local_time(self, ticks: int) -> SystemTime:
        """按当前系统时区将 UTC 文件时间转换为本地年月日时分秒。"""

        value = make_filetime(ticks)
        universal, local = SystemTime(), SystemTime()
        self._check(self.api.FileTimeToSystemTime(ctypes.byref(value), ctypes.byref(universal)))
        self._check(self.api.SystemTimeToTzSpecificLocalTime(
            None, ctypes.byref(universal), ctypes.byref(local)
        ))
        return local

    def replace_date(self, ticks: int, target_date: date) -> int:
        """只替换本地年月日，保留时分秒和全部小数秒；拒绝不可表示的时间。"""

        local = self.local_time(ticks)
        if (local.year, local.month, local.day) == (
            target_date.year, target_date.month, target_date.day
        ):
            return ticks
        local.year, local.month, local.day = target_date.year, target_date.month, target_date.day
        universal, value = SystemTime(), wintypes.FILETIME()
        self._check(self.api.TzSpecificLocalTimeToSystemTime(
            None, ctypes.byref(local), ctypes.byref(universal)
        ))
        self._check(self.api.SystemTimeToFileTime(ctypes.byref(universal), ctypes.byref(value)))
        # SYSTEMTIME 只有毫秒精度，补回原始 FILETIME 中不足一毫秒的部分。
        result = filetime_ticks(value) + ticks % 10000
        checked = self.local_time(result)
        if any(getattr(checked, field) != getattr(local, field) for field in (
            "year", "month", "day", "hour", "minute", "second", "milliseconds"
        )):
            raise ValueError("目标日期的本地时间不可表示，无法保留原时分秒")
        return result

    def set_creation_time(self, handle, ticks: int) -> None:
        """仅传创建时间指针；访问及修改时间传 NULL，保持原值。"""

        value = make_filetime(ticks)
        self._check(self.api.SetFileTime(handle, ctypes.byref(value), None, None))

    def set_file_dates(self, handle, creation_ticks: int, modified_ticks: int) -> None:
        """一次写入创建/修改时间；访问时间传 NULL，保持原值。"""

        creation, modified = make_filetime(creation_ticks), make_filetime(modified_ticks)
        self._check(self.api.SetFileTime(handle, ctypes.byref(creation), None, ctypes.byref(modified)))
