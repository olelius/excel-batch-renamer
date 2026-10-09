"""创建/修改日期任务的预检、递归边界、失败统计及真实 Windows 属性测试。"""

import ctypes
import os
import subprocess
import tempfile
import unittest
from contextlib import contextmanager
from ctypes import wintypes
from datetime import date
from pathlib import Path
from unittest.mock import patch

from excel_batch_renamer.infrastructure.windows_file_time import (
    SystemTime, WindowsFileTime, filetime_ticks, make_filetime,
)
from excel_batch_renamer.update_file_dates import (
    FileDateExecutionError, collect_date_files,
    parse_file_date, update_file_dates,
)


class FileDateTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def test_valid_date_and_leap_year(self):
        self.assertEqual(parse_file_date(" 2024-02-29 "), date(2024, 2, 29))

    def test_invalid_dates_are_rejected_before_scanning(self):
        for text in ("", "2023-02-29", "2026-04-31", "2026-13-01", "2026-1-02",
                     "2026/10/09", "1500-01-01", "2026-10-09 12:00:00"):
            with self.subTest(text=text), patch(
                "excel_batch_renamer.update_file_dates.collect_date_files"
            ) as scan, self.assertRaises(ValueError):
                update_file_dates(self.root, text)
            scan.assert_not_called()

    def test_recursive_case_insensitive_types_and_ignored_files(self):
        nested = self.root / "档案" / "子文件夹"
        nested.mkdir(parents=True)
        expected = []
        for index, suffix in enumerate((".JPG", ".jpeg", ".PDF", ".xls", ".xlsx", ".xlsm", ".xlsb")):
            path = (self.root if index == 0 else nested) / ("文件" + suffix)
            path.write_bytes(b"content")
            expected.append(path)
        (self.root / "ignore.txt").write_bytes(b"ignore")
        (nested / "ignore.png").write_bytes(b"ignore")
        (self.root / "目录.xlsx").mkdir()
        self.assertEqual(set(collect_date_files(self.root)), set(expected))

    def test_invalid_directory(self):
        for path in (self.root / "missing", self.root / "file.txt"):
            if path.name == "file.txt":
                path.write_bytes(b"x")
            with self.assertRaisesRegex(ValueError, "不存在或不是文件夹"):
                collect_date_files(path)

    def test_scan_permission_error_is_not_silently_ignored(self):
        def failing_walk(*args, **kwargs):
            kwargs["onerror"](PermissionError("无法扫描子目录"))
        with patch("excel_batch_renamer.update_file_dates.os.walk", side_effect=failing_walk):
            with self.assertRaises(PermissionError):
                collect_date_files(self.root)

    def _fake_backend(self, fail_preflight=False, fail_write=False):
        class FakeBackend:
            def __init__(self):
                self.written = []
                self.reads = []

            @contextmanager
            def open_file(self, path):
                if fail_preflight and path.name == "b.pdf":
                    raise PermissionError("文件被占用")
                yield path

            def read_times(self, handle):
                self.reads.append(handle)
                return (10, 20, 30)

            def replace_date(self, ticks, target):
                return 11

            def set_file_dates(self, handle, creation, modified):
                if fail_write and handle.name == "b.pdf":
                    raise PermissionError("无法写入")
                self.written.append(handle)
        return FakeBackend()

    def _three_files(self):
        for name in ("a.jpg", "b.pdf", "c.xlsx"):
            (self.root / name).write_bytes(b"x")

    def test_preflight_error_does_not_write_any_file(self):
        self._three_files()
        backend = self._fake_backend(fail_preflight=True)
        with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
            with self.assertRaisesRegex(ValueError, "b.pdf.*未修改任何文件"):
                update_file_dates(self.root, "2026-10-09")
        self.assertEqual(backend.written, [])

    def test_write_failure_stops_with_completed_count_without_rollback(self):
        self._three_files()
        backend = self._fake_backend(fail_write=True)
        with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
            with self.assertRaises(FileDateExecutionError) as raised:
                update_file_dates(self.root, "2026-10-09")
        self.assertEqual(raised.exception.updated, 1)
        self.assertEqual(raised.exception.total, 3)
        self.assertEqual(raised.exception.failed_path, self.root / "b.pdf")
        self.assertEqual(backend.written, [self.root / "a.jpg"])

    def test_preflight_reads_all_files_before_writing_then_refreshes_times(self):
        self._three_files()
        backend = self._fake_backend()
        original_write = backend.set_file_dates
        def check_write(handle, creation, modified):
            self.assertEqual(set(backend.reads), set(collect_date_files(self.root)))
            original_write(handle, creation, modified)
        backend.set_file_dates = check_write
        with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
            result = update_file_dates(self.root, "2026-10-09")
        self.assertEqual((result.total, result.updated, result.unchanged), (3, 3, 0))
        self.assertEqual(len(backend.reads), 6)

    def test_only_one_date_needs_changing_still_counts_file_as_updated(self):
        (self.root / "file.xlsx").write_bytes(b"x")
        for unchanged_ticks in (10, 30):
            with self.subTest(unchanged_ticks=unchanged_ticks):
                backend = self._fake_backend()
                backend.replace_date = lambda ticks, target: ticks if ticks == unchanged_ticks else ticks + 1
                with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
                    result = update_file_dates(self.root, "2026-10-09")
                self.assertEqual((result.updated, result.unchanged), (1, 0))

    def test_modified_time_preflight_error_prevents_any_writes(self):
        self._three_files()
        backend = self._fake_backend()
        def invalid_modified(ticks, target):
            if ticks == 30:
                raise ValueError("修改时间转换失败")
            return ticks + 1
        backend.replace_date = invalid_modified
        with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
            with self.assertRaisesRegex(ValueError, "未修改任何文件.*修改时间转换失败"):
                update_file_dates(self.root, "2026-10-09")
        self.assertEqual(backend.written, [])

    def test_no_matching_files_reports_zero(self):
        (self.root / "ignore.txt").write_bytes(b"x")
        backend = self._fake_backend()
        with patch("excel_batch_renamer.update_file_dates.WindowsFileTime", return_value=backend):
            result = update_file_dates(self.root, "2026-10-09")
        self.assertEqual(result.total, 0)
        self.assertIn("已修改 0 个", result.status_text)
        self.assertEqual(backend.written, [])


@unittest.skipUnless(os.name == "nt", "真实创建时间属性测试仅支持 Windows")
class WindowsFileDateTests(unittest.TestCase):
    setUp = FileDateTests.setUp
    _three_files = FileDateTests._three_files

    def _seed(self, path, year=2024, month=12, day=31, hour=23, minute=59, second=58,
              milliseconds=321, submillisecond=9876, modified=False):
        backend = WindowsFileTime()
        local = SystemTime(year, month, 0, day, hour, minute, second, milliseconds)
        universal, value = SystemTime(), wintypes.FILETIME()
        backend._check(backend.api.TzSpecificLocalTimeToSystemTime(
            None, ctypes.byref(local), ctypes.byref(universal)
        ))
        backend._check(backend.api.SystemTimeToFileTime(ctypes.byref(universal), ctypes.byref(value)))
        ticks = filetime_ticks(value) + submillisecond
        with backend.open_file(path) as handle:
            if modified:
                modified_value = make_filetime(ticks)
                backend._check(backend.api.SetFileTime(handle, None, None, ctypes.byref(modified_value)))
            else:
                backend.set_creation_time(handle, ticks)
        return ticks

    def _times(self, path):
        backend = WindowsFileTime()
        with backend.open_file(path) as handle:
            return backend.read_times(handle)

    def test_real_recursive_update_preserves_hms_fraction_content_and_other_times(self):
        nested = self.root / "中文子目录"
        nested.mkdir()
        paths = [self.root / "图像.JPG", nested / "文件.pdf", nested / "任务.xlsx"]
        originals = {}
        for path in paths:
            path.write_bytes(b"do not rewrite file bytes")
            self._seed(path)
            self._seed(path, year=2025, month=6, day=15, hour=1, minute=2, second=3,
                       milliseconds=654, submillisecond=1234, modified=True)
            originals[path] = self._times(path)
        excluded = nested / "ignore.txt"
        excluded.write_bytes(b"untouched")
        excluded_before = self._times(excluded)
        folder_before = nested.stat().st_ctime_ns
        result = update_file_dates(self.root, "2024-02-29")
        self.assertEqual((result.total, result.updated, result.unchanged), (3, 3, 0))
        backend = WindowsFileTime()
        for path in paths:
            after = self._times(path)
            for index, expected_hms in ((0, (23, 59, 58)), (2, (1, 2, 3))):
                local = backend.local_time(after[index])
                self.assertEqual((local.year, local.month, local.day), (2024, 2, 29))
                self.assertEqual((local.hour, local.minute, local.second), expected_hms)
                self.assertEqual(after[index] % 10000000, originals[path][index] % 10000000)
            self.assertEqual(after[1], originals[path][1])
            self.assertEqual(path.read_bytes(), b"do not rewrite file bytes")
        self.assertEqual(self._times(excluded), excluded_before)
        self.assertEqual(nested.stat().st_ctime_ns, folder_before)
        repeated = update_file_dates(self.root, "2024-02-29")
        self.assertEqual((repeated.updated, repeated.unchanged), (0, 3))

    def test_cross_year_and_midnight_use_local_date_not_utc_date(self):
        path = self.root / "午夜.xls"
        path.write_bytes(b"x")
        self._seed(path, year=2025, month=1, day=1, hour=0, minute=2, second=3)
        update_file_dates(self.root, "2026-12-31")
        local = WindowsFileTime().local_time(self._times(path)[0])
        self.assertEqual((local.year, local.month, local.day, local.hour, local.minute, local.second),
                         (2026, 12, 31, 0, 2, 3))

    def test_content_lock_does_not_require_reading_content_and_handle_is_released(self):
        self._three_files()
        backend = WindowsFileTime()
        locked_path = self.root / "b.pdf"
        handle = backend.api.CreateFileW(str(locked_path), 0x80000000, 0, None, 3, 0, None)
        self.assertNotEqual(handle, ctypes.c_void_p(-1).value)
        try:
            # Windows 的文件共享锁约束内容访问，不一定阻止仅属性权限的句柄。
            result = update_file_dates(self.root, "2024-02-29")
            self.assertEqual(result.updated, 3)
        finally:
            backend.api.CloseHandle(handle)
        # 能删除和重新创建文件，说明没有遗留属性句柄。
        locked_path.unlink()
        locked_path.write_bytes(b"reopened")

    def test_directory_junction_is_not_followed(self):
        with tempfile.TemporaryDirectory() as external:
            outside = Path(external) / "outside.jpg"
            outside.write_bytes(b"outside")
            link = self.root / "junction"
            process = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), external],
                                     stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.assertEqual(process.returncode, 0, process.stderr)
            try:
                self.assertEqual(collect_date_files(self.root), ())
                with self.assertRaisesRegex(ValueError, "实际文件夹"):
                    collect_date_files(link)
            finally:
                # 非递归删除目录联接本身，不触及外部文件夹。
                os.rmdir(str(link))

    def test_filetime_integer_roundtrip(self):
        ticks = 134000000003219876
        self.assertEqual(filetime_ticks(make_filetime(ticks)), ticks)
