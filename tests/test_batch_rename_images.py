import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from excel_batch_renamer.batch_rename_images import (
    BatchImageRenameExecutionError,
    batch_rename_images,
    build_batch_image_rename_plan,
)


class BatchRenameImagesTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.workbook_path = self.root / "图片任务.xlsx"

    def tearDown(self):
        self.temporary_directory.cleanup()

    def _write_workbook(self, sheets):
        workbook = Workbook()
        workbook.remove(workbook.active)
        for sheet_name, rows in sheets:
            worksheet = workbook.create_sheet(sheet_name)
            worksheet.append(["图片档案目录"])
            worksheet.append(["制表说明"])
            worksheet.append(["序号", "文件题名", "页次"])
            for index, (title, page_reference) in enumerate(rows, start=1):
                worksheet.append([index, title, page_reference])
        workbook.save(self.workbook_path)

    def _make_folder(self, name, pages):
        folder = self.root / name
        folder.mkdir()
        for page in pages:
            (folder / "{:03d}.jpg".format(page)).write_bytes(b"jpg")
        return folder

    def test_renames_all_numeric_worksheets_in_matching_folders(self):
        self._write_workbook(
            [
                ("1", [("甲", "001-002")]),
                ("说明", []),
                ("2", [("乙", "001-001")]),
            ]
        )
        first = self._make_folder("001——", [1, 2])
        second = self._make_folder("002——材料", [1])

        result = batch_rename_images(self.workbook_path, self.root)

        self.assertEqual(
            (result.folders, result.total, result.renamed, result.unchanged),
            (2, 3, 3, 0),
        )
        self.assertTrue((first / "001甲.jpg").exists())
        self.assertTrue((first / "002甲.jpg").exists())
        self.assertTrue((second / "001乙.jpg").exists())

    def test_non_numeric_worksheets_are_ignored(self):
        self._write_workbook(
            [("说明", []), ("1", [("文件", "001-001")])]
        )
        folder = self._make_folder("001——", [1])

        plans = build_batch_image_rename_plan(self.workbook_path, self.root)

        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].worksheet_name, "1")
        self.assertEqual(plans[0].folder_path, folder)

    def test_rejects_workbook_without_numeric_worksheet(self):
        self._write_workbook([("说明", [])])

        with self.assertRaisesRegex(ValueError, "没有可处理的数字工作表"):
            build_batch_image_rename_plan(self.workbook_path, self.root)

    def test_duplicate_numeric_worksheet_sequence_is_rejected(self):
        self._write_workbook(
            [
                ("1", [("甲", "001-001")]),
                ("01", [("乙", "001-001")]),
            ]
        )
        folder = self._make_folder("001——", [1])

        with self.assertRaisesRegex(ValueError, "对应同一文件夹序号"):
            batch_rename_images(self.workbook_path, self.root)

        self.assertTrue((folder / "001.jpg").exists())

    def test_missing_folder_blocks_entire_batch_before_rename(self):
        self._write_workbook(
            [
                ("1", [("甲", "001-001")]),
                ("2", [("乙", "001-001")]),
            ]
        )
        first = self._make_folder("001——", [1])

        with self.assertRaisesRegex(ValueError, "工作表 2 未找到"):
            batch_rename_images(self.workbook_path, self.root)

        self.assertTrue((first / "001.jpg").exists())
        self.assertFalse((first / "001甲.jpg").exists())

    def test_duplicate_matching_folders_block_entire_batch(self):
        self._write_workbook([("1", [("甲", "001-001")])])
        first = self._make_folder("001——甲", [1])
        second = self._make_folder("001——乙", [1])

        with self.assertRaisesRegex(ValueError, "匹配到多个直属文件夹"):
            batch_rename_images(self.workbook_path, self.root)

        self.assertTrue((first / "001.jpg").exists())
        self.assertTrue((second / "001.jpg").exists())

    def test_later_folder_validation_failure_blocks_earlier_folder(self):
        self._write_workbook(
            [
                ("1", [("甲", "001-001")]),
                ("2", [("乙", "001-002")]),
            ]
        )
        first = self._make_folder("001——", [1])
        self._make_folder("002——", [1])

        with self.assertRaisesRegex(ValueError, "缺少页码.*002"):
            batch_rename_images(self.workbook_path, self.root)

        self.assertTrue((first / "001.jpg").exists())

    def test_repeat_execution_reports_unchanged_images(self):
        self._write_workbook([("1", [("甲", "001-002")])])
        self._make_folder("001——", [1, 2])
        batch_rename_images(self.workbook_path, self.root)

        result = batch_rename_images(self.workbook_path, self.root)

        self.assertEqual((result.renamed, result.unchanged), (0, 2))

    def test_filesystem_failure_stops_later_folders_without_rollback(self):
        self._write_workbook(
            [
                ("1", [("甲", "001-001")]),
                ("2", [("乙", "001-001")]),
            ]
        )
        first = self._make_folder("001——", [1])
        second = self._make_folder("002——", [1])
        original_rename = Path.rename
        calls = []

        def fail_on_third_rename(path, target):
            calls.append((path, target))
            if len(calls) == 3:
                raise PermissionError("文件被占用")
            return original_rename(path, target)

        with patch.object(Path, "rename", new=fail_on_third_rename):
            with self.assertRaises(BatchImageRenameExecutionError) as captured:
                batch_rename_images(self.workbook_path, self.root)

        error = captured.exception
        self.assertEqual(error.worksheet_name, "2")
        self.assertEqual(error.result.folders, 1)
        self.assertEqual(error.result.renamed, 1)
        self.assertTrue((first / "001甲.jpg").exists())
        self.assertTrue((second / "001.jpg").exists())


if __name__ == "__main__":
    unittest.main()
