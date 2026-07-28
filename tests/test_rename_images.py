import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from excel_batch_renamer.rename_images import (
    ImageRenameExecutionError,
    build_image_rename_plan,
    get_worksheet_names,
    rename_images,
    suggest_worksheet_name,
)


class RenameImagesTests(unittest.TestCase):
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
            worksheet.append(["序号", "文件题名", "页次", "其他列"])
            for index, (title, page_reference) in enumerate(rows, start=1):
                worksheet.append([index, title, page_reference, "忽略"])
        workbook.save(self.workbook_path)

    def _make_folder(self, name="001——"):
        folder = self.root / name
        folder.mkdir()
        return folder

    def _make_images(self, folder, pages, suffix=".jpg"):
        for page in pages:
            (folder / "{:03d}{}".format(page, suffix)).write_bytes(b"jpg")

    def test_enumerates_worksheets_and_suggests_matching_name(self):
        self._write_workbook([("1", [("文件A", "001-001")]), ("2", [])])
        folder = self._make_folder()

        names = get_worksheet_names(self.workbook_path)

        self.assertEqual(names, ["1", "2"])
        self.assertEqual(suggest_worksheet_name(folder, names), "1")
        self.assertIsNone(suggest_worksheet_name(self.root / "003——", names))

    def test_rejects_binding_mismatch_before_modifying_images(self):
        self._write_workbook(
            [
                ("1", [("文件A", "001-001")]),
                ("2", [("文件B", "001-001")]),
            ]
        )
        folder = self._make_folder()
        self._make_images(folder, [1])

        with self.assertRaisesRegex(ValueError, "序号不匹配"):
            rename_images(self.workbook_path, "2", folder)

        self.assertTrue((folder / "001.jpg").exists())

    def test_renames_acceptance_range_and_returns_statistics(self):
        self._write_workbook(
            [
                (
                    "1",
                    [
                        ("文件A", "001"),
                        ("文件B", "005"),
                        ("文件C", "010-017"),
                    ],
                )
            ]
        )
        folder = self._make_folder()
        self._make_images(folder, range(1, 18))

        result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual((result.total, result.renamed, result.unchanged), (17, 17, 0))
        self.assertTrue((folder / "001文件A.jpg").exists())
        self.assertTrue((folder / "004文件A.jpg").exists())
        self.assertTrue((folder / "005文件B.jpg").exists())
        self.assertTrue((folder / "009文件B.jpg").exists())
        self.assertTrue((folder / "010文件C.jpg").exists())
        self.assertTrue((folder / "017文件C.jpg").exists())

    def test_repeat_execution_recalculates_titles_and_unchanged_items(self):
        self._write_workbook([("1", [("旧题名", "001-002")])])
        folder = self._make_folder()
        self._make_images(folder, [1, 2])
        rename_images(self.workbook_path, "1", folder)

        first_repeat = rename_images(self.workbook_path, "1", folder)
        self.assertEqual(
            (first_repeat.renamed, first_repeat.unchanged),
            (0, 2),
        )

        self._write_workbook([("1", [("新题名", "001-002")])])
        second_repeat = rename_images(self.workbook_path, "1", folder)

        self.assertEqual(
            (second_repeat.renamed, second_repeat.unchanged),
            (2, 0),
        )
        self.assertTrue((folder / "001新题名.jpg").exists())
        self.assertTrue((folder / "002新题名.jpg").exists())

    def test_scans_jpg_case_insensitively_and_ignores_subfolders(self):
        self._write_workbook([("1", [("文件A", "001-001")])])
        folder = self._make_folder()
        self._make_images(folder, [1], suffix=".JPG")
        nested = folder / "子目录"
        nested.mkdir()
        (nested / "002.jpg").write_bytes(b"nested")

        result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual(result.renamed, 1)
        self.assertTrue((folder / "001文件A.jpg").exists())
        self.assertTrue((nested / "002.jpg").exists())

    def test_missing_page_blocks_all_renames(self):
        self._write_workbook([("1", [("文件A", "001-002")])])
        folder = self._make_folder()
        self._make_images(folder, [1])

        with self.assertRaisesRegex(ValueError, "缺少页码.*002"):
            rename_images(self.workbook_path, "1", folder)

        self.assertTrue((folder / "001.jpg").exists())

    def test_extra_page_blocks_all_renames(self):
        self._write_workbook([("1", [("文件A", "001-001")])])
        folder = self._make_folder()
        self._make_images(folder, [1, 2])

        with self.assertRaisesRegex(ValueError, "未覆盖.*002"):
            rename_images(self.workbook_path, "1", folder)

        self.assertTrue((folder / "001.jpg").exists())
        self.assertTrue((folder / "002.jpg").exists())

    def test_duplicate_page_prefix_blocks_all_renames(self):
        self._write_workbook([("1", [("文件A", "001-001")])])
        folder = self._make_folder()
        (folder / "001.jpg").write_bytes(b"first")
        (folder / "001旧题名.jpg").write_bytes(b"second")

        with self.assertRaisesRegex(ValueError, "同一页码"):
            rename_images(self.workbook_path, "1", folder)

        self.assertTrue((folder / "001.jpg").exists())
        self.assertTrue((folder / "001旧题名.jpg").exists())

    def test_existing_target_directory_blocks_all_renames(self):
        self._write_workbook([("1", [("文件A", "001-001")])])
        folder = self._make_folder()
        self._make_images(folder, [1])
        (folder / "001文件A.jpg").mkdir()

        with self.assertRaisesRegex(ValueError, "目标名称已被占用"):
            build_image_rename_plan(self.workbook_path, "1", folder)

        self.assertTrue((folder / "001.jpg").exists())

    def test_filesystem_error_stops_without_rollback(self):
        self._write_workbook([("1", [("文件A", "001-002")])])
        folder = self._make_folder()
        self._make_images(folder, [1, 2])
        original_rename = Path.rename
        calls = []

        def fail_on_third_rename(path, target):
            calls.append((path, target))
            if len(calls) == 3:
                raise PermissionError("文件被占用")
            return original_rename(path, target)

        with patch.object(Path, "rename", new=fail_on_third_rename):
            with self.assertRaises(ImageRenameExecutionError) as captured:
                rename_images(self.workbook_path, "1", folder)

        error = captured.exception
        self.assertEqual(error.failed_path.name, "002.jpg")
        self.assertEqual(error.result.renamed, 1)
        self.assertTrue((folder / "001文件A.jpg").exists())
        self.assertFalse((folder / "001.jpg").exists())
        self.assertTrue((folder / "002.jpg").exists())
        self.assertFalse((folder / "002文件A.jpg").exists())


if __name__ == "__main__":
    unittest.main()
