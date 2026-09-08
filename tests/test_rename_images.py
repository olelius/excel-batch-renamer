import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from PIL import Image
from pypdf import PdfReader

from excel_batch_renamer.infrastructure.pdf_writer import write_jpeg_pdf
from excel_batch_renamer.infrastructure.xlsx_reader import read_image_task
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
            worksheet.append(["图片档案目录"])
            worksheet.append(["制表说明"])
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
            with Image.new("RGB", (24 + page, 32), (page % 256, 80, 160)) as image:
                image.save(folder / "{:03d}{}".format(page, suffix), format="JPEG")

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
        self.assertEqual(result.generated_pdfs, 3)
        for title, page_count in [("文件A", 4), ("文件B", 5), ("文件C", 8)]:
            self.assertEqual(len(PdfReader(folder / (title + ".pdf")).pages), page_count)

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
        self.assertEqual(first_repeat.generated_pdfs, 1)
        self.assertEqual(second_repeat.generated_pdfs, 1)
        self.assertTrue((folder / "旧题名.pdf").is_file())
        self.assertEqual(len(PdfReader(folder / "新题名.pdf").pages), 2)

    def test_acceptance_pages_six_to_ten_generate_five_page_pdf(self):
        self._write_workbook([("1", [("验收文件", "6-10")])])
        folder = self._make_folder()
        self._make_images(folder, range(6, 11))

        with patch(
            "excel_batch_renamer.rename_images.read_image_task", wraps=read_image_task
        ) as reader:
            result = rename_images(self.workbook_path, "1", folder)

        reader.assert_called_once_with(self.workbook_path, "1")
        self.assertEqual((result.total, result.renamed, result.generated_pdfs), (5, 5, 1))
        self.assertEqual(len(PdfReader(folder / "验收文件.pdf").pages), 5)
        self.assertEqual(
            sorted(path.name for path in folder.glob("*.jpg")),
            ["{:03d}验收文件.jpg".format(page) for page in range(6, 11)],
        )

    def test_non_adjacent_rows_with_same_title_share_one_pdf_in_page_order(self):
        self._write_workbook(
            [("1", [("验收文件", "6"), ("附件", "8"), ("验收文件", "9-10")])]
        )
        folder = self._make_folder()
        self._make_images(folder, [10, 8, 6, 9, 7])

        with patch(
            "excel_batch_renamer.rename_images.write_jpeg_pdf", wraps=write_jpeg_pdf
        ) as writer:
            result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual(result.generated_pdfs, 2)
        self.assertEqual(
            writer.call_args_list[0].args,
            (
                [folder / "{:03d}验收文件.jpg".format(page) for page in [6, 7, 9, 10]],
                folder / "验收文件.pdf",
                "验收文件",
            ),
        )
        self.assertEqual(len(PdfReader(folder / "验收文件.pdf").pages), 4)
        self.assertEqual(len(PdfReader(folder / "附件.pdf").pages), 1)

    def test_repeat_overwrites_existing_pdf_even_when_all_images_unchanged(self):
        self._write_workbook([("1", [("验收文件", "6-10")])])
        folder = self._make_folder()
        self._make_images(folder, range(6, 11))
        rename_images(self.workbook_path, "1", folder)
        target_pdf = folder / "验收文件.pdf"
        target_pdf.write_bytes(b"previous PDF content")

        result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual((result.renamed, result.unchanged, result.generated_pdfs), (0, 5, 1))
        self.assertEqual(len(PdfReader(target_pdf).pages), 5)

    def test_repeat_rebuilds_pdfs_from_current_ranges_and_image_collection(self):
        self._write_workbook([("1", [("甲", "1"), ("乙", "3-4")])])
        folder = self._make_folder()
        self._make_images(folder, range(1, 5))
        rename_images(self.workbook_path, "1", folder)
        (folder / "004乙.jpg").unlink()
        self._write_workbook([("1", [("甲", "1"), ("乙", "2-3")])])

        result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual((result.total, result.renamed, result.unchanged), (3, 1, 2))
        self.assertEqual(result.generated_pdfs, 2)
        self.assertEqual(len(PdfReader(folder / "甲.pdf").pages), 1)
        self.assertEqual(len(PdfReader(folder / "乙.pdf").pages), 2)
        self.assertTrue((folder / "002乙.jpg").exists())

    def test_case_only_unchanged_jpg_is_included_in_pdf(self):
        self._write_workbook([("1", [("文件A", "1-1")])])
        folder = self._make_folder()
        self._make_images(folder, [1])
        (folder / "001.jpg").rename(folder / "001文件a.JPG")

        result = rename_images(self.workbook_path, "1", folder)

        self.assertEqual((result.renamed, result.unchanged, result.generated_pdfs), (0, 1, 1))
        self.assertEqual(len(PdfReader(folder / "文件A.pdf").pages), 1)

    def test_case_ambiguous_titles_block_all_renames(self):
        self._write_workbook([("1", [("文件A", "1"), ("文件a", "2-2")])])
        folder = self._make_folder()
        self._make_images(folder, [1, 2])

        with self.assertRaisesRegex(ValueError, "Windows 同名歧义"):
            rename_images(self.workbook_path, "1", folder)

        self.assertEqual(sorted(path.name for path in folder.iterdir()), ["001.jpg", "002.jpg"])

    def test_invalid_jpg_blocks_all_renames_and_pdf_writes(self):
        self._write_workbook([("1", [("文件A", "1-2")])])
        folder = self._make_folder()
        self._make_images(folder, [1])
        (folder / "002.jpg").write_bytes(b"corrupt JPEG")

        with self.assertRaisesRegex(ValueError, "JPG 图片校验失败.*002.jpg"):
            rename_images(self.workbook_path, "1", folder)

        self.assertEqual(sorted(path.name for path in folder.iterdir()), ["001.jpg", "002.jpg"])

    def test_existing_pdf_target_directory_blocks_all_renames(self):
        self._write_workbook([("1", [("文件A", "1-1")])])
        folder = self._make_folder()
        self._make_images(folder, [1])
        (folder / "文件A.pdf").mkdir()

        with self.assertRaisesRegex(ValueError, "PDF 目标名称已被目录占用"):
            rename_images(self.workbook_path, "1", folder)

        self.assertTrue((folder / "001.jpg").exists())
        self.assertTrue((folder / "文件A.pdf").is_dir())

    def test_pdf_failure_preserves_completed_images_and_pdf_progress(self):
        self._write_workbook([("1", [("甲", "1"), ("乙", "2-3")])])
        folder = self._make_folder()
        self._make_images(folder, [1, 2, 3])
        (folder / "001.jpg").rename(folder / "001甲.jpg")

        def fail_second_pdf(image_paths, target_path, title):
            if title == "乙":
                raise PermissionError("PDF 被占用")
            write_jpeg_pdf(image_paths, target_path, title)

        with patch(
            "excel_batch_renamer.rename_images.write_jpeg_pdf", side_effect=fail_second_pdf
        ):
            with self.assertRaises(ImageRenameExecutionError) as captured:
                rename_images(self.workbook_path, "1", folder)

        error = captured.exception
        self.assertEqual(error.operation, "生成 PDF")
        self.assertEqual(error.failed_path, folder / "乙.pdf")
        self.assertEqual(
            (error.result.total, error.result.renamed, error.result.unchanged,
             error.result.generated_pdfs),
            (3, 2, 1, 1),
        )
        self.assertIn("已生成 PDF 1 个", str(error))
        self.assertEqual(len(PdfReader(folder / "甲.pdf").pages), 1)
        self.assertTrue((folder / "002乙.jpg").exists())
        self.assertTrue((folder / "003乙.jpg").exists())
        self.assertFalse((folder / "乙.pdf").exists())

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
