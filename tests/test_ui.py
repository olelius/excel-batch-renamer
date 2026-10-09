import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from excel_batch_renamer.catalog_reindex import CatalogReindexResult
from excel_batch_renamer.rename_images import (
    ImageRenameExecutionError,
    ImageRenameResult,
)
from excel_batch_renamer.ui.main_window import MainWindow
from excel_batch_renamer.ui.rename_images_tab import RenameImagesTab
from excel_batch_renamer.update_file_dates import FileDateResult, FileDateExecutionError


class UiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.window = MainWindow()
            cls.window.withdraw()
        except tk.TclError as error:
            raise unittest.SkipTest(
                "当前环境不能创建 Tkinter 窗口：{}".format(error)
            )

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "window"):
            cls.window.destroy()

    def test_main_window_has_five_independent_directory_variables(self):
        create_tab = self.window.create_folders_tab
        folder_tab = self.window.rename_folders_tab
        image_tab = self.window.rename_images_tab
        batch_image_tab = self.window.batch_rename_images_tab
        date_tab = self.window.update_file_dates_tab

        create_tab.directory_variable.set("C:/create")
        folder_tab.directory_variable.set("C:/rename")
        image_tab.directory_variable.set("C:/images")
        batch_image_tab.directory_variable.set("C:/batch-images")
        date_tab.directory_variable.set("C:/dates")

        self.assertEqual(create_tab.directory_variable.get(), "C:/create")
        self.assertEqual(folder_tab.directory_variable.get(), "C:/rename")
        self.assertEqual(image_tab.directory_variable.get(), "C:/images")
        self.assertEqual(date_tab.directory_variable.get(), "C:/dates")
        notebook = date_tab.master
        self.assertEqual(len(notebook.tabs()), 5)
        self.assertEqual(notebook.tab(date_tab, "text"), "更新文件日期")
        self.assertEqual(
            batch_image_tab.directory_variable.get(),
            "C:/batch-images",
        )

    def test_date_tab_passes_own_directory_and_input_date(self):
        tab = self.window.update_file_dates_tab
        tab.directory_variable.set("C:/档案")
        tab.date_variable.set("2024-02-29")
        with patch(
            "excel_batch_renamer.ui.update_file_dates_tab.update_file_dates",
            return_value=FileDateResult(3, 2, 1),
        ) as service, patch("tkinter.messagebox.showinfo") as popup:
            tab._execute()
        service.assert_called_once_with(Path("C:/档案"), "2024-02-29")
        self.assertIn("已修改 2 个，未变化 1 个", tab.status_variable.get())
        popup.assert_not_called()

    def test_date_tab_requires_own_directory(self):
        tab = self.window.update_file_dates_tab
        tab.directory_variable.set("")
        with patch("excel_batch_renamer.ui.update_file_dates_tab.update_file_dates") as service:
            tab._execute()
        service.assert_not_called()
        self.assertIn("请先选择任务文件夹", tab.status_variable.get())

    def test_date_tab_failure_shows_failed_object_and_progress_without_popup(self):
        tab = self.window.update_file_dates_tab
        tab.directory_variable.set("C:/档案")
        error = FileDateExecutionError(Path("C:/档案/文件.pdf"), 2, 1, 5, PermissionError("拒绝访问"))
        with patch(
            "excel_batch_renamer.ui.update_file_dates_tab.update_file_dates", side_effect=error
        ), patch("tkinter.messagebox.showerror") as popup:
            tab._execute()
        self.assertIn("文件.pdf", tab.status_variable.get())
        self.assertIn("已修改 2 个，未变化 1 个", tab.status_variable.get())
        popup.assert_not_called()

    def test_batch_image_tab_passes_three_workbooks_and_parent_to_service(self):
        tab = self.window.batch_rename_images_tab
        tab.file_workbook_variable.set("C:/项目文件目录.xlsx")
        tab.drawing_workbook_variable.set("C:/项目图纸目录.xlsx")
        tab.archive_name_workbook_variable.set("C:/档号名称对应表.xlsx")
        tab.directory_variable.set("C:/parent")
        expected = CatalogReindexResult(
            worksheets=7,
            folders_renamed=7,
            folders_unchanged=0,
            images=10,
            images_renamed=9,
            images_unchanged=1,
            generated_pdfs=2,
            mapping_path=Path("C:/parent/档号与新序号对应表.xlsx"),
        )

        with patch(
            "excel_batch_renamer.ui.batch_rename_images_tab.organize_catalog_images",
            return_value=expected,
        ) as service:
            status = tab._rename_images()

        service.assert_called_once_with(
            Path("C:/项目文件目录.xlsx"),
            Path("C:/项目图纸目录.xlsx"),
            Path("C:/档号名称对应表.xlsx"),
            Path("C:/parent"),
        )
        self.assertIn("已编排 7 个工作表", status)
        self.assertIn("文件夹已重命名 7 个", status)
        self.assertIn("未变化 1 张", status)
        self.assertIn("已生成 PDF 2 个", status)

    def test_image_tab_loads_sheets_then_auto_selects_folder_match(self):
        tab = self.window.rename_images_tab
        tab.workbook_variable.set("C:/tasks.xlsx")
        tab.directory_variable.set("C:/001——")

        with patch(
            "excel_batch_renamer.ui.rename_images_tab.get_worksheet_names",
            return_value=["1", "2"],
        ), patch(
            "excel_batch_renamer.ui.rename_images_tab.suggest_worksheet_name",
            return_value="1",
        ):
            tab._load_worksheets()

        self.assertEqual(tab.worksheet_names, ["1", "2"])
        self.assertEqual(tab.worksheet_variable.get(), "1")
        self.assertIn("自动选择工作表 1", tab.status_variable.get())

    def test_image_tab_clears_sheet_when_folder_has_no_match(self):
        tab = self.window.rename_images_tab
        tab.directory_variable.set("C:/003——")
        tab.worksheet_names = ["1", "2"]
        tab.worksheet_variable.set("1")

        with patch(
            "excel_batch_renamer.ui.rename_images_tab.suggest_worksheet_name",
            return_value=None,
        ):
            tab._select_matching_worksheet()

        self.assertEqual(tab.worksheet_variable.get(), "")
        self.assertIn("未找到", tab.status_variable.get())

    def test_manual_sheet_selection_is_passed_to_service(self):
        tab = self.window.rename_images_tab
        tab.workbook_variable.set("C:/tasks.xlsx")
        tab.directory_variable.set("C:/001——")
        tab.worksheet_variable.set("2")
        expected = ImageRenameResult(total=3, renamed=2, unchanged=1, generated_pdfs=1)

        with patch(
            "excel_batch_renamer.ui.rename_images_tab.rename_images",
            return_value=expected,
        ) as service:
            status = tab._rename_images()

        service.assert_called_once_with(
            Path("C:/tasks.xlsx"),
            "2",
            Path("C:/001——"),
        )
        self.assertIn("共处理 3 张", status)
        self.assertIn("未变化 1 张", status)
        self.assertIn("已生成 PDF 1 个", status)
        self.assertIn("保存在原图片文件夹", status)
        self.assertIn("JPG 已保留", status)

    def test_single_pdf_failure_shows_image_and_pdf_progress_in_existing_tab(self):
        tab = self.window.rename_images_tab
        tab.workbook_variable.set("C:/tasks.xlsx")
        tab.directory_variable.set("C:/001——")
        tab.worksheet_variable.set("1")
        error = ImageRenameExecutionError(
            Path("C:/001——/002乙.pdf"),
            PermissionError("PDF 被占用"),
            ImageRenameResult(total=3, renamed=2, unchanged=1, generated_pdfs=1),
            operation="生成 PDF",
        )

        with patch(
            "excel_batch_renamer.ui.rename_images_tab.rename_images", side_effect=error
        ), patch("tkinter.messagebox.showinfo") as success_popup, patch(
            "tkinter.messagebox.showerror"
        ) as error_popup:
            tab._execute()

        status = tab.status_variable.get()
        self.assertIn("002乙.pdf", status)
        self.assertIn("PDF 被占用", status)
        self.assertIn("已重命名 2 张，未变化 1 张，已生成 PDF 1 个", status)
        success_popup.assert_not_called()
        error_popup.assert_not_called()

    def test_task_failure_is_reported_in_tab_without_raising(self):
        tab = self.window.create_folders_tab
        operation = Mock(side_effect=PermissionError("拒绝访问"))

        tab.run_and_report(operation, "创建文件夹")

        self.assertEqual(tab.status_variable.get(), "失败：拒绝访问")


if __name__ == "__main__":
    unittest.main()
