import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from excel_batch_renamer.batch_rename_images import BatchImageRenameResult
from excel_batch_renamer.rename_images import ImageRenameResult
from excel_batch_renamer.ui.main_window import MainWindow
from excel_batch_renamer.ui.rename_images_tab import RenameImagesTab


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

    def test_main_window_has_four_independent_directory_variables(self):
        create_tab = self.window.create_folders_tab
        folder_tab = self.window.rename_folders_tab
        image_tab = self.window.rename_images_tab
        batch_image_tab = self.window.batch_rename_images_tab

        create_tab.directory_variable.set("C:/create")
        folder_tab.directory_variable.set("C:/rename")
        image_tab.directory_variable.set("C:/images")
        batch_image_tab.directory_variable.set("C:/batch-images")

        self.assertEqual(create_tab.directory_variable.get(), "C:/create")
        self.assertEqual(folder_tab.directory_variable.get(), "C:/rename")
        self.assertEqual(image_tab.directory_variable.get(), "C:/images")
        self.assertEqual(
            batch_image_tab.directory_variable.get(),
            "C:/batch-images",
        )

    def test_batch_image_tab_passes_workbook_and_parent_to_service(self):
        tab = self.window.batch_rename_images_tab
        tab.workbook_variable.set("C:/tasks.xlsx")
        tab.directory_variable.set("C:/parent")
        expected = BatchImageRenameResult(
            folders=2,
            total=3,
            renamed=2,
            unchanged=1,
            skipped_worksheets=("3", "5"),
        )

        with patch(
            "excel_batch_renamer.ui.batch_rename_images_tab.batch_rename_images",
            return_value=expected,
        ) as service:
            status = tab._rename_images()

        service.assert_called_once_with(
            Path("C:/tasks.xlsx"),
            Path("C:/parent"),
        )
        self.assertIn("已处理 2 个文件夹", status)
        self.assertIn("未变化 1 张", status)
        self.assertIn("已跳过空工作表 2 个（3、5）", status)

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
        expected = ImageRenameResult(total=3, renamed=2, unchanged=1)

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

    def test_task_failure_is_reported_in_tab_without_raising(self):
        tab = self.window.create_folders_tab
        operation = Mock(side_effect=PermissionError("拒绝访问"))

        tab.run_and_report(operation, "创建文件夹")

        self.assertEqual(tab.status_variable.get(), "失败：拒绝访问")


if __name__ == "__main__":
    unittest.main()
