"""按多个数字工作表批量重命名对应文件夹图片的标签页。"""

import tkinter as tk

from excel_batch_renamer.batch_rename_images import batch_rename_images
from excel_batch_renamer.ui.common import TaskTab


class BatchRenameImagesTab(TaskTab):
    """选择工作簿和父目录后自动处理全部数字工作表。"""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.workbook_variable = tk.StringVar()
        self.directory_variable = tk.StringVar()

        self.add_path_picker(
            "Excel 表格",
            self.workbook_variable,
            lambda: self.choose_xlsx(self.workbook_variable),
        )
        self.add_path_picker(
            "父文件夹",
            self.directory_variable,
            lambda: self.choose_directory(self.directory_variable),
        )
        self.add_execute_area(self._execute)

    def _execute(self) -> None:
        self.run_and_report(self._rename_images, "批量重命名图片")

    def _rename_images(self) -> str:
        workbook = self.require_path(self.workbook_variable.get(), "Excel 表格")
        directory = self.require_path(self.directory_variable.get(), "父文件夹")
        result = batch_rename_images(workbook, directory)
        return (
            "完成：已处理 {} 个文件夹、共 {} 张图片，已重命名 {} 张，"
            "未变化 {} 张"
        ).format(
            result.folders,
            result.total,
            result.renamed,
            result.unchanged,
        )
