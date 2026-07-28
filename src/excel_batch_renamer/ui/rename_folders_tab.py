"""重命名文件夹标签页。"""

import tkinter as tk

from excel_batch_renamer.rename_folders import rename_folders
from excel_batch_renamer.ui.common import TaskTab


class RenameFoldersTab(TaskTab):
    """按表格序号补全或更新任务文件夹名称。"""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.workbook_variable = tk.StringVar()
        self.directory_variable = tk.StringVar()

        self.add_path_picker(
            "Excel 表格",
            self.workbook_variable,
            self._browse_workbook,
        )
        self.add_path_picker(
            "任务文件夹",
            self.directory_variable,
            self._browse_directory,
        )
        self.add_execute_area(self._execute)

    def _browse_workbook(self) -> None:
        self.choose_xlsx(self.workbook_variable)

    def _browse_directory(self) -> None:
        self.choose_directory(self.directory_variable)

    def _execute(self) -> None:
        self.run_and_report(self._rename_folders, "重命名文件夹")

    def _rename_folders(self) -> str:
        workbook = self.require_path(self.workbook_variable.get(), "Excel 表格")
        directory = self.require_path(self.directory_variable.get(), "任务文件夹")
        result = rename_folders(workbook, directory)
        return "完成：已重命名 {} 个文件夹，未变化 {} 个".format(
            result.renamed_count,
            result.unchanged_count,
        )
