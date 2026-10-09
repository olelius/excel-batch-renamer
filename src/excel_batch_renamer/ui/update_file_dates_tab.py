"""独立的文件创建/修改日期更新标签页。"""

import tkinter as tk
from datetime import date
from tkinter import ttk

from excel_batch_renamer.ui.common import TaskTab
from excel_batch_renamer.update_file_dates import update_file_dates


class UpdateFileDatesTab(TaskTab):
    """输入日期并选择自己的目录，一次更新其全部子目录中的支持文件。"""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.directory_variable = tk.StringVar()
        self.date_variable = tk.StringVar(value=date.today().isoformat())
        self.add_path_picker("任务文件夹", self.directory_variable, self._browse_directory)
        row = self._next_row
        ttk.Label(self, text="目标日期", width=12).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.date_variable, width=16).grid(
            row=row, column=1, sticky="w", pady=6
        )
        ttk.Label(self, text="格式：YYYY-MM-DD，例如 2026-10-09").grid(
            row=row + 1, column=1, sticky="w", pady=(0, 6)
        )
        ttk.Label(
            self,
            text="包含所有子文件夹：JPG/JPEG、PDF、Excel（xls/xlsx/xlsm/xlsb）。\n"
                 "同时修改创建和修改日期，各自保留原时分秒；不改访问时间或文件内容。\n"
                 "不修改文件夹本身，不跟随符号链接或目录联接。",
            wraplength=680,
            justify="left",
        ).grid(row=row + 2, column=0, columnspan=3, sticky="w", pady=6)
        self._next_row += 3
        self.add_execute_area(self._execute)

    def _browse_directory(self) -> None:
        """只更新本标签页的任务目录。"""

        self.choose_directory(self.directory_variable)

    def _execute(self) -> None:
        """直接执行并把成功数量或失败原因留在当前页。"""

        self.run_and_report(self._update_file_dates, "更新文件日期")

    def _update_file_dates(self) -> str:
        """校验必选目录，将日期原样交给应用服务进行合法性校验。"""

        directory = self.require_path(self.directory_variable.get(), "任务文件夹")
        return update_file_dates(directory, self.date_variable.get()).status_text
