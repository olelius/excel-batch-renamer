"""重命名 JPG 图片标签页。"""

import logging
import tkinter as tk
from tkinter import ttk

from excel_batch_renamer.rename_images import (
    get_worksheet_names,
    rename_images,
    suggest_worksheet_name,
)
from excel_batch_renamer.ui.common import TaskTab


LOGGER = logging.getLogger(__name__)


class RenameImagesTab(TaskTab):
    """按 Excel、文件夹、工作表的固定顺序组织图片任务。"""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.workbook_variable = tk.StringVar()
        self.directory_variable = tk.StringVar()
        self.worksheet_variable = tk.StringVar()
        self.worksheet_names = []

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
        self._add_worksheet_picker()
        self.add_execute_area(self._execute)

    def _add_worksheet_picker(self) -> None:
        row = self._next_row
        ttk.Label(self, text="工作表", width=12).grid(
            row=row,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        self.worksheet_combobox = ttk.Combobox(
            self,
            textvariable=self.worksheet_variable,
            state="readonly",
            values=(),
        )
        self.worksheet_combobox.grid(
            row=row,
            column=1,
            columnspan=2,
            sticky="ew",
            pady=6,
        )
        self._next_row += 1

    def _browse_workbook(self) -> None:
        if not self.choose_xlsx(self.workbook_variable):
            return
        self._load_worksheets()

    def _browse_directory(self) -> None:
        if not self.choose_directory(self.directory_variable):
            return
        self._select_matching_worksheet()

    def _load_worksheets(self) -> None:
        try:
            workbook = self.require_path(
                self.workbook_variable.get(),
                "Excel 表格",
            )
            self.worksheet_names = get_worksheet_names(workbook)
        except Exception as error:
            LOGGER.exception("加载工作表失败")
            self.worksheet_names = []
            self.worksheet_combobox.configure(values=())
            self.worksheet_variable.set("")
            self.status_variable.set("失败：{}".format(error))
            return

        self.worksheet_combobox.configure(values=self.worksheet_names)
        self.worksheet_variable.set("")
        self.status_variable.set(
            "已加载 {} 个工作表，请选择任务文件夹。".format(
                len(self.worksheet_names)
            )
        )
        if self.directory_variable.get().strip():
            self._select_matching_worksheet()

    def _select_matching_worksheet(self) -> None:
        if not self.directory_variable.get().strip():
            return
        try:
            directory = self.require_path(
                self.directory_variable.get(),
                "任务文件夹",
            )
            suggested = suggest_worksheet_name(
                directory,
                self.worksheet_names,
            )
        except Exception as error:
            LOGGER.exception("自动选择工作表失败")
            self.worksheet_variable.set("")
            self.status_variable.set("失败：{}".format(error))
            return

        if suggested is None:
            self.worksheet_variable.set("")
            self.status_variable.set("未找到与任务文件夹编号对应的工作表。")
            return
        self.worksheet_variable.set(suggested)
        self.status_variable.set("已自动选择工作表 {}。".format(suggested))

    def _execute(self) -> None:
        self.run_and_report(self._rename_images, "重命名图片")

    def _rename_images(self) -> str:
        workbook = self.require_path(self.workbook_variable.get(), "Excel 表格")
        directory = self.require_path(self.directory_variable.get(), "任务文件夹")
        worksheet = self.worksheet_variable.get().strip()
        if not worksheet:
            raise ValueError("请先选择工作表")
        result = rename_images(workbook, worksheet, directory)
        return "完成：共处理 {} 张，已重命名 {} 张，未变化 {} 张".format(
            result.total,
            result.renamed,
            result.unchanged,
        )
