"""标签页共用的 Tkinter 控件与状态处理。"""

import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Callable


LOGGER = logging.getLogger(__name__)
XLSX_FILE_TYPES = (("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*"))


class TaskTab(ttk.Frame):
    """为任务标签页提供统一布局、路径选择和状态显示。"""

    def __init__(self, master) -> None:
        super().__init__(master, padding=16)
        self.columnconfigure(1, weight=1)
        self.status_variable = tk.StringVar(value="请选择任务输入。")
        self._next_row = 0

    def add_path_picker(
        self,
        label_text: str,
        variable: tk.StringVar,
        browse_command: Callable[[], None],
    ) -> None:
        """按统一样式添加一个路径输入和浏览按钮。"""

        row = self._next_row
        ttk.Label(self, text=label_text, width=12).grid(
            row=row,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        ttk.Entry(
            self,
            textvariable=variable,
            state="readonly",
        ).grid(
            row=row,
            column=1,
            sticky="ew",
            pady=6,
        )
        ttk.Button(self, text="选择…", command=browse_command).grid(
            row=row,
            column=2,
            padx=(8, 0),
            pady=6,
        )
        self._next_row += 1

    def choose_xlsx(self, variable: tk.StringVar) -> bool:
        """选择 `.xlsx` 文件，返回用户是否完成了选择。"""

        selected = filedialog.askopenfilename(
            parent=self,
            title="选择 Excel 工作簿",
            filetypes=XLSX_FILE_TYPES,
        )
        if not selected:
            return False
        variable.set(selected)
        return True

    def choose_directory(self, variable: tk.StringVar) -> bool:
        """选择任务目录，返回用户是否完成了选择。"""

        selected = filedialog.askdirectory(
            parent=self,
            title="选择任务文件夹",
            mustexist=True,
        )
        if not selected:
            return False
        variable.set(selected)
        return True

    def add_execute_area(self, execute_command: Callable[[], None]) -> None:
        """添加直接执行按钮和标签页底部状态栏。"""

        row = self._next_row
        ttk.Button(
            self,
            text="直接执行",
            command=execute_command,
        ).grid(row=row, column=0, columnspan=3, pady=(18, 12))
        ttk.Separator(self).grid(
            row=row + 1,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(0, 10),
        )
        ttk.Label(
            self,
            textvariable=self.status_variable,
            wraplength=680,
            justify="left",
        ).grid(row=row + 2, column=0, columnspan=3, sticky="ew")
        self.rowconfigure(row + 2, weight=1)
        self._next_row += 3

    def require_path(self, value: str, label: str) -> Path:
        """将必选路径转换为 ``Path``，空值直接阻止执行。"""

        text = value.strip()
        if not text:
            raise ValueError("请先选择{}".format(label))
        return Path(text)

    def run_and_report(
        self,
        operation: Callable[[], str],
        operation_name: str,
    ) -> None:
        """同步执行任务，把成功或失败结果留在当前标签页。"""

        self.status_variable.set("正在{}…".format(operation_name))
        self.update_idletasks()
        LOGGER.info("开始%s", operation_name)
        try:
            status = operation()
        except Exception as error:
            LOGGER.exception("%s失败", operation_name)
            self.status_variable.set("失败：{}".format(error))
            return

        LOGGER.info("%s完成：%s", operation_name, status)
        self.status_variable.set(status)
