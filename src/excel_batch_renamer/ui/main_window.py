"""应用主窗口。"""

import tkinter as tk
from tkinter import ttk

from excel_batch_renamer.ui.batch_rename_images_tab import BatchRenameImagesTab
from excel_batch_renamer.ui.create_folders_tab import CreateFoldersTab
from excel_batch_renamer.ui.rename_folders_tab import RenameFoldersTab
from excel_batch_renamer.ui.rename_images_tab import RenameImagesTab


class MainWindow(tk.Tk):
    """包含四个相互独立任务标签页的单一主窗口。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 批量文件整理工具")
        self.geometry("780x400")
        self.minsize(700, 360)

        notebook = ttk.Notebook(self, padding=8)
        notebook.pack(fill="both", expand=True)

        self.create_folders_tab = CreateFoldersTab(notebook)
        self.rename_folders_tab = RenameFoldersTab(notebook)
        self.rename_images_tab = RenameImagesTab(notebook)
        self.batch_rename_images_tab = BatchRenameImagesTab(notebook)

        notebook.add(self.create_folders_tab, text="创建文件夹")
        notebook.add(self.rename_folders_tab, text="重命名文件夹")
        notebook.add(self.rename_images_tab, text="重命名图片")
        notebook.add(self.batch_rename_images_tab, text="批量重命名图片")
