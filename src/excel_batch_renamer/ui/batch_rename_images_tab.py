"""按文件目录、图纸目录顺序编排后批量整理图片的标签页。"""

import tkinter as tk

from excel_batch_renamer.catalog_reindex import organize_catalog_images
from excel_batch_renamer.ui.common import TaskTab


class BatchRenameImagesTab(TaskTab):
    """选择两个目录工作簿和父目录后统一生成新序号并执行。"""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.file_workbook_variable = tk.StringVar()
        self.drawing_workbook_variable = tk.StringVar()
        self.archive_name_workbook_variable = tk.StringVar()
        self.directory_variable = tk.StringVar()

        self.add_path_picker(
            "文件目录 Excel",
            self.file_workbook_variable,
            lambda: self.choose_xlsx(self.file_workbook_variable),
        )
        self.add_path_picker(
            "图纸目录 Excel",
            self.drawing_workbook_variable,
            lambda: self.choose_xlsx(self.drawing_workbook_variable),
        )
        self.add_path_picker(
            "档号名称 Excel",
            self.archive_name_workbook_variable,
            lambda: self.choose_xlsx(self.archive_name_workbook_variable),
        )
        self.add_path_picker(
            "父文件夹",
            self.directory_variable,
            lambda: self.choose_directory(self.directory_variable),
        )
        self.add_execute_area(self._execute)

    def _execute(self) -> None:
        self.run_and_report(self._rename_images, "编排目录并批量整理图片")

    def _rename_images(self) -> str:
        file_workbook = self.require_path(
            self.file_workbook_variable.get(),
            "文件目录 Excel",
        )
        drawing_workbook = self.require_path(
            self.drawing_workbook_variable.get(),
            "图纸目录 Excel",
        )
        archive_name_workbook = self.require_path(
            self.archive_name_workbook_variable.get(),
            "档号名称 Excel",
        )
        directory = self.require_path(self.directory_variable.get(), "父文件夹")
        result = organize_catalog_images(
            file_workbook,
            drawing_workbook,
            archive_name_workbook,
            directory,
        )
        return (
            "完成：已编排 {} 个工作表，文件夹已重命名 {} 个、未变化 {} 个；"
            "共处理 {} 张图片，已重命名 {} 张、未变化 {} 张，已生成 PDF {} 个；"
            "档号映射表：{}"
        ).format(
            result.worksheets,
            result.folders_renamed,
            result.folders_unchanged,
            result.images,
            result.images_renamed,
            result.images_unchanged,
            result.generated_pdfs,
            result.mapping_path,
        )
