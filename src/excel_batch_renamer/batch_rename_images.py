"""按数字工作表批量匹配多个项目文件夹并重命名图片。"""

from dataclasses import dataclass
from pathlib import Path
from typing import List, Sequence

from excel_batch_renamer.core.naming import extract_folder_sequence
from excel_batch_renamer.infrastructure.xlsx_reader import list_worksheet_names
from excel_batch_renamer.rename_images import (
    ImageRenameExecutionError,
    ImageRenamePlanItem,
    build_image_rename_plan,
    execute_image_rename_plan,
)


@dataclass(frozen=True)
class BatchImageFolderPlan:
    """一个数字工作表、对应文件夹及其已校验图片计划。"""

    worksheet_name: str
    folder_path: Path
    items: Sequence[ImageRenamePlanItem]


@dataclass(frozen=True)
class BatchImageRenameResult:
    """一次多工作表图片任务的汇总统计。"""

    folders: int
    total: int
    renamed: int
    unchanged: int


class BatchImageRenameExecutionError(RuntimeError):
    """批量执行期间的文件系统错误及已经完成的汇总统计。"""

    def __init__(
        self,
        worksheet_name: str,
        folder_path: Path,
        failed_path: Path,
        reason: BaseException,
        result: BatchImageRenameResult,
    ) -> None:
        self.worksheet_name = worksheet_name
        self.folder_path = folder_path
        self.failed_path = failed_path
        self.reason = reason
        self.result = result
        super().__init__(
            "批量重命名图片失败：工作表 {}，文件夹 {}，失败对象 {}；原因：{}".format(
                worksheet_name,
                folder_path,
                failed_path,
                reason,
            )
        )


def build_batch_image_rename_plan(
    workbook_path: Path,
    parent_directory: Path,
) -> List[BatchImageFolderPlan]:
    """先验证全部数字工作表及对应直属文件夹，再返回整批计划。"""

    workbook_path = Path(workbook_path)
    parent_directory = Path(parent_directory)
    if not parent_directory.is_dir():
        raise ValueError(
            "批量图片任务父目录不存在或不是文件夹：{}".format(parent_directory)
        )

    worksheets_by_sequence = {}
    for worksheet_name in list_worksheet_names(workbook_path):
        text = str(worksheet_name).strip()
        if not text.isdigit() or int(text) < 1:
            continue
        sequence = int(text)
        if sequence in worksheets_by_sequence:
            raise ValueError(
                "多个数字工作表对应同一文件夹序号：{}、{}".format(
                    worksheets_by_sequence[sequence],
                    worksheet_name,
                )
            )
        worksheets_by_sequence[sequence] = worksheet_name

    if not worksheets_by_sequence:
        raise ValueError("图片任务表中没有可处理的数字工作表")

    folders_by_sequence = {}
    for path in parent_directory.iterdir():
        if not path.is_dir():
            continue
        try:
            sequence = extract_folder_sequence(path.name)
        except ValueError:
            continue
        folders_by_sequence.setdefault(sequence, []).append(path)

    plans = []
    for sequence in sorted(worksheets_by_sequence):
        worksheet_name = worksheets_by_sequence[sequence]
        matches = folders_by_sequence.get(sequence, [])
        if not matches:
            raise ValueError(
                "工作表 {} 未找到前三位序号匹配的直属文件夹".format(
                    worksheet_name
                )
            )
        if len(matches) > 1:
            raise ValueError(
                "工作表 {} 匹配到多个直属文件夹：{}".format(
                    worksheet_name,
                    "、".join(path.name for path in sorted(matches)),
                )
            )

        folder_path = matches[0]
        items = tuple(
            build_image_rename_plan(
                workbook_path,
                worksheet_name,
                folder_path,
            )
        )
        plans.append(
            BatchImageFolderPlan(
                worksheet_name=worksheet_name,
                folder_path=folder_path,
                items=items,
            )
        )

    return plans


def batch_rename_images(
    workbook_path: Path,
    parent_directory: Path,
) -> BatchImageRenameResult:
    """按工作表编号处理全部匹配文件夹；执行失败时立即停止且不回滚。"""

    plans = build_batch_image_rename_plan(workbook_path, parent_directory)
    total = sum(len(folder_plan.items) for folder_plan in plans)
    completed_folders = 0
    renamed = 0
    unchanged = 0

    for folder_plan in plans:
        try:
            result = execute_image_rename_plan(folder_plan.items)
        except ImageRenameExecutionError as error:
            raise BatchImageRenameExecutionError(
                worksheet_name=folder_plan.worksheet_name,
                folder_path=folder_plan.folder_path,
                failed_path=error.failed_path,
                reason=error.reason,
                result=BatchImageRenameResult(
                    folders=completed_folders,
                    total=total,
                    renamed=renamed + error.result.renamed,
                    unchanged=unchanged + error.result.unchanged,
                ),
            ) from error

        completed_folders += 1
        renamed += result.renamed
        unchanged += result.unchanged

    return BatchImageRenameResult(
        folders=completed_folders,
        total=total,
        renamed=renamed,
        unchanged=unchanged,
    )
