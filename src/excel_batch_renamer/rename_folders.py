"""按文件夹任务表批量重命名直属文件夹。"""

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

from excel_batch_renamer.core.models import FolderTaskRow
from excel_batch_renamer.core.naming import (
    build_folder_name,
    extract_folder_sequence,
)
from excel_batch_renamer.infrastructure.xlsx_reader import read_folder_task


@dataclass(frozen=True)
class FolderRenameOperation:
    """一次已经过完整验证的文件夹重命名操作。"""

    source: Path
    target: Path


@dataclass(frozen=True)
class FolderRenamePlan:
    """重命名执行前生成的不可变计划。"""

    operations: Tuple[FolderRenameOperation, ...]
    unchanged_count: int


@dataclass(frozen=True)
class FolderRenameResult:
    """供 UI 显示的文件夹重命名统计。"""

    renamed_count: int
    unchanged_count: int


class FolderRenameExecutionError(OSError):
    """文件系统重命名失败，并保留失败对象与已完成数量。"""

    def __init__(
        self,
        operation: FolderRenameOperation,
        renamed_count: int,
        unchanged_count: int,
        original_error: OSError,
    ) -> None:
        self.operation = operation
        self.renamed_count = renamed_count
        self.unchanged_count = unchanged_count
        self.original_error = original_error
        super().__init__(
            "重命名文件夹失败：{} -> {}；原因：{}".format(
                operation.source,
                operation.target,
                original_error,
            )
        )


def rename_folders(workbook_path: Path, task_directory: Path) -> FolderRenameResult:
    """读取任务表、完整规划并执行文件夹重命名。"""

    rows = read_folder_task(Path(workbook_path))
    plan = plan_folder_renames(rows, Path(task_directory))
    return execute_folder_rename_plan(plan)


def plan_folder_renames(
    rows: Sequence[FolderTaskRow],
    task_directory: Path,
) -> FolderRenamePlan:
    """在不修改文件系统的前提下生成并验证完整重命名计划。"""

    directory = Path(task_directory)
    if not directory.is_dir():
        raise ValueError("任务目录不存在或不是文件夹：{}".format(directory))

    rows_by_sequence = _validate_task_sequences(rows)
    folders_by_sequence = _find_project_folders(directory)
    operations: List[FolderRenameOperation] = []
    unchanged_count = 0
    target_names = set()

    for sequence in sorted(rows_by_sequence):
        matches = folders_by_sequence.get(sequence, [])
        if not matches:
            raise ValueError(
                "任务目录缺少序号 {:03d} 对应的文件夹".format(sequence)
            )
        if len(matches) > 1:
            names = "、".join(sorted(path.name for path in matches))
            raise ValueError(
                "任务目录中序号 {:03d} 匹配到多个文件夹：{}".format(
                    sequence,
                    names,
                )
            )

        source = matches[0]
        row = rows_by_sequence[sequence]
        target_name = build_folder_name(sequence, row.folder_name)
        if target_name in target_names:
            raise ValueError("重命名计划包含重复目标：{}".format(target_name))
        target_names.add(target_name)

        # 空名称行表示保留现有目录，不把已经补全的名称反向改回占位名。
        if not row.folder_name or source.name == target_name:
            unchanged_count += 1
            continue

        target = directory / target_name
        if target.exists():
            raise ValueError("目标路径已存在：{}".format(target))
        operations.append(FolderRenameOperation(source=source, target=target))

    return FolderRenamePlan(
        operations=tuple(operations),
        unchanged_count=unchanged_count,
    )


def execute_folder_rename_plan(plan: FolderRenamePlan) -> FolderRenameResult:
    """顺序执行计划；首次文件系统失败时停止，不回滚已完成操作。"""

    renamed_count = 0
    for operation in plan.operations:
        try:
            _rename_path(operation.source, operation.target)
        except OSError as error:
            raise FolderRenameExecutionError(
                operation=operation,
                renamed_count=renamed_count,
                unchanged_count=plan.unchanged_count,
                original_error=error,
            ) from error
        renamed_count += 1

    return FolderRenameResult(
        renamed_count=renamed_count,
        unchanged_count=plan.unchanged_count,
    )


def _validate_task_sequences(
    rows: Sequence[FolderTaskRow],
) -> Dict[int, FolderTaskRow]:
    if not rows:
        raise ValueError("文件夹任务表没有有效数据")

    rows_by_sequence: Dict[int, FolderTaskRow] = {}
    duplicates = set()
    for row in rows:
        if row.sequence < 1:
            raise ValueError("文件夹任务表序号必须从 1 开始")
        if row.sequence in rows_by_sequence:
            duplicates.add(row.sequence)
        else:
            rows_by_sequence[row.sequence] = row

    if duplicates:
        values = "、".join("{:03d}".format(value) for value in sorted(duplicates))
        raise ValueError("文件夹任务表包含重复序号：{}".format(values))

    maximum = max(rows_by_sequence)
    missing = sorted(set(range(1, maximum + 1)) - set(rows_by_sequence))
    if missing:
        values = "、".join("{:03d}".format(value) for value in missing)
        raise ValueError("文件夹任务表缺少连续序号：{}".format(values))

    return rows_by_sequence


def _find_project_folders(task_directory: Path) -> Dict[int, List[Path]]:
    folders_by_sequence: Dict[int, List[Path]] = {}
    for child in task_directory.iterdir():
        if not child.is_dir():
            continue
        try:
            sequence = extract_folder_sequence(child.name)
        except ValueError:
            continue
        folders_by_sequence.setdefault(sequence, []).append(child)
    return folders_by_sequence


def _rename_path(source: Path, target: Path) -> None:
    """隔离实际文件系统调用，便于验证无变化项和失败停止语义。"""

    source.rename(target)
