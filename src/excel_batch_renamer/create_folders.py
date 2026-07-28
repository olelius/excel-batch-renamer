"""根据文件夹任务表规划并创建直属文件夹。"""

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple, Union

from excel_batch_renamer.core.naming import build_folder_name
from excel_batch_renamer.infrastructure.xlsx_reader import read_folder_task


PathLike = Union[str, Path]


@dataclass(frozen=True)
class FolderCreationItem:
    """创建计划中的一个目标文件夹。"""

    sequence: int
    target_path: Path


@dataclass(frozen=True)
class FolderCreationPlan:
    """已完成全部静态校验、可供执行的创建计划。"""

    workbook_path: Path
    target_directory: Path
    items: Tuple[FolderCreationItem, ...]

    @property
    def planned_count(self) -> int:
        """返回计划创建的文件夹数量。"""

        return len(self.items)


@dataclass(frozen=True)
class FolderCreationResult:
    """一次成功执行的统计结果，供 UI 直接展示。"""

    planned_count: int
    created_count: int

    @property
    def status_text(self) -> str:
        """返回适合标签页底部展示的完成信息。"""

        return "完成：已创建 {} 个文件夹".format(self.created_count)


class FolderCreationExecutionError(RuntimeError):
    """创建过程中的首个文件系统错误，包含已完成数量和失败对象。"""

    def __init__(
        self,
        failed_path: Path,
        created_count: int,
        planned_count: int,
        cause: OSError,
    ) -> None:
        self.failed_path = failed_path
        self.created_count = created_count
        self.planned_count = planned_count
        self.cause = cause
        super().__init__(
            "创建文件夹失败：{}；已创建 {}/{}；原因：{}".format(
                failed_path,
                created_count,
                planned_count,
                cause,
            )
        )


def build_folder_creation_plan(
    workbook_path: PathLike,
    target_directory: PathLike,
) -> FolderCreationPlan:
    """读取任务表并在任何创建操作前生成、验证完整计划。

    目标目录必须已经存在。任务序号必须按 1 到 n 连续排列；同一序号或
    Windows 语义下重名的目标均视为计划重复。任何目标已经存在时也会拒绝
    整个计划，确保预检失败不会留下部分创建结果。
    """

    workbook = Path(workbook_path)
    target = Path(target_directory)
    if not target.exists():
        raise ValueError("任务目录不存在：{}".format(target))
    if not target.is_dir():
        raise ValueError("任务目录不是文件夹：{}".format(target))

    rows = read_folder_task(workbook)
    items = []
    seen_sequences = set()
    seen_target_names = set()

    for expected_sequence, row in enumerate(rows, start=1):
        if row.sequence in seen_sequences:
            raise ValueError("创建计划包含重复序号：{}".format(row.sequence))
        if row.sequence != expected_sequence:
            raise ValueError(
                "文件夹序号必须从 1 连续排列：第 {} 条数据应为 {}，实际为 {}".format(
                    expected_sequence,
                    expected_sequence,
                    row.sequence,
                )
            )

        target_name = build_folder_name(row.sequence, row.folder_name)
        normalized_name = target_name.casefold()
        if normalized_name in seen_target_names:
            raise ValueError("创建计划包含重复目标：{}".format(target_name))

        target_path = target / target_name
        if target_path.exists():
            raise ValueError("目标文件夹已存在：{}".format(target_path))

        seen_sequences.add(row.sequence)
        seen_target_names.add(normalized_name)
        items.append(
            FolderCreationItem(
                sequence=row.sequence,
                target_path=target_path,
            )
        )

    return FolderCreationPlan(
        workbook_path=workbook,
        target_directory=target,
        items=tuple(items),
    )


def execute_folder_creation_plan(
    plan: FolderCreationPlan,
) -> FolderCreationResult:
    """按顺序执行已验证计划，首次失败即停止且不回滚。"""

    created_count = 0
    for item in plan.items:
        try:
            item.target_path.mkdir()
        except OSError as error:
            raise FolderCreationExecutionError(
                failed_path=item.target_path,
                created_count=created_count,
                planned_count=plan.planned_count,
                cause=error,
            ) from error
        created_count += 1

    return FolderCreationResult(
        planned_count=plan.planned_count,
        created_count=created_count,
    )


def create_folders_from_xlsx(
    workbook_path: PathLike,
    target_directory: PathLike,
) -> FolderCreationResult:
    """从 `.xlsx` 读取任务、完成预检并创建文件夹。"""

    plan = build_folder_creation_plan(workbook_path, target_directory)
    return execute_folder_creation_plan(plan)
