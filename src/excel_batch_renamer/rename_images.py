"""图片重命名应用服务。

本模块负责把工作簿、文件夹和纯业务规则组织成一次可执行任务。所有可预见
的数据问题都会在第一个文件改名之前完成验证；真正的文件系统错误则按项目
约定立即停止，不回滚已经完成的操作。
"""

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence
from uuid import uuid4

from excel_batch_renamer.core.naming import (
    build_image_name,
    extract_folder_sequence,
    worksheet_matches_folder,
    worksheet_name_for_folder,
)
from excel_batch_renamer.core.page_ranges import build_page_title_map
from excel_batch_renamer.infrastructure.xlsx_reader import (
    list_worksheet_names,
    read_image_task,
)


@dataclass(frozen=True)
class ImageRenamePlanItem:
    """一张图片的当前路径、目标路径和页码。"""

    page: int
    source: Path
    target: Path


@dataclass(frozen=True)
class ImageRenameResult:
    """一次图片重命名任务的统计结果。"""

    total: int
    renamed: int
    unchanged: int


class ImageRenameExecutionError(RuntimeError):
    """文件系统改名失败，并携带失败对象与已完成统计。"""

    def __init__(
        self,
        failed_path: Path,
        reason: BaseException,
        result: ImageRenameResult,
    ) -> None:
        self.failed_path = failed_path
        self.reason = reason
        self.result = result
        super().__init__(
            "重命名图片失败：{}；原因：{}".format(failed_path, reason)
        )


def get_worksheet_names(workbook_path: Path) -> List[str]:
    """枚举图片任务表的工作表，供 UI 下拉框使用。"""

    return list_worksheet_names(Path(workbook_path))


def suggest_worksheet_name(
    folder_path: Path,
    worksheet_names: Sequence[str],
) -> Optional[str]:
    """按文件夹三位序号建议工作表；不存在时返回 ``None``。"""

    suggested = worksheet_name_for_folder(Path(folder_path).name)
    return suggested if suggested in worksheet_names else None


def build_image_rename_plan(
    workbook_path: Path,
    worksheet_name: str,
    folder_path: Path,
) -> List[ImageRenamePlanItem]:
    """读取当前状态、完成全部验证并生成不可变的重命名计划。"""

    workbook_path = Path(workbook_path)
    folder_path = Path(folder_path)
    selected_worksheet = str(worksheet_name)

    if not folder_path.is_dir():
        raise ValueError("图片任务目录不存在或不是文件夹：{}".format(folder_path))

    # 提取序号也会验证目录名称是否属于本项目规定的格式。
    extract_folder_sequence(folder_path.name)
    if not worksheet_matches_folder(selected_worksheet, folder_path.name):
        raise ValueError(
            "工作表 {} 与文件夹 {} 的序号不匹配".format(
                selected_worksheet,
                folder_path.name,
            )
        )

    rows = read_image_task(workbook_path, selected_worksheet)
    page_titles = build_page_title_map(rows)
    if not page_titles:
        raise ValueError("工作表没有可执行的图片任务")

    images_by_page = {}
    for path in folder_path.iterdir():
        if not path.is_file() or path.suffix.lower() != ".jpg":
            continue

        prefix = path.name[:3]
        if len(prefix) != 3 or not prefix.isdigit() or int(prefix) < 1:
            raise ValueError("JPG 文件名必须以三位正数页码开头：{}".format(path.name))

        page = int(prefix)
        if page in images_by_page:
            raise ValueError(
                "同一页码匹配到多个 JPG：{}、{}".format(
                    images_by_page[page].name,
                    path.name,
                )
            )
        images_by_page[page] = path

    expected_pages = set(page_titles)
    actual_pages = set(images_by_page)
    missing_pages = sorted(expected_pages - actual_pages)
    if missing_pages:
        raise ValueError(
            "缺少页码对应的 JPG：{}".format(_format_pages(missing_pages))
        )

    extra_pages = sorted(actual_pages - expected_pages)
    if extra_pages:
        raise ValueError(
            "存在工作表未覆盖的 JPG 页码：{}".format(_format_pages(extra_pages))
        )

    plan = []
    target_owners = {}
    source_keys = {_windows_name_key(path.name) for path in images_by_page.values()}
    for page in sorted(expected_pages):
        source = images_by_page[page]
        target = folder_path / build_image_name(page, page_titles[page])
        target_key = _windows_name_key(target.name)

        if target_key in target_owners:
            raise ValueError(
                "多个图片将使用同一目标名称：{}".format(target.name)
            )
        target_owners[target_key] = source

        # 允许目标是本计划中的当前源文件；执行阶段使用临时名隔离互相占用。
        if target.exists() and target_key not in source_keys:
            raise ValueError("目标名称已被占用：{}".format(target))

        plan.append(ImageRenamePlanItem(page, source, target))

    return plan


def rename_images(
    workbook_path: Path,
    worksheet_name: str,
    folder_path: Path,
) -> ImageRenameResult:
    """校验并执行图片重命名，返回适合 UI 展示的完成统计。"""

    plan = build_image_rename_plan(workbook_path, worksheet_name, folder_path)
    return execute_image_rename_plan(plan)


def execute_image_rename_plan(
    plan: Sequence[ImageRenamePlanItem],
) -> ImageRenameResult:
    """执行已经完整校验的图片计划，供单目录和多目录任务复用。"""

    renamed = 0
    unchanged = 0

    for item in plan:
        if _windows_name_key(item.source.name) == _windows_name_key(item.target.name):
            # Windows 文件名不区分大小写；题名相同即视为已经达到目标。
            unchanged += 1
            continue

        temporary = _unused_temporary_path(item.source.parent)
        try:
            item.source.rename(temporary)
            temporary.rename(item.target)
        except OSError as error:
            raise ImageRenameExecutionError(
                failed_path=item.source,
                reason=error,
                result=ImageRenameResult(
                    total=len(plan),
                    renamed=renamed,
                    unchanged=unchanged,
                ),
            ) from error
        renamed += 1

    return ImageRenameResult(
        total=len(plan),
        renamed=renamed,
        unchanged=unchanged,
    )


def _windows_name_key(name: str) -> str:
    """用 Windows 不区分大小写的语义比较同目录文件名。"""

    return name.casefold()


def _unused_temporary_path(folder_path: Path) -> Path:
    """生成同目录临时名，避免任何源文件被直接覆盖。"""

    while True:
        candidate = folder_path / (
            ".__excel_batch_renamer_{}.tmp".format(uuid4().hex)
        )
        if not candidate.exists():
            return candidate


def _format_pages(pages: Sequence[int]) -> str:
    return "、".join("{:03d}".format(page) for page in pages)
