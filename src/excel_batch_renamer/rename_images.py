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
from excel_batch_renamer.infrastructure.pdf_writer import (
    validate_jpeg,
    write_jpeg_pdf,
)
from excel_batch_renamer.infrastructure.xlsx_reader import (
    list_worksheet_names,
    read_image_task,
)


@dataclass(frozen=True)
class ImageRenamePlanItem:
    """一张图片的路径、页码和本次 Excel 题名，供改名与 PDF 分组复用。"""

    page: int
    source: Path
    target: Path
    file_title: str


@dataclass(frozen=True)
class ImageRenameResult:
    """一次图片改名及自动生成 PDF 任务的统计结果。"""

    total: int
    renamed: int
    unchanged: int
    generated_pdfs: int = 0


class ImageRenameExecutionError(RuntimeError):
    """改名或 PDF 生成失败，并携带失败步骤、对象与已完成统计。"""

    def __init__(
        self,
        failed_path: Path,
        reason: BaseException,
        result: ImageRenameResult,
        operation: str = "重命名图片",
    ) -> None:
        self.failed_path = failed_path
        self.reason = reason
        self.result = result
        self.operation = operation
        super().__init__(
            (
                "{}失败：{}；原因：{}；已重命名 {} 张，未变化 {} 张，"
                "已生成 PDF {} 个（保存在原图片文件夹）"
            ).format(
                operation,
                failed_path,
                reason,
                result.renamed,
                result.unchanged,
                result.generated_pdfs,
            )
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

        plan.append(ImageRenamePlanItem(page, source, target, page_titles[page]))

    _validate_pdf_inputs(plan)
    return plan


def rename_images(
    workbook_path: Path,
    worksheet_name: str,
    folder_path: Path,
) -> ImageRenameResult:
    """校验、重命名 JPG 并按题名自动生成 PDF，保留原图片。"""

    plan = build_image_rename_plan(workbook_path, worksheet_name, folder_path)
    return execute_image_rename_plan(plan)


def execute_image_rename_plan(
    plan: Sequence[ImageRenamePlanItem],
) -> ImageRenameResult:
    """执行已完整校验的图片计划，再从本次改名后路径生成同题名 PDF。

    同一题名在不同 Excel 行中出现时仍按页码升序归入同一个 PDF。
    每次重跑都重新生成同名 PDF，不依赖执行记录或删除旧题名的 PDF。
    """

    renamed = 0
    unchanged = 0
    generated_pdfs = 0

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

    # 题名保留在本次计划中，不重新读取工作簿，也不从已有 PDF 推断分组。
    groups = {}
    for item in sorted(plan, key=lambda planned: planned.page):
        groups.setdefault(item.file_title, []).append(item)
    for title, items in groups.items():
        target_pdf = _pdf_target_path(items)
        image_paths = [
            item.source
            if _windows_name_key(item.source.name) == _windows_name_key(item.target.name)
            else item.target
            for item in items
        ]
        try:
            write_jpeg_pdf(image_paths, target_pdf, title)
        except Exception as error:
            raise ImageRenameExecutionError(
                failed_path=target_pdf,
                reason=error,
                result=ImageRenameResult(
                    total=len(plan),
                    renamed=renamed,
                    unchanged=unchanged,
                    generated_pdfs=generated_pdfs,
                ),
                operation="生成 PDF",
            ) from error
        generated_pdfs += 1

    return ImageRenameResult(
        total=len(plan),
        renamed=renamed,
        unchanged=unchanged,
        generated_pdfs=generated_pdfs,
    )


def _validate_pdf_inputs(plan: Sequence[ImageRenamePlanItem]) -> None:
    """首次改名前检查 PDF 名称占用、题名大小写歧义及每张源 JPG。

    已存在的同名普通文件由 PDF 适配器原子更新；同名目录直接阻止任务。
    批量服务在执行之前构建全部文件夹计划，因此后续文件夹中的坏图也
    会阻止整批任务改动前面的文件夹。
    """

    groups = {}
    title_keys = {}
    for item in plan:
        title_key = _windows_name_key(item.file_title)
        if title_key in title_keys and title_keys[title_key] != item.file_title:
            raise ValueError(
                "文件题名存在 Windows 同名歧义：{}、{}".format(
                    title_keys[title_key],
                    item.file_title,
                )
            )
        title_keys[title_key] = item.file_title
        groups.setdefault(item.file_title, []).append(item)

    target_titles = {}
    for title, items in groups.items():
        target = _pdf_target_path(items)
        target_key = _windows_name_key(target.name)
        if target_key in target_titles:
            if target_titles[target_key] != title:
                raise ValueError(
                    "文件题名存在 Windows 同名歧义：{}、{}".format(
                        target_titles[target_key], title
                    )
                )
            continue
        target_titles[target_key] = title
        if target.exists() and not target.is_file():
            raise ValueError("PDF 目标名称已被目录占用：{}".format(target))

    for item in plan:
        try:
            validate_jpeg(item.source)
        except (OSError, ValueError) as error:
            raise ValueError(
                "JPG 图片校验失败：{}；原因：{}".format(item.source, error)
            ) from error


def _pdf_target_path(items: Sequence[ImageRenamePlanItem]) -> Path:
    """使用同题名全部页中的最小页码作为 PDF 起始页前缀。"""

    first = min(items, key=lambda item: item.page)
    return first.target.parent / "{:03d}{}.pdf".format(
        first.page,
        first.file_title,
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
