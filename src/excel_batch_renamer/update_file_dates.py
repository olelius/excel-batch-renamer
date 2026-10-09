"""递归修改指定类型文件的创建/修改日期，预检失败零写入，执行失败立即停止。"""

import os
import re
import stat
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Tuple, Union

from excel_batch_renamer.infrastructure.windows_file_time import WindowsFileTime


SUPPORTED_EXTENSIONS = frozenset((".jpg", ".jpeg", ".pdf", ".xls", ".xlsx", ".xlsm", ".xlsb"))


def parse_file_date(text: str) -> date:
    """校验 YYYY-MM-DD 日期；Windows 文件时间从 1601 年开始。"""

    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", text.strip()):
        raise ValueError("请输入 YYYY-MM-DD 格式的目标日期，例如 2026-10-09")
    try:
        value = date.fromisoformat(text.strip())
    except ValueError as error:
        raise ValueError("目标日期无效，请检查年月日（包括闰年和每月天数）") from error
    if value.year < 1601:
        raise ValueError("目标日期年份不能早于 1601 年")
    return value


@dataclass(frozen=True)
class FileDateResult:
    """一次成功任务中匹配、已修改和无需修改的文件数量。"""

    total: int
    updated: int
    unchanged: int

    @property
    def status_text(self) -> str:
        """向当前标签页汇报递归范围及完成数量。"""

        return "完成：含子文件夹共匹配 {} 个文件，已修改 {} 个，未变化 {} 个；创建/修改日期已更新，各自原时分秒不变".format(
            self.total, self.updated, self.unchanged
        )


class FileDateExecutionError(RuntimeError):
    """包含失败步骤、对象、原因和已完成数量，不回滚先前修改。"""

    def __init__(self, path: Path, updated: int, unchanged: int, total: int, cause: Exception):
        self.failed_path = path
        self.updated = updated
        self.unchanged = unchanged
        self.total = total
        self.cause = cause
        super().__init__(
            "写入文件日期失败：{}；已修改 {} 个，未变化 {} 个，共 {} 个；原因：{}".format(
                path, updated, unchanged, total, cause
            )
        )


def _is_reparse(path: Path) -> bool:
    """不跟随符号链接、目录联接等重解析点，避免越出所选目录或产生循环。"""

    attributes = path.lstat()
    return stat.S_ISLNK(attributes.st_mode) or bool(
        getattr(attributes, "st_file_attributes", 0) & 0x400
    )


def collect_date_files(directory: Path) -> Tuple[Path, ...]:
    """确定性扫描全部子文件夹，仅收集支持后缀的普通文件，扫描错误直接上报。"""

    if not directory.is_dir():
        raise ValueError("任务文件夹不存在或不是文件夹：{}".format(directory))
    if _is_reparse(directory):
        raise ValueError("请选择实际文件夹，不使用符号链接或目录联接：{}".format(directory))

    def fail(error):
        raise error

    files = []
    for root, subdirectories, names in os.walk(str(directory), followlinks=False, onerror=fail):
        parent = Path(root)
        subdirectories[:] = sorted(
            (name for name in subdirectories if not _is_reparse(parent / name)),
            key=str.casefold,
        )
        for name in sorted(names, key=str.casefold):
            path = parent / name
            if path.suffix.lower() in SUPPORTED_EXTENSIONS and not _is_reparse(path):
                if stat.S_ISREG(path.lstat().st_mode):
                    files.append(path)
    return tuple(files)


def update_file_dates(directory: Union[str, Path], date_text: str) -> FileDateResult:
    """先校验日期、递归扫描并读取全部文件时间，再写创建/修改日期。

    日期以执行机器的系统本地时区解释，不读取 Excel 内容或图片元数据。
    创建和修改时间各自保留原时分秒；访问时间不变。
    写入前重新读取每个文件，确保保留当前时分秒；失败立即停止且不自动回滚。
    """

    target_date = parse_file_date(date_text)
    files = collect_date_files(Path(directory))
    backend = WindowsFileTime()
    for path in files:
        try:
            with backend.open_file(path) as handle:
                creation, _, modified = backend.read_times(handle)
                backend.replace_date(creation, target_date)
                backend.replace_date(modified, target_date)
        except (OSError, ValueError) as error:
            raise ValueError("预检文件日期失败：{}；未修改任何文件；原因：{}".format(path, error)) from error

    updated = unchanged = 0
    for path in files:
        try:
            with backend.open_file(path) as handle:
                creation, _, modified = backend.read_times(handle)
                target_creation = backend.replace_date(creation, target_date)
                target_modified = backend.replace_date(modified, target_date)
                if creation == target_creation and modified == target_modified:
                    unchanged += 1
                else:
                    backend.set_file_dates(handle, target_creation, target_modified)
                    updated += 1
        except (OSError, ValueError) as error:
            raise FileDateExecutionError(path, updated, unchanged, len(files), error) from error
    return FileDateResult(len(files), updated, unchanged)
