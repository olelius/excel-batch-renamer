"""图片任务表页次区间推导规则。"""

import re
from typing import List, Optional, Tuple

from .models import ImageTaskRow, PageRange


_PAGE_REFERENCE_PATTERN = re.compile(
    r"^\s*(?P<start>\d+)\s*(?:-\s*(?P<end>\d+)\s*)?$"
)


def parse_page_reference(page_reference: str) -> Tuple[int, Optional[int]]:
    """解析单独页码或“起始页-结束页”格式。"""

    match = _PAGE_REFERENCE_PATTERN.match(str(page_reference))
    if match is None:
        raise ValueError("页次必须是单独页码或“起始页-结束页”")

    start_page = int(match.group("start"))
    end_text = match.group("end")
    end_page = int(end_text) if end_text is not None else None
    if start_page < 1 or (end_page is not None and end_page < start_page):
        raise ValueError("页次范围必须从正数开始且结束页不小于起始页")
    return start_page, end_page


def derive_page_ranges(rows: List[ImageTaskRow]) -> List[PageRange]:
    """按相邻起始页和末行显式范围推导闭区间页次。"""

    if not rows:
        return []

    parsed = [parse_page_reference(row.page_reference) for row in rows]
    ranges = []

    for index, row in enumerate(rows):
        start_page, explicit_end = parsed[index]
        if index + 1 < len(rows):
            next_start = parsed[index + 1][0]
            if next_start <= start_page:
                raise ValueError("页次起始页必须严格递增")

            inferred_end = next_start - 1
            if explicit_end is not None and explicit_end != inferred_end:
                raise ValueError("非末行显式结束页必须与下一行起始页连续")
            end_page = inferred_end
        else:
            if explicit_end is None:
                raise ValueError("最后一行页次必须使用“起始页-结束页”格式")
            end_page = explicit_end

        ranges.append(
            PageRange(
                file_title=row.file_title,
                start_page=start_page,
                end_page=end_page,
            )
        )

    return ranges


def build_page_title_map(rows: List[ImageTaskRow]) -> dict:
    """展开页次区间，返回图片页码到文件题名的映射。"""

    page_titles = {}
    for page_range in derive_page_ranges(rows):
        for page in range(page_range.start_page, page_range.end_page + 1):
            page_titles[page] = page_range.file_title
    return page_titles

