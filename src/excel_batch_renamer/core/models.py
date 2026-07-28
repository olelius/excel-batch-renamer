"""核心业务数据结构。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FolderTaskRow:
    """文件夹任务表中的一条指令。

    Attributes:
        sequence: 任务表中的正整数序号。
        folder_name: 不包含序号和分隔符的文件夹名称，允许为空。
    """

    sequence: int
    folder_name: str


@dataclass(frozen=True)
class ImageTaskRow:
    """图片任务表中参与重命名的一条指令。

    任务表中的“序号”不参与图片命名，因此不保存到该模型。
    """

    file_title: str
    page_reference: str


@dataclass(frozen=True)
class PageRange:
    """文件题名对应的闭区间页次。"""

    file_title: str
    start_page: int
    end_page: int

