"""文件夹、图片名称和工作表绑定规则。"""

import re
from typing import Union


FOLDER_NAME_SEPARATOR = "——"
_FOLDER_PREFIX_PATTERN = re.compile(r"^(?P<sequence>\d{3})——")


def _as_positive_integer(value: Union[int, str], field_name: str) -> int:
    """将规范输入转换为正整数，并为编程错误提供明确异常。"""

    if isinstance(value, bool):
        raise ValueError("{}必须是正整数".format(field_name))

    text = str(value).strip()
    if not text.isdigit():
        raise ValueError("{}必须是正整数".format(field_name))

    number = int(text)
    if number < 1:
        raise ValueError("{}必须是正整数".format(field_name))
    return number


def format_three_digits(value: Union[int, str]) -> str:
    """将序号或页码按至少三位、左侧补零的形式输出。"""

    return "{:03d}".format(_as_positive_integer(value, "序号或页码"))


def build_folder_name(sequence: Union[int, str], folder_name: str = "") -> str:
    """生成标准文件夹名或占位文件夹名。"""

    name = "" if folder_name is None else str(folder_name)
    return "{}{}{}".format(
        format_three_digits(sequence),
        FOLDER_NAME_SEPARATOR,
        name,
    )


def extract_folder_sequence(folder_name: str) -> int:
    """从本项目文件夹名中提取三位序号。"""

    match = _FOLDER_PREFIX_PATTERN.match(str(folder_name))
    if match is None:
        raise ValueError("文件夹名称必须以三位序号和两个中文破折号开头")
    sequence = int(match.group("sequence"))
    if sequence < 1:
        raise ValueError("文件夹序号必须是正整数")
    return sequence


def worksheet_name_for_folder(folder_name: str) -> str:
    """返回文件夹自动匹配的、不补零数字工作表名称。"""

    return str(extract_folder_sequence(folder_name))


def worksheet_matches_folder(worksheet_name: str, folder_name: str) -> bool:
    """判断数字工作表名称是否与文件夹前三位序号绑定。"""

    text = str(worksheet_name).strip()
    if not text.isdigit():
        return False
    return int(text) == extract_folder_sequence(folder_name)


def build_image_name(
    page: Union[int, str],
    file_title: str,
    extension: str = ".jpg",
) -> str:
    """生成“图片页码+文件题名+扩展名”的目标图片名。"""

    suffix = extension if extension.startswith(".") else ".{}".format(extension)
    return "{}{}{}".format(format_three_digits(page), file_title, suffix)
