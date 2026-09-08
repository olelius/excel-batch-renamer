"""将有序 JPG 原始数据嵌入 PDF，隔离图像、排版及文件写入依赖。"""

import logging
import math
import tempfile
from pathlib import Path
from typing import Sequence, Tuple

from PIL import Image
from reportlab.pdfgen.canvas import Canvas


LOGGER = logging.getLogger(__name__)


def validate_jpeg(path: Path) -> None:
    """在改名前验证真实 JPEG 及完整解码，失败信息包含源图片路径。"""

    try:
        with Image.open(str(path)) as picture:
            if picture.format != "JPEG":
                raise ValueError("文件内容不是 JPEG 图片")
            picture.verify()
        # verify 主要检查结构；load 另外检查截断的 JPEG 像素数据。
        with Image.open(str(path)) as picture:
            picture.load()
    except (OSError, ValueError, SyntaxError, Image.DecompressionBombError) as error:
        raise ValueError("读取 JPG 失败：{}；原因：{}".format(path, error)) from error


def _page_geometry(path: Path) -> Tuple[float, float, int]:
    """返回旋转前的页面点尺寸和 EXIF 方向；无有效 DPI 时使用 72。"""

    with Image.open(str(path)) as picture:
        width, height = picture.size
        dpi = picture.info.get("dpi", (72, 72))
        try:
            x_dpi, y_dpi = float(dpi[0]), float(dpi[1])
            if not all(math.isfinite(value) and value > 0 for value in (x_dpi, y_dpi)):
                raise ValueError("无效 DPI")
        except (TypeError, ValueError, IndexError, OverflowError):
            x_dpi = y_dpi = 72.0
        orientation = picture.getexif().get(274, 1)
        if orientation not in range(1, 9):
            orientation = 1
        return width * 72.0 / x_dpi, height * 72.0 / y_dpi, orientation


def _draw_jpeg_page(canvas: Canvas, path: Path) -> None:
    """按 EXIF 方向放置原始 JPEG，一张图片一页，不重新压缩图片。"""

    width, height, orientation = _page_geometry(path)
    page_size = (height, width) if orientation >= 5 else (width, height)
    # PDF 原点位于左下角；矩阵在该坐标系实现 EXIF 的八种旋转/镜像。
    transforms = {
        1: (1, 0, 0, 1, 0, 0),
        2: (-1, 0, 0, 1, width, 0),
        3: (-1, 0, 0, -1, width, height),
        4: (1, 0, 0, -1, 0, height),
        5: (0, -1, -1, 0, height, width),
        6: (0, -1, 1, 0, 0, width),
        7: (0, 1, 1, 0, 0, 0),
        8: (0, 1, -1, 0, height, 0),
    }
    canvas.setPageSize(page_size)
    canvas.saveState()
    canvas.transform(*transforms[orientation])
    # 传文件路径而非转换后的位图，ReportLab 直接采用 JPEG DCT 数据流。
    canvas.drawImage(str(path), 0, 0, width=width, height=height)
    canvas.restoreState()
    canvas.showPage()


def write_jpeg_pdf(image_paths: Sequence[Path], target_path: Path, title: str) -> None:
    """按调用方给定页序生成 PDF，完整写入后才替换同名目标。

    输入必须来自已校验的图片计划。本函数保留全部 JPG；失败时清理未完成的
    临时 PDF，保留已有同名 PDF，不撤销此前完成的图片改名或其他 PDF。
    """

    if not image_paths:
        raise ValueError("PDF 至少需要一张 JPG 图片")
    target_path = Path(target_path)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w+b", prefix=".__excel_batch_pdf_", suffix=".tmp",
            dir=str(target_path.parent), delete=False,
        ) as stream:
            temporary_path = Path(stream.name)
            canvas = Canvas(stream, pageCompression=1)
            canvas.setTitle(title)
            canvas.setCreator("ExcelBatchRenamer")
            for image_path in image_paths:
                try:
                    _draw_jpeg_page(canvas, Path(image_path))
                except Exception as error:
                    raise ValueError(
                        "生成 PDF 时读取图片失败：{}；原因：{}".format(image_path, error)
                    ) from error
            canvas.save()
        temporary_path.replace(target_path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                LOGGER.exception("清理未完成 PDF 临时文件失败：%s", temporary_path)
