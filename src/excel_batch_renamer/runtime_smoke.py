"""便携构建自检：实际调用随包图像和 PDF 依赖，不依赖外部程序。"""

import tempfile
from pathlib import Path

from PIL import Image

from excel_batch_renamer.infrastructure.pdf_writer import validate_jpeg, write_jpeg_pdf


def check_pdf_runtime() -> None:
    """在临时目录生成两张 JPEG 和一个 PDF，检查冻结依赖实际可用。"""

    with tempfile.TemporaryDirectory(prefix="excel-batch-pdf-smoke-") as temporary:
        root = Path(temporary)
        images = []
        for page, color in enumerate(("red", "blue"), start=1):
            path = root / "{:03d}.jpg".format(page)
            with Image.new("RGB", (32, 48), color=color) as picture:
                picture.save(str(path), "JPEG")
            validate_jpeg(path)
            images.append(path)
        target = root / "验收文件.pdf"
        write_jpeg_pdf(images, target, "验收文件")
        data = target.read_bytes()
        if not (data.startswith(b"%PDF-") and data.rstrip().endswith(b"%%EOF")
                and b"/Count 2" in data and b"/DCTDecode" in data):
            raise RuntimeError("便携 PDF 生成自检失败")
