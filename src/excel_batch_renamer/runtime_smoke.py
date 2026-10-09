"""便携构建自检：实际调用随包图像和 PDF 依赖，不依赖外部程序。"""

import tempfile
from datetime import date
from pathlib import Path

from PIL import Image

from excel_batch_renamer.infrastructure.pdf_writer import validate_jpeg, write_jpeg_pdf
from excel_batch_renamer.infrastructure.windows_file_time import WindowsFileTime
from excel_batch_renamer.update_file_dates import update_file_dates


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
        nested = root / "子目录"
        nested.mkdir()
        workbook = nested / "创建日期.xlsx"
        workbook.write_bytes(b"timestamp smoke test does not parse workbook content")
        backend = WindowsFileTime()
        paths = images + [target, workbook]
        before = {}
        for path in paths:
            with backend.open_file(path) as handle:
                before[path] = backend.read_times(handle)
        result = update_file_dates(root, "2024-02-29")
        if result.total != 4:
            raise RuntimeError("便携创建日期递归扫描自检失败")
        for path in paths:
            with backend.open_file(path) as handle:
                after = backend.read_times(handle)
            for index in (0, 2):
                original_local = backend.local_time(before[path][index])
                current_local = backend.local_time(after[index])
                if date(current_local.year, current_local.month, current_local.day) != date(2024, 2, 29):
                    raise RuntimeError("便携文件日期写入自检失败")
                for field in ("hour", "minute", "second", "milliseconds"):
                    if getattr(original_local, field) != getattr(current_local, field):
                        raise RuntimeError("便携文件日期未保留原时分秒")
                if before[path][index] % 10000 != after[index] % 10000:
                    raise RuntimeError("便携文件日期意外改变小数秒")
            if before[path][1] != after[1]:
                raise RuntimeError("便携文件日期意外改变访问时间")
