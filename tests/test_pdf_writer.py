"""以真实 JPEG 和独立 PDF 解析器检查页面、原始数据和失败边界。"""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pypdf import PdfReader

from excel_batch_renamer.infrastructure.pdf_writer import (
    validate_jpeg,
    write_jpeg_pdf,
)


class PdfWriterTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.target = self.root / "验收文件.pdf"

    def _jpeg(self, name="006验收文件.jpg", size=(120, 80), mode="RGB", **kwargs):
        """创建可解码的 JPEG，可指定 DPI、EXIF 及颜色模式。"""

        path = self.root / name
        with Image.new(mode, size) as picture:
            picture.save(str(path), "JPEG", **kwargs)
        return path

    def test_five_pages_preserve_input_order_and_original_jpeg_bytes(self):
        paths = [
            self._jpeg("{:03d}验收文件.jpg".format(page), (page * 20, 100))
            for page in range(6, 11)
        ]
        originals = [path.read_bytes() for path in paths]

        write_jpeg_pdf(paths, self.target, "验收文件")

        reader = PdfReader(self.target)
        self.assertEqual(reader.metadata.title, "验收文件")
        self.assertEqual(len(reader.pages), 5)
        for page, original, path in zip(reader.pages, originals, paths):
            objects = page["/Resources"]["/XObject"].get_object()
            self.assertEqual(len(objects), 1)
            picture = next(iter(objects.values())).get_object()
            self.assertIn("/DCTDecode", picture["/Filter"])
            self.assertEqual(picture.get_data(), original)
            self.assertEqual(path.read_bytes(), original)
        self.assertEqual([float(page.mediabox.width) for page in reader.pages],
                         [120, 140, 160, 180, 200])

    def test_dpi_controls_page_size_without_resampling(self):
        path = self._jpeg(size=(600, 900), dpi=(300, 300))
        write_jpeg_pdf([path], self.target, "验收文件")
        page = PdfReader(self.target).pages[0]
        self.assertAlmostEqual(float(page.mediabox.width), 144)
        self.assertAlmostEqual(float(page.mediabox.height), 216)

    def test_all_exif_orientations_use_expected_page_size_and_matrix(self):
        matrices = {
            1: "1 0 0 1 0 0 cm", 2: "-1 0 0 1 120 0 cm",
            3: "-1 0 0 -1 120 80 cm", 4: "1 0 0 -1 0 80 cm",
            5: "0 -1 -1 0 80 120 cm", 6: "0 -1 1 0 0 120 cm",
            7: "0 1 1 0 0 0 cm", 8: "0 1 -1 0 80 0 cm",
        }
        for orientation in range(1, 9):
            with self.subTest(orientation=orientation):
                exif = Image.Exif()
                exif[274] = orientation
                path = self._jpeg(exif=exif)
                write_jpeg_pdf([path], self.target, "验收文件")
                page = PdfReader(self.target).pages[0]
                size = (120, 80) if orientation < 5 else (80, 120)
                self.assertEqual((float(page.mediabox.width),
                                  float(page.mediabox.height)), size)
                self.assertIn(matrices[orientation].encode(), page.get_contents().get_data())

    def test_grayscale_cmyk_and_progressive_jpeg_are_supported(self):
        for mode in ("L", "RGB", "CMYK"):
            with self.subTest(mode=mode):
                path = self._jpeg(mode=mode, progressive=True)
                validate_jpeg(path)
                write_jpeg_pdf([path], self.target, "验收文件")
                self.assertEqual(len(PdfReader(self.target).pages), 1)

    def test_invalid_or_disguised_image_is_rejected_with_source_path(self):
        path = self.root / "006.jpg"
        path.write_bytes(b"not a jpeg")
        with self.assertRaisesRegex(ValueError, "读取 JPG 失败.*006.jpg"):
            validate_jpeg(path)
        with Image.new("RGB", (10, 10)) as picture:
            picture.save(str(path), "PNG")
        with self.assertRaisesRegex(ValueError, "文件内容不是 JPEG"):
            validate_jpeg(path)

    def test_truncated_jpeg_is_rejected_before_execution(self):
        path = self._jpeg(size=(1000, 1000))
        path.write_bytes(path.read_bytes()[:-1000])
        with self.assertRaisesRegex(ValueError, "读取 JPG 失败"):
            validate_jpeg(path)

    def test_repeat_replaces_pdf_with_current_pages(self):
        paths = [self._jpeg("006.jpg"), self._jpeg("007.jpg")]
        write_jpeg_pdf(paths, self.target, "验收文件")
        write_jpeg_pdf(paths[:1], self.target, "验收文件")
        self.assertEqual(len(PdfReader(self.target).pages), 1)
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_failed_write_keeps_existing_pdf_and_cleans_temporary(self):
        path = self._jpeg()
        self.target.write_bytes(b"existing pdf")
        with patch("excel_batch_renamer.infrastructure.pdf_writer.Canvas.save",
                   side_effect=OSError("磁盘空间不足")):
            with self.assertRaisesRegex(OSError, "磁盘空间不足"):
                write_jpeg_pdf([path], self.target, "验收文件")
        self.assertEqual(self.target.read_bytes(), b"existing pdf")
        self.assertTrue(path.exists())
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_failed_replace_keeps_existing_pdf_and_cleans_temporary(self):
        path = self._jpeg()
        self.target.write_bytes(b"existing pdf")
        with patch.object(Path, "replace", side_effect=PermissionError("PDF 正在使用")):
            with self.assertRaisesRegex(PermissionError, "PDF 正在使用"):
                write_jpeg_pdf([path], self.target, "验收文件")
        self.assertEqual(self.target.read_bytes(), b"existing pdf")
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_empty_input_creates_no_pdf(self):
        with self.assertRaisesRegex(ValueError, "至少需要一张"):
            write_jpeg_pdf([], self.target, "验收文件")
        self.assertFalse(self.target.exists())


if __name__ == "__main__":
    unittest.main()
