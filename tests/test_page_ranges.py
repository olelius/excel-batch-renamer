import unittest

from excel_batch_renamer.core.models import ImageTaskRow, PageRange
from excel_batch_renamer.core.page_ranges import (
    build_page_title_map,
    derive_page_ranges,
    parse_page_reference,
)


class PageRangeTests(unittest.TestCase):
    def setUp(self):
        self.rows = [
            ImageTaskRow(file_title="文件A", page_reference="001"),
            ImageTaskRow(file_title="文件B", page_reference="005"),
            ImageTaskRow(file_title="文件C", page_reference="010-017"),
        ]

    def test_parse_single_and_explicit_range(self):
        self.assertEqual(parse_page_reference("001"), (1, None))
        self.assertEqual(parse_page_reference("010-017"), (10, 17))

    def test_derive_ranges_from_adjacent_start_pages(self):
        self.assertEqual(
            derive_page_ranges(self.rows),
            [
                PageRange("文件A", 1, 4),
                PageRange("文件B", 5, 9),
                PageRange("文件C", 10, 17),
            ],
        )

    def test_expand_page_title_map(self):
        page_titles = build_page_title_map(self.rows)
        self.assertEqual(page_titles[1], "文件A")
        self.assertEqual(page_titles[4], "文件A")
        self.assertEqual(page_titles[5], "文件B")
        self.assertEqual(page_titles[9], "文件B")
        self.assertEqual(page_titles[10], "文件C")
        self.assertEqual(page_titles[17], "文件C")
        self.assertEqual(len(page_titles), 17)

    def test_last_row_requires_explicit_range(self):
        with self.assertRaisesRegex(ValueError, "最后一行"):
            derive_page_ranges(
                [ImageTaskRow(file_title="文件A", page_reference="001")]
            )

    def test_start_pages_must_increase(self):
        with self.assertRaisesRegex(ValueError, "严格递增"):
            derive_page_ranges(
                [
                    ImageTaskRow(file_title="文件A", page_reference="005"),
                    ImageTaskRow(file_title="文件B", page_reference="003-010"),
                ]
            )


if __name__ == "__main__":
    unittest.main()

