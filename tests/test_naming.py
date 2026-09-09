import unittest

from excel_batch_renamer.core.naming import (
    build_folder_name,
    build_image_name,
    extract_folder_sequence,
    format_three_digits,
    worksheet_matches_folder,
    worksheet_name_for_folder,
)


class NamingTests(unittest.TestCase):
    def test_format_three_digits(self):
        self.assertEqual(format_three_digits(1), "001")
        self.assertEqual(format_three_digits("12"), "012")

    def test_build_placeholder_folder_name(self):
        self.assertEqual(build_folder_name(1, ""), "001——")
        self.assertEqual(build_folder_name(1, None), "001——")

    def test_build_standard_folder_name(self):
        self.assertEqual(
            build_folder_name(1, "规划管理文件材料"),
            "001——规划管理文件材料",
        )

    def test_extract_folder_sequence(self):
        self.assertEqual(extract_folder_sequence("001——"), 1)
        self.assertEqual(extract_folder_sequence("012——某文件夹"), 12)

    def test_invalid_folder_prefix_is_rejected(self):
        with self.assertRaises(ValueError):
            extract_folder_sequence("1——某文件夹")
        with self.assertRaises(ValueError):
            extract_folder_sequence("000——某文件夹")

    def test_worksheet_binding_and_auto_selection(self):
        self.assertEqual(worksheet_name_for_folder("001——"), "1")
        self.assertEqual(worksheet_name_for_folder("012——某文件夹"), "12")
        self.assertTrue(worksheet_matches_folder("1", "001——"))
        self.assertFalse(worksheet_matches_folder("2", "001——"))
        self.assertFalse(worksheet_matches_folder("项目一", "001——"))

    def test_coded_folder_uses_final_numeric_segment(self):
        """编码目录取末段序号，工作表仍用数字名且保留原三位规则。"""
        for folder_name, sequence in (
            ("I74-6-256——", 256),
            ("I74-6-006——已有名称", 6),
            ("74-6-256——", 256),
            ("001——", 1),
        ):
            with self.subTest(folder_name=folder_name):
                self.assertEqual(extract_folder_sequence(folder_name), sequence)
                self.assertEqual(worksheet_name_for_folder(folder_name), str(sequence))
                self.assertTrue(worksheet_matches_folder(str(sequence), folder_name))

        self.assertFalse(worksheet_matches_folder("I74-6-256", "I74-6-256——"))
        self.assertFalse(worksheet_matches_folder("255", "I74-6-256——"))
        for folder_name in (
            "I74--256——",
            "I74-6-0——",
            "I74-6-XYZ——",
            "I74-6-256",
            "1——",
        ):
            with self.subTest(invalid_folder_name=folder_name):
                with self.assertRaises(ValueError):
                    extract_folder_sequence(folder_name)

    def test_build_image_name(self):
        self.assertEqual(
            build_image_name(1, "项目建议批复文件及项目建议书"),
            "001项目建议批复文件及项目建议书.jpg",
        )


if __name__ == "__main__":
    unittest.main()
