import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from excel_batch_renamer.core.models import FolderTaskRow
from excel_batch_renamer.rename_folders import (
    FolderRenameExecutionError,
    FolderRenameOperation,
    FolderRenamePlan,
    execute_folder_rename_plan,
    plan_folder_renames,
    rename_folders,
)


class RenameFoldersTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temporary_directory.name)

    def tearDown(self):
        self.temporary_directory.cleanup()

    def test_rename_folders_returns_renamed_and_unchanged_counts(self):
        workbook_path = self._save_folder_workbook(
            [
                (1, "规划管理文件材料"),
                (2, None),
            ]
        )
        (self.base_path / "001——").mkdir()
        (self.base_path / "002——").mkdir()

        result = rename_folders(workbook_path, self.base_path)

        self.assertEqual(result.renamed_count, 1)
        self.assertEqual(result.unchanged_count, 1)
        self.assertTrue((self.base_path / "001——规划管理文件材料").is_dir())
        self.assertTrue((self.base_path / "002——").is_dir())

    def test_unchanged_folder_does_not_call_filesystem_rename(self):
        rows = [FolderTaskRow(sequence=1, folder_name="")]
        (self.base_path / "001——").mkdir()

        plan = plan_folder_renames(rows, self.base_path)
        with patch(
            "excel_batch_renamer.rename_folders._rename_path"
        ) as rename_path:
            result = execute_folder_rename_plan(plan)

        rename_path.assert_not_called()
        self.assertEqual(result.renamed_count, 0)
        self.assertEqual(result.unchanged_count, 1)

    def test_blank_name_preserves_an_already_named_folder(self):
        rows = [FolderTaskRow(sequence=1, folder_name="")]
        existing = self.base_path / "001——已有名称"
        existing.mkdir()

        plan = plan_folder_renames(rows, self.base_path)
        with patch(
            "excel_batch_renamer.rename_folders._rename_path"
        ) as rename_path:
            result = execute_folder_rename_plan(plan)

        rename_path.assert_not_called()
        self.assertEqual(result.renamed_count, 0)
        self.assertEqual(result.unchanged_count, 1)
        self.assertTrue(existing.is_dir())
        self.assertFalse((self.base_path / "001——").exists())

    def test_missing_matching_folder_blocks_all_renames(self):
        rows = [
            FolderTaskRow(sequence=1, folder_name="新名称一"),
            FolderTaskRow(sequence=2, folder_name="新名称二"),
        ]
        (self.base_path / "001——旧名称").mkdir()

        with self.assertRaisesRegex(ValueError, "缺少序号 002"):
            plan_folder_renames(rows, self.base_path)

        self.assertTrue((self.base_path / "001——旧名称").is_dir())
        self.assertFalse((self.base_path / "001——新名称一").exists())

    def test_duplicate_task_sequence_blocks_planning(self):
        rows = [
            FolderTaskRow(sequence=1, folder_name="名称一"),
            FolderTaskRow(sequence=1, folder_name="名称二"),
        ]
        (self.base_path / "001——").mkdir()

        with self.assertRaisesRegex(ValueError, "重复序号：001"):
            plan_folder_renames(rows, self.base_path)

    def test_non_contiguous_task_sequence_blocks_planning(self):
        rows = [
            FolderTaskRow(sequence=1, folder_name="名称一"),
            FolderTaskRow(sequence=3, folder_name="名称三"),
        ]
        (self.base_path / "001——").mkdir()
        (self.base_path / "003——").mkdir()

        with self.assertRaisesRegex(ValueError, "缺少连续序号：002"):
            plan_folder_renames(rows, self.base_path)

    def test_duplicate_directory_sequence_blocks_all_renames(self):
        rows = [FolderTaskRow(sequence=1, folder_name="新名称")]
        (self.base_path / "001——旧名称甲").mkdir()
        (self.base_path / "001——旧名称乙").mkdir()

        with self.assertRaisesRegex(ValueError, "匹配到多个文件夹"):
            plan_folder_renames(rows, self.base_path)

    def test_duplicate_targets_block_all_renames(self):
        rows = [
            FolderTaskRow(sequence=1, folder_name="名称一"),
            FolderTaskRow(sequence=2, folder_name="名称二"),
        ]
        (self.base_path / "001——旧名称").mkdir()
        (self.base_path / "002——旧名称").mkdir()

        with patch(
            "excel_batch_renamer.rename_folders.build_folder_name",
            return_value="重复目标",
        ):
            with self.assertRaisesRegex(ValueError, "重复目标"):
                plan_folder_renames(rows, self.base_path)

        self.assertTrue((self.base_path / "001——旧名称").is_dir())
        self.assertTrue((self.base_path / "002——旧名称").is_dir())

    def test_existing_target_file_blocks_all_renames(self):
        rows = [FolderTaskRow(sequence=1, folder_name="新名称")]
        (self.base_path / "001——旧名称").mkdir()
        (self.base_path / "001——新名称").write_text("冲突", encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "目标路径已存在"):
            plan_folder_renames(rows, self.base_path)

        self.assertTrue((self.base_path / "001——旧名称").is_dir())

    def test_non_project_and_nested_folders_are_not_matches(self):
        rows = [FolderTaskRow(sequence=1, folder_name="新名称")]
        (self.base_path / "普通文件夹").mkdir()
        nested_parent = self.base_path / "其他"
        nested_parent.mkdir()
        (nested_parent / "001——嵌套目录").mkdir()

        with self.assertRaisesRegex(ValueError, "缺少序号 001"):
            plan_folder_renames(rows, self.base_path)

    def test_first_filesystem_failure_stops_without_rollback(self):
        source_one = self.base_path / "001——旧名称"
        source_two = self.base_path / "002——旧名称"
        source_three = self.base_path / "003——旧名称"
        source_one.mkdir()
        source_two.mkdir()
        source_three.mkdir()
        operations = (
            FolderRenameOperation(
                source_one,
                self.base_path / "001——新名称",
            ),
            FolderRenameOperation(
                source_two,
                self.base_path / "002——新名称",
            ),
            FolderRenameOperation(
                source_three,
                self.base_path / "003——新名称",
            ),
        )
        plan = FolderRenamePlan(operations=operations, unchanged_count=2)

        def rename_with_second_failure(source, target):
            if source == source_two:
                raise PermissionError("模拟占用")
            source.rename(target)

        with patch(
            "excel_batch_renamer.rename_folders._rename_path",
            side_effect=rename_with_second_failure,
        ):
            with self.assertRaises(FolderRenameExecutionError) as raised:
                execute_folder_rename_plan(plan)

        self.assertEqual(raised.exception.renamed_count, 1)
        self.assertEqual(raised.exception.unchanged_count, 2)
        self.assertEqual(raised.exception.operation.source, source_two)
        self.assertTrue((self.base_path / "001——新名称").is_dir())
        self.assertTrue(source_two.is_dir())
        self.assertTrue(source_three.is_dir())
        self.assertFalse((self.base_path / "003——新名称").exists())

    def _save_folder_workbook(self, rows):
        workbook_path = self.base_path / "folders.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["序号", "文件夹名称"])
        for row in rows:
            worksheet.append(row)
        workbook.save(str(workbook_path))
        workbook.close()
        return workbook_path


if __name__ == "__main__":
    unittest.main()
