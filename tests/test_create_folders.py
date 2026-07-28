import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from openpyxl import Workbook

from excel_batch_renamer.create_folders import (
    FolderCreationExecutionError,
    FolderCreationItem,
    FolderCreationPlan,
    build_folder_creation_plan,
    create_folders_from_xlsx,
    execute_folder_creation_plan,
)


class CreateFoldersTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temporary_directory.name)
        self.target_directory = self.base_path / "target"
        self.target_directory.mkdir()

    def tearDown(self):
        self.temporary_directory.cleanup()

    def _write_workbook(self, rows):
        workbook_path = self.base_path / "folders.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["序号", "文件夹名称"])
        for row in rows:
            worksheet.append(row)
        workbook.save(str(workbook_path))
        workbook.close()
        return workbook_path

    def test_create_placeholder_and_named_folders_from_xlsx(self):
        workbook_path = self._write_workbook(
            [
                [1, None],
                [2, "规划管理文件材料"],
            ]
        )

        result = create_folders_from_xlsx(
            workbook_path,
            self.target_directory,
        )

        self.assertTrue((self.target_directory / "001——").is_dir())
        self.assertTrue(
            (self.target_directory / "002——规划管理文件材料").is_dir()
        )
        self.assertEqual(result.planned_count, 2)
        self.assertEqual(result.created_count, 2)
        self.assertEqual(result.status_text, "完成：已创建 2 个文件夹")

    def test_existing_target_rejects_entire_plan_before_creation(self):
        workbook_path = self._write_workbook(
            [
                [1, "第一项"],
                [2, "已存在"],
            ]
        )
        (self.target_directory / "002——已存在").mkdir()

        with self.assertRaisesRegex(ValueError, "目标文件夹已存在"):
            create_folders_from_xlsx(
                workbook_path,
                self.target_directory,
            )

        self.assertFalse((self.target_directory / "001——第一项").exists())

    def test_duplicate_sequence_rejects_entire_plan_before_creation(self):
        workbook_path = self._write_workbook(
            [
                [1, "第一项"],
                [1, "重复项"],
            ]
        )

        with self.assertRaisesRegex(ValueError, "创建计划包含重复序号"):
            create_folders_from_xlsx(
                workbook_path,
                self.target_directory,
            )

        self.assertEqual(list(self.target_directory.iterdir()), [])

    def test_non_contiguous_sequence_is_rejected(self):
        workbook_path = self._write_workbook(
            [
                [1, "第一项"],
                [3, "第三项"],
            ]
        )

        with self.assertRaisesRegex(ValueError, "序号必须从 1 连续排列"):
            build_folder_creation_plan(
                workbook_path,
                self.target_directory,
            )

    def test_target_directory_must_exist(self):
        workbook_path = self._write_workbook([[1, "第一项"]])

        with self.assertRaisesRegex(ValueError, "任务目录不存在"):
            build_folder_creation_plan(
                workbook_path,
                self.base_path / "missing",
            )

    def test_empty_task_returns_zero_statistics(self):
        workbook_path = self._write_workbook([])

        result = create_folders_from_xlsx(
            workbook_path,
            self.target_directory,
        )

        self.assertEqual(result.planned_count, 0)
        self.assertEqual(result.created_count, 0)
        self.assertEqual(result.status_text, "完成：已创建 0 个文件夹")

    def test_runtime_failure_stops_and_reports_partial_progress(self):
        successful_target = Mock()
        successful_target.mkdir.return_value = None
        failed_target = Mock()
        failed_target.mkdir.side_effect = PermissionError("拒绝访问")
        untouched_target = Mock()
        plan = FolderCreationPlan(
            workbook_path=Path("folders.xlsx"),
            target_directory=Path("target"),
            items=(
                FolderCreationItem(1, successful_target),
                FolderCreationItem(2, failed_target),
                FolderCreationItem(3, untouched_target),
            ),
        )

        with self.assertRaises(FolderCreationExecutionError) as context:
            execute_folder_creation_plan(plan)

        error = context.exception
        self.assertIs(error.failed_path, failed_target)
        self.assertEqual(error.created_count, 1)
        self.assertEqual(error.planned_count, 3)
        self.assertIsInstance(error.cause, PermissionError)
        successful_target.mkdir.assert_called_once_with()
        failed_target.mkdir.assert_called_once_with()
        untouched_target.mkdir.assert_not_called()


if __name__ == "__main__":
    unittest.main()
