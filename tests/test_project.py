import tempfile
import unittest
from pathlib import Path

from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow
from pyexcel_lite.project import project_from_payload, project_to_payload, scan_project_folder
from pyexcel_lite.settings import StartupSettings


class ProjectDataTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_scan_project_folder_keeps_nested_structure_and_openable_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "Client Work"
            nested = root / "2026" / "June"
            nested.mkdir(parents=True)
            (nested / "budget.xlsx").write_text("placeholder", encoding="utf-8")
            (root / "raw.csv").write_text("a,b\n1,2\n", encoding="utf-8")
            (root / "readme.md").write_text("# notes", encoding="utf-8")

            project = scan_project_folder(root)

            self.assertEqual(project.name, "Client Work")
            self.assertIn("2026", project.folders)
            self.assertIn("2026/June", project.folders)
            self.assertEqual(project.openable_count, 2)
            self.assertEqual(project.file_by_relative_path("2026/June/budget.xlsx").kind, "workbook")
            self.assertFalse(project.file_by_relative_path("readme.md").openable)

    def test_project_payload_round_trip_marks_remote_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "Team"
            root.mkdir()
            (root / "sheet.csv").write_text("x\n1\n", encoding="utf-8")

            payload = project_to_payload(scan_project_folder(root))
            restored = project_from_payload(payload, remote=True)

            self.assertTrue(restored.remote)
            self.assertEqual(restored.name, "Team")
            self.assertEqual(restored.file_by_relative_path("sheet.csv").kind, "csv")

    def test_window_reopens_last_project_from_startup_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "Last Project"
            root.mkdir()
            (root / "sheet.csv").write_text("x\n1\n", encoding="utf-8")
            window = SpreadsheetWindow()
            try:
                window.startup_settings = StartupSettings(last_project_path=str(root))
                window.apply_last_project_settings()

                self.assertEqual(window.project.name, "Last Project")
                self.assertEqual(window.startup_settings.last_project_path, str(root))
                self.assertIsNotNone(window.project.file_by_relative_path("sheet.csv"))
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()
