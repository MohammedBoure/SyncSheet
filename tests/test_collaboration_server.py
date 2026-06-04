import json
import tempfile
import unittest
from pathlib import Path

from PySide6.QtCore import QCoreApplication

from pyexcel_lite.collaboration_server import CollaborationWorkbookServer
from pyexcel_lite.network import cell_update_message, sheet_message, structure_message
from pyexcel_lite.project import project_snapshot_message, scan_project_folder


class CollaborationWorkbookServerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QCoreApplication.instance() or QCoreApplication([])

    def test_cell_updates_are_applied_and_persisted(self):
        with tempfile.TemporaryDirectory() as directory:
            state_path = Path(directory) / "state.json"
            service = CollaborationWorkbookServer(state_path=state_path)

            service.apply_message(cell_update_message(0, "Sheet1", [(0, 0, "12"), (0, 1, "=A1*2")]))

            self.assertEqual(service.workbook.sheets[0].raw_value(0, 0), "12")
            self.assertEqual(service.workbook.sheets[0].raw_value(0, 1), "=A1*2")
            payload = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["workbook"]["sheets"][0]["cells"][0]["value"], "12")

            restored = CollaborationWorkbookServer(state_path=state_path)
            self.assertEqual(restored.workbook.sheets[0].raw_value(0, 1), "=A1*2")

    def test_sheet_and_structure_messages_update_authoritative_workbook(self):
        service = CollaborationWorkbookServer(state_path=None)

        service.apply_message(sheet_message("sheet_add", 1, "Team"))
        service.apply_message(cell_update_message(1, "Team", [(2, 0, "moved")]))
        service.apply_message(structure_message("insert_rows", 1, "Team", 1, 2))
        service.apply_message(sheet_message("sheet_rename", 1, "Shared"))

        self.assertEqual(service.workbook.sheet_names(), ["Sheet1", "Shared"])
        self.assertEqual(service.workbook.sheets[1].raw_value(4, 0), "moved")

        service.apply_message(sheet_message("sheet_delete", 1, "Shared"))

        self.assertEqual(service.workbook.sheet_names(), ["Sheet1"])

    def test_project_snapshots_are_saved_and_restored(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "project"
            reports = root / "reports"
            reports.mkdir(parents=True)
            (reports / "team.csv").write_text("a,b\n1,2\n", encoding="utf-8")
            (root / "notes.txt").write_text("notes", encoding="utf-8")
            state_path = Path(directory) / "state.json"
            project = scan_project_folder(root)
            service = CollaborationWorkbookServer(state_path=state_path)

            service.apply_message(project_snapshot_message(project))

            self.assertEqual(service.project.name, "project")
            self.assertEqual(service.project.openable_count, 1)
            payload = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["project"]["files"][0]["relative_path"], "notes.txt")

            restored = CollaborationWorkbookServer(state_path=state_path)
            self.assertEqual(restored.project.file_by_relative_path("reports/team.csv").kind, "csv")


if __name__ == "__main__":
    unittest.main()
