import os
import time
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow
from pyexcel_lite.network import (
    CollaborationClient,
    CollaborationEndpoint,
    CollaborationServer,
    cell_update_message,
    local_join_addresses,
    workbook_from_payload,
    workbook_to_payload,
)
from pyexcel_lite.workbook import WorkbookData


class RecordingEndpoint(CollaborationEndpoint):
    def __init__(self):
        super().__init__()
        self._running = True
        self.sent = []

    def send(self, message: dict) -> None:
        self.sent.append(message)


class CollaborationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_until(self, predicate, timeout: float = 4.0) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self.app.processEvents()
            if predicate():
                return
            time.sleep(0.01)
        self.fail("Timed out waiting for socket synchronization")

    def test_workbook_payload_round_trips_cells_styles_and_sheets(self):
        workbook = WorkbookData()
        sheet = workbook.active_sheet
        sheet.set_value(0, 0, "10")
        sheet.set_value(0, 1, "=A1*2")
        sheet.set_style(0, 1, bold=True, fill_color="#fff2cc")
        source = workbook.add_sheet("Source")
        source.set_value(4, 3, "remote")

        payload = workbook_to_payload(workbook)
        restored = workbook_from_payload(payload)

        self.assertEqual(restored.sheet_names(), ["Sheet1", "Source"])
        self.assertEqual(restored.sheets[0].raw_value(0, 0), "10")
        self.assertEqual(restored.sheets[0].raw_value(0, 1), "=A1*2")
        self.assertTrue(restored.sheets[0].get_cell(0, 1).style.bold)
        self.assertEqual(restored.sheets[1].raw_value(4, 3), "remote")

    def test_local_cell_edit_sends_collaboration_message(self):
        window = SpreadsheetWindow()
        endpoint = RecordingEndpoint()
        try:
            window.collaboration = endpoint
            window.current_model.setData(window.current_model.index(0, 0), "42", Qt.EditRole)

            self.assertEqual(len(endpoint.sent), 1)
            self.assertEqual(endpoint.sent[0]["type"], "cell_update")
            self.assertEqual(endpoint.sent[0]["values"][0], {"row": 0, "column": 0, "value": "42"})
        finally:
            window.close()

    def test_remote_cell_update_applies_without_undo_or_echo(self):
        window = SpreadsheetWindow()
        endpoint = RecordingEndpoint()
        try:
            window.collaboration = endpoint
            message = cell_update_message(0, "Sheet1", [(0, 0, "99"), (0, 1, "=A1*2")])

            window.on_collaboration_message(message)

            self.assertEqual(window.current_sheet.raw_value(0, 0), "99")
            self.assertEqual(window.current_model.data(window.current_model.index(0, 1), Qt.DisplayRole), "198")
            self.assertFalse(window.current_model.can_undo())
            self.assertEqual(endpoint.sent, [])
        finally:
            window.close()

    def test_remote_sheet_messages_keep_sheet_list_in_sync(self):
        window = SpreadsheetWindow()
        try:
            window.on_collaboration_message({"type": "sheet_add", "sheet_index": 1, "sheet_name": "Team"})
            self.assertEqual(window.workbook.sheet_names(), ["Sheet1", "Team"])

            window.on_collaboration_message({"type": "sheet_rename", "sheet_index": 1, "sheet_name": "Shared"})
            self.assertEqual(window.workbook.sheet_names(), ["Sheet1", "Shared"])

            window.on_collaboration_message({"type": "sheet_delete", "sheet_index": 1, "sheet_name": "Shared"})
            self.assertEqual(window.workbook.sheet_names(), ["Sheet1"])
        finally:
            window.close()

    def test_local_join_addresses_include_port_for_clients(self):
        addresses = local_join_addresses(9012)

        self.assertTrue(addresses)
        self.assertTrue(all(address.endswith(":9012") for address in addresses))
        self.assertNotIn("0.0.0.0:9012", addresses)

    def test_socket_server_client_windows_sync_cell_edits_both_ways(self):
        server_window = SpreadsheetWindow()
        client_window = SpreadsheetWindow()
        server = CollaborationServer(
            host="127.0.0.1",
            port=0,
            snapshot_provider=lambda: workbook_to_payload(server_window.workbook),
        )
        client = None
        try:
            server_window.current_model.setData(server_window.current_model.index(0, 0), "initial", Qt.EditRole)
            server_window.attach_collaboration(server, "Host")
            server.start()

            client = CollaborationClient("127.0.0.1", server.port)
            client_window.attach_collaboration(client, "Client")
            client.start()

            self.wait_until(lambda: client_window.current_sheet.raw_value(0, 0) == "initial")

            server_window.current_model.setData(server_window.current_model.index(0, 0), "from-server", Qt.EditRole)
            self.wait_until(lambda: client_window.current_sheet.raw_value(0, 0) == "from-server")

            client_window.current_model.setData(client_window.current_model.index(0, 1), "from-client", Qt.EditRole)
            self.wait_until(lambda: server_window.current_sheet.raw_value(0, 1) == "from-client")
        finally:
            if client is not None:
                client.stop()
            server.stop()
            client_window.close()
            server_window.close()


if __name__ == "__main__":
    unittest.main()
