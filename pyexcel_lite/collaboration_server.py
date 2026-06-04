"""Standalone shared-workbook server for LAN collaboration."""

from __future__ import annotations

import argparse
import json
import signal
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QObject, Qt, QTimer, Slot

from .network import (
    DEFAULT_PORT,
    CollaborationServer,
    local_join_addresses,
    workbook_from_payload,
    workbook_to_payload,
)
from .workbook import WorkbookData, WorksheetData


class CollaborationWorkbookServer(QObject):
    """Keep an authoritative workbook behind the socket relay server."""

    def __init__(
        self,
        host: str = "0.0.0.0",
        port: int = DEFAULT_PORT,
        state_path: Path | str | None = "shared_workbook_state.json",
        workbook: WorkbookData | None = None,
    ):
        super().__init__()
        self.state_path = Path(state_path) if state_path else None
        self._lock = threading.RLock()
        self.workbook = workbook or self._load_state() or WorkbookData()
        self.server = CollaborationServer(host=host, port=port, snapshot_provider=self.snapshot_payload)
        self.server.message_received.connect(self.apply_message, Qt.ConnectionType.DirectConnection)

    @property
    def port(self) -> int:
        return self.server.port

    def start(self) -> None:
        self.server.start()
        self.save_state()

    def stop(self) -> None:
        self.save_state()
        self.server.stop()

    def snapshot_payload(self) -> dict:
        with self._lock:
            return workbook_to_payload(self.workbook)

    @Slot(dict)
    def apply_message(self, message: dict) -> None:
        if self.apply_workbook_message(message):
            self.save_state()

    def apply_workbook_message(self, message: dict) -> bool:
        message_type = str(message.get("type") or "")
        if message_type in {"ping", "pong"}:
            return False
        with self._lock:
            if message_type == "snapshot":
                payload = message.get("workbook")
                if not isinstance(payload, dict):
                    return False
                self.workbook = workbook_from_payload(payload)
                return True
            if message_type == "cell_update":
                return self._apply_cell_update(message)
            if message_type == "sheet_add":
                return self._apply_sheet_add(message)
            if message_type == "sheet_rename":
                return self._apply_sheet_rename(message)
            if message_type == "sheet_delete":
                return self._apply_sheet_delete(message)
            if message_type in {"insert_rows", "remove_rows", "insert_columns", "remove_columns"}:
                return self._apply_structure_update(message)
        return False

    def _apply_cell_update(self, message: dict) -> bool:
        sheet = self._resolve_sheet(message, create=True)
        if sheet is None:
            return False
        changed = False
        for item in message.get("values", []):
            row = self._safe_int(item.get("row"), 0)
            column = self._safe_int(item.get("column"), 0)
            changed = sheet.set_value(row, column, item.get("value", ""), touch=False) or changed
        if changed:
            sheet.bump_revision()
        return changed

    def _apply_sheet_add(self, message: dict) -> bool:
        name = str(message.get("sheet_name") or self.workbook.unique_sheet_name())
        if self.workbook.sheet_by_name(name) is not None:
            return False
        index = self._clamped_sheet_insert_index(message.get("sheet_index"))
        self.workbook.sheets.insert(index, WorksheetData(name=name))
        self.workbook.active_sheet_index = min(self.workbook.active_sheet_index, len(self.workbook.sheets) - 1)
        return True

    def _apply_sheet_rename(self, message: dict) -> bool:
        sheet = self._resolve_sheet(message)
        if sheet is None:
            return False
        name = str(message.get("sheet_name") or sheet.name).strip()
        if not name or name == sheet.name:
            return False
        sheet.name = name
        sheet.bump_revision()
        return True

    def _apply_sheet_delete(self, message: dict) -> bool:
        if len(self.workbook.sheets) == 1:
            return False
        sheet = self._resolve_sheet(message)
        if sheet is None:
            return False
        self.workbook.remove_sheet(self.workbook.sheets.index(sheet))
        return True

    def _apply_structure_update(self, message: dict) -> bool:
        sheet = self._resolve_sheet(message)
        if sheet is None:
            return False
        start = max(0, self._safe_int(message.get("start"), 0))
        count = max(1, self._safe_int(message.get("count"), 1))
        message_type = message.get("type")
        if message_type == "insert_rows":
            sheet.insert_rows(start, count)
        elif message_type == "remove_rows":
            sheet.remove_rows(start, count)
        elif message_type == "insert_columns":
            sheet.insert_columns(start, count)
        elif message_type == "remove_columns":
            sheet.remove_columns(start, count)
        else:
            return False
        return True

    def _resolve_sheet(self, message: dict, *, create: bool = False) -> WorksheetData | None:
        sheet_index = self._safe_int(message.get("sheet_index"), -1)
        if 0 <= sheet_index < len(self.workbook.sheets):
            return self.workbook.sheets[sheet_index]
        sheet_name = str(message.get("sheet_name") or "").strip()
        if sheet_name:
            sheet = self.workbook.sheet_by_name(sheet_name)
            if sheet is not None:
                return sheet
        if not create:
            return None
        sheet = WorksheetData(name=sheet_name or self.workbook.unique_sheet_name())
        index = self._clamped_sheet_insert_index(sheet_index)
        self.workbook.sheets.insert(index, sheet)
        return sheet

    def _clamped_sheet_insert_index(self, value) -> int:
        return max(0, min(self._safe_int(value, len(self.workbook.sheets)), len(self.workbook.sheets)))

    def _load_state(self) -> WorkbookData | None:
        if self.state_path is None or not self.state_path.exists():
            return None
        try:
            payload = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        if not isinstance(payload, dict):
            return None
        return workbook_from_payload(payload)

    def save_state(self) -> None:
        if self.state_path is None:
            return
        with self._lock:
            self.state_path.parent.mkdir(parents=True, exist_ok=True)
            self.state_path.write_text(json.dumps(workbook_to_payload(self.workbook), indent=2), encoding="utf-8")

    @staticmethod
    def _safe_int(value, default: int) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return default


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run a PyExcel Lite shared workbook server.")
    parser.add_argument("--host", default="0.0.0.0", help="IP address to bind. Use 0.0.0.0 for the LAN.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="TCP port for clients.")
    parser.add_argument("--state", default="shared_workbook_state.json", help="JSON file used to persist workbook state.")
    parser.add_argument("--memory-only", action="store_true", help="Do not save workbook state to disk.")
    args = parser.parse_args(argv)

    app = QCoreApplication.instance() or QCoreApplication(sys.argv[:1])
    state_path = None if args.memory_only else args.state
    service = CollaborationWorkbookServer(host=args.host, port=args.port, state_path=state_path)
    service.server.status_changed.connect(lambda text: print(text, flush=True))
    service.server.error_occurred.connect(lambda text: print(f"Error: {text}", flush=True))

    try:
        service.start()
    except OSError as exc:
        print(f"Could not start shared server: {exc}", file=sys.stderr)
        return 1

    print("PyExcel Lite shared server is running.", flush=True)
    print(f"Listening on {args.host}:{service.port}", flush=True)
    print("Client addresses:", flush=True)
    for address in local_join_addresses(service.port):
        print(f"  {address}", flush=True)

    def request_stop(*_args) -> None:
        app.quit()

    signal.signal(signal.SIGINT, request_stop)
    signal.signal(signal.SIGTERM, request_stop)
    app.aboutToQuit.connect(service.stop)
    timer = QTimer(app)
    timer.timeout.connect(lambda: None)
    timer.start(250)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
