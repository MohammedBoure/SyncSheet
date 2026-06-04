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
    normalize_workbook_id,
    workbook_from_payload,
    workbook_to_payload,
)
from .io_xlsx import load_csv, load_xlsx
from .project import ProjectData, project_from_payload, project_to_payload
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
        state_workbook, state_project, state_workbooks = self._load_state()
        self.workbook = workbook or state_workbook or WorkbookData()
        self.project = state_project or ProjectData()
        self.workbooks = state_workbooks
        self.server = CollaborationServer(host=host, port=port, snapshot_message_provider=self.snapshot_message)
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

    def snapshot_message(self) -> dict:
        with self._lock:
            message = {
                "type": "snapshot",
                "workbook": workbook_to_payload(self.workbook),
                "project": project_to_payload(self.project),
            }
            if self.workbooks:
                message["workbooks"] = {
                    workbook_id: workbook_to_payload(workbook)
                    for workbook_id, workbook in sorted(self.workbooks.items())
                }
            return message

    @Slot(dict)
    def apply_message(self, message: dict) -> None:
        if self.respond_to_workbook_request(message):
            return
        if self.apply_workbook_message(message):
            self.save_state()

    def apply_workbook_message(self, message: dict) -> bool:
        message_type = str(message.get("type") or "")
        if message_type in {"ping", "pong"}:
            return False
        with self._lock:
            if message_type == "snapshot":
                payload = message.get("workbook")
                changed = self._apply_project_workbooks_payload(message.get("workbooks"))
                if isinstance(payload, dict):
                    workbook = workbook_from_payload(payload)
                    workbook_id = normalize_workbook_id(message.get("workbook_id"))
                    if workbook_id:
                        self.workbooks[workbook_id] = workbook
                    else:
                        self.workbook = workbook
                    changed = True
                if isinstance(message.get("project"), dict):
                    self.project = project_from_payload(message.get("project"))
                    changed = True
                return changed
            if message_type == "project_snapshot":
                payload = message.get("project")
                changed = self._apply_project_workbooks_payload(message.get("workbooks"))
                if isinstance(payload, dict):
                    self.project = project_from_payload(payload)
                    changed = True
                return changed
            if message_type == "workbook_request":
                return False
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

    def respond_to_workbook_request(self, message: dict) -> bool:
        if str(message.get("type") or "") != "workbook_request":
            return False
        workbook_id = normalize_workbook_id(message.get("workbook_id"))
        if not workbook_id:
            return True
        workbook = self._load_project_workbook_for_request(workbook_id)
        if workbook is None:
            return True
        self.server.send(
            {
                "type": "snapshot",
                "workbook_id": workbook_id,
                "workbook": workbook_to_payload(workbook),
                "project": project_to_payload(self.project),
            }
        )
        return True

    def _apply_cell_update(self, message: dict) -> bool:
        workbook = self._resolve_workbook(message, create=True)
        sheet = self._resolve_sheet(workbook, message, create=True)
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
        workbook = self._resolve_workbook(message, create=True)
        name = str(message.get("sheet_name") or workbook.unique_sheet_name())
        if workbook.sheet_by_name(name) is not None:
            return False
        index = self._clamped_sheet_insert_index(workbook, message.get("sheet_index"))
        workbook.sheets.insert(index, WorksheetData(name=name))
        workbook.active_sheet_index = min(workbook.active_sheet_index, len(workbook.sheets) - 1)
        return True

    def _apply_sheet_rename(self, message: dict) -> bool:
        workbook = self._resolve_workbook(message)
        if workbook is None:
            return False
        sheet = self._resolve_sheet(workbook, message)
        if sheet is None:
            return False
        name = str(message.get("sheet_name") or sheet.name).strip()
        if not name or name == sheet.name:
            return False
        sheet.name = name
        sheet.bump_revision()
        return True

    def _apply_sheet_delete(self, message: dict) -> bool:
        workbook = self._resolve_workbook(message)
        if workbook is None:
            return False
        if len(workbook.sheets) == 1:
            return False
        sheet = self._resolve_sheet(workbook, message)
        if sheet is None:
            return False
        workbook.remove_sheet(workbook.sheets.index(sheet))
        return True

    def _apply_structure_update(self, message: dict) -> bool:
        workbook = self._resolve_workbook(message, create=True)
        if workbook is None:
            return False
        sheet = self._resolve_sheet(workbook, message)
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

    def _resolve_workbook(self, message: dict, *, create: bool = False) -> WorkbookData | None:
        workbook_id = normalize_workbook_id(message.get("workbook_id"))
        if not workbook_id:
            return self.workbook
        if workbook_id not in self.workbooks and create:
            self.workbooks[workbook_id] = WorkbookData()
        return self.workbooks.get(workbook_id)

    def _load_project_workbook_for_request(self, workbook_id: str) -> WorkbookData | None:
        if workbook_id in self.workbooks:
            return self.workbooks[workbook_id]
        project_file = self.project.file_by_relative_path(workbook_id)
        if project_file is None or not project_file.openable:
            return None
        path = self.project.absolute_path_for(project_file)
        if path is None or not path.exists():
            return None
        try:
            workbook = load_xlsx(path) if project_file.extension == ".xlsx" else load_csv(path)
        except Exception:
            return None
        self.workbooks[workbook_id] = workbook
        self.save_state()
        return workbook

    def _resolve_sheet(self, workbook: WorkbookData, message: dict, *, create: bool = False) -> WorksheetData | None:
        sheet_index = self._safe_int(message.get("sheet_index"), -1)
        if 0 <= sheet_index < len(workbook.sheets):
            return workbook.sheets[sheet_index]
        sheet_name = str(message.get("sheet_name") or "").strip()
        if sheet_name:
            sheet = workbook.sheet_by_name(sheet_name)
            if sheet is not None:
                return sheet
        if not create:
            return None
        sheet = WorksheetData(name=sheet_name or workbook.unique_sheet_name())
        index = self._clamped_sheet_insert_index(workbook, sheet_index)
        workbook.sheets.insert(index, sheet)
        return sheet

    def _clamped_sheet_insert_index(self, workbook: WorkbookData, value) -> int:
        return max(0, min(self._safe_int(value, len(workbook.sheets)), len(workbook.sheets)))

    def _apply_project_workbooks_payload(self, payload) -> bool:
        if not isinstance(payload, dict):
            return False
        changed = False
        for raw_workbook_id, workbook_payload in payload.items():
            workbook_id = normalize_workbook_id(raw_workbook_id)
            if not workbook_id or not isinstance(workbook_payload, dict):
                continue
            self.workbooks[workbook_id] = workbook_from_payload(workbook_payload)
            changed = True
        return changed

    def _load_state(self) -> tuple[WorkbookData | None, ProjectData | None, dict[str, WorkbookData]]:
        if self.state_path is None or not self.state_path.exists():
            return None, None, {}
        try:
            payload = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None, None, {}
        if not isinstance(payload, dict):
            return None, None, {}
        workbooks = {
            normalize_workbook_id(workbook_id): workbook_from_payload(workbook_payload)
            for workbook_id, workbook_payload in payload.get("workbooks", {}).items()
            if normalize_workbook_id(workbook_id) and isinstance(workbook_payload, dict)
        } if isinstance(payload.get("workbooks"), dict) else {}
        if isinstance(payload.get("workbook"), dict):
            workbook = workbook_from_payload(payload["workbook"])
            project = project_from_payload(payload.get("project")) if isinstance(payload.get("project"), dict) else None
            return workbook, project, workbooks
        return workbook_from_payload(payload), None, workbooks

    def save_state(self) -> None:
        if self.state_path is None:
            return
        with self._lock:
            self.state_path.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "workbook": workbook_to_payload(self.workbook),
                "project": project_to_payload(self.project),
                "workbooks": {
                    workbook_id: workbook_to_payload(workbook)
                    for workbook_id, workbook in sorted(self.workbooks.items())
                },
            }
            self.state_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

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
