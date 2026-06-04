"""Lightweight realtime collaboration over TCP JSON messages."""

from __future__ import annotations

import json
import socket
import threading
import time
from dataclasses import asdict
from typing import Callable

from PySide6.QtCore import QObject, Signal

from .workbook import CellData, CellStyle, WorkbookData, WorksheetData

DEFAULT_PORT = 8765
PROTOCOL_VERSION = 1
RECONNECT_DELAY_SECONDS = 1.5
HEARTBEAT_INTERVAL_SECONDS = 10.0
STYLE_FIELDS = set(CellStyle.__dataclass_fields__)
PING_MESSAGE = {"type": "ping"}
PONG_MESSAGE = {"type": "pong"}


def configure_collaboration_socket(sock: socket.socket) -> None:
    sock.settimeout(None)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
    if hasattr(socket, "TCP_NODELAY"):
        sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)


def local_ipv4_addresses() -> list[str]:
    addresses: list[str] = []

    def add_address(address: str) -> None:
        if address and address != "0.0.0.0" and not address.startswith("127.") and address not in addresses:
            addresses.append(address)

    try:
        for item in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            add_address(item[4][0])
    except OSError:
        pass

    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("192.168.1.1", 1))
        add_address(probe.getsockname()[0])
    except OSError:
        pass
    finally:
        probe.close()

    return addresses or ["127.0.0.1"]


def local_join_addresses(port: int = DEFAULT_PORT) -> list[str]:
    return [f"{address}:{port}" for address in local_ipv4_addresses()]


def workbook_to_payload(workbook: WorkbookData) -> dict:
    return {
        "version": PROTOCOL_VERSION,
        "active_sheet_index": workbook.active_sheet_index,
        "sheets": [worksheet_to_payload(sheet) for sheet in workbook.sheets],
    }


def workbook_from_payload(payload: dict) -> WorkbookData:
    sheets = [worksheet_from_payload(item) for item in payload.get("sheets", [])]
    workbook = WorkbookData(sheets=sheets or [WorksheetData(name="Sheet1")])
    active_index = int(payload.get("active_sheet_index", 0) or 0)
    workbook.active_sheet_index = max(0, min(active_index, len(workbook.sheets) - 1))
    return workbook


def worksheet_to_payload(sheet: WorksheetData) -> dict:
    return {
        "name": sheet.name,
        "row_count": sheet.row_count,
        "column_count": sheet.column_count,
        "revision": sheet.revision,
        "column_widths": {str(column): width for column, width in sheet.column_widths.items()},
        "row_heights": {str(row): height for row, height in sheet.row_heights.items()},
        "cells": [
            {
                "row": row,
                "column": column,
                "value": json_safe_value(cell.value),
                "style": asdict(cell.style),
            }
            for row, column, cell in sheet.iter_used_cells()
        ],
    }


def worksheet_from_payload(payload: dict) -> WorksheetData:
    sheet = WorksheetData(
        name=str(payload.get("name") or "Sheet"),
        row_count=max(1, int(payload.get("row_count", 200) or 200)),
        column_count=max(1, int(payload.get("column_count", 52) or 52)),
    )
    for item in payload.get("cells", []):
        row = int(item.get("row", 0) or 0)
        column = int(item.get("column", 0) or 0)
        value = item.get("value", "")
        style = style_from_payload(item.get("style", {}))
        sheet.ensure_size(row, column)
        sheet.cells[(row, column)] = CellData(value="" if value is None else value, style=style)
        sheet.track_formula_cell(row, column, value)
    sheet.column_widths = {int(column): int(width) for column, width in payload.get("column_widths", {}).items()}
    sheet.row_heights = {int(row): int(height) for row, height in payload.get("row_heights", {}).items()}
    sheet.revision = int(payload.get("revision", 0) or 0)
    return sheet


def style_from_payload(payload: dict) -> CellStyle:
    if not isinstance(payload, dict):
        return CellStyle()
    return CellStyle(**{key: value for key, value in payload.items() if key in STYLE_FIELDS})


def json_safe_value(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def cell_update_message(sheet_index: int, sheet_name: str, values: list[tuple[int, int, object]]) -> dict:
    return {
        "type": "cell_update",
        "sheet_index": sheet_index,
        "sheet_name": sheet_name,
        "values": [
            {"row": row, "column": column, "value": json_safe_value(value)}
            for row, column, value in values
        ],
    }


def sheet_message(message_type: str, sheet_index: int, sheet_name: str = "") -> dict:
    return {"type": message_type, "sheet_index": sheet_index, "sheet_name": sheet_name}


def structure_message(message_type: str, sheet_index: int, sheet_name: str, start: int, count: int = 1) -> dict:
    return {
        "type": message_type,
        "sheet_index": sheet_index,
        "sheet_name": sheet_name,
        "start": start,
        "count": count,
    }


class _JsonConnection:
    def __init__(self, sock: socket.socket, label: str):
        self.sock = sock
        self.label = label
        self.lock = threading.Lock()
        self.closed = False

    def send(self, message: dict) -> None:
        data = (json.dumps(message, separators=(",", ":")) + "\n").encode("utf-8")
        with self.lock:
            if self.closed:
                return
            self.sock.sendall(data)

    def close(self) -> None:
        with self.lock:
            if self.closed:
                return
            self.closed = True
        try:
            self.sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            self.sock.close()
        except OSError:
            pass


class CollaborationEndpoint(QObject):
    message_received = Signal(dict)
    status_changed = Signal(str)
    error_occurred = Signal(str)
    client_count_changed = Signal(int)

    def __init__(self):
        super().__init__()
        self._running = False

    @property
    def running(self) -> bool:
        return self._running

    def send(self, message: dict) -> None:
        raise NotImplementedError

    def stop(self) -> None:
        self._running = False


class CollaborationServer(CollaborationEndpoint):
    def __init__(
        self,
        host: str = "0.0.0.0",
        port: int = DEFAULT_PORT,
        snapshot_provider: Callable[[], dict] | None = None,
        snapshot_message_provider: Callable[[], dict] | None = None,
        heartbeat_interval: float = HEARTBEAT_INTERVAL_SECONDS,
    ):
        super().__init__()
        self.host = host
        self.port = port
        self.snapshot_provider = snapshot_provider
        self.snapshot_message_provider = snapshot_message_provider
        self.heartbeat_interval = heartbeat_interval
        self._server_socket: socket.socket | None = None
        self._clients: list[_JsonConnection] = []
        self._clients_lock = threading.Lock()
        self._accept_thread: threading.Thread | None = None
        self._heartbeat_thread: threading.Thread | None = None

    def start(self) -> None:
        if self._running:
            return
        self._server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._server_socket.bind((self.host, self.port))
        self._server_socket.listen()
        self._server_socket.settimeout(0.5)
        self.port = int(self._server_socket.getsockname()[1])
        self._running = True
        self._accept_thread = threading.Thread(target=self._accept_loop, daemon=True)
        self._accept_thread.start()
        self._heartbeat_thread = threading.Thread(target=self._heartbeat_loop, daemon=True)
        self._heartbeat_thread.start()
        self.status_changed.emit(f"Server ready. Clients join: {', '.join(local_join_addresses(self.port))}")

    def send(self, message: dict) -> None:
        self._broadcast(message)

    def stop(self) -> None:
        super().stop()
        if self._server_socket is not None:
            try:
                self._server_socket.close()
            except OSError:
                pass
            self._server_socket = None
        with self._clients_lock:
            clients = list(self._clients)
            self._clients.clear()
        for client in clients:
            client.close()
        self.client_count_changed.emit(0)
        self.status_changed.emit("Offline")

    def _accept_loop(self) -> None:
        while self._running and self._server_socket is not None:
            try:
                sock, address = self._server_socket.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            configure_collaboration_socket(sock)
            client = _JsonConnection(sock, f"{address[0]}:{address[1]}")
            with self._clients_lock:
                self._clients.append(client)
                client_count = len(self._clients)
            self.client_count_changed.emit(client_count)
            self.status_changed.emit(f"Client connected: {client.label}")
            self._send_snapshot(client)
            threading.Thread(target=self._client_loop, args=(client,), daemon=True).start()

    def _send_snapshot(self, client: _JsonConnection) -> None:
        if self.snapshot_message_provider is None and self.snapshot_provider is None:
            return
        try:
            if self.snapshot_message_provider is not None:
                client.send(self.snapshot_message_provider())
            elif self.snapshot_provider is not None:
                client.send({"type": "snapshot", "workbook": self.snapshot_provider()})
        except Exception as exc:
            self.error_occurred.emit(f"Snapshot failed: {exc}")

    def _client_loop(self, client: _JsonConnection) -> None:
        try:
            for message in _iter_json_messages(client.sock):
                if self._handle_heartbeat_message(client, message):
                    continue
                self.message_received.emit(message)
                self._broadcast(message, exclude=client)
        except OSError as exc:
            if self._running:
                self.error_occurred.emit(f"Client error: {exc}")
        finally:
            self._remove_client(client)

    def _heartbeat_loop(self) -> None:
        while self._running:
            self._wait_for_heartbeat()
            if self._running:
                self._broadcast(PING_MESSAGE)

    def _wait_for_heartbeat(self) -> None:
        deadline = time.monotonic() + max(0.1, self.heartbeat_interval)
        while self._running and time.monotonic() < deadline:
            time.sleep(0.05)

    def _handle_heartbeat_message(self, client: _JsonConnection, message: dict) -> bool:
        message_type = message.get("type")
        if message_type == "ping":
            client.send(PONG_MESSAGE)
            return True
        if message_type == "pong":
            return True
        return False

    def _broadcast(self, message: dict, exclude: _JsonConnection | None = None) -> None:
        with self._clients_lock:
            clients = list(self._clients)
        stale: list[_JsonConnection] = []
        for client in clients:
            if client is exclude:
                continue
            try:
                client.send(message)
            except OSError:
                stale.append(client)
        for client in stale:
            self._remove_client(client)

    def _remove_client(self, client: _JsonConnection) -> None:
        client.close()
        with self._clients_lock:
            if client in self._clients:
                self._clients.remove(client)
            client_count = len(self._clients)
        self.client_count_changed.emit(client_count)


class CollaborationClient(CollaborationEndpoint):
    def __init__(self, host: str, port: int = DEFAULT_PORT, heartbeat_interval: float = HEARTBEAT_INTERVAL_SECONDS):
        super().__init__()
        self.host = host
        self.port = port
        self.heartbeat_interval = heartbeat_interval
        self._connection: _JsonConnection | None = None
        self._thread: threading.Thread | None = None
        self._heartbeat_thread: threading.Thread | None = None

    def start(self) -> None:
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._heartbeat_thread = threading.Thread(target=self._heartbeat_loop, daemon=True)
        self._heartbeat_thread.start()

    def send(self, message: dict) -> None:
        if self._connection is None:
            return
        self._connection.send(message)

    def stop(self) -> None:
        super().stop()
        if self._connection is not None:
            self._connection.close()
            self._connection = None
        self.status_changed.emit("Offline")

    def _run(self) -> None:
        while self._running:
            try:
                sock = socket.create_connection((self.host, self.port), timeout=5)
                configure_collaboration_socket(sock)
                connection = _JsonConnection(sock, f"{self.host}:{self.port}")
                self._connection = connection
                self.status_changed.emit(f"Connected to {self.host}:{self.port}")
                for message in _iter_json_messages(sock):
                    if self._handle_heartbeat_message(message):
                        continue
                    self.message_received.emit(message)
                    if not self._running:
                        break
                if self._running:
                    self.status_changed.emit(f"Disconnected. Reconnecting to {self.host}:{self.port}")
            except OSError as exc:
                if self._running:
                    self.error_occurred.emit(f"Connection failed: {exc}")
            finally:
                if self._connection is not None:
                    self._connection.close()
                    self._connection = None
            if self._running:
                self._wait_before_reconnect()
        self.status_changed.emit("Offline")

    def _wait_before_reconnect(self) -> None:
        deadline = time.monotonic() + RECONNECT_DELAY_SECONDS
        while self._running and time.monotonic() < deadline:
            time.sleep(0.05)

    def _heartbeat_loop(self) -> None:
        while self._running:
            self._wait_for_heartbeat()
            connection = self._connection
            if not self._running or connection is None:
                continue
            try:
                connection.send(PING_MESSAGE)
            except OSError:
                connection.close()

    def _wait_for_heartbeat(self) -> None:
        deadline = time.monotonic() + max(0.1, self.heartbeat_interval)
        while self._running and time.monotonic() < deadline:
            time.sleep(0.05)

    def _handle_heartbeat_message(self, message: dict) -> bool:
        message_type = message.get("type")
        if message_type == "ping":
            connection = self._connection
            if connection is not None:
                connection.send(PONG_MESSAGE)
            return True
        if message_type == "pong":
            return True
        return False


def _iter_json_messages(sock: socket.socket):
    buffer = ""
    while True:
        try:
            data = sock.recv(65536)
        except socket.timeout:
            continue
        if not data:
            break
        buffer += data.decode("utf-8")
        while "\n" in buffer:
            line, buffer = buffer.split("\n", 1)
            if line.strip():
                try:
                    yield json.loads(line)
                except json.JSONDecodeError:
                    continue
