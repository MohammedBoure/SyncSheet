"""Application startup settings."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from .network import DEFAULT_PORT

SETTINGS_PATH = Path.cwd() / "pyexcel_lite_settings.json"
STARTUP_MODES = {"manual", "shared_client", "local_server"}


@dataclass
class StartupSettings:
    startup_mode: str = "manual"
    shared_server_host: str = "127.0.0.1"
    shared_server_port: int = DEFAULT_PORT
    local_server_port: int = DEFAULT_PORT

    def normalized(self) -> "StartupSettings":
        mode = self.startup_mode if self.startup_mode in STARTUP_MODES else "manual"
        return StartupSettings(
            startup_mode=mode,
            shared_server_host=self.shared_server_host.strip() or "127.0.0.1",
            shared_server_port=valid_port(self.shared_server_port),
            local_server_port=valid_port(self.local_server_port),
        )


def shared_client_startup(settings: StartupSettings, host: str, port: int) -> StartupSettings:
    current = settings.normalized()
    return StartupSettings(
        startup_mode="shared_client",
        shared_server_host=host,
        shared_server_port=port,
        local_server_port=current.local_server_port,
    ).normalized()


def local_server_startup(settings: StartupSettings, port: int) -> StartupSettings:
    current = settings.normalized()
    return StartupSettings(
        startup_mode="local_server",
        shared_server_host=current.shared_server_host,
        shared_server_port=current.shared_server_port,
        local_server_port=port,
    ).normalized()


def without_startup_mode(settings: StartupSettings, mode: str) -> StartupSettings:
    current = settings.normalized()
    if current.startup_mode != mode:
        return current
    return StartupSettings(
        startup_mode="manual",
        shared_server_host=current.shared_server_host,
        shared_server_port=current.shared_server_port,
        local_server_port=current.local_server_port,
    ).normalized()


def valid_port(value: int) -> int:
    try:
        port = int(value)
    except (TypeError, ValueError):
        return DEFAULT_PORT
    return port if 1 <= port <= 65535 else DEFAULT_PORT


def load_startup_settings(path: Path = SETTINGS_PATH) -> StartupSettings:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return StartupSettings()
    if not isinstance(payload, dict):
        return StartupSettings()
    return StartupSettings(
        startup_mode=str(payload.get("startup_mode", "manual")),
        shared_server_host=str(payload.get("shared_server_host", "127.0.0.1")),
        shared_server_port=valid_port(payload.get("shared_server_port", DEFAULT_PORT)),
        local_server_port=valid_port(payload.get("local_server_port", DEFAULT_PORT)),
    ).normalized()


def save_startup_settings(settings: StartupSettings, path: Path = SETTINGS_PATH) -> None:
    normalized = settings.normalized()
    path.write_text(json.dumps(asdict(normalized), indent=2), encoding="utf-8")
