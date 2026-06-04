"""Project/workspace support for grouped spreadsheet folders."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

PROJECT_PROTOCOL_VERSION = 1
OPENABLE_EXTENSIONS = {".xlsx", ".csv"}
IGNORED_DIRECTORY_NAMES = {".git", "__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache", "venv", ".venv"}


@dataclass(frozen=True)
class ProjectFile:
    relative_path: str
    kind: str
    size: int = 0
    modified: float = 0.0
    openable: bool = False

    @property
    def name(self) -> str:
        return Path(self.relative_path).name

    @property
    def extension(self) -> str:
        return Path(self.relative_path).suffix.lower()


@dataclass
class ProjectData:
    name: str = "No Project"
    root_path: str | None = None
    folders: list[str] = field(default_factory=list)
    files: list[ProjectFile] = field(default_factory=list)
    remote: bool = False

    @property
    def is_open(self) -> bool:
        return bool(self.root_path or self.files or self.folders)

    @property
    def openable_count(self) -> int:
        return sum(1 for item in self.files if item.openable)

    @property
    def total_entries(self) -> int:
        return len(self.folders) + len(self.files)

    def file_by_relative_path(self, relative_path: str) -> ProjectFile | None:
        normalized = normalize_relative_path(relative_path)
        for item in self.files:
            if item.relative_path == normalized:
                return item
        return None

    def absolute_path_for(self, item: ProjectFile) -> Path | None:
        if self.remote or not self.root_path:
            return None
        return Path(self.root_path) / Path(item.relative_path)


def scan_project_folder(root_path: str | Path, max_files: int = 5000) -> ProjectData:
    root = Path(root_path).resolve()
    folders: list[str] = []
    files: list[ProjectFile] = []
    for current_root, directory_names, file_names in os.walk(root):
        directory_names[:] = sorted(
            name for name in directory_names if name not in IGNORED_DIRECTORY_NAMES
        )
        current_path = Path(current_root)
        if current_path != root:
            folders.append(normalize_relative_path(current_path.relative_to(root)))
        for file_name in sorted(file_names):
            file_path = current_path / file_name
            try:
                stat = file_path.stat()
            except OSError:
                continue
            relative_path = normalize_relative_path(file_path.relative_to(root))
            extension = file_path.suffix.lower()
            openable = extension in OPENABLE_EXTENSIONS
            kind = "workbook" if extension == ".xlsx" else "csv" if extension == ".csv" else "file"
            files.append(
                ProjectFile(
                    relative_path=relative_path,
                    kind=kind,
                    size=stat.st_size,
                    modified=stat.st_mtime,
                    openable=openable,
                )
            )
            if len(files) >= max_files:
                break
        if len(files) >= max_files:
            break
    return ProjectData(
        name=root.name,
        root_path=str(root),
        folders=sorted(folders),
        files=sorted(files, key=lambda item: item.relative_path.lower()),
    )


def normalize_relative_path(path: str | Path) -> str:
    return Path(path).as_posix().strip("/")


def project_to_payload(project: ProjectData) -> dict:
    return {
        "version": PROJECT_PROTOCOL_VERSION,
        "name": project.name,
        "root_path": project.root_path,
        "remote": project.remote,
        "folders": list(project.folders),
        "files": [
            {
                "relative_path": item.relative_path,
                "kind": item.kind,
                "size": item.size,
                "modified": item.modified,
                "openable": item.openable,
            }
            for item in project.files
        ],
    }


def project_from_payload(payload: dict | None, *, remote: bool = False) -> ProjectData:
    if not isinstance(payload, dict):
        return ProjectData(remote=remote)
    files: list[ProjectFile] = []
    for item in payload.get("files", []):
        if not isinstance(item, dict):
            continue
        relative_path = normalize_relative_path(str(item.get("relative_path") or ""))
        if not relative_path:
            continue
        files.append(
            ProjectFile(
                relative_path=relative_path,
                kind=str(item.get("kind") or "file"),
                size=safe_int(item.get("size"), 0),
                modified=safe_float(item.get("modified"), 0.0),
                openable=bool(item.get("openable")),
            )
        )
    folders = [
        normalize_relative_path(str(item))
        for item in payload.get("folders", [])
        if normalize_relative_path(str(item))
    ]
    return ProjectData(
        name=str(payload.get("name") or "Shared Project"),
        root_path=payload.get("root_path") if isinstance(payload.get("root_path"), str) else None,
        folders=sorted(set(folders)),
        files=sorted(files, key=lambda item: item.relative_path.lower()),
        remote=remote or bool(payload.get("remote")),
    )


def project_snapshot_message(project: ProjectData) -> dict:
    return {"type": "project_snapshot", "project": project_to_payload(project)}


def safe_int(value, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def safe_float(value, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default
