import json
import tempfile
import unittest
from pathlib import Path

from pyexcel_lite.network import DEFAULT_PORT
from pyexcel_lite.settings import (
    StartupSettings,
    forget_last_project,
    load_startup_settings,
    local_server_startup,
    remember_last_project,
    save_startup_settings,
    shared_client_startup,
    without_startup_mode,
)


class StartupSettingsTest(unittest.TestCase):
    def test_settings_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            save_startup_settings(
                StartupSettings(
                    startup_mode="shared_client",
                    shared_server_host="192.168.1.25",
                    shared_server_port=9001,
                    local_server_port=9002,
                    last_project_path=str(Path(directory) / "Project"),
                ),
                path,
            )

            settings = load_startup_settings(path)

            self.assertEqual(settings.startup_mode, "shared_client")
            self.assertEqual(settings.shared_server_host, "192.168.1.25")
            self.assertEqual(settings.shared_server_port, 9001)
            self.assertEqual(settings.local_server_port, 9002)
            self.assertEqual(settings.last_project_path, str(Path(directory) / "Project"))

    def test_invalid_settings_fall_back_to_safe_defaults(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            path.write_text(
                json.dumps(
                    {
                        "startup_mode": "invalid",
                        "shared_server_host": "  ",
                        "shared_server_port": "bad",
                        "local_server_port": 999999,
                    }
                ),
                encoding="utf-8",
            )

            settings = load_startup_settings(path)

            self.assertEqual(settings.startup_mode, "manual")
            self.assertEqual(settings.shared_server_host, "127.0.0.1")
            self.assertEqual(settings.shared_server_port, DEFAULT_PORT)
            self.assertEqual(settings.local_server_port, DEFAULT_PORT)
            self.assertEqual(settings.last_project_path, "")

    def test_shared_client_startup_preserves_local_server_port(self):
        settings = shared_client_startup(
            StartupSettings(local_server_port=9010),
            "192.168.1.50",
            9020,
        )

        self.assertEqual(settings.startup_mode, "shared_client")
        self.assertEqual(settings.shared_server_host, "192.168.1.50")
        self.assertEqual(settings.shared_server_port, 9020)
        self.assertEqual(settings.local_server_port, 9010)

    def test_local_server_startup_preserves_shared_server_target(self):
        settings = local_server_startup(
            StartupSettings(shared_server_host="192.168.1.60", shared_server_port=9021),
            9030,
        )

        self.assertEqual(settings.startup_mode, "local_server")
        self.assertEqual(settings.shared_server_host, "192.168.1.60")
        self.assertEqual(settings.shared_server_port, 9021)
        self.assertEqual(settings.local_server_port, 9030)

    def test_without_startup_mode_only_disables_matching_mode(self):
        shared = StartupSettings(startup_mode="shared_client", shared_server_host="192.168.1.70")
        self.assertEqual(without_startup_mode(shared, "shared_client").startup_mode, "manual")
        self.assertEqual(without_startup_mode(shared, "local_server").startup_mode, "shared_client")

    def test_last_project_helpers_preserve_network_settings(self):
        base = StartupSettings(
            startup_mode="shared_client",
            shared_server_host="192.168.1.25",
            shared_server_port=9001,
            local_server_port=9002,
        )

        project_path = Path("C:/Work/Gold")
        remembered = remember_last_project(base, project_path)
        cleared = forget_last_project(remembered)

        self.assertEqual(remembered.last_project_path, str(project_path))
        self.assertEqual(cleared.last_project_path, "")
        self.assertEqual(cleared.startup_mode, "shared_client")
        self.assertEqual(cleared.shared_server_host, "192.168.1.25")


if __name__ == "__main__":
    unittest.main()
