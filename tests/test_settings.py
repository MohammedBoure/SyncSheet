import json
import tempfile
import unittest
from pathlib import Path

from pyexcel_lite.network import DEFAULT_PORT
from pyexcel_lite.settings import StartupSettings, load_startup_settings, save_startup_settings


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
                ),
                path,
            )

            settings = load_startup_settings(path)

            self.assertEqual(settings.startup_mode, "shared_client")
            self.assertEqual(settings.shared_server_host, "192.168.1.25")
            self.assertEqual(settings.shared_server_port, 9001)
            self.assertEqual(settings.local_server_port, 9002)

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


if __name__ == "__main__":
    unittest.main()
