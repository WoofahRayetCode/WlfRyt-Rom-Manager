"""
Phase 5 Week 2 + Phase 4 Week 1 Extended Tests

Covers:
  - DragDropQueuePanel (headless, mocked tk)
  - ToolStatusDisplay (headless)
  - ConfigValidationDisplay (headless)
  - ConfigExportImport (headless)

All tests are headless-safe (no display required).
"""

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch


# ── DragDropQueuePanel tests ──────────────────────────────────────────────────

from dragdrop_queue_panel import (
    DragDropQueuePanel, create_dragdrop_queue_panel, TKDND_AVAILABLE
)


class TestDragDropQueuePanel(unittest.TestCase):

    def test_create_panel_no_crash(self):
        """Instantiate panel without crashing."""
        with patch("dragdrop_queue_panel.ttk.Frame.__init__", return_value=None), \
             patch("dragdrop_queue_panel.DragDropQueuePanel._build"):
            panel = DragDropQueuePanel(MagicMock())
        self.assertIsNotNone(panel)

    def test_add_files_updates_queue(self):
        """Adding files should update internal order and map."""
        with tempfile.TemporaryDirectory() as tmp:
            # Create test files
            test_files = [
                os.path.join(tmp, f"game{i}.iso") for i in range(3)
            ]
            for path in test_files:
                Path(path).touch()

            panel = object.__new__(DragDropQueuePanel)
            panel._lock = __import__("threading").Lock()
            panel._queue_order = []
            panel._file_map = {}
            panel.on_files_added = None
            panel.on_order_changed = None
            panel.on_file_removed = None
            panel._refresh_tree = lambda: None  # Mock refresh

            panel.add_files(test_files)

            self.assertEqual(len(panel._queue_order), 3)
            self.assertEqual(len(panel._file_map), 3)

    def test_remove_file_by_id(self):
        """Removing a file should update queue and map."""
        with tempfile.TemporaryDirectory() as tmp:
            test_file = os.path.join(tmp, "game.iso")
            Path(test_file).touch()

            panel = object.__new__(DragDropQueuePanel)
            panel._lock = __import__("threading").Lock()
            panel._queue_order = []
            panel._file_map = {}
            panel.on_files_added = None
            panel.on_order_changed = None
            panel.on_file_removed = None
            panel._refresh_tree = lambda: None  # Mock refresh

            panel.add_files([test_file])
            file_id = panel._queue_order[0]
            panel.remove_file(file_id)

            self.assertEqual(len(panel._queue_order), 0)
            self.assertEqual(len(panel._file_map), 0)

    def test_get_queue_order(self):
        """get_queue_order should return file paths in current order."""
        with tempfile.TemporaryDirectory() as tmp:
            test_files = [
                os.path.join(tmp, f"game{i}.iso") for i in range(2)
            ]
            for path in test_files:
                Path(path).touch()

            panel = object.__new__(DragDropQueuePanel)
            panel._lock = __import__("threading").Lock()
            panel._queue_order = []
            panel._file_map = {}
            panel.on_files_added = None
            panel._refresh_tree = lambda: None  # Mock refresh

            panel.add_files(test_files)
            order = panel.get_queue_order()

            self.assertEqual(len(order), 2)
            for path in test_files:
                self.assertIn(path, order)

    def test_clear_queue(self):
        """clear_queue should empty both order and map."""
        with tempfile.TemporaryDirectory() as tmp:
            test_file = os.path.join(tmp, "game.iso")
            Path(test_file).touch()

            panel = object.__new__(DragDropQueuePanel)
            panel._lock = __import__("threading").Lock()
            panel._queue_order = []
            panel._file_map = {}
            panel.on_files_added = None
            panel._refresh_tree = lambda: None  # Mock refresh

            panel.add_files([test_file])
            panel.clear_queue()

            self.assertEqual(len(panel._queue_order), 0)
            self.assertEqual(len(panel._file_map), 0)

    def test_set_queue_order_reorders(self):
        """set_queue_order should rearrange the queue."""
        with tempfile.TemporaryDirectory() as tmp:
            test_files = [
                os.path.join(tmp, f"game{i}.iso") for i in range(2)
            ]
            for path in test_files:
                Path(path).touch()

            panel = object.__new__(DragDropQueuePanel)
            panel._lock = __import__("threading").Lock()
            panel._queue_order = []
            panel._file_map = {}
            panel.on_files_added = None
            panel._refresh_tree = lambda: None  # Mock refresh

            panel.add_files(test_files)
            original_order = panel.get_queue_order()

            # Reverse order
            panel.set_queue_order(list(reversed(original_order)))
            new_order = panel.get_queue_order()

            self.assertEqual(new_order, list(reversed(original_order)))

    def test_scan_roms_finds_files(self):
        """_scan_roms should recursively find ROM files."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "game1.iso").touch()
            (root / "game2.bin").touch()
            (root / "game3.txt").touch()  # non-ROM
            subdir = root / "roms"
            subdir.mkdir()
            (subdir / "game4.chd").touch()

            found = DragDropQueuePanel._scan_roms(root)
            found_names = {f.name for f in found}

            self.assertIn("game1.iso", found_names)
            self.assertIn("game2.bin", found_names)
            self.assertIn("game4.chd", found_names)
            self.assertNotIn("game3.txt", found_names)

    def test_factory(self):
        """Factory function should create a panel."""
        with patch("dragdrop_queue_panel.ttk.Frame.__init__", return_value=None), \
             patch("dragdrop_queue_panel.DragDropQueuePanel._build"):
            panel = create_dragdrop_queue_panel(MagicMock())
        self.assertIsInstance(panel, DragDropQueuePanel)


# ── ConfigManagerUI tests ──────────────────────────────────────────────────────

from config_manager_ui import (
    ToolStatusDisplay, ConfigValidationDisplay, ConfigExportImport,
    create_tool_status_display, create_config_validation_display, create_config_export_import,
)


class TestToolStatusDisplay(unittest.TestCase):

    def test_create_display_no_crash(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ToolStatusDisplay._build"):
            display = ToolStatusDisplay(MagicMock())
        self.assertIsNotNone(display)

    def test_add_tool_creates_row(self):
        """add_tool should create a row in _tool_rows."""
        display = object.__new__(ToolStatusDisplay)
        display._tool_rows = {}
        display._scroll_frame = MagicMock()
        display.tool_registry = None

        # Mock row widgets
        mock_row = MagicMock()
        with patch("config_manager_ui.ttk.Frame", return_value=mock_row):
            with patch("config_manager_ui.ttk.Label") as mock_label:
                display.add_tool("chdman")

        self.assertIn("chdman", display._tool_rows)

    def test_refresh_with_no_registry(self):
        """refresh with no registry should not crash."""
        display = object.__new__(ToolStatusDisplay)
        display.tool_registry = None
        display.refresh()  # Should not raise

    def test_factory(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ToolStatusDisplay._build"):
            display = create_tool_status_display(MagicMock())
        self.assertIsInstance(display, ToolStatusDisplay)


class TestConfigValidationDisplay(unittest.TestCase):

    def test_create_display_no_crash(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ConfigValidationDisplay._build"):
            display = ConfigValidationDisplay(MagicMock())
        self.assertIsNotNone(display)

    def test_validate_with_no_adapter(self):
        """validate with no adapter should display message."""
        display = object.__new__(ConfigValidationDisplay)
        display.config_adapter = None
        display._issues = []
        display._text = MagicMock()
        display._display = MagicMock()

        display.validate()
        display._display.assert_called()

    def test_factory(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ConfigValidationDisplay._build"):
            display = create_config_validation_display(MagicMock())
        self.assertIsInstance(display, ConfigValidationDisplay)


class TestConfigExportImport(unittest.TestCase):

    def test_create_panel_no_crash(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ConfigExportImport._build"):
            panel = ConfigExportImport(MagicMock())
        self.assertIsNotNone(panel)

    def test_export_to_json(self):
        """Export should write a valid JSON file."""
        with tempfile.TemporaryDirectory() as tmp:
            out_path = os.path.join(tmp, "config.json")

            prefs = MagicMock()
            prefs.source_dir = "/home/roms"
            prefs.ps1_output_format = "CHD"
            prefs.ps2_output_format = "CSO"
            prefs.psp_output_format = "CSO"
            prefs.delete_originals = False
            prefs.move_to_backup = True
            prefs.max_workers = 4
            prefs.chunk_size_mb = 8
            prefs.memory_threshold_pct = 80
            prefs.cpu_threshold_pct = 95
            prefs.retry_count = 3
            prefs.circuit_breaker_threshold = 5

            adapter = MagicMock()
            adapter.get_preferences.return_value = prefs

            panel = object.__new__(ConfigExportImport)
            panel.config_adapter = adapter
            panel._status_var = MagicMock()

            with patch("config_manager_ui.filedialog.asksaveasfilename", return_value=out_path):
                panel._export()

            self.assertTrue(os.path.exists(out_path))
            with open(out_path, "r") as f:
                data = json.load(f)
            self.assertEqual(data["ps1_output_format"], "CHD")

    def test_import_from_json(self):
        """Import should load and apply settings from JSON."""
        with tempfile.TemporaryDirectory() as tmp:
            config_data = {
                "source_dir": "/test",
                "ps1_output_format": "CHD",
                "ps2_output_format": "ISO",
                "max_workers": 8,
            }
            config_path = os.path.join(tmp, "config.json")
            with open(config_path, "w") as f:
                json.dump(config_data, f)

            adapter = MagicMock()
            panel = object.__new__(ConfigExportImport)
            panel.config_adapter = adapter
            panel._status_var = MagicMock()

            with patch("config_manager_ui.filedialog.askopenfilename", return_value=config_path), \
                 patch("config_manager_ui.messagebox.showinfo"):
                panel._import()

            adapter.update_preferences.assert_called_once()

    def test_factory(self):
        with patch("config_manager_ui.ttk.LabelFrame.__init__", return_value=None), \
             patch("config_manager_ui.ConfigExportImport._build"):
            panel = create_config_export_import(MagicMock())
        self.assertIsInstance(panel, ConfigExportImport)


if __name__ == "__main__":
    unittest.main()
