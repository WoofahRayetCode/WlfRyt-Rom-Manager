"""
Phase 4 Config Manager UI Enhancements

Provides visual tools for:
- Tool status display (detected, version, path)
- Configuration validation feedback
- Export/import settings functionality

Phase 4 Week 1 – Integration UI Track
"""

import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Any, Dict

logger = logging.getLogger(__name__)


class ToolStatusDisplay(ttk.LabelFrame):
    """
    Visual display of external tool status.

    Shows tool name, detected status (✅/⚠️/❌), version, and path.
    """

    def __init__(
        self,
        parent: tk.Widget,
        tool_registry: Optional[Any] = None,
        **kwargs,
    ):
        super().__init__(parent, text="🔧 Tool Status", **kwargs)
        self.tool_registry = tool_registry
        self._tool_rows: Dict[str, Dict[str, tk.Widget]] = {}

        self._build()
        if tool_registry:
            self.refresh()

    def _build(self) -> None:
        """Build the status grid."""
        # Header
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(header, text="Tool", width=15, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Status", width=15, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Version", width=15, font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="Path", font=("", 9, "bold")).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Scrollable frame for tools
        canvas = tk.Canvas(self, bg="white", highlightthickness=0, height=150)
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=canvas.scroll)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6, pady=4)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._scroll_frame = scrollable_frame

        # Button bar
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh).pack(side=tk.LEFT, padx=2)

    def add_tool(self, tool_name: str) -> None:
        """Add a tool status row (for manual initialization)."""
        row = ttk.Frame(self._scroll_frame)
        row.pack(fill=tk.X, padx=4, pady=2)

        name_lbl = ttk.Label(row, text=tool_name, width=15)
        name_lbl.pack(side=tk.LEFT)

        status_lbl = ttk.Label(row, text="⏳", width=15)
        status_lbl.pack(side=tk.LEFT)

        version_lbl = ttk.Label(row, text="—", width=15)
        version_lbl.pack(side=tk.LEFT)

        path_lbl = ttk.Label(row, text="", font=("Courier", 8))
        path_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._tool_rows[tool_name] = {
            "status": status_lbl,
            "version": version_lbl,
            "path": path_lbl,
        }

    def refresh(self) -> None:
        """Update status for all tools from ToolRegistry."""
        if not self.tool_registry:
            return

        # Clear existing rows
        for widget in self._scroll_frame.winfo_children():
            widget.destroy()
        self._tool_rows.clear()

        # Known tools to display
        known_tools = ["chdman", "maxcso", "7zip", "ps3-dumper", "extract-xiso", "ndecrypt"]
        for tool_name in known_tools:
            self.add_tool(tool_name)

        # Populate from registry
        for tool_name, row_widgets in self._tool_rows.items():
            try:
                # Try to get tool from registry
                tool_mgr = self.tool_registry.tools.get(tool_name)
                if tool_mgr:
                    is_avail = tool_mgr.is_available()
                    status_icon = "✅" if is_avail else "⚠️"
                    version = getattr(tool_mgr, "version", "unknown") if is_avail else "—"
                    path = getattr(tool_mgr, "path", "not found") if is_avail else "—"
                    row_widgets["status"].config(text=status_icon)
                    row_widgets["version"].config(text=version)
                    row_widgets["path"].config(text=str(path)[:60])
                else:
                    row_widgets["status"].config(text="❌")
                    row_widgets["version"].config(text="—")
                    row_widgets["path"].config(text="not installed")
            except Exception as exc:
                logger.debug(f"Tool status update error for {tool_name}: {exc}")
                row_widgets["status"].config(text="❌")


class ConfigValidationDisplay(ttk.LabelFrame):
    """
    Display configuration validation issues and warnings.

    Shows:
    - Missing required settings
    - Invalid values
    - Recommendations
    """

    def __init__(
        self,
        parent: tk.Widget,
        config_adapter: Optional[Any] = None,
        **kwargs,
    ):
        super().__init__(parent, text="⚠️ Configuration Status", **kwargs)
        self.config_adapter = config_adapter
        self._issues: list = []

        self._build()
        if config_adapter:
            self.validate()

    def _build(self) -> None:
        """Build the validation display."""
        # Text widget for issues
        self._text = tk.Text(
            self,
            height=8,
            bg="#fff9e6",
            fg="#333",
            font=("Courier", 9),
            relief="flat",
            state="disabled",
        )
        self._text.pack(fill="both", expand=True, padx=6, pady=6)

        # Button bar
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(btn_frame, text="✓ Revalidate", command=self.validate).pack(side=tk.LEFT, padx=2)

    def validate(self) -> None:
        """Run validation checks and display results."""
        self._issues.clear()

        if not self.config_adapter:
            self._display("No config adapter available.")
            return

        try:
            prefs = self.config_adapter.get_preferences()

            # Check source_dir
            if not prefs.source_dir:
                self._issues.append("⚠️  No source directory configured")
            elif not Path(prefs.source_dir).exists():
                self._issues.append(f"❌ Source directory not found: {prefs.source_dir}")

            # Check output formats
            if not prefs.ps1_output_format:
                self._issues.append("⚠️  PS1 output format not set (default: CHD)")
            if not prefs.ps2_output_format:
                self._issues.append("⚠️  PS2 output format not set (default: CHD)")

            # Check parallel settings
            if prefs.max_workers and prefs.max_workers < 1:
                self._issues.append("❌ max_workers must be ≥ 1")

            if not self._issues:
                self._display("✅ All settings valid!\n\n(No issues detected)")
            else:
                self._display("\n".join(self._issues))
        except Exception as exc:
            logger.debug(f"Config validation error: {exc}")
            self._display(f"Validation error: {exc}")

    def _display(self, text: str) -> None:
        """Update the text widget."""
        self._text.config(state="normal")
        self._text.delete("1.0", tk.END)
        self._text.insert("1.0", text)
        self._text.config(state="disabled")


class ConfigExportImport(ttk.LabelFrame):
    """
    Export and import configuration settings.

    Allows users to:
    - Save config to JSON file
    - Load config from JSON file
    """

    def __init__(
        self,
        parent: tk.Widget,
        config_adapter: Optional[Any] = None,
        **kwargs,
    ):
        super().__init__(parent, text="💾 Configuration Files", **kwargs)
        self.config_adapter = config_adapter

        self._build()

    def _build(self) -> None:
        """Build the export/import buttons."""
        frame = ttk.Frame(self)
        frame.pack(fill=tk.X, padx=6, pady=6)

        ttk.Button(
            frame, text="📥 Export Settings",
            command=self._export,
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            frame, text="📤 Import Settings",
            command=self._import,
        ).pack(side=tk.LEFT, padx=4)

        self._status_var = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self._status_var, foreground="gray").pack(
            side=tk.RIGHT, padx=4
        )

    def _export(self) -> None:
        """Export current config to JSON."""
        if not self.config_adapter:
            messagebox.showwarning("Export", "No config adapter available.", parent=self)
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON config", "*.json")],
            title="Export Configuration",
        )
        if not path:
            return

        try:
            prefs = self.config_adapter.get_preferences()
            config_dict = {
                "source_dir": prefs.source_dir,
                "ps1_output_format": prefs.ps1_output_format,
                "ps2_output_format": prefs.ps2_output_format,
                "psp_output_format": prefs.psp_output_format,
                "delete_originals": prefs.delete_originals,
                "move_to_backup": prefs.move_to_backup,
                "max_workers": prefs.max_workers,
                "chunk_size_mb": prefs.chunk_size_mb,
                "memory_threshold_pct": prefs.memory_threshold_pct,
                "cpu_threshold_pct": prefs.cpu_threshold_pct,
                "retry_count": prefs.retry_count,
                "circuit_breaker_threshold": prefs.circuit_breaker_threshold,
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(config_dict, f, indent=2)
            self._status_var.set(f"✅ Exported to {Path(path).name}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc), parent=self)

    def _import(self) -> None:
        """Import config from JSON."""
        if not self.config_adapter:
            messagebox.showwarning("Import", "No config adapter available.", parent=self)
            return

        path = filedialog.askopenfilename(
            filetypes=[("JSON config", "*.json")],
            title="Import Configuration",
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                config_dict = json.load(f)

            self.config_adapter.update_preferences(**config_dict)
            self._status_var.set(f"✅ Imported from {Path(path).name}")
            messagebox.showinfo("Import", f"Configuration imported successfully.", parent=self)
        except Exception as exc:
            messagebox.showerror("Import Error", str(exc), parent=self)


from pathlib import Path


def create_tool_status_display(
    parent: tk.Widget,
    tool_registry: Optional[Any] = None,
) -> ToolStatusDisplay:
    """Factory for ToolStatusDisplay."""
    return ToolStatusDisplay(parent, tool_registry=tool_registry)


def create_config_validation_display(
    parent: tk.Widget,
    config_adapter: Optional[Any] = None,
) -> ConfigValidationDisplay:
    """Factory for ConfigValidationDisplay."""
    return ConfigValidationDisplay(parent, config_adapter=config_adapter)


def create_config_export_import(
    parent: tk.Widget,
    config_adapter: Optional[Any] = None,
) -> ConfigExportImport:
    """Factory for ConfigExportImport."""
    return ConfigExportImport(parent, config_adapter=config_adapter)
