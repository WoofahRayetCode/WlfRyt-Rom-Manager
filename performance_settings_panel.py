"""
Performance Settings Panel

Provides a tkinter UI panel exposing Phase 4 performance knobs:
- Parallel worker count
- Chunk size for streaming
- Memory threshold (RAM %)
- Circuit breaker failure threshold
- Retry count per file

Settings are persisted via ConfigAdapter / UserPreferences.

Phase 5 Week 2 – UI Track
"""

import logging
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable, Any

logger = logging.getLogger(__name__)


# Slider configuration: (label, attr_name, min, max, default, step, unit)
SLIDERS = [
    ("Worker Threads",   "max_workers",          1,  32,   4,    1,  ""),
    ("Chunk Size (MB)",  "chunk_size_mb",         1,  64,   8,    1,  "MB"),
    ("Memory Limit (%)", "memory_threshold_pct", 50,  95,  80,    1,  "%"),
    ("CPU Limit (%)",    "cpu_threshold_pct",    50, 100,  95,    1,  "%"),
    ("Retry Count",      "retry_count",           0,  10,   3,    1,  ""),
    ("Circuit Breaker",  "circuit_breaker_threshold", 1, 20, 5,   1,  "failures"),
]


class PerformanceSettingsPanel(ttk.LabelFrame):
    """
    Collapsible settings panel for performance tuning.

    Embed inside an existing Settings tab:

        panel = PerformanceSettingsPanel(parent, config_adapter=self.config_adapter,
                                         on_apply=self._on_settings_apply)
        panel.pack(fill=tk.X, padx=8, pady=4)
    """

    def __init__(
        self,
        parent: tk.Widget,
        config_adapter: Optional[Any] = None,
        on_apply: Optional[Callable[[dict], None]] = None,
        **kwargs,
    ):
        super().__init__(parent, text="⚙ Performance Settings", **kwargs)
        self.config_adapter = config_adapter
        self.on_apply = on_apply

        self._vars: dict[str, tk.IntVar] = {}
        self._collapsed = False

        self._build()
        self._load_from_config()

    # ------------------------------------------------------------------ build

    def _build(self) -> None:
        """Build slider rows and button bar."""
        # Toggle header button
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 4))
        self._toggle_btn = ttk.Button(
            header, text="▼ Show", width=8, command=self._toggle
        )
        self._toggle_btn.pack(side=tk.LEFT)
        ttk.Label(header, text="Tune parallel workers, memory limits, and retries").pack(
            side=tk.LEFT, padx=8
        )

        # Collapsible body
        self._body = ttk.Frame(self)
        self._body.pack(fill=tk.X, expand=True)

        for label, attr, lo, hi, default, step, unit in self.SLIDERS:
            self._add_slider_row(self._body, label, attr, lo, hi, default, unit)

        # Temp dir row
        temp_row = ttk.Frame(self._body)
        temp_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(temp_row, text="Temp Dir", width=22, anchor="w").pack(side=tk.LEFT)
        self._temp_var = tk.StringVar(value="")
        ttk.Entry(temp_row, textvariable=self._temp_var, width=36).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(temp_row, text="Browse…", command=self._browse_temp).pack(side=tk.LEFT)

        # Button bar
        btn_bar = ttk.Frame(self._body)
        btn_bar.pack(fill=tk.X, padx=4, pady=(6, 2))
        ttk.Button(btn_bar, text="✔ Apply", command=self._apply).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="↺ Reset Defaults", command=self._reset_defaults).pack(
            side=tk.LEFT, padx=2
        )

        # Start collapsed
        self._collapse()

    def _add_slider_row(
        self, parent: tk.Widget, label: str, attr: str,
        lo: int, hi: int, default: int, unit: str,
    ) -> None:
        var = tk.IntVar(value=default)
        self._vars[attr] = var

        row = ttk.Frame(parent)
        row.pack(fill=tk.X, padx=4, pady=2)

        ttk.Label(row, text=label, width=22, anchor="w").pack(side=tk.LEFT)

        slider = ttk.Scale(
            row, from_=lo, to=hi, orient=tk.HORIZONTAL,
            variable=var, length=200,
            command=lambda v, a=attr, u=unit: self._on_slider(a, u),
        )
        slider.pack(side=tk.LEFT, padx=4)

        self._val_label = ttk.Label(row, text=f"{default} {unit}", width=12)
        self._val_label.pack(side=tk.LEFT)
        # Store label reference keyed by attr
        setattr(self, f"_lbl_{attr}", self._val_label)

    def _on_slider(self, attr: str, unit: str) -> None:
        lbl = getattr(self, f"_lbl_{attr}", None)
        if lbl:
            lbl.config(text=f"{self._vars[attr].get()} {unit}")

    # ------------------------------------------------------------ collapse/expand

    def _toggle(self) -> None:
        if self._collapsed:
            self._expand()
        else:
            self._collapse()

    def _collapse(self) -> None:
        self._body.pack_forget()
        self._toggle_btn.config(text="▶ Show")
        self._collapsed = True

    def _expand(self) -> None:
        self._body.pack(fill=tk.X, expand=True)
        self._toggle_btn.config(text="▼ Hide")
        self._collapsed = False

    # --------------------------------------------------------------- actions

    def _browse_temp(self) -> None:
        from tkinter import filedialog
        path = filedialog.askdirectory(title="Select Temp Directory")
        if path:
            self._temp_var.set(path)

    def _apply(self) -> None:
        settings = self._collect()
        self._save_to_config(settings)
        if self.on_apply:
            self.on_apply(settings)
        messagebox.showinfo("Settings", "Performance settings applied.", parent=self)

    def _reset_defaults(self) -> None:
        for label, attr, lo, hi, default, step, unit in self.SLIDERS:
            self._vars[attr].set(default)
            self._on_slider(attr, unit)
        self._temp_var.set("")

    # --------------------------------------------------------------- config I/O

    def _collect(self) -> dict:
        result = {attr: var.get() for attr, var in self._vars.items()}
        result["temp_dir"] = self._temp_var.get()
        return result

    def _load_from_config(self) -> None:
        if not self.config_adapter:
            return
        try:
            prefs = self.config_adapter.get_preferences()
            for label, attr, lo, hi, default, step, unit in self.SLIDERS:
                val = getattr(prefs, attr, default)
                if isinstance(val, (int, float)):
                    self._vars[attr].set(int(val))
                    self._on_slider(attr, unit)
            temp = getattr(prefs, "temp_dir", "") or ""
            self._temp_var.set(str(temp))
        except Exception as exc:
            logger.debug("Could not load perf settings from config: %s", exc)

    def _save_to_config(self, settings: dict) -> None:
        if not self.config_adapter:
            return
        try:
            self.config_adapter.update_preferences(**settings)
        except Exception as exc:
            logger.debug("Could not save perf settings to config: %s", exc)

    # --------------------------------------------------------------- public API

    def get_settings(self) -> dict:
        """Return current slider values as a dict."""
        return self._collect()

    def set_settings(self, settings: dict) -> None:
        """Programmatically set slider values."""
        for attr, var in self._vars.items():
            if attr in settings:
                try:
                    var.set(int(settings[attr]))
                    unit = next(
                        (u for _, a, *_, u in SLIDERS if a == attr), ""
                    )
                    self._on_slider(attr, unit)
                except (ValueError, TypeError):
                    pass
        if "temp_dir" in settings:
            self._temp_var.set(settings["temp_dir"])


def create_performance_settings_panel(
    parent: tk.Widget,
    config_adapter: Optional[Any] = None,
    on_apply: Optional[Callable[[dict], None]] = None,
) -> PerformanceSettingsPanel:
    """Factory for PerformanceSettingsPanel."""
    return PerformanceSettingsPanel(parent, config_adapter=config_adapter, on_apply=on_apply)
