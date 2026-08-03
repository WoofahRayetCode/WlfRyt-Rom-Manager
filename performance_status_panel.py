"""
Performance Status Panel

A tkinter widget that displays live performance metrics in the ROM Converter UI:
- Active conversion count and worker slots
- Current CPU / RAM usage with color-coded bars
- Throughput and ETA for the active batch
- Quick access to the full PerformanceAnalyzer report
- Circuit-breaker and error-handler status

Phase 4 Week 6: Final Integration & UI
"""

import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable

try:
    from performance_optimizer import PerformanceOptimizer
    OPTIMIZER_AVAILABLE = True
except ImportError:
    OPTIMIZER_AVAILABLE = False

try:
    from conversion_error_handler import ConversionErrorHandler
    ERROR_HANDLER_AVAILABLE = True
except ImportError:
    ERROR_HANDLER_AVAILABLE = False

try:
    from performance_analyzer import PerformanceAnalyzer
    ANALYZER_AVAILABLE = True
except ImportError:
    ANALYZER_AVAILABLE = False


# Colour constants
COLOR_OK     = "#2ecc71"
COLOR_WARN   = "#f39c12"
COLOR_CRIT   = "#e74c3c"
COLOR_IDLE   = "#95a5a6"
COLOR_BG     = "#2c2c2c"
COLOR_FG     = "#ecf0f1"
COLOR_PANEL  = "#3a3a3a"


def _cpu_color(pct: float) -> str:
    if pct >= 95:
        return COLOR_CRIT
    if pct >= 75:
        return COLOR_WARN
    return COLOR_OK


def _ram_color(pct: float) -> str:
    if pct >= 85:
        return COLOR_CRIT
    if pct >= 70:
        return COLOR_WARN
    return COLOR_OK


class ResourceBar(tk.Frame):
    """Compact horizontal bar showing a resource usage percentage."""

    def __init__(
        self,
        parent: tk.Widget,
        label: str,
        width: int = 160,
        height: int = 14,
        **kwargs,
    ):
        super().__init__(parent, bg=COLOR_PANEL, **kwargs)
        self._label_text = label
        self._bar_width   = width
        self._bar_height  = height

        lbl = tk.Label(self, text=label, bg=COLOR_PANEL, fg=COLOR_FG,
                       font=("Helvetica", 8), width=5, anchor="w")
        lbl.pack(side=tk.LEFT, padx=(2, 4))

        self._canvas = tk.Canvas(self, width=width, height=height,
                                 bg="#1a1a1a", highlightthickness=0)
        self._canvas.pack(side=tk.LEFT)

        self._pct_label = tk.Label(self, text="0%", bg=COLOR_PANEL, fg=COLOR_FG,
                                   font=("Helvetica", 8), width=5)
        self._pct_label.pack(side=tk.LEFT, padx=2)

        self._bar_id = self._canvas.create_rectangle(
            0, 0, 0, height, fill=COLOR_OK, outline=""
        )

    def update_value(self, pct: float, color: Optional[str] = None) -> None:
        """Update bar to display a new percentage (0–100)."""
        pct = max(0.0, min(100.0, pct))
        fill_w = int(self._bar_width * pct / 100)
        fill   = color or _cpu_color(pct)
        self._canvas.coords(self._bar_id, 0, 0, fill_w, self._bar_height)
        self._canvas.itemconfig(self._bar_id, fill=fill)
        self._pct_label.config(text=f"{pct:.0f}%")


class PerformanceStatusPanel(tk.LabelFrame):
    """
    Embeddable tkinter panel showing live performance metrics.

    Usage:
        panel = PerformanceStatusPanel(parent, optimizer=..., error_handler=..., analyzer=...)
        panel.pack(fill=tk.X, padx=8, pady=4)
        panel.start_updates()          # begin background polling
        # ... when done:
        panel.stop_updates()
    """

    POLL_INTERVAL_MS = 1000  # Refresh every second

    def __init__(
        self,
        parent: tk.Widget,
        optimizer: Optional["PerformanceOptimizer"]   = None,
        error_handler: Optional["ConversionErrorHandler"] = None,
        analyzer: Optional["PerformanceAnalyzer"]     = None,
        log_callback: Optional[Callable[[str], None]] = None,
        **kwargs,
    ):
        super().__init__(
            parent,
            text="⚡ Performance",
            bg=COLOR_PANEL,
            fg=COLOR_FG,
            font=("Helvetica", 9, "bold"),
            **kwargs,
        )
        self.optimizer     = optimizer
        self.error_handler = error_handler
        self.analyzer      = analyzer
        self.log_callback  = log_callback

        self._running = False
        self._poll_job: Optional[str] = None  # after() job id

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Build all child widgets."""
        # ── Row 0: resource bars ──
        bars_row = tk.Frame(self, bg=COLOR_PANEL)
        bars_row.pack(fill=tk.X, padx=6, pady=(4, 2))

        self._cpu_bar = ResourceBar(bars_row, "CPU")
        self._cpu_bar.pack(side=tk.LEFT, padx=(0, 12))

        self._ram_bar = ResourceBar(bars_row, "RAM")
        self._ram_bar.pack(side=tk.LEFT)

        # ── Row 1: conversion stats ──
        stats_row = tk.Frame(self, bg=COLOR_PANEL)
        stats_row.pack(fill=tk.X, padx=6, pady=2)

        self._workers_lbl = self._stat_label(stats_row, "Workers: —")
        self._workers_lbl.pack(side=tk.LEFT, padx=(0, 16))

        self._throughput_lbl = self._stat_label(stats_row, "Throughput: —")
        self._throughput_lbl.pack(side=tk.LEFT, padx=(0, 16))

        self._eta_lbl = self._stat_label(stats_row, "ETA: —")
        self._eta_lbl.pack(side=tk.LEFT)

        # ── Row 2: error/circuit-breaker status ──
        err_row = tk.Frame(self, bg=COLOR_PANEL)
        err_row.pack(fill=tk.X, padx=6, pady=2)

        self._errors_lbl = self._stat_label(err_row, "Errors: 0")
        self._errors_lbl.pack(side=tk.LEFT, padx=(0, 12))

        self._circuit_lbl = self._stat_label(err_row, "Circuit: OK ✅")
        self._circuit_lbl.pack(side=tk.LEFT)

        # ── Row 3: analyse button ──
        btn_row = tk.Frame(self, bg=COLOR_PANEL)
        btn_row.pack(fill=tk.X, padx=6, pady=(2, 6))

        tk.Button(
            btn_row,
            text="📊 Full Analysis",
            command=self._show_full_analysis,
            bg="#555",
            fg=COLOR_FG,
            font=("Helvetica", 8),
            relief=tk.FLAT,
            padx=6,
            pady=2,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(
            btn_row,
            text="🗑 Clear Errors",
            command=self._clear_errors,
            bg="#555",
            fg=COLOR_FG,
            font=("Helvetica", 8),
            relief=tk.FLAT,
            padx=6,
            pady=2,
            cursor="hand2",
        ).pack(side=tk.LEFT)

    @staticmethod
    def _stat_label(parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(
            parent, text=text,
            bg=COLOR_PANEL, fg=COLOR_FG,
            font=("Helvetica", 8),
        )

    # ------------------------------------------------------------------
    # Polling & updates
    # ------------------------------------------------------------------

    def start_updates(self) -> None:
        """Begin background polling. Call after the panel is visible."""
        if not self._running:
            self._running = True
            self._schedule_poll()

    def stop_updates(self) -> None:
        """Stop background polling."""
        self._running = False
        if self._poll_job is not None:
            try:
                self.after_cancel(self._poll_job)
            except Exception:
                pass
            self._poll_job = None

    def _schedule_poll(self) -> None:
        if self._running:
            self._poll_job = self.after(self.POLL_INTERVAL_MS, self._poll)

    def _poll(self) -> None:
        """Gather metrics and refresh widgets, then reschedule."""
        try:
            self._refresh_resources()
            self._refresh_conversion_stats()
            self._refresh_error_status()
        except Exception:
            pass  # Never crash the UI loop
        finally:
            self._schedule_poll()

    def _refresh_resources(self) -> None:
        """Update CPU and RAM bars."""
        if self.optimizer and self.optimizer.resource_monitor:
            mon = self.optimizer.resource_monitor
            cpu = mon.get_cpu_percent()
            ram = mon.get_memory_percent()
        else:
            # Fallback: try psutil directly
            try:
                import psutil
                cpu = psutil.cpu_percent(interval=None)
                ram = psutil.virtual_memory().percent
            except ImportError:
                cpu, ram = 0.0, 0.0

        self._cpu_bar.update_value(cpu, color=_cpu_color(cpu))
        self._ram_bar.update_value(ram, color=_ram_color(ram))

    def _refresh_conversion_stats(self) -> None:
        """Update workers, throughput, and ETA labels."""
        if not self.optimizer:
            return

        pm = self.optimizer.parallel_manager
        active  = pm.active_conversions
        maximum = pm.max_workers
        self._workers_lbl.config(text=f"Workers: {active}/{maximum}")

        avg_tp = pm.get_average_throughput()
        if avg_tp > 0:
            self._throughput_lbl.config(text=f"Throughput: {avg_tp:.1f} MB/s")

        total     = getattr(self.optimizer, "total_files", 0) or 0
        processed = getattr(self.optimizer, "processed_files", 0) or 0
        remaining = total - processed
        if remaining > 0:
            eta = pm.estimate_time_remaining(remaining)
            if eta is not None:
                total_secs = int(eta.total_seconds())
                h, rem = divmod(total_secs, 3600)
                m, s   = divmod(rem, 60)
                self._eta_lbl.config(
                    text=f"ETA: {h:02d}:{m:02d}:{s:02d}" if h else f"ETA: {m:02d}:{s:02d}"
                )
        elif processed > 0 and total > 0 and remaining == 0:
            self._eta_lbl.config(text="ETA: Done ✅")

    def _refresh_error_status(self) -> None:
        """Update error count and circuit breaker status labels."""
        if not self.error_handler:
            return

        stats = self.error_handler.get_stats()
        total_errors = stats.get("total_failures", 0)
        color = COLOR_CRIT if total_errors > 0 else COLOR_OK
        self._errors_lbl.config(
            text=f"Errors: {total_errors}",
            fg=color,
        )

        open_cbs = stats.get("open_circuit_breakers", [])
        if open_cbs:
            self._circuit_lbl.config(
                text=f"Circuit: {len(open_cbs)} OPEN ⛔",
                fg=COLOR_CRIT,
            )
        else:
            self._circuit_lbl.config(text="Circuit: OK ✅", fg=COLOR_OK)

    # ------------------------------------------------------------------
    # Button actions
    # ------------------------------------------------------------------

    def _show_full_analysis(self) -> None:
        """Open a popup window with the full PerformanceAnalyzer report."""
        if not self.analyzer:
            messagebox.showinfo(
                "Performance Analysis",
                "PerformanceAnalyzer is not available.\n"
                "Ensure metrics_store and performance_analyzer modules are present.",
            )
            return

        try:
            report = self.analyzer.run_full_analysis()
            text   = report.format_report()
        except Exception as exc:
            text = f"Analysis failed: {exc}"

        win = tk.Toplevel(self)
        win.title("📊 Performance Analysis Report")
        win.configure(bg=COLOR_BG)
        win.resizable(True, True)

        txt = tk.Text(
            win, wrap=tk.WORD, bg=COLOR_BG, fg=COLOR_FG,
            font=("Courier", 9), relief=tk.FLAT,
        )
        scrollbar = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        txt.insert(tk.END, text)
        txt.config(state=tk.DISABLED)

        tk.Button(
            win, text="Close", command=win.destroy,
            bg="#555", fg=COLOR_FG, relief=tk.FLAT,
        ).pack(pady=(0, 8))

        win.geometry("680x500")

    def _clear_errors(self) -> None:
        """Clear the error handler's failure log."""
        if self.error_handler:
            self.error_handler.clear()
            self._errors_lbl.config(text="Errors: 0", fg=COLOR_FG)
            self._circuit_lbl.config(text="Circuit: OK ✅", fg=COLOR_OK)
            if self.log_callback:
                self.log_callback("Error log cleared by user")

    # ------------------------------------------------------------------
    # Manual refresh (call from conversion code)
    # ------------------------------------------------------------------

    def refresh_now(self) -> None:
        """Trigger an immediate refresh (safe to call from any thread)."""
        try:
            self.after(0, self._poll)
        except Exception:
            pass


def create_performance_panel(
    parent: tk.Widget,
    optimizer=None,
    error_handler=None,
    analyzer=None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> PerformanceStatusPanel:
    """
    Factory function to create and return a PerformanceStatusPanel.

    Args:
        parent: Parent tkinter widget
        optimizer: PerformanceOptimizer instance (optional)
        error_handler: ConversionErrorHandler instance (optional)
        analyzer: PerformanceAnalyzer instance (optional)
        log_callback: Optional logging callback

    Returns:
        PerformanceStatusPanel (not yet packed — caller controls layout)
    """
    return PerformanceStatusPanel(
        parent,
        optimizer=optimizer,
        error_handler=error_handler,
        analyzer=analyzer,
        log_callback=log_callback,
    )
