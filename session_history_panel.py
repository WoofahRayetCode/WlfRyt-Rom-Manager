"""
Session History & Analysis Browser

Provides a tkinter notebook tab that reads from MetricsStore and displays:
- List of past sessions (date, files, success rate, avg throughput)
- Per-session breakdown with per-file rows
- Throughput trend chart (ASCII sparkline)
- CSV export

Phase 5 Week 2 – UI Track
"""

import csv
import logging
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Any, List

logger = logging.getLogger(__name__)

# ASCII sparkline characters (8-level)
_SPARK = " ▁▂▃▄▅▆▇█"


def _sparkline(values: List[float], width: int = 20) -> str:
    """Render a list of floats as a fixed-width ASCII sparkline."""
    if not values:
        return " " * width
    lo, hi = min(values), max(values)
    span = hi - lo or 1.0
    chars = [_SPARK[min(8, int((v - lo) / span * 8))] for v in values]
    # Resample to width
    if len(chars) > width:
        step = len(chars) / width
        chars = [chars[int(i * step)] for i in range(width)]
    elif len(chars) < width:
        chars = chars + [" "] * (width - len(chars))
    return "".join(chars)


def _fmt_time(ts: float) -> str:
    return time.strftime("%Y-%m-%d %H:%M", time.localtime(ts)) if ts else "—"


def _fmt_mb(b: int) -> str:
    return f"{b / 1_048_576:.1f} MB" if b else "0 MB"


class SessionHistoryPanel(ttk.Frame):
    """
    History & analysis browser tab.

    Usage:
        panel = SessionHistoryPanel(notebook, metrics_store=self.metrics_store)
        notebook.add(panel, text="📜 History")
    """

    def __init__(
        self,
        parent: tk.Widget,
        metrics_store: Optional[Any] = None,
        **kwargs,
    ):
        super().__init__(parent, **kwargs)
        self.metrics_store = metrics_store
        self._sessions: list = []        # list of SessionSummary
        self._selected_session: str = ""

        self._build()
        self.refresh()

    # ------------------------------------------------------------------ build

    def _build(self) -> None:
        # Top toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(toolbar, text="🔄 Refresh", command=self.refresh).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📥 Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑 Clear History", command=self._clear_history).pack(
            side=tk.LEFT, padx=2
        )
        self._status_var = tk.StringVar(value="No data loaded")
        ttk.Label(toolbar, textvariable=self._status_var, foreground="gray").pack(
            side=tk.RIGHT, padx=6
        )

        # Paned layout: session list (left) | detail (right)
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=2)

        # --- Session list pane ---
        left = ttk.Frame(paned)
        paned.add(left, weight=1)

        ttk.Label(left, text="Sessions", font=("", 10, "bold")).pack(anchor="w")

        cols = ("date", "files", "ok", "rate", "throughput")
        self._sess_tree = ttk.Treeview(
            left, columns=cols, show="headings", selectmode="browse", height=20
        )
        for col, hdr, w in [
            ("date",       "Date",        140),
            ("files",      "Files",         45),
            ("ok",         "OK",            35),
            ("rate",       "Rate",          50),
            ("throughput", "MB/s avg",      70),
        ]:
            self._sess_tree.heading(col, text=hdr)
            self._sess_tree.column(col, width=w, anchor="center")
        self._sess_tree.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self._sess_tree.yview)
        self._sess_tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._sess_tree.bind("<<TreeviewSelect>>", self._on_session_select)

        # --- Detail pane ---
        right = ttk.Frame(paned)
        paned.add(right, weight=2)

        # Trend sparkline
        spark_frame = ttk.LabelFrame(right, text="Throughput Trend (MB/s)")
        spark_frame.pack(fill=tk.X, padx=4, pady=2)
        self._spark_var = tk.StringVar(value="Select a session →")
        ttk.Label(
            spark_frame, textvariable=self._spark_var,
            font=("Courier", 11), anchor="w"
        ).pack(fill=tk.X, padx=6, pady=4)

        # Summary labels
        info_frame = ttk.LabelFrame(right, text="Session Summary")
        info_frame.pack(fill=tk.X, padx=4, pady=2)
        self._info_vars: dict[str, tk.StringVar] = {}
        for key in ("Session ID", "Started", "Ended", "Total Files",
                    "Successful", "Failed", "Input Size", "Output Size",
                    "Avg Throughput"):
            row = ttk.Frame(info_frame)
            row.pack(fill=tk.X, padx=6, pady=1)
            ttk.Label(row, text=f"{key}:", width=18, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar(value="—")
            ttk.Label(row, textvariable=var, anchor="w").pack(side=tk.LEFT)
            self._info_vars[key] = var

        # File-level breakdown
        file_frame = ttk.LabelFrame(right, text="File Breakdown")
        file_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)

        fcols = ("file", "status", "input", "output", "duration", "throughput")
        self._file_tree = ttk.Treeview(
            file_frame, columns=fcols, show="headings", height=10
        )
        for col, hdr, w in [
            ("file",       "File",        180),
            ("status",     "Status",       60),
            ("input",      "Input",        80),
            ("output",     "Output",       80),
            ("duration",   "Sec",          50),
            ("throughput", "MB/s",         60),
        ]:
            self._file_tree.heading(col, text=hdr)
            self._file_tree.column(col, width=w, anchor="center")
        self._file_tree.pack(fill=tk.BOTH, expand=True)
        fsb = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self._file_tree.yview)
        self._file_tree.configure(yscrollcommand=fsb.set)
        fsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Tag colors
        self._file_tree.tag_configure("ok",  foreground="green")
        self._file_tree.tag_configure("err", foreground="red")

    # ------------------------------------------------------------ data loading

    def refresh(self) -> None:
        """Reload sessions from MetricsStore."""
        if not self.metrics_store:
            self._status_var.set("MetricsStore not available")
            return
        try:
            sessions = self.metrics_store.get_all_sessions()
            self._sessions = sessions or []
            self._populate_session_list()
            self._status_var.set(f"{len(self._sessions)} sessions loaded")
        except Exception as exc:
            logger.debug("History refresh error: %s", exc)
            self._status_var.set(f"Error: {exc}")

    def _populate_session_list(self) -> None:
        for row in self._sess_tree.get_children():
            self._sess_tree.delete(row)
        for s in reversed(self._sessions):  # newest first
            rate = f"{s.success_rate * 100:.0f}%"
            self._sess_tree.insert(
                "", tk.END, iid=s.session_id,
                values=(
                    _fmt_time(s.start_timestamp),
                    s.total_files,
                    s.successful,
                    rate,
                    f"{s.avg_throughput_mbps:.2f}",
                ),
            )

    def _on_session_select(self, _event: tk.Event) -> None:
        sel = self._sess_tree.selection()
        if not sel:
            return
        session_id = sel[0]
        self._selected_session = session_id
        self._load_session_detail(session_id)

    def _load_session_detail(self, session_id: str) -> None:
        if not self.metrics_store:
            return
        try:
            # Get summary
            summary = next(
                (s for s in self._sessions if s.session_id == session_id), None
            )
            if summary:
                self._info_vars["Session ID"].set(summary.session_id)
                self._info_vars["Started"].set(_fmt_time(summary.start_timestamp))
                self._info_vars["Ended"].set(_fmt_time(summary.end_timestamp))
                self._info_vars["Total Files"].set(str(summary.total_files))
                self._info_vars["Successful"].set(str(summary.successful))
                self._info_vars["Failed"].set(str(summary.failed))
                self._info_vars["Input Size"].set(_fmt_mb(summary.total_input_bytes))
                self._info_vars["Output Size"].set(_fmt_mb(summary.total_output_bytes))
                self._info_vars["Avg Throughput"].set(
                    f"{summary.avg_throughput_mbps:.2f} MB/s"
                )

            # Get records for sparkline and file list
            records = self.metrics_store.get_session_records(session_id)
            throughputs = [r.throughput_mbps for r in records]
            self._spark_var.set(_sparkline(throughputs, width=40))

            # Populate file tree
            for row in self._file_tree.get_children():
                self._file_tree.delete(row)
            for rec in records:
                status_icon = "✅" if rec.success else "❌"
                tag = "ok" if rec.success else "err"
                self._file_tree.insert(
                    "", tk.END, tags=(tag,),
                    values=(
                        rec.file_name,
                        status_icon,
                        _fmt_mb(rec.input_size_bytes),
                        _fmt_mb(rec.output_size_bytes),
                        f"{rec.duration_seconds:.1f}",
                        f"{rec.throughput_mbps:.2f}",
                    ),
                )
        except Exception as exc:
            logger.debug("Session detail load error: %s", exc)

    # --------------------------------------------------------------- actions

    def _export_csv(self) -> None:
        if not self.metrics_store:
            messagebox.showwarning("Export", "MetricsStore not available.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Export History",
        )
        if not path:
            return
        try:
            all_records = []
            for s in self._sessions:
                all_records.extend(self.metrics_store.get_session_records(s.session_id))
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "session_id", "file_name", "input_format", "output_format",
                    "input_mb", "output_mb", "duration_s", "throughput_mbps",
                    "success", "tool", "error", "timestamp",
                ])
                for r in all_records:
                    writer.writerow([
                        r.session_id, r.file_name, r.input_format, r.output_format,
                        f"{r.input_size_bytes / 1_048_576:.3f}",
                        f"{r.output_size_bytes / 1_048_576:.3f}",
                        f"{r.duration_seconds:.3f}",
                        f"{r.throughput_mbps:.3f}",
                        r.success, r.tool_name, r.error_message,
                        _fmt_time(r.timestamp),
                    ])
            messagebox.showinfo("Export", f"Exported {len(all_records)} records to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc), parent=self)

    def _clear_history(self) -> None:
        if not self.metrics_store:
            return
        if not messagebox.askyesno(
            "Clear History",
            "Delete ALL conversion history?\nThis cannot be undone.",
            parent=self,
        ):
            return
        try:
            self.metrics_store.clear_all()
            self.refresh()
        except Exception as exc:
            messagebox.showerror("Error", str(exc), parent=self)


def create_session_history_panel(
    parent: tk.Widget,
    metrics_store: Optional[Any] = None,
) -> SessionHistoryPanel:
    """Factory for SessionHistoryPanel."""
    return SessionHistoryPanel(parent, metrics_store=metrics_store)
