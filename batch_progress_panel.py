"""
Batch Progress Panel — Phase 5 UI Track

A tkinter widget that replaces the single progress bar with:
- Overall batch progress bar + percentage + ETA
- Per-file progress list (scrollable, shows status icons)
- Pause / Resume and Cancel Individual File controls
- Live throughput readout
"""

import threading
import time
import tkinter as tk
from tkinter import ttk
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional, Callable, Dict, List

# Reuse colour constants pattern from the app
_BG_DARK   = "#2c2c2c"
_BG_MED    = "#3a3a3a"
_BG_LIGHT  = "#4a4a4a"
_FG_PRI    = "#ecf0f1"
_FG_SEC    = "#bdc3c7"
_FG_MUT    = "#7f8c8d"
_GREEN     = "#2ecc71"
_ORANGE    = "#f39c12"
_RED       = "#e74c3c"
_BLUE      = "#3498db"
_YELLOW    = "#f1c40f"


class FileStatus(Enum):
    QUEUED     = "queued"
    CONVERTING = "converting"
    DONE       = "done"
    FAILED     = "failed"
    CANCELLED  = "cancelled"
    VERIFYING  = "verifying"


STATUS_ICON = {
    FileStatus.QUEUED:     "⏳",
    FileStatus.CONVERTING: "⚙️",
    FileStatus.DONE:       "✅",
    FileStatus.FAILED:     "❌",
    FileStatus.CANCELLED:  "🚫",
    FileStatus.VERIFYING:  "🔍",
}

STATUS_COLOR = {
    FileStatus.QUEUED:     _FG_MUT,
    FileStatus.CONVERTING: _YELLOW,
    FileStatus.DONE:       _GREEN,
    FileStatus.FAILED:     _RED,
    FileStatus.CANCELLED:  _FG_MUT,
    FileStatus.VERIFYING:  _BLUE,
}


@dataclass
class FileEntry:
    """State for a single file in the queue."""
    file_id: str
    name: str
    status: FileStatus = FileStatus.QUEUED
    progress: float = 0.0        # 0.0–100.0
    size_mb: float = 0.0
    throughput_mbps: float = 0.0
    start_time: Optional[float] = None
    end_time: Optional[float] = None
    error: str = ""
    cancel_requested: bool = False

    @property
    def duration_s(self) -> float:
        if self.start_time is None:
            return 0.0
        end = self.end_time or time.time()
        return end - self.start_time


class BatchProgressPanel(tk.LabelFrame):
    """
    Embeddable batch progress panel.

    Usage:
        panel = BatchProgressPanel(parent, log_callback=self.log)
        panel.pack(fill=tk.BOTH, expand=True)

        panel.start_batch(file_names)        # initialise queue
        panel.update_file(id, status, pct)   # call from conversion threads
        panel.finish_batch()                 # mark done
    """

    POLL_MS = 250   # UI refresh interval

    def __init__(
        self,
        parent: tk.Widget,
        log_callback: Optional[Callable[[str], None]] = None,
        on_pause_resume: Optional[Callable[[bool], None]] = None,
        on_cancel_file: Optional[Callable[[str], None]] = None,
        **kwargs,
    ):
        super().__init__(
            parent,
            text="📋 Conversion Queue",
            bg=_BG_DARK,
            fg=_FG_PRI,
            font=("Helvetica", 9, "bold"),
            **kwargs,
        )
        self.log_callback    = log_callback
        self.on_pause_resume = on_pause_resume
        self.on_cancel_file  = on_cancel_file

        self._lock         = threading.Lock()
        self._files: Dict[str, FileEntry] = {}
        self._order: List[str] = []
        self._paused       = False
        self._batch_start  = 0.0
        self._total        = 0
        self._done_count   = 0
        self._poll_job: Optional[str] = None

        self._build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ── Overall progress row ──
        top = tk.Frame(self, bg=_BG_DARK)
        top.pack(fill=tk.X, padx=6, pady=(4, 2))

        self._overall_lbl = tk.Label(
            top, text="No batch running", bg=_BG_DARK, fg=_FG_SEC,
            font=("Helvetica", 9),
        )
        self._overall_lbl.pack(side=tk.LEFT)

        self._eta_lbl = tk.Label(
            top, text="", bg=_BG_DARK, fg=_FG_SEC, font=("Helvetica", 9),
        )
        self._eta_lbl.pack(side=tk.RIGHT)

        self._overall_bar = ttk.Progressbar(
            self, mode="determinate", length=400,
        )
        self._overall_bar.pack(fill=tk.X, padx=6, pady=(0, 4))

        # ── Control buttons ──
        btn_row = tk.Frame(self, bg=_BG_DARK)
        btn_row.pack(fill=tk.X, padx=6, pady=(0, 4))

        self._pause_btn = tk.Button(
            btn_row, text="⏸ Pause", command=self._toggle_pause,
            bg=_BG_LIGHT, fg=_FG_PRI, font=("Helvetica", 8),
            relief=tk.FLAT, padx=8, pady=2, cursor="hand2",
        )
        self._pause_btn.pack(side=tk.LEFT, padx=(0, 6))

        self._cancel_sel_btn = tk.Button(
            btn_row, text="🚫 Cancel Selected", command=self._cancel_selected,
            bg=_BG_LIGHT, fg=_FG_PRI, font=("Helvetica", 8),
            relief=tk.FLAT, padx=8, pady=2, cursor="hand2", state=tk.DISABLED,
        )
        self._cancel_sel_btn.pack(side=tk.LEFT)

        self._throughput_lbl = tk.Label(
            btn_row, text="", bg=_BG_DARK, fg=_FG_SEC, font=("Helvetica", 8),
        )
        self._throughput_lbl.pack(side=tk.RIGHT)

        # ── Per-file list ──
        list_frame = tk.Frame(self, bg=_BG_DARK)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self._listbox = tk.Listbox(
            list_frame,
            bg=_BG_MED, fg=_FG_PRI, selectbackground=_BLUE,
            font=("Courier", 9), relief=tk.FLAT, activestyle="none",
            yscrollcommand=scrollbar.set, height=10,
        )
        scrollbar.config(command=self._listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._listbox.bind("<<ListboxSelect>>", self._on_select)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def start_batch(self, file_names: List[str]) -> None:
        """
        Initialise a new batch.

        Args:
            file_names: Ordered list of file names (shown in queue)
        """
        with self._lock:
            self._files.clear()
            self._order.clear()
            self._paused       = False
            self._batch_start  = time.time()
            self._total        = len(file_names)
            self._done_count   = 0
            for name in file_names:
                fid = name
                self._files[fid] = FileEntry(file_id=fid, name=name)
                self._order.append(fid)

        self.after(0, self._rebuild_list)
        self._start_poll()

    def update_file(
        self,
        file_id: str,
        status: FileStatus,
        progress: float = 0.0,
        throughput_mbps: float = 0.0,
        error: str = "",
    ) -> None:
        """
        Update a file's status. Thread-safe — call from any thread.

        Args:
            file_id: File identifier (name passed to start_batch)
            status: New FileStatus
            progress: 0–100
            throughput_mbps: Current throughput
            error: Error message (if failed)
        """
        with self._lock:
            if file_id not in self._files:
                return
            entry = self._files[file_id]
            entry.status          = status
            entry.progress        = progress
            entry.throughput_mbps = throughput_mbps
            if error:
                entry.error = error
            if status == FileStatus.CONVERTING and entry.start_time is None:
                entry.start_time = time.time()
            if status in (FileStatus.DONE, FileStatus.FAILED, FileStatus.CANCELLED):
                entry.end_time   = time.time()
                self._done_count += 1

        # Schedule UI update on main thread
        self.after(0, self._refresh_ui)

    def finish_batch(self) -> None:
        """Mark batch as complete and stop polling."""
        self._stop_poll()
        self.after(0, self._refresh_ui)

    def is_cancelled(self, file_id: str) -> bool:
        """Check if a file has been cancel-requested."""
        with self._lock:
            entry = self._files.get(file_id)
            return entry.cancel_requested if entry else False

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _start_poll(self) -> None:
        if self._poll_job is None:
            self._poll_job = self.after(self.POLL_MS, self._poll)

    def _stop_poll(self) -> None:
        if self._poll_job:
            try:
                self.after_cancel(self._poll_job)
            except Exception:
                pass
            self._poll_job = None

    def _poll(self) -> None:
        self._refresh_ui()
        if self._poll_job is not None:
            self._poll_job = self.after(self.POLL_MS, self._poll)

    def _refresh_ui(self) -> None:
        with self._lock:
            total    = self._total
            done     = self._done_count
            files    = dict(self._files)
            elapsed  = time.time() - self._batch_start if self._batch_start else 0
            paused   = self._paused

        if total == 0:
            return

        pct = (done / total) * 100
        self._overall_bar["value"] = pct

        active = [e for e in files.values() if e.status == FileStatus.CONVERTING]
        avg_tp = sum(e.throughput_mbps for e in active) / max(len(active), 1) if active else 0.0
        self._throughput_lbl.config(
            text=f"{avg_tp:.1f} MB/s" if avg_tp > 0 else ""
        )

        # ETA
        if done > 0 and elapsed > 0 and done < total:
            eta_s = int((elapsed / done) * (total - done))
            h, r  = divmod(eta_s, 3600)
            m, s  = divmod(r, 60)
            eta_str = f"ETA {h:02d}:{m:02d}:{s:02d}" if h else f"ETA {m:02d}:{s:02d}"
        elif done == total and total > 0:
            eta_str = f"Done in {elapsed:.0f}s ✅"
        else:
            eta_str = ""
        self._eta_lbl.config(text=eta_str)

        status_icon = "⏸" if paused else "⚙️" if done < total else "✅"
        self._overall_lbl.config(
            text=f"{status_icon}  {done}/{total} files  ({pct:.1f}%)"
        )

        self._pause_btn.config(text="▶ Resume" if paused else "⏸ Pause")

        self._update_list(files)

    def _rebuild_list(self) -> None:
        """Full rebuild of the listbox."""
        self._listbox.delete(0, tk.END)
        with self._lock:
            order = list(self._order)
            files = dict(self._files)
        for fid in order:
            entry = files[fid]
            self._listbox.insert(tk.END, self._format_entry(entry))
            self._listbox.itemconfig(
                tk.END, fg=STATUS_COLOR[entry.status]
            )

    def _update_list(self, files: dict) -> None:
        """Update list items in-place (faster than full rebuild)."""
        with self._lock:
            order = list(self._order)
        for idx, fid in enumerate(order):
            entry = files.get(fid)
            if entry is None:
                continue
            text = self._format_entry(entry)
            try:
                self._listbox.delete(idx)
                self._listbox.insert(idx, text)
                self._listbox.itemconfig(idx, fg=STATUS_COLOR[entry.status])
            except tk.TclError:
                break

    @staticmethod
    def _format_entry(entry: FileEntry) -> str:
        icon = STATUS_ICON[entry.status]
        name = entry.name[:45].ljust(46)
        if entry.status == FileStatus.CONVERTING:
            bar_filled = int(entry.progress / 10)
            bar = "█" * bar_filled + "░" * (10 - bar_filled)
            tp  = f" {entry.throughput_mbps:.1f}MB/s" if entry.throughput_mbps > 0 else ""
            return f"{icon} {name} [{bar}]{tp}"
        if entry.status == FileStatus.FAILED:
            return f"{icon} {name} {entry.error[:30]}"
        if entry.status == FileStatus.DONE:
            dur = f"{entry.duration_s:.1f}s"
            return f"{icon} {name} {dur}"
        return f"{icon} {name}"

    def _on_select(self, _event) -> None:
        sel = self._listbox.curselection()
        state = tk.NORMAL if sel else tk.DISABLED
        self._cancel_sel_btn.config(state=state)

    def _toggle_pause(self) -> None:
        with self._lock:
            self._paused = not self._paused
            paused = self._paused
        if self.on_pause_resume:
            self.on_pause_resume(paused)
        if self.log_callback:
            self.log_callback("⏸ Batch paused" if paused else "▶ Batch resumed")

    def _cancel_selected(self) -> None:
        sel = self._listbox.curselection()
        if not sel:
            return
        with self._lock:
            idx = sel[0]
            if idx < len(self._order):
                fid = self._order[idx]
                entry = self._files.get(fid)
                if entry and entry.status in (FileStatus.QUEUED, FileStatus.CONVERTING):
                    entry.cancel_requested = True
        if self.on_cancel_file:
            with self._lock:
                fid = self._order[sel[0]] if sel[0] < len(self._order) else None
            if fid:
                self.on_cancel_file(fid)
        self.after(0, self._refresh_ui)


def create_batch_progress_panel(
    parent: tk.Widget,
    log_callback: Optional[Callable[[str], None]] = None,
    on_pause_resume: Optional[Callable[[bool], None]] = None,
    on_cancel_file:  Optional[Callable[[str], None]] = None,
) -> BatchProgressPanel:
    """Factory function for BatchProgressPanel."""
    return BatchProgressPanel(
        parent,
        log_callback=log_callback,
        on_pause_resume=on_pause_resume,
        on_cancel_file=on_cancel_file,
    )
