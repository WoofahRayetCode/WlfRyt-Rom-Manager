"""
Drag-and-Drop File Queue Manager

Extends the BatchProgressPanel with tkinter drag-and-drop support for:
- Drag ROM files / folders onto the queue
- Reorder queue by dragging
- Add/remove files while batch is running
- Visual feedback on drag enter/over/leave

Phase 5 Week 2 – UI Track (drag-and-drop enhancement)
"""

import os
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import Optional, Callable, List, Set

# Try to import tkinterdnd2 (optional — graceful fallback if not installed)
try:
    from tkinterdnd2 import DND_FILES, DND_TEXT
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False
    DND_FILES = None
    DND_TEXT = None


class DragDropQueuePanel(ttk.Frame):
    """
    Enhanced queue panel with drag-and-drop support.

    Wrap this around an existing BatchProgressPanel or use standalone.
    If tkinterdnd2 is not installed, falls back to buttons.

    Usage:
        panel = DragDropQueuePanel(parent, on_files_added=self._on_files_added,
                                   on_order_changed=self._on_order_changed)
        panel.pack(fill="both", expand=True)
    """

    def __init__(
        self,
        parent: tk.Widget,
        on_files_added: Optional[Callable[[List[str]], None]] = None,
        on_order_changed: Optional[Callable[[List[str]], None]] = None,
        on_file_removed: Optional[Callable[[str], None]] = None,
        **kwargs,
    ):
        super().__init__(parent, **kwargs)
        self.on_files_added = on_files_added
        self.on_order_changed = on_order_changed
        self.on_file_removed = on_file_removed

        self._lock = threading.Lock()
        self._queue_order: List[str] = []      # file IDs in order
        self._file_map: dict = {}               # file_id → (name, path, size_mb)
        self._dnd_target_idx: Optional[int] = None

        self._build()
        if TKDND_AVAILABLE:
            self._setup_drag_drop()

    # ----------------------------------------------------------------- build

    def _build(self) -> None:
        """Build the UI."""
        # Toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=6, pady=4)

        ttk.Button(
            toolbar, text="📁 Add Files", command=self._on_add_files
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            toolbar, text="📂 Add Folder", command=self._on_add_folder
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            toolbar, text="🗑 Remove Selected", command=self._on_remove_selected
        ).pack(side=tk.LEFT, padx=2)

        self._dnd_label = ttk.Label(
            toolbar,
            text="💡 Drag files here to add to queue" if TKDND_AVAILABLE else "",
            foreground="gray",
        )
        self._dnd_label.pack(side=tk.RIGHT, padx=6)

        # Queue list
        list_frame = ttk.LabelFrame(self, text="📋 Conversion Queue")
        list_frame.pack(fill="both", expand=True, padx=6, pady=2)

        cols = ("name", "status", "size")
        self._queue_tree = ttk.Treeview(
            list_frame, columns=cols, show="headings", selectmode="browse"
        )
        for col, hdr, w in [
            ("name", "File", 280),
            ("status", "Status", 80),
            ("size", "Size (MB)", 70),
        ]:
            self._queue_tree.heading(col, text=hdr)
            self._queue_tree.column(col, width=w, anchor="w")
        self._queue_tree.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self._queue_tree.yview)
        self._queue_tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind selection
        self._queue_tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Reorder controls
        reorder_frame = ttk.Frame(self)
        reorder_frame.pack(fill=tk.X, padx=6, pady=(2, 4))
        ttk.Button(
            reorder_frame, text="⬆ Move Up", command=self._move_up
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            reorder_frame, text="⬇ Move Down", command=self._move_down
        ).pack(side=tk.LEFT, padx=2)

        self._selected_iid: Optional[str] = None

    def _setup_drag_drop(self) -> None:
        """Register drag-and-drop if tkinterdnd2 is available."""
        if not TKDND_AVAILABLE:
            return

        self.drop_target_register(DND_FILES, DND_TEXT)
        self.dnd_bind("<<Drop>>", self._on_drop)
        self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DragLeave>>", self._on_drag_leave)

    # --------------------------------------------------------------- drag-drop

    def _on_drag_enter(self, event) -> str:
        if TKDND_AVAILABLE:
            self._dnd_label.config(foreground="yellow", text="✨ Drop to add to queue")
            return event.action
        return "none"

    def _on_drag_leave(self, event) -> str:
        if TKDND_AVAILABLE:
            self._dnd_label.config(
                foreground="gray",
                text="💡 Drag files here to add to queue",
            )
        return event.action

    def _on_drop(self, event) -> str:
        """Handle dropped files."""
        if not TKDND_AVAILABLE or not hasattr(event, "data"):
            return "none"

        # Parse the dropped data (may be TCL list format with braces)
        data = event.data
        files = []
        if data.startswith("{") or " " in data:
            # TCL list format — split carefully
            try:
                # Try simple split on spaces
                parts = data.replace("{", "").replace("}", "").split()
                files = [p for p in parts if os.path.exists(p)]
            except Exception:
                pass
        else:
            files = [data] if os.path.exists(data) else []

        if files:
            self.add_files(files)
            self._on_drag_leave(event)
        return "copy" if files else "none"

    # ------------------------------------------------------------ queue mgmt

    def add_files(self, paths: List[str]) -> None:
        """Add ROM files to the queue."""
        with self._lock:
            for path in paths:
                p = Path(path)
                if p.is_file():
                    file_id = p.stem
                    size_mb = p.stat().st_size / (1 << 20)
                    self._file_map[file_id] = (p.name, str(p), size_mb)
                    if file_id not in self._queue_order:
                        self._queue_order.append(file_id)
                elif p.is_dir():
                    # Recursive scan for ROM extensions
                    for rom in self._scan_roms(p):
                        file_id = rom.stem
                        size_mb = rom.stat().st_size / (1 << 20)
                        self._file_map[file_id] = (rom.name, str(rom), size_mb)
                        if file_id not in self._queue_order:
                            self._queue_order.append(file_id)
            self._refresh_tree()
            if self.on_files_added:
                self.on_files_added([self._file_map[fid][1] for fid in self._queue_order])

    def remove_file(self, file_id: str) -> None:
        """Remove a file from the queue."""
        with self._lock:
            if file_id in self._queue_order:
                self._queue_order.remove(file_id)
            if file_id in self._file_map:
                del self._file_map[file_id]
            self._refresh_tree()
            if self.on_file_removed:
                self.on_file_removed(file_id)
            if self.on_order_changed:
                self.on_order_changed([self._file_map[fid][1] for fid in self._queue_order])

    def get_queue_order(self) -> List[str]:
        """Return list of file paths in current queue order."""
        with self._lock:
            return [self._file_map[fid][1] for fid in self._queue_order]

    def set_queue_order(self, file_paths: List[str]) -> None:
        """Programmatically set the queue order."""
        with self._lock:
            # Rebuild _queue_order based on paths
            new_order = []
            for path in file_paths:
                for file_id, (name, fpath, size) in self._file_map.items():
                    if fpath == path:
                        new_order.append(file_id)
                        break
            if new_order:
                self._queue_order = new_order
                self._refresh_tree()

    def clear_queue(self) -> None:
        """Remove all files from queue."""
        with self._lock:
            self._queue_order.clear()
            self._file_map.clear()
            self._refresh_tree()

    # --------------------------------------------------------------- UI actions

    def _on_add_files(self) -> None:
        from tkinter.filedialog import askopenfilenames
        files = askopenfilenames(
            title="Add ROM Files",
            filetypes=[("All ROMs", "*.iso *.cue *.bin *.chd *.cso *.wad *.cdi *.gdi")],
        )
        if files:
            self.add_files(list(files))

    def _on_add_folder(self) -> None:
        from tkinter.filedialog import askdirectory
        folder = askdirectory(title="Add ROM Folder")
        if folder:
            self.add_files([folder])

    def _on_remove_selected(self) -> None:
        if self._selected_iid:
            self.remove_file(self._selected_iid)

    def _on_tree_select(self, event: tk.Event) -> None:
        sel = self._queue_tree.selection()
        if sel:
            self._selected_iid = sel[0]

    def _move_up(self) -> None:
        if not self._selected_iid:
            return
        with self._lock:
            try:
                idx = self._queue_order.index(self._selected_iid)
                if idx > 0:
                    self._queue_order[idx], self._queue_order[idx - 1] = (
                        self._queue_order[idx - 1],
                        self._queue_order[idx],
                    )
                    self._refresh_tree()
                    if self.on_order_changed:
                        self.on_order_changed(
                            [self._file_map[fid][1] for fid in self._queue_order]
                        )
            except ValueError:
                pass

    def _move_down(self) -> None:
        if not self._selected_iid:
            return
        with self._lock:
            try:
                idx = self._queue_order.index(self._selected_iid)
                if idx < len(self._queue_order) - 1:
                    self._queue_order[idx], self._queue_order[idx + 1] = (
                        self._queue_order[idx + 1],
                        self._queue_order[idx],
                    )
                    self._refresh_tree()
                    if self.on_order_changed:
                        self.on_order_changed(
                            [self._file_map[fid][1] for fid in self._queue_order]
                        )
            except ValueError:
                pass

    # ------------------------------------------------------------ helpers

    def _refresh_tree(self) -> None:
        """Rebuild the listbox from current queue order."""
        for row in self._queue_tree.get_children():
            self._queue_tree.delete(row)
        for file_id in self._queue_order:
            if file_id in self._file_map:
                name, path, size = self._file_map[file_id]
                self._queue_tree.insert(
                    "", tk.END, iid=file_id,
                    values=(name, "⏳ Queued", f"{size:.1f}"),
                )

    @staticmethod
    def _scan_roms(directory: Path) -> List[Path]:
        """Recursively find ROM files in a directory."""
        roms = []
        rom_exts = {
            ".iso", ".cue", ".bin", ".chd", ".cso", ".wad", ".cdi", ".gdi",
            ".img", ".nrg", ".mdf", ".ape", ".flac", ".wav", ".zip", ".7z",
        }
        for item in directory.rglob("*"):
            if item.is_file() and item.suffix.lower() in rom_exts:
                roms.append(item)
        return roms


def create_dragdrop_queue_panel(
    parent: tk.Widget,
    on_files_added: Optional[Callable[[List[str]], None]] = None,
    on_order_changed: Optional[Callable[[List[str]], None]] = None,
    on_file_removed: Optional[Callable[[str], None]] = None,
) -> DragDropQueuePanel:
    """Factory for DragDropQueuePanel."""
    return DragDropQueuePanel(
        parent,
        on_files_added=on_files_added,
        on_order_changed=on_order_changed,
        on_file_removed=on_file_removed,
    )
