"""
PhotosUploader.py
A GUI application for photo processing workflows.
Requires: pip install Pillow tkinterdnd2 piexif
"""

import os
import re
import json
import shutil
import tempfile
import threading
import logging
import tkinter as tk
from typing import Any
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


try:
    from PIL import Image, ImageTk, ExifTags, IptcImagePlugin
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow not installed. Run: pip install Pillow")


try:
    import piexif
    PIEXIF_AVAILABLE = True
except ImportError:
    piexif = None  # type: ignore[assignment]
    PIEXIF_AVAILABLE = False
    print("WARNING: piexif not installed. Run: pip install piexif")

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    print("WARNING: tkinterdnd2 not installed. Run: pip install tkinterdnd2")

import DownloadAlbumStructure

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# State persistence
# ---------------------------------------------------------------------------
STATE_FILE = DownloadAlbumStructure.get_data_dir() / "PhotosUploader State.json"


def _window_is_on_a_monitor(hwnd: int) -> bool:
    """Return True if any part of the window is on a connected monitor."""
    try:
        import ctypes
        MONITOR_DEFAULTTONULL = 0
        return bool(ctypes.windll.user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL))
    except Exception:
        return True  # assume visible if the API fails


def load_state() -> dict[str, Any]:
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_state(state: dict[str, Any]) -> None:
    with open(STATE_FILE, "w", encoding='utf-8') as f:
        json.dump(state, f, indent=2)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
                    '.tif', '.webp', '.heic', '.heif'}

EXIF_TAG_NAMES = {
    'DateTime': 'Date/Time',
    'DateTimeOriginal': 'Date Created',
    'Make': 'Camera Make',
    'Model': 'Camera Model',
    'ImageWidth': 'Width',
    'ImageLength': 'Height',
    'GPSInfo': 'GPS Info',
    'Software': 'Software',
    'Artist': 'Artist',
    'Copyright': 'Copyright',
    'ImageDescription': 'Description',
    'XResolution': 'X Resolution',
    'YResolution': 'Y Resolution',
    'Orientation': 'Orientation',
    'ExposureTime': 'Exposure Time',
    'FNumber': 'F-Number',
    'ISOSpeedRatings': 'ISO Speed',
    'FocalLength': 'Focal Length',
}

CUSTOM_FIELDS = [
    ('output_filename', 'Output Filename'),
    ('photo_source', 'Photographer/Source'),
    ('date_of_photo', 'Date of Photo'),
    ('comments', 'Caption'),
    ('tags', 'Tags (comma-separated)'),
]


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def is_image(path: str) -> bool:
    return Path(path).suffix.lower() in IMAGE_EXTENSIONS


def parse_dnd_paths(data: str) -> list[str]:
    """Parse drag-and-drop path string into a list of file paths."""
    paths = []
    # tkinterdnd2 wraps paths with spaces in braces
    raw = data.strip()
    i = 0
    while i < len(raw):
        if raw[i] == '{':
            end = raw.index('}', i)
            paths.append(raw[i+1:end])
            i = end + 1
        else:
            # Space-separated
            end = raw.find(' ', i)
            if end == -1:
                paths.append(raw[i:])
                break
            paths.append(raw[i:end])
            i = end + 1
    return [p for p in paths if p]


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class PhotosUploader:
    def __init__(self, root):
        self.root = root
        self.root.title("PhotosUploader")
        self.root.geometry("1400x820")
        self.root.minsize(1000, 600)

        # State
        self.input_paths = []       # list of str
        self.current_photo = None   # str path
        self.photo_image = None     # ImageTk reference
        self._cached_image = None           # PIL Image (full-size, not thumbnailed)
        self._cached_image_path = None      # path matching _cached_image
        self.custom_data = {}       # path -> dict of custom field values
        self._field_links = []      # list of link descriptors; see _register_field_link
        self.persist_vars = {}      # {field_key: BooleanVar} for persist checkboxes
        self._exif_data   = {}      # display-name -> value for the current photo
        self._field_validity = {'date': False, 'caption': False, 'filename': False}
        self.status_var = tk.StringVar(value="Ready.")
        self.upload_album_var = tk.StringVar(value="(none)")
        self.upload_album_id  = 0
        self.state_data = load_state()
        if self.state_data.get("upload_album"):
            self.upload_album_var.set(self.state_data["upload_album"])
        if self.state_data.get("upload_album_id"):
            self.upload_album_id = int(self.state_data["upload_album_id"])

        self._build_ui()
        self._validate_caption_field()         # start pink since caption is empty
        self._validate_date_field()            # start pink since date is empty
        self._validate_output_filename_field() # start pink since filename is empty
        self._bind_shortcuts()
        self.root.update_idletasks()
        self._restore_geometry()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -----------------------------------------------------------------------
    # UI Construction
    # -----------------------------------------------------------------------
    def _build_ui(self):
        # ── Top toolbar ──────────────────────────────────────────────────
        toolbar = ttk.Frame(self.root, padding=4)
        toolbar.pack(side="top", fill="x")

        ttk.Button(toolbar, text="Add New Album", command=self._add_new_album).pack(side="left", padx=2)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(toolbar, text="Download Album Hierarchy", command=self._download_album_hierarchy).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Download File List", command=self._download_file_list).pack(side="left", padx=2)

        # ── Main three-panel area ─────────────────────────────────────────
        main_pane = ttk.PanedWindow(self.root, orient="horizontal")
        main_pane.pack(side="top", fill="both", expand=True, padx=4, pady=4)

        # LEFT: Input queue
        left_frame = self._build_queue_panel(main_pane, "Input Queue")
        main_pane.add(left_frame, weight=1)

        # CENTER: Photo viewer + fields
        center_frame = self._build_center_panel(main_pane)
        main_pane.add(center_frame, weight=3)

        # ── Status bar ───────────────────────────────────────────────────
        status_bar = ttk.Frame(self.root, relief="sunken")
        status_bar.pack(side="bottom", fill="x")
        ttk.Label(status_bar, textvariable=self.status_var, anchor="w").pack(
            side="left", padx=6, pady=2)

    def _build_queue_panel(self, parent, title: str) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text=title, padding=4)

        # Buttons
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill="x", pady=(0, 4))

        ttk.Button(btn_row, text="Add…", command=self.add_photos_dialog).pack(side="left", padx=2)
        self.input_remove_btn = ttk.Button(btn_row, text="Remove",
                                           command=self.remove_selected_input,
                                           state="disabled")
        self.input_remove_btn.pack(side="left", padx=2)
        ttk.Button(btn_row, text="Remove All",
                   command=self.remove_all_input).pack(side="left", padx=2)
        ttk.Button(btn_row, text="↑", width=2, command=lambda: self._move_item(self.input_list, -1)).pack(side="left")
        ttk.Button(btn_row, text="↓", width=2, command=lambda: self._move_item(self.input_list, 1)).pack(side="left")
        ttk.Button(btn_row, text="Sort", command=self.sort_input).pack(side="left", padx=2)

        # Count label
        self.input_count_var = tk.StringVar(value="0 items")  # kept in sync by _update_counts
        ttk.Label(btn_row, textvariable=self.input_count_var).pack(side="right", padx=4)

        # Listbox with scrollbars
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True)

        yscroll = ttk.Scrollbar(list_frame, orient="vertical")
        xscroll = ttk.Scrollbar(list_frame, orient="horizontal")

        lb = tk.Listbox(list_frame, selectmode="extended",
                        yscrollcommand=yscroll.set,
                        xscrollcommand=xscroll.set,
                        activestyle='dotbox',
                        font=('Consolas', 9))
        yscroll.config(command=lb.yview)
        xscroll.config(command=lb.xview)

        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        lb.pack(side="left", fill="both", expand=True)

        self.input_list = lb
        lb.bind('<<ListboxSelect>>', self._on_input_select)
        if DND_AVAILABLE:
            lb.drop_target_register(DND_FILES)
            lb.dnd_bind('<<Drop>>', self._on_drop)

        return frame

    def _build_center_panel(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent)

        # Vertical paned window: viewer (top, smaller) / fields row (bottom)
        vpane = ttk.PanedWindow(frame, orient="vertical")
        vpane.pack(fill="both", expand=True)

        # ── Photo display (top pane) ──────────────────────────────────────
        viewer_frame = ttk.LabelFrame(vpane, text="Photo Viewer", padding=4)
        vpane.add(viewer_frame, weight=1)

        # ── Top bar: Album selection ───────────────────────────────────────
        top_bar = ttk.Frame(viewer_frame)
        top_bar.pack(fill="x", pady=(0, 2))

        upload_to_label = ttk.Label(top_bar, text="Upload to:")
        upload_to_label.pack(side="left", padx=(2, 4))

        album_display_var = tk.StringVar(value="(none)")
        album_label = ttk.Label(top_bar, textvariable=album_display_var,
                                foreground='blue', anchor="w")
        album_label.pack(side="left", padx=(0, 4), fill="x", expand=True)

        def _refresh_album_display(*_):
            from tkinter.font import nametofont
            full = self.upload_album_var.get()
            # Set color: blue if album selected, gray if none
            color = 'blue' if full and full != '(none)' else 'gray'
            album_label.configure(foreground=color)
            try:
                font = nametofont("TkDefaultFont")
                avail = album_label.winfo_width() - 4
            except Exception:
                album_display_var.set(full)
                return
            if avail < 20 or not full or full == "(none)":
                album_display_var.set(full)
                return
            if font.measure(full) <= avail:
                album_display_var.set(full)
                return
            for i in range(len(full)):
                candidate = "\u2026" + full[i:]
                if font.measure(candidate) <= avail:
                    album_display_var.set(candidate)
                    return
            album_display_var.set("\u2026")

        album_label.bind("<Configure>", _refresh_album_display)
        self.upload_album_var.trace_add("write", _refresh_album_display)

        # ── Button row: Change Upload Album ────────────────────────────────
        button_bar = ttk.Frame(viewer_frame)
        button_bar.pack(fill="x", pady=(0, 4))

        # Spacer frame to align button with album_label above
        spacer = tk.Frame(button_bar)
        spacer.pack(side="left", padx=(2, 4))

        # Measure and sync spacer width to match upload_to_label
        def _sync_widths():
            w = upload_to_label.winfo_reqwidth()
            spacer.configure(width=max(w, 1), height=1)

        button_bar.after(50, _sync_widths)

        ttk.Button(button_bar, text="Change Upload Album",
                   command=self.open_output_folder).pack(side="left", padx=2)

        # ── Left column: nav buttons, filename, dims, path ───────────────
        left_col = ttk.Frame(viewer_frame)
        left_col.pack(side="left", fill="y", padx=(0, 6))

        self.upload_photo_btn = tk.Button(left_col, text="⬆ Upload to Piwigo",
                                          command=self._upload_current_photo,
                                          background="#add8e6",
                                          font=("TkDefaultFont", 10, "bold"),
                                          state="disabled")
        self.upload_photo_btn.pack(fill="x", pady=(0, 4))

        nav = ttk.Frame(left_col)
        nav.pack(fill="x", pady=(0, 6))
        self.skip_btn = ttk.Button(nav, text="⊘ Skip",
                                   command=self._skip_photo,
                                   state="disabled")
        self.skip_btn.pack(side="left", padx=2)
        self.revert_btn = ttk.Button(nav, text="↺ Revert",
                                     command=self._revert_photo,
                                     state="disabled")
        self.revert_btn.pack(side="left", padx=2)

        self.photo_label_var = tk.StringVar(value="No photo selected")
        ttk.Label(left_col, textvariable=self.photo_label_var,
                  font=('TkDefaultFont', 9, 'italic'),
                  anchor="w").pack(fill="x", pady=(0, 2))

        self.photo_dim_var = tk.StringVar(value="")
        ttk.Label(left_col, textvariable=self.photo_dim_var,
                  anchor="w").pack(fill="x", pady=(0, 6))

        # Path display — wraps to the actual column width
        self.path_var = tk.StringVar(value="")
        path_label = ttk.Label(left_col, textvariable=self.path_var,
                               font=('TkDefaultFont', 9),
                               anchor="nw", justify="left", wraplength=220)
        path_label.pack(fill="x")

        def _update_wraplength(event):
            path_label.configure(wraplength=max(event.width - 4, 50))
        path_label.bind('<Configure>', _update_wraplength)

        # ── Canvas — right side, narrower ────────────────────────────────
        self.canvas = tk.Canvas(viewer_frame, bg='#1a1a1a', cursor='crosshair',
                                height=200)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind('<Configure>', self._on_canvas_resize)

        # ── Bottom pane: Custom Fields / EXIF stacked vertically ─────────
        hpane = ttk.PanedWindow(vpane, orient="vertical")
        vpane.add(hpane, weight=2)

        # ── Custom Fields (left) ──────────────────────────────────────────
        custom_frame = ttk.LabelFrame(hpane, text="Custom Fields", padding=6)
        hpane.add(custom_frame, weight=1)

        # "Persist" label above the checkbox column
        ttk.Label(custom_frame, text="Persist", anchor="center").grid(
            row=0, column=1, sticky="ew", pady=2)

        self.custom_vars = {}
        self.persist_vars = {}  # Track which fields should persist across photos
        for i, (key, label) in enumerate(CUSTOM_FIELDS):
            row = i + 1  # Offset by 1 to account for header row
            ttk.Label(custom_frame, text=label + ":", width=22, anchor="e").grid(
                row=row, column=0, sticky="e", pady=2, padx=(0, 4))

            # Checkbox for persist (output_filename has no persist option)
            if key != 'output_filename':
                persist_var = tk.BooleanVar(value=False)
                ttk.Checkbutton(custom_frame, variable=persist_var).grid(
                    row=row, column=1, sticky="w", padx=4, pady=2)
                self.persist_vars[key] = persist_var

            if key == 'comments':
                txt = tk.Text(custom_frame, height=3, width=40, wrap="word",
                              font=('TkDefaultFont', 9))
                txt.grid(row=row, column=2, sticky="ew", pady=2)
                self.custom_vars[key] = txt
            else:
                var = tk.StringVar()
                if key == 'date_of_photo':
                    entry = tk.Entry(custom_frame, textvariable=var, width=40)
                    self.date_entry = entry
                elif key == 'output_filename':
                    entry = tk.Entry(custom_frame, textvariable=var, width=40)
                    self.output_filename_entry = entry
                else:
                    entry = ttk.Entry(custom_frame, textvariable=var, width=40)
                entry.grid(row=row, column=2, sticky="ew", pady=2)
                self.custom_vars[key] = var
        custom_frame.columnconfigure(2, weight=1)

        # ── Bidirectional field links (custom field ↔ EXIF ↔ IPTC) ──────────
        self._register_field_link(
            custom_key  = 'photo_source',
            exif_key    = 'Artist',
            iptc_tag    = (2, 80),   # By-line
            validate_fn = None,      # any string is valid
        )
        self._register_field_link(
            custom_key  = 'date_of_photo',
            exif_key    = 'Date Created',
            iptc_tag    = (2, 55),   # Date Created
            validate_fn = self._parse_date,
        )
        self.custom_vars['date_of_photo'].trace_add(
            'write', lambda *_: self._validate_date_field()
        )
        self.custom_vars['output_filename'].trace_add(
            'write', lambda *_: self._validate_output_filename_field()
        )
        self._register_field_link(
            custom_key  = 'tags',
            exif_key    = None,      # no EXIF tag for keywords
            iptc_tag    = (2, 25),   # Keywords (multi-valued)
            validate_fn = None,      # any string is valid
        )

        # comments uses tk.Text (not StringVar) so it cannot use _register_field_link
        def _on_comments_changed(_event=None):
            self._validate_caption_field()
            widget = self.custom_vars['comments']
            value = widget.get('1.0', "end").strip()
            self._exif_data['Description'] = value
            self._refresh_exif_tree()
            widget.edit_modified(False)

        self.custom_vars['comments'].bind('<<Modified>>', _on_comments_changed)

        # ── EXIF / Metadata (right) ───────────────────────────────────────
        exif_frame = ttk.LabelFrame(hpane, text="EXIF / Metadata", padding=6)
        hpane.add(exif_frame, weight=1)

        # Edit row packed first (side=BOTTOM) so the tree fills remaining space
        edit_row = ttk.Frame(exif_frame)
        edit_row.pack(side="bottom", fill="x", pady=(4, 0))
        ttk.Label(edit_row, text="Edit selected value:").pack(side="left", padx=(0, 4))
        self.exif_edit_var = tk.StringVar()
        ttk.Entry(edit_row, textvariable=self.exif_edit_var, width=30).pack(
            side="left", padx=2)
        ttk.Button(edit_row, text="Apply", command=self._apply_exif_edit).pack(
            side="left", padx=2)

        tree_frame = ttk.Frame(exif_frame)
        tree_frame.pack(side="top", fill="both", expand=True)

        exif_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        self.exif_tree = ttk.Treeview(tree_frame, columns=('key', 'value'),
                                      show='headings',
                                      yscrollcommand=exif_scroll.set)
        exif_scroll.config(command=self.exif_tree.yview)
        self.exif_tree.heading('key', text='Field')
        self.exif_tree.heading('value', text='Value')
        self.exif_tree.column('key', width=160)
        self.exif_tree.column('value', width=260)
        self.exif_tree.pack(side="left", fill="both", expand=True)
        exif_scroll.pack(side="left", fill="y")
        self.exif_tree.bind('<<TreeviewSelect>>', self._on_exif_select)

        return frame

    # -----------------------------------------------------------------------
    # Drop-and-drop handler
    # -----------------------------------------------------------------------
    def _on_drop(self, event):
        paths = parse_dnd_paths(event.data)
        batch_state = {}
        added = 0
        for p in paths:
            p = p.strip()
            if os.path.isdir(p):
                added += self._add_folder(p, batch_state)
            elif is_image(p):
                if self._add_single_image(p, batch_state):
                    added += 1
        self._update_counts()
        self.set_status(f"Dropped {added} image(s) into input queue.")

    # -----------------------------------------------------------------------
    # Add photos / folders
    # -----------------------------------------------------------------------
    def add_photos_dialog(self):
        filetypes = [
            ("Image files", " ".join(f"*{e}" for e in IMAGE_EXTENSIONS)),
            ("All files", "*.*"),
        ]
        initialdir = self.state_data.get("input_path") or None
        paths = filedialog.askopenfilenames(title="Select Images", filetypes=filetypes,
                                            initialdir=initialdir)
        if paths:
            self.state_data["input_path"] = os.path.dirname(paths[-1])
        batch_state = {}
        added = 0
        for p in paths:
            if self._add_single_image(p, batch_state):
                added += 1
        self._update_counts()
        self.set_status(f"Added {added} image(s) to input queue.")


    def _add_folder(self, folder: str, batch_state: dict[str, str] | None = None) -> int:
        if batch_state is None:
            batch_state = {}
        added = 0
        for root_dir, _, files in os.walk(folder):
            for f in sorted(files):
                if Path(f).suffix.lower() in IMAGE_EXTENSIONS:
                    full = os.path.join(root_dir, f)
                    if self._add_single_image(full, batch_state):
                        added += 1
        return added

    def _find_name_conflict(self, path: str):
        """Return the existing input path whose basename matches path, or None."""
        name = os.path.basename(path)
        for existing in self.input_paths:
            if os.path.basename(existing) == name:
                return existing
        return None

    def _show_conflict_dialog(self, new_path: str, existing_path: str,
                              batch_state: dict[str, str]) -> str:
        """Ask the user how to resolve a filename conflict.

        Returns 'skip' or 'replace'.  Sets batch_state['all'] to the chosen
        base action when the user picks Skip All or Replace All.
        """
        if batch_state.get('all'):
            return batch_state['all']

        name = os.path.basename(new_path)
        dlg = tk.Toplevel(self.root)
        dlg.title("Filename Conflict")
        dlg.resizable(False, False)
        dlg.grab_set()

        dlg.geometry("500x150")
        self._center_dialog(dlg)

        msg = (f'A file named "{name}" is already in the input queue.\n\n'
               f'Existing:  {existing_path}\n'
               f'New:         {new_path}')
        ttk.Label(dlg, text=msg, padding=(12, 10, 12, 6),
                  wraplength=476, justify="left").pack()

        result = tk.StringVar(value='skip')

        btn_frame = ttk.Frame(dlg, padding=(12, 0, 12, 12))
        btn_frame.pack()

        def choose(action):
            result.set(action)
            if action in ('skip_all', 'replace_all'):
                batch_state['all'] = action.replace('_all', '')
            dlg.destroy()

        ttk.Button(btn_frame, text="Skip",
                   command=lambda: choose('skip')).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Skip All",
                   command=lambda: choose('skip_all')).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Replace",
                   command=lambda: choose('replace')).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Replace All",
                   command=lambda: choose('replace_all')).pack(side="left", padx=4)

        dlg.wait_window()
        action = result.get()
        return 'skip' if action in ('skip', 'skip_all') else 'replace'

    def _add_single_image(self, path: str, batch_state: dict[str, str]) -> bool:
        """Add one image to the input queue, handling name conflicts.

        Returns True if the image was added.
        """
        if path in self.input_paths:
            return False  # exact duplicate — skip silently

        existing = self._find_name_conflict(path)
        if existing:
            action = self._show_conflict_dialog(path, existing, batch_state)
            if action == 'skip':
                return False
            # replace: remove the existing entry
            idx = self.input_paths.index(existing)
            self.input_paths.pop(idx)
            self.input_list.delete(idx)

        self.input_paths.append(path)
        self.input_list.insert("end", os.path.basename(path))
        return True

    # -----------------------------------------------------------------------
    # Queue management
    # -----------------------------------------------------------------------
    def remove_selected_input(self):
        sel = list(self.input_list.curselection())
        if not sel:
            return
        first_removed = sel[0]
        for i in reversed(sel):
            self.input_paths.pop(i)
            self.input_list.delete(i)
        self._update_counts()
        if self.input_paths:
            next_idx = min(first_removed, len(self.input_paths) - 1)
            self.input_list.selection_clear(0, "end")
            self.input_list.selection_set(next_idx)
            self.input_list.see(next_idx)
            self._save_current_custom_fields()
            self.current_photo = self.input_paths[next_idx]
            self._load_photo(self.current_photo)
        else:
            self._clear_viewer()

    def remove_all_input(self):
        self.input_paths.clear()
        self.input_list.delete(0, "end")
        self._clear_viewer()
        self._update_counts()

    def sort_input(self):
        if not self.input_paths:
            return
        self.input_paths.sort(key=lambda p: os.path.basename(p).lower())
        self.input_list.delete(0, "end")
        for p in self.input_paths:
            self.input_list.insert("end", os.path.basename(p))
        # Re-select the current photo if there is one
        if self.current_photo and self.current_photo in self.input_paths:
            idx = self.input_paths.index(self.current_photo)
            self.input_list.selection_set(idx)
            self.input_list.see(idx)

    def _move_item(self, listbox, direction):
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + direction
        if j < 0 or j >= listbox.size():
            return
        self.input_paths[i], self.input_paths[j] = self.input_paths[j], self.input_paths[i]
        text_i = listbox.get(i)
        text_j = listbox.get(j)
        listbox.delete(i)
        listbox.insert(i, text_j)
        listbox.delete(j)
        listbox.insert(j, text_i)
        listbox.selection_clear(0, "end")
        listbox.selection_set(j)

    def open_output_folder(self):
        def on_select(album_id, fullname):
            self.upload_album_id = album_id
            self.upload_album_var.set(fullname)
            self.set_status(f"Upload album set to: {fullname}")
        DownloadAlbumStructure.pick_album(self.root, self.set_status, on_select)

    # -----------------------------------------------------------------------
    # Photo selection & display
    # -----------------------------------------------------------------------
    def _on_input_select(self, event):
        sel = self.input_list.curselection()
        if sel:
            self._save_current_custom_fields()
            self.current_photo = self.input_paths[sel[0]]
            self._load_photo(self.current_photo)
        else:
            self._update_button_states()

    def _skip_photo(self):
        """Skip the current photo and load the next one from the input queue."""
        if not self.current_photo or not self.input_paths:
            return
        try:
            idx = self.input_paths.index(self.current_photo)
        except ValueError:
            return

        # Remove the current photo from the input queue
        self.input_paths.pop(idx)
        self.input_list.delete(idx)
        self._update_counts()

        # Load the next photo if one exists
        if self.input_paths:
            next_idx = min(idx, len(self.input_paths) - 1)
            self.input_list.selection_clear(0, "end")
            self.input_list.selection_set(next_idx)
            self.input_list.see(next_idx)
            self.current_photo = self.input_paths[next_idx]
            self._load_photo(self.current_photo)
        else:
            self._clear_viewer()
            self.set_status("All photos have been skipped.")
        self._persist_state()

    def _revert_photo(self):
        """Create a new, unchanged copy of the photo (prefer original from input queue).

        Finds a matching filename in the input queue and copies that file; if none
        is found, duplicates the currently displayed file. The new file is added
        to the input queue and selected for viewing.
        """
        if not self.current_photo:
            self.set_status("No photo selected to revert.")
            return

        base_name = os.path.basename(self.current_photo)

        # Prefer original from input queue (left column)
        source = None
        for p in self.input_paths:
            if os.path.basename(p) == base_name:
                source = p
                break

        # If not found, warn that the copy may not be clean and ask to proceed
        if not source:
            if not messagebox.askyesno(
                "Revert",
                "No unmodified original was found in the input queue.\n\n"
                "The current file may already have had metadata written to it.\n\n"
                "Copy it anyway?",
                parent=self.root,
            ):
                return
            source = self.current_photo

        src_dir = os.path.dirname(source) or '.'
        name, ext = os.path.splitext(base_name)
        # Generate a unique filename in the same directory
        i = 1
        while True:
            candidate = f"{name} (revert){'' if i == 1 else f' {i}'}{ext}"
            dest = os.path.join(src_dir, candidate)
            if not os.path.exists(dest):
                break
            i += 1

        try:
            shutil.copy2(source, dest)
        except Exception as e:
            messagebox.showerror("Revert Failed",
                                 f"Failed to create revert copy:\n{e}")
            self.set_status("Revert failed.")
            return

        # Add to input queue and select
        if dest not in self.input_paths:
            self.input_paths.append(dest)
            self.input_list.insert("end", os.path.basename(dest))
        new_idx = self.input_paths.index(dest)
        self.input_list.selection_clear(0, "end")
        self.input_list.selection_set(new_idx)
        self.input_list.see(new_idx)
        self._update_counts()
        self.set_status(f"Revert: added {os.path.basename(dest)} to input queue.")
        self._save_current_custom_fields()
        self.current_photo = dest
        self._load_photo(self.current_photo)

    def _load_photo(self, path: str):
        self.photo_label_var.set(os.path.basename(path))
        self._display_photo(path)
        self._load_exif(path)
        self._load_iptc(path)
        self._load_custom_fields(path)
        self.path_var.set(path)
        self._validate_caption_field()
        self._validate_date_field()
        self._validate_output_filename_field()

    def _display_photo(self, path: str):
        if not PIL_AVAILABLE:
            self.canvas.delete('all')
            self.canvas.create_text(200, 150, text="Pillow not installed", fill='white')
            return
        try:
            if path != self._cached_image_path:
                self._cached_image = Image.open(path)  # type: ignore[possibly-undefined]
                self._cached_image_path = path
                self.photo_dim_var.set(
                    f"{self._cached_image.width} × {self._cached_image.height} px"
                    f"  |  {self._cached_image.mode}")
            img = self._cached_image
            if img is None:
                return
            cw = max(self.canvas.winfo_width(), 100)
            ch = max(self.canvas.winfo_height(), 100)
            thumb = img.copy()
            thumb.thumbnail((cw, ch), Image.Resampling.LANCZOS)  # type: ignore[possibly-undefined]
            self.photo_image = ImageTk.PhotoImage(thumb)
            self.canvas.delete('all')
            self.canvas.create_image(cw // 2, ch // 2,
                                     anchor="center",
                                     image=self.photo_image)
        except Exception as e:
            self._cached_image = None
            self._cached_image_path = None
            self.canvas.delete('all')
            self.canvas.create_text(10, 10, anchor="nw",
                                    text=f"Cannot display image:\n{e}",
                                    fill='red', font=('TkDefaultFont', 10))

    def _on_canvas_resize(self, event):
        if self.current_photo:
            self._display_photo(self.current_photo)

    # -----------------------------------------------------------------------
    # EXIF
    # -----------------------------------------------------------------------
    def _register_field_link(self, custom_key: str, exif_key: str | None,
                             iptc_tag: tuple, validate_fn=None):
        """Register a link between a custom StringVar field, an EXIF
        display-name key, and an IPTC tag tuple (e.g. (2, 80)).

        Changes to the custom field propagate into _exif_data and refresh
        the metadata tree.  validate_fn(value: str) -> truthy is called on
        non-empty values; falsy result suppresses the update.

        Must be called after self.custom_vars has been built.
        Descriptor is also used by _load_iptc for initial population.
        """
        link = {
            'custom_key':  custom_key,
            'exif_key':    exif_key,
            'iptc_tag':    iptc_tag,
            'validate_fn': validate_fn,
            'syncing':     False,
        }
        self._field_links.append(link)

        def _on_custom_changed(*_):
            if link['syncing']:
                return
            value = self.custom_vars[custom_key].get().strip()
            # Empty value always clears the EXIF field; non-empty must pass validation
            if value and validate_fn is not None and not validate_fn(value):
                return
            link['syncing'] = True
            try:
                if exif_key is not None:
                    self._exif_data[exif_key] = value
                self._refresh_exif_tree()
            finally:
                link['syncing'] = False

        self.custom_vars[custom_key].trace_add('write', _on_custom_changed)

    def _load_exif(self, path: str):
        self.exif_tree.delete(*self.exif_tree.get_children())
        self._exif_data = {}
        if not PIL_AVAILABLE:
            return
        try:
            img = (self._cached_image if self._cached_image_path == path
                   else Image.open(path))  # type: ignore[possibly-undefined]
            if img is None:
                return
            raw = img._getexif() if hasattr(img, '_getexif') else None
            if raw:
                for tag_id, val in raw.items():
                    tag = ExifTags.TAGS.get(tag_id, str(tag_id))
                    if isinstance(val, bytes):
                        val = val.decode('utf-8', errors='replace')
                    elif isinstance(val, tuple) and len(val) == 2:
                        val = f"{val[0]}/{val[1]} ({val[0]/val[1]:.4f})"
                    display_tag = EXIF_TAG_NAMES.get(tag, tag)
                    self._exif_data[display_tag] = str(val)
                    self.exif_tree.insert('', "end", values=(display_tag, str(val)[:120]))
            else:
                self.exif_tree.insert('', "end", values=('(No EXIF data)', ''))
        except Exception as e:
            self.exif_tree.insert('', "end", values=('Error reading EXIF', str(e)))

    _MONTH_MAP: dict[str, int] = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12,
    }

    @staticmethod
    def _expand_year(yy: int) -> int:
        """00–35 → 2000+yy; 36–99 → 1900+yy (historical photos)."""
        return (2000 if yy <= 35 else 1900) + yy

    def _parse_date(self, text: str) -> datetime | None:
        """Parse a user-supplied date string.  Returns a datetime or None.

        Accepted formats (mm/dd both accept 1 or 2 digits; year accepts
        4 digits or 2 digits with century inferred by _expand_year):
          mm/dd/yyyy   mm/dd/yy   (dd/mm tried as fallback)
          dd month yyyy           (e.g. 15 Dec 2023)
          month dd, yyyy          (e.g. December 15, 2023)
          mm/yyyy   mm/yy         (day defaults to 1)
          month yyyy              (e.g. Dec 2023 — day defaults to 1)
          yyyy                    (month/day default to 1)
          EXIF: YYYY:MM:DD HH:MM:SS  and  YYYY:MM:DD
        """
        text = re.sub(r'\s+', ' ', text.strip())
        if not text:
            return None

        def yr(s: str) -> int:
            v = int(s)
            return v if len(s) == 4 else self._expand_year(v)

        def make(year: int, month: int, day: int = 1) -> datetime | None:
            try:
                return datetime(year, month, day)
            except ValueError:
                return None

        # EXIF native: YYYY:MM:DD HH:MM:SS  or  YYYY:MM:DD
        for fmt in ('%Y:%m:%d %H:%M:%S', '%Y:%m:%d'):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass

        # mm/dd/yyyy  or  mm-dd-yyyy  (with optional 2-digit year)
        m = re.fullmatch(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{4}|\d{2})', text)
        if m:
            a, b, y = int(m.group(1)), int(m.group(2)), yr(m.group(3))
            return make(y, a, b) or make(y, b, a)  # mm/dd then dd/mm fallback

        # mm/yyyy  or  mm/yy  (month + year, no day)
        m = re.fullmatch(r'(\d{1,2})[/\-](\d{4}|\d{2})', text)
        if m:
            return make(yr(m.group(2)), int(m.group(1)))

        # yyyy alone
        m = re.fullmatch(r'(\d{4})', text)
        if m:
            return make(int(m.group(1)), 1)

        # dd month yyyy  — e.g. "15 December 2023" or "15 Dec 23"
        m = re.fullmatch(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4}|\d{2})', text)
        if m:
            month = self._MONTH_MAP.get(m.group(2).lower())
            if month:
                return make(yr(m.group(3)), month, int(m.group(1)))

        # month dd, yyyy  — e.g. "December 15, 2023" or "Dec 15 23"
        m = re.fullmatch(r'([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4}|\d{2})', text)
        if m:
            month = self._MONTH_MAP.get(m.group(1).lower())
            if month:
                return make(yr(m.group(3)), month, int(m.group(2)))

        # month yyyy  — e.g. "December 2023" or "Dec 23"
        m = re.fullmatch(r'([A-Za-z]+)\s+(\d{4}|\d{2})', text)
        if m:
            month = self._MONTH_MAP.get(m.group(1).lower())
            if month:
                return make(yr(m.group(2)), month)

        return None

    def _load_iptc(self, path: str):
        """Read IPTC Date Created (2#055) and populate the date_of_photo field.

        If the IPTC tag is missing but DateTimeOriginal is present in EXIF,
        the EXIF value is used to fill in the IPTC date (in memory).
        Ensures 'Date Created' appears in _exif_data so the tree always shows
        the row and bidirectional sync works.
        """
        if not PIL_AVAILABLE:
            return
        iptc = {}
        try:
            img = (self._cached_image if self._cached_image_path == path
                   else Image.open(path))  # type: ignore[possibly-undefined]
            if img is not None:
                iptc = IptcImagePlugin.getiptcinfo(img) or {}  # type: ignore[possibly-undefined]
        except Exception:
            pass

        for link in self._field_links:
            # Read IPTC tag — may be a single bytes value or a list (e.g. Keywords)
            raw = iptc.get(link['iptc_tag'], b'')
            if isinstance(raw, list):
                # Multi-valued: decode each entry and join with ", "
                parts = [bytes(r).decode('utf-8', errors='replace').strip()
                         for r in raw if r]
                value = ', '.join(p for p in parts if p)
            else:
                value = bytes(raw).decode('utf-8', errors='replace').strip() if raw else ''

            # Fall back to EXIF if IPTC tag is absent and an EXIF key is defined
            if not value and link['exif_key'] is not None:
                value = self._exif_data.get(link['exif_key'], '')

            # Keep _exif_data in sync so the tree always shows the row
            if link['exif_key'] is not None:
                self._exif_data[link['exif_key']] = value

            # Don't overwrite a persisted field
            if self.persist_vars[link['custom_key']].get():
                continue

            # Populate the custom field under the reentrancy guard
            link['syncing'] = True
            try:
                self.custom_vars[link['custom_key']].set(value)
            finally:
                link['syncing'] = False

        self._refresh_exif_tree()

    def _on_exif_select(self, event):
        sel = self.exif_tree.selection()
        if sel:
            vals = self.exif_tree.item(sel[0], 'values')
            if vals:
                # Read the full (untruncated) value from _exif_data, not the display cell
                self.exif_edit_var.set(self._exif_data.get(vals[0], vals[1] if len(vals) > 1 else ''))

    def _refresh_exif_tree(self, reselect_key: str | None = None):
        """Rebuild the EXIF treeview from self._exif_data in place."""
        self.exif_tree.delete(*self.exif_tree.get_children())
        reselect_iid = None
        for key, val in self._exif_data.items():
            iid = self.exif_tree.insert('', "end", values=(key, str(val)[:120]))
            if key == reselect_key:
                reselect_iid = iid
        if reselect_iid:
            self.exif_tree.selection_set(reselect_iid)
            self.exif_tree.see(reselect_iid)

    def _apply_exif_edit(self):
        sel = self.exif_tree.selection()
        if not sel:
            self.set_status("Select an EXIF row to edit first.")
            return
        item = sel[0]
        vals = self.exif_tree.item(item, 'values')
        new_val = self.exif_edit_var.get()
        if vals:
            self._exif_data[vals[0]] = new_val
        self._refresh_exif_tree(reselect_key=vals[0] if vals else "")
        self.set_status(f"EXIF field '{vals[0]}' updated (pending save).")

    # -----------------------------------------------------------------------
    # Custom fields
    # -----------------------------------------------------------------------
    def _load_custom_fields(self, path: str):
        # Keys populated from IPTC on every load — don't clear these when
        # there is no saved data, as _load_iptc has already set them.
        iptc_keys = {link['custom_key'] for link in self._field_links}

        data = self.custom_data.get(path)
        for key, widget in self.custom_vars.items():
            if key == 'output_filename':
                continue  # handled separately below
            # Never overwrite a persisted field
            if self.persist_vars[key].get():
                continue
            if data is None:
                # No saved data: clear fields that have no IPTC source so
                # they don't retain values from the previous photo.
                if key not in iptc_keys:
                    if isinstance(widget, tk.Text):
                        widget.delete('1.0', "end")
                    else:
                        widget.set('')
                continue

            val = data.get(key, '')
            if isinstance(widget, tk.Text):
                widget.delete('1.0', "end")
                widget.insert('1.0', val)
            else:
                widget.set(val)

        # Set output filename: use saved value if present, else default to basename.
        saved_fn = data.get('output_filename') if data else None
        self.custom_vars['output_filename'].set(saved_fn if saved_fn else os.path.basename(path))

    def _save_current_custom_fields(self):
        if not self.current_photo:
            return
        data = {}
        for key, widget in self.custom_vars.items():
            if isinstance(widget, tk.Text):
                data[key] = widget.get('1.0', "end").strip()
            else:
                data[key] = widget.get()
        self.custom_data[self.current_photo] = data

    # -----------------------------------------------------------------------
    # Processing
    # -----------------------------------------------------------------------
    # Characters forbidden in Windows filenames; also covers Linux/macOS
    _ILLEGAL_FILENAME_CHARS: re.Pattern[str] = re.compile(r'[\\/:*?"<>|]')
    # Windows reserved names (case-insensitive, with or without extension)
    _RESERVED_NAMES: re.Pattern[str] = re.compile(
        r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.[^.]*)?$', re.IGNORECASE)

    def _load_file_dict(self) -> dict[str, list[dict[str, Any]]]:
        """Load FileDict.json and return it, or an empty dict if unavailable."""
        p = DownloadAlbumStructure._file_index_file()
        if not p.exists():
            return {}
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _write_exif_fields(self, path: str):
        """Write custom field values into the file's EXIF block before upload.

        Fields written:
          Comments         → ImageDescription (0x010E)
          Photographer/Source → Artist        (0x013B)
          Output Filename  → DocumentName     (0x010D)
          Date of Photo    → DateTimeOriginal (0x9003) in YYYY:MM:DD HH:MM:SS format

        Only writes tags that have a non-empty value. Date must be parseable.
        Silently skips if piexif is unavailable or the write fails.
        """
        if not PIEXIF_AVAILABLE:
            return
        assert piexif is not None
        comments = self.custom_vars['comments'].get('1.0', 'end').strip()
        source   = self.custom_vars['photo_source'].get().strip()
        out_name = self.custom_vars['output_filename'].get().strip()
        raw_date = self.custom_vars['date_of_photo'].get().strip()
        parsed_date = self._parse_date(raw_date) if raw_date else None
        if not any([comments, source, out_name, parsed_date]):
            return
        try:
            exif_dict = piexif.load(path)
            ifd0 = exif_dict.setdefault('0th', {})
            exif = exif_dict.setdefault('Exif', {})
            if comments:
                ifd0[piexif.ImageIFD.ImageDescription] = comments.encode('utf-8')
            if source:
                ifd0[piexif.ImageIFD.Artist] = source.encode('utf-8')
            if out_name:
                ifd0[piexif.ImageIFD.DocumentName] = out_name.encode('utf-8')
            if parsed_date:
                # EXIF datetime format: "YYYY:MM:DD HH:MM:SS"
                exif_date = parsed_date.strftime('%Y:%m:%d %H:%M:%S').encode('utf-8')
                exif[piexif.ExifIFD.DateTimeOriginal] = exif_date
                exif[piexif.ExifIFD.DateTimeDigitized] = exif_date
            piexif.insert(piexif.dump(exif_dict), path)
        except Exception as e:
            self.set_status(f"Warning: could not write EXIF to {os.path.basename(path)}: {e}")

    def _clear_viewer(self):
        """Reset the photo viewer, EXIF tree, and all custom fields to empty."""
        self.current_photo = None
        self.photo_image   = None
        self._cached_image = None
        self._cached_image_path = None
        self.photo_label_var.set("No photo selected")
        self.photo_dim_var.set("")
        self.path_var.set("")
        self.canvas.delete('all')
        self._exif_data = {}
        self.exif_tree.delete(*self.exif_tree.get_children())
        for widget in self.custom_vars.values():
            if isinstance(widget, tk.Text):
                widget.delete('1.0', "end")
            else:
                widget.set('')
        self._validate_caption_field()
        self._validate_date_field()
        self._validate_output_filename_field()

    # -----------------------------------------------------------------------
    # Utilities
    # -----------------------------------------------------------------------
    def _validate_output_filename_field(self):
        """Require a legal filename with a recognised image extension; update bg and button state."""
        name = self.custom_vars['output_filename'].get().strip()
        _, dot_ext = os.path.splitext(name)
        ext = dot_ext.lstrip('.').lower()
        valid = (
            bool(name)
            and not self._ILLEGAL_FILENAME_CHARS.search(name)
            and not self._RESERVED_NAMES.match(name)
            and not name.endswith('.')
            and ('.' + ext) in IMAGE_EXTENSIONS
        )
        self._field_validity['filename'] = valid
        self.output_filename_entry.config(bg='pink' if not valid else 'white')
        self._update_button_states()

    def _validate_caption_field(self):
        """Require at least one non-blank character in Caption; update bg and button state."""
        widget = self.custom_vars['comments']
        value = widget.get('1.0', "end").strip()
        valid = len(value) > 0
        self._field_validity['caption'] = valid
        widget.config(bg='pink' if not valid else 'white')
        self._update_button_states()

    def _validate_date_field(self):
        """Validate Date of Photo; color the entry pink and gray the upload button if invalid."""
        text = self.custom_vars['date_of_photo'].get()
        parsed = self._parse_date(text)
        valid = parsed is not None and 1926 <= parsed.year <= 2050
        self._field_validity['date'] = valid
        self.date_entry.config(bg='pink' if not valid else 'white')
        self._update_button_states()

    def _update_button_states(self):
        self.skip_btn.config(state="normal" if self.current_photo and self.input_paths else "disabled")
        self.upload_photo_btn.config(
            state="normal" if self.current_photo and all(self._field_validity.values()) else "disabled")
        self.revert_btn.config(
            state="normal" if self.current_photo else "disabled")
        has_input_sel = bool(self.input_list.curselection())
        self.input_remove_btn.config(
            state="normal" if has_input_sel else "disabled")

    def _update_counts(self):
        self.input_count_var.set(f"{len(self.input_paths)} items")
        self._update_button_states()

    def _center_dialog(self, dlg: tk.Toplevel):
        """Centre dlg over the main window."""
        self.root.update_idletasks()
        dlg.update_idletasks()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        dw, dh = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")

    def set_status(self, msg: str):
        self.status_var.set(msg)


    def _add_new_album(self):
        DownloadAlbumStructure.add_album(self.root, self.set_status)

    # -----------------------------------------------------------------------
    # Piwigo
    # -----------------------------------------------------------------------

    def _record_upload(self, path: str, album_fullname: str,
                       album_id: int = 0, file_id: int = 0):
        """Call after a file has been successfully uploaded to Piwigo."""
        filename = os.path.basename(path)
        DownloadAlbumStructure.record_uploaded_file(
            filename, album_fullname, album_id=album_id, file_id=file_id)

    def _prepare_upload_copy(self, path: str, params: dict) -> str | None:
        """Create a temp copy of the file for upload, optionally resized and with EXIF stripped.

        If 'max_upload_pixels' is set in params (e.g. 2000000 for 2 MP), the image is
        downsampled so that width*height does not exceed that value.  Aspect ratio is
        preserved.  The original file is never modified.

        Returns the path to the temp file, or None if preparation fails (caller should
        fall back to the original file).  Caller is responsible for deleting the temp file.
        """
        if not PIL_AVAILABLE:
            return None

        try:
            _, ext = os.path.splitext(path)
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                temp_path = tmp.name

            max_pixels = params.get('max_upload_pixels')
            if max_pixels:
                img = Image.open(path)  # type: ignore[possibly-undefined]
                w, h = img.size
                if w * h > max_pixels:
                    scale = (max_pixels / (w * h)) ** 0.5
                    new_w = max(1, int(w * scale))
                    new_h = max(1, int(h * scale))
                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)  # type: ignore[possibly-undefined]
                img.save(temp_path)
            else:
                shutil.copy2(path, temp_path)

            # Strip EXIF from temp copy
            if PIEXIF_AVAILABLE:
                try:
                    piexif.insert(piexif.dump({}), temp_path)
                except Exception:
                    pass

            return temp_path
        except Exception as e:
            self.set_status(f"Warning: could not create upload copy: {e}")
            return None

    def _upload_current_photo(self):
        """Upload the currently displayed photo to the selected Piwigo album."""
        if not self.current_photo:
            self.set_status("No photo selected.")
            return

        album    = self.upload_album_var.get()
        album_id = self.upload_album_id
        if not album or album == "(none)" or album_id == 0:
            messagebox.showerror(
                "No Album Selected",
                "Please select an upload album before uploading.\n\n"
                "Use the 'Change Upload Album' button to choose an album.",
                parent=self.root,
            )
            return

        path = self.current_photo
        original_path = path  # preserved for queue removal even if path is renamed

        try:
            params = DownloadAlbumStructure.load_params()
        except (FileNotFoundError, ValueError) as exc:
            messagebox.showerror("Configuration error", str(exc), parent=self.root)
            return

        # Gather custom field values for this file (read from widgets, not saved data)
        original_filename = os.path.basename(path)
        output_filename = self.custom_vars['output_filename'].get().strip() or original_filename
        author   = self.custom_vars['photo_source'].get().strip()
        comment  = self.custom_vars['comments'].get('1.0', "end").strip()
        tags     = self.custom_vars['tags'].get().strip()

        # Parse and format date for Piwigo API (MySQL format: YYYY-MM-DD HH:MM:SS)
        raw_date = self.custom_vars['date_of_photo'].get().strip()
        date_creation = ''
        if raw_date:
            parsed_date = self._parse_date(raw_date)
            if parsed_date:
                date_creation = parsed_date.strftime('%Y-%m-%d %H:%M:%S')

        # Check against the Piwigo file dictionary — ask before overwriting
        file_dict = self._load_file_dict()
        if output_filename in file_dict:
            entries = file_dict[output_filename]
            album_list = "\n  • ".join(
                e.get("fullname", str(e)) if isinstance(e, dict) else str(e)
                for e in entries
            )
            proceed = messagebox.askyesno(
                "File Already on Piwigo",
                f'"{output_filename}" already exists on Piwigo in:\n  • {album_list}\n\n'
                "Do you want to overwrite it?",
                parent=self.root,
            )
            if not proceed:
                self.set_status(f"Upload cancelled: {output_filename} already exists on Piwigo.")
                return

        # Rename the file if output filename differs from current name
        _, current_ext = os.path.splitext(original_filename)
        if output_filename and output_filename != original_filename:
            _, out_ext = os.path.splitext(output_filename)
            new_name = output_filename if out_ext else output_filename + current_ext
            new_path = os.path.join(os.path.dirname(path), new_name)
            if os.path.exists(new_path) and new_path != path:
                messagebox.showwarning(
                    "Rename Failed",
                    f'Cannot rename to "{new_name}": a file with that name already exists.',
                    parent=self.root,
                )
                self.set_status("Upload blocked: rename target already exists.")
                return
            try:
                os.rename(path, new_path)
                path = new_path
                self.current_photo = new_path
            except Exception as e:
                messagebox.showwarning(
                    "Rename Failed",
                    f'Could not rename "{original_filename}" to "{new_name}":\n{e}',
                    parent=self.root,
                )
                self.set_status("Upload blocked: rename failed.")
                return

        self.set_status(f"Uploading {output_filename}…")

        # Progress dialog
        progress_dlg = tk.Toplevel(self.root)
        progress_dlg.title("Uploading…")
        progress_dlg.resizable(False, False)
        progress_dlg.grab_set()
        progress_dlg.protocol("WM_DELETE_WINDOW", lambda: None)  # prevent close
        ttk.Label(progress_dlg, text=f"Uploading  {output_filename}",
                  padding=(16, 12, 16, 4)).pack()
        ttk.Label(progress_dlg, text=f"to  '{album}'",
                  padding=(16, 0, 16, 8)).pack()
        pbar = ttk.Progressbar(progress_dlg, mode='indeterminate', length=320)
        pbar.pack(padx=16, pady=(0, 8))
        pbar.start(12)
        progress_stage_var = tk.StringVar(value="Preparing image…")
        ttk.Label(progress_dlg, textvariable=progress_stage_var,
                  foreground="gray", padding=(16, 0, 16, 12)).pack()
        self._center_dialog(progress_dlg)

        def set_stage(msg: str):
            self.root.after(0, lambda: progress_stage_var.set(msg))

        def close_progress():
            pbar.stop()
            progress_dlg.grab_release()
            progress_dlg.destroy()

        def worker():
            # Prepare upload copy (resize + EXIF strip) inside the thread so the
            # UI is never blocked and the progress dialog covers the full operation.
            temp_path = self._prepare_upload_copy(path, params)
            upload_path = temp_path if temp_path else path
            logger.debug(f"Uploading to Piwigo: album_id={album_id}, album={album}, path={upload_path}")

            client = DownloadAlbumStructure.PiwigoClient(
                params['url'], params['username'], params['password'],
                verify_ssl=params.get('verify_ssl', True),
            )
            try:
                set_stage("Logging in…")
                client.login(params['username'], params['password'])
                set_stage(f"Uploading {output_filename}…")
                result = client.upload_image(
                    upload_path, album_id,
                    name=output_filename, author=author, comment=comment,
                    tags=tags, date_creation=date_creation,
                )
                image_id = int(result.get('image_id', 0))
                if params.get('sync_metadata', True):
                    set_stage("Syncing metadata (pwg.images.syncMetadata)…")
                    try:
                        client.sync_metadata(image_id)
                    except Exception as e:
                        logger.warning(f"syncMetadata failed (non-fatal): {e}")
                if params.get('refresh_representative', True):
                    set_stage("Refreshing album thumbnail (pwg.categories.refreshRepresentative)…")
                    try:
                        client.refresh_representative(album_id)
                    except Exception as e:
                        logger.warning(f"refreshRepresentative failed (non-fatal): {e}")
                set_stage("Done.")
                self.root.after(0, lambda: finish_ok(image_id))
            except Exception as exc:
                err = str(exc)
                self.root.after(0, lambda: finish_err(err))
            finally:
                client.logout()
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

        def finish_ok(image_id):
            close_progress()
            # Remove the uploaded file from the input list and move to next
            queue_path = original_path if original_path in self.input_paths else path
            if queue_path in self.input_paths:
                idx = self.input_paths.index(queue_path)
                self.input_paths.pop(idx)
                self.input_list.delete(idx)
                self._update_counts()
                if self.input_paths:
                    next_idx = min(idx, len(self.input_paths) - 1)
                    self.input_list.selection_clear(0, "end")
                    self.input_list.selection_set(next_idx)
                    self.input_list.see(next_idx)
                    self.current_photo = self.input_paths[next_idx]
                    self._load_photo(self.current_photo)
                else:
                    self._clear_viewer()

            self._record_upload(path, album, album_id=album_id, file_id=image_id)
            self.set_status(
                f"Uploaded {output_filename} → '{album}' (image id {image_id}).")

        def finish_err(err):
            close_progress()
            messagebox.showerror("Upload Failed",
                                 f"Could not upload {output_filename}:\n\n{err}",
                                 parent=self.root)
            self.set_status(f"Upload failed: {output_filename}")

        threading.Thread(target=worker, daemon=True).start()

    def _download_album_hierarchy(self):
        DownloadAlbumStructure.run(self.root, self.set_status)

    def _download_file_list(self):
        DownloadAlbumStructure.download_file_index(self.root, self.set_status)

    # -----------------------------------------------------------------------
    # Window geometry persistence
    # -----------------------------------------------------------------------

    def _resolve_startup_geometry(self):
        """Parse saved geometry and return (x, y, w, h), or defaults."""
        geom = self.state_data.get("geometry", "")
        normalised = geom.replace("+-", "-").replace("--", "+")
        m = re.fullmatch(r'(\d+)x(\d+)([+-]\d+)([+-]\d+)', normalised)
        if m:
            return int(m.group(3)), int(m.group(4)), int(m.group(1)), int(m.group(2))
        return 100, 100, 1400, 820  # defaults when no geometry is saved

    def _restore_geometry(self):
        """Apply saved position/size; snap to primary screen if off all monitors."""
        if not self.state_data.get("geometry"):
            return
        x, y, w, h = self._resolve_startup_geometry()
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.update_idletasks()
        if not _window_is_on_a_monitor(self.root.winfo_id()):
            min_w, min_h = self.root.minsize()
            self.root.geometry(f"{max(w, min_w)}x{max(h, min_h)}+100+100")

    def _persist_state(self):
        """Update state_data with current runtime state and write to disk."""
        album = self.upload_album_var.get()
        self.state_data["upload_album"]    = album if album != "(none)" else ""
        self.state_data["upload_album_id"] = self.upload_album_id
        # input_path is updated in-place by add_photos_dialog; just ensure the key exists
        self.state_data.setdefault("input_path", "")
        save_state(self.state_data)

    def _on_close(self):
        self._save_current_custom_fields()
        self.root.update_idletasks()
        x, y = self.root.winfo_x(), self.root.winfo_y()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        self.state_data["geometry"] = f"{w}x{h}+{x}+{y}"
        self._persist_state()
        self.root.destroy()

    def _bind_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self.add_photos_dialog())
        self.root.bind('<Control-u>', lambda e: self._upload_current_photo())


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("NOTE: Drag-and-drop disabled. Install tkinterdnd2 to enable.")

    app = PhotosUploader(root)
    root.mainloop()


if __name__ == '__main__':
    main()
