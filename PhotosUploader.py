"""
PhotosUploader.py
A GUI application for photo processing workflows.
Requires: pip install Pillow tkinterdnd2
"""

import os
import re
import json
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


try:
    from PIL import Image, ImageTk, ExifTags
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow not installed. Run: pip install Pillow")


try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    print("WARNING: tkinterdnd2 not installed. Run: pip install tkinterdnd2")

import DownloadAlbumStructure

# ---------------------------------------------------------------------------
# State persistence
# ---------------------------------------------------------------------------
STATE_FILE = Path(".") / "PhotosUploader State.json"


def _window_is_on_a_monitor(hwnd: int) -> bool:
    """Return True if any part of the window is on a connected monitor."""
    try:
        import ctypes
        MONITOR_DEFAULTTONULL = 0
        return bool(ctypes.windll.user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL))
    except Exception:
        return True  # assume visible if the API fails


def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_state(state: dict):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
                    '.tif', '.webp', '.heic', '.heif', '.raw', '.cr2',
                    '.nef', '.arw', '.dng'}

EXIF_TAG_NAMES = {
    'DateTime': 'Date/Time',
    'DateTimeOriginal': 'Date Taken',
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
    ('caption', 'Caption'),
    ('photo_source', 'Photographer/Source'),
    ('comments', 'Comments'),
    ('tags', 'Tags (comma-separated)'),
]


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def is_image(path: str) -> bool:
    return Path(path).suffix.lower() in IMAGE_EXTENSIONS


def parse_dnd_paths(data: str) -> list:
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
        self.output_paths = []      # list of str
        self.current_photo = None   # str path
        self.photo_image = None     # ImageTk reference
        self.custom_data = {}       # path -> dict of custom field values
        self.status_var = tk.StringVar(value="Ready.")
        self.upload_album_var = tk.StringVar(value="(none)")
        self.state_data = load_state()
        if self.state_data.get("upload_album"):
            self.upload_album_var.set(self.state_data["upload_album"])

        self._build_ui()
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
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(toolbar, text="Add New Album", command=self._add_new_album).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)
        ttk.Button(toolbar, text="Download Album Hierarchy", command=self._download_album_hierarchy).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Download File List", command=self._download_file_list).pack(side=tk.LEFT, padx=2)

        # ── Main three-panel area ─────────────────────────────────────────
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=4, pady=4)

        # LEFT: Input queue
        left_frame = self._build_queue_panel(main_pane, "Input Queue",
                                             is_input=True)
        main_pane.add(left_frame, weight=1)

        # CENTER: Photo viewer + fields
        center_frame = self._build_center_panel(main_pane)
        main_pane.add(center_frame, weight=3)

        # RIGHT: Output queue
        right_frame = self._build_queue_panel(main_pane, "Upload Queue",
                                              is_input=False)
        main_pane.add(right_frame, weight=1)

        # ── Status bar ───────────────────────────────────────────────────
        status_bar = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(status_bar, textvariable=self.status_var, anchor=tk.W).pack(
            side=tk.LEFT, padx=6, pady=2)
        self.progress = ttk.Progressbar(status_bar, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=6, pady=2)

    def _build_queue_panel(self, parent, title: str, is_input: bool) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text=title, padding=4)

        if not is_input:
            # Album selection sits at the very top of the upload panel
            album_btn_row = ttk.Frame(frame)
            album_btn_row.pack(fill=tk.X, pady=(0, 2))
            self.upload_queue_btn = ttk.Button(album_btn_row, text="Upload Queue",
                                               command=self._upload_queue,
                                               state=tk.DISABLED)
            self.upload_queue_btn.pack(side=tk.LEFT, padx=2)

            album_display_row = ttk.Frame(frame)
            album_display_row.pack(fill=tk.X, pady=(0, 2))
            ttk.Label(album_display_row, text="Album:").pack(side=tk.LEFT, padx=(2, 4))

            album_display_var = tk.StringVar(value="(none)")
            album_label = ttk.Label(album_display_row, textvariable=album_display_var,
                                    foreground='gray', anchor=tk.W)
            album_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            def _refresh_album_display(*_):
                from tkinter.font import nametofont
                full = self.upload_album_var.get()
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

            ttk.Button(frame, text="Change Upload Album",
                       command=self.open_output_folder).pack(fill=tk.X, padx=2, pady=(0, 4))

        # Buttons
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        if is_input:
            ttk.Button(btn_row, text="Add…", command=self.add_photos_dialog).pack(side=tk.LEFT, padx=2)
            self.input_remove_btn = ttk.Button(btn_row, text="Remove",
                                               command=self.remove_selected_input,
                                               state=tk.DISABLED)
            self.input_remove_btn.pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_row, text="↑", width=2, command=lambda: self._move_item(self.input_list, -1)).pack(side=tk.LEFT)
            ttk.Button(btn_row, text="↓", width=2, command=lambda: self._move_item(self.input_list, 1)).pack(side=tk.LEFT)
        else:
            ttk.Button(btn_row, text="Remove", command=self.remove_selected_output).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_row, text="← Return", command=self.return_to_input).pack(side=tk.LEFT, padx=2)

        # Count label
        if is_input:
            self.input_count_var = tk.StringVar(value="0 items")
            ttk.Label(btn_row, textvariable=self.input_count_var).pack(side=tk.RIGHT, padx=4)
        else:
            self.output_count_var = tk.StringVar(value="0 items")
            ttk.Label(btn_row, textvariable=self.output_count_var).pack(side=tk.RIGHT, padx=4)

        # Listbox with scrollbars
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)

        lb = tk.Listbox(list_frame, selectmode=tk.EXTENDED,
                        yscrollcommand=yscroll.set,
                        xscrollcommand=xscroll.set,
                        activestyle='dotbox',
                        font=('Consolas', 9))
        yscroll.config(command=lb.yview)
        xscroll.config(command=lb.xview)

        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if is_input:
            self.input_list = lb
            lb.bind('<<ListboxSelect>>', self._on_input_select)
            lb.bind('<Double-Button-1>', lambda e: self._queue_for_upload())
            if DND_AVAILABLE:
                lb.drop_target_register(DND_FILES)
                lb.dnd_bind('<<Drop>>', self._on_drop)
        else:
            self.output_list = lb
            lb.bind('<<ListboxSelect>>', self._on_output_select)

        return frame

    def _build_center_panel(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent)

        # Vertical paned window: viewer (top, smaller) / fields row (bottom)
        vpane = ttk.PanedWindow(frame, orient=tk.VERTICAL)
        vpane.pack(fill=tk.BOTH, expand=True)

        # ── Photo display (top pane) ──────────────────────────────────────
        viewer_frame = ttk.LabelFrame(vpane, text="Photo Viewer", padding=4)
        vpane.add(viewer_frame, weight=1)

        # ── Left column: nav buttons, filename, dims, path ───────────────
        left_col = ttk.Frame(viewer_frame)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        nav = ttk.Frame(left_col)
        nav.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(nav, text="◀ Prev",   command=self.prev_photo).pack(side=tk.LEFT, padx=2)
        self.revert_btn = ttk.Button(nav, text="↺ Revert",
                                     command=self._revert_photo,
                                     state=tk.DISABLED)
        self.revert_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(nav, text="Next ▶",   command=self.next_photo).pack(side=tk.LEFT, padx=2)

        ttk.Button(left_col, text="☁ Queue for Upload",
                   command=self._queue_for_upload).pack(fill=tk.X, pady=(0, 6))

        self.photo_label_var = tk.StringVar(value="No photo selected")
        ttk.Label(left_col, textvariable=self.photo_label_var,
                  font=('TkDefaultFont', 9, 'italic'),
                  anchor=tk.W).pack(fill=tk.X, pady=(0, 2))

        self.photo_dim_var = tk.StringVar(value="")
        ttk.Label(left_col, textvariable=self.photo_dim_var,
                  anchor=tk.W).pack(fill=tk.X, pady=(0, 6))

        # Path display — wraps to the actual column width
        self.path_var = tk.StringVar(value="")
        path_label = ttk.Label(left_col, textvariable=self.path_var,
                               font=('TkDefaultFont', 9),
                               anchor=tk.NW, justify=tk.LEFT, wraplength=220)
        path_label.pack(fill=tk.X)

        def _update_wraplength(event):
            path_label.configure(wraplength=max(event.width - 4, 50))
        path_label.bind('<Configure>', _update_wraplength)

        # ── Canvas — right side, narrower ────────────────────────────────
        self.canvas = tk.Canvas(viewer_frame, bg='#1a1a1a', cursor='crosshair',
                                height=200)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.bind('<Configure>', self._on_canvas_resize)

        # ── Bottom pane: Custom Fields / EXIF stacked vertically ─────────
        hpane = ttk.PanedWindow(vpane, orient=tk.VERTICAL)
        vpane.add(hpane, weight=2)

        # ── Custom Fields (left) ──────────────────────────────────────────
        custom_frame = ttk.LabelFrame(hpane, text="Custom Fields", padding=6)
        hpane.add(custom_frame, weight=1)

        self.custom_vars = {}
        for i, (key, label) in enumerate(CUSTOM_FIELDS):
            ttk.Label(custom_frame, text=label + ":", width=22, anchor=tk.E).grid(
                row=i, column=0, sticky=tk.E, pady=2, padx=(0, 4))
            if key == 'comments':
                txt = tk.Text(custom_frame, height=3, width=40, wrap=tk.WORD,
                              font=('TkDefaultFont', 9))
                txt.grid(row=i, column=1, sticky=tk.EW, pady=2)
                self.custom_vars[key] = txt
            else:
                var = tk.StringVar()
                entry = ttk.Entry(custom_frame, textvariable=var, width=40)
                entry.grid(row=i, column=1, sticky=tk.EW, pady=2)
                self.custom_vars[key] = var
        custom_frame.columnconfigure(1, weight=1)

        def _on_photo_source_changed(*_):
            if not hasattr(self, '_exif_data'):
                return
            if getattr(self, '_exif_has_artist', False):
                return  # original EXIF already has Artist — don't overwrite
            value = self.custom_vars['photo_source'].get().strip()
            self._exif_data['Artist'] = value
            self._refresh_exif_tree()

        self.custom_vars['photo_source'].trace_add('write', _on_photo_source_changed)

        def _on_comments_changed(_event=None):
            if not hasattr(self, '_exif_data'):
                return
            if getattr(self, '_exif_has_description', False):
                return  # original EXIF already has Description — don't overwrite
            widget = self.custom_vars['comments']
            value = widget.get('1.0', tk.END).strip()
            self._exif_data['Description'] = value
            self._refresh_exif_tree()
            widget.edit_modified(False)  # reset the modified flag so next change fires again

        self.custom_vars['comments'].bind('<<Modified>>', _on_comments_changed)

        # ── EXIF / Metadata (right) ───────────────────────────────────────
        exif_frame = ttk.LabelFrame(hpane, text="EXIF / Metadata", padding=6)
        hpane.add(exif_frame, weight=1)

        # Edit row packed first (side=BOTTOM) so the tree fills remaining space
        edit_row = ttk.Frame(exif_frame)
        edit_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        ttk.Label(edit_row, text="Edit selected value:").pack(side=tk.LEFT, padx=(0, 4))
        self.exif_edit_var = tk.StringVar()
        ttk.Entry(edit_row, textvariable=self.exif_edit_var, width=30).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(edit_row, text="Apply", command=self._apply_exif_edit).pack(
            side=tk.LEFT, padx=2)

        tree_frame = ttk.Frame(exif_frame)
        tree_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        exif_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        self.exif_tree = ttk.Treeview(tree_frame, columns=('key', 'value'),
                                      show='headings',
                                      yscrollcommand=exif_scroll.set)
        exif_scroll.config(command=self.exif_tree.yview)
        self.exif_tree.heading('key', text='Field')
        self.exif_tree.heading('value', text='Value')
        self.exif_tree.column('key', width=160)
        self.exif_tree.column('value', width=260)
        self.exif_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        exif_scroll.pack(side=tk.LEFT, fill=tk.Y)
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
                    self.input_list.itemconfig(tk.END, {'foreground': '#000'})
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
        paths = filedialog.askopenfilenames(title="Select Images", filetypes=filetypes)
        batch_state = {}
        added = 0
        for p in paths:
            if self._add_single_image(p, batch_state):
                added += 1
        self._update_counts()
        self.set_status(f"Added {added} image(s) to input queue.")


    def _add_folder(self, folder: str, batch_state: dict | None = None) -> int:
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
                              batch_state: dict) -> str:
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

        self.root.update_idletasks()
        rx = self.root.winfo_x() + self.root.winfo_width()  // 2
        ry = self.root.winfo_y() + self.root.winfo_height() // 2
        dlg.geometry(f"500x150+{rx - 250}+{ry - 75}")

        msg = (f'A file named "{name}" is already in the input queue.\n\n'
               f'Existing:  {existing_path}\n'
               f'New:         {new_path}')
        ttk.Label(dlg, text=msg, padding=(12, 10, 12, 6),
                  wraplength=476, justify=tk.LEFT).pack()

        result = tk.StringVar(value='skip')

        btn_frame = ttk.Frame(dlg, padding=(12, 0, 12, 12))
        btn_frame.pack()

        def choose(action):
            result.set(action)
            if action in ('skip_all', 'replace_all'):
                batch_state['all'] = action.replace('_all', '')
            dlg.destroy()

        ttk.Button(btn_frame, text="Skip",
                   command=lambda: choose('skip')).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Skip All",
                   command=lambda: choose('skip_all')).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Replace",
                   command=lambda: choose('replace')).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Replace All",
                   command=lambda: choose('replace_all')).pack(side=tk.LEFT, padx=4)

        dlg.wait_window()
        action = result.get()
        return 'skip' if action in ('skip', 'skip_all') else 'replace'

    def _add_single_image(self, path: str, batch_state: dict) -> bool:
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
        self.input_list.insert(tk.END, os.path.basename(path))
        return True

    # -----------------------------------------------------------------------
    # Queue management
    # -----------------------------------------------------------------------
    def remove_selected_input(self):
        sel = list(self.input_list.curselection())
        for i in reversed(sel):
            self.input_paths.pop(i)
            self.input_list.delete(i)
        self._update_counts()

    def remove_selected_output(self):
        sel = list(self.output_list.curselection())
        for i in reversed(sel):
            self.output_paths.pop(i)
            self.output_list.delete(i)
        self._update_counts()

    def clear_input(self):
        if messagebox.askyesno("Clear Input", "Remove all items from the input queue?"):
            self.input_paths.clear()
            self.input_list.delete(0, tk.END)
            self._update_counts()

    def clear_output(self):
        if messagebox.askyesno("Clear Output", "Remove all items from the output queue?"):
            self.output_paths.clear()
            self.output_list.delete(0, tk.END)
            self._update_counts()

    def return_to_input(self):
        sel = list(self.output_list.curselection())
        for i in reversed(sel):
            p = self.output_paths.pop(i)
            self.output_list.delete(i)
            if p not in self.input_paths:
                self.input_paths.append(p)
                self.input_list.insert(tk.END, os.path.basename(p))
        self._update_counts()

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
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(j)

    def open_output_folder(self):
        def on_select(_album_id, fullname):
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

    def _on_output_select(self, event):
        sel = self.output_list.curselection()
        if sel:
            self._save_current_custom_fields()
            self.current_photo = self.output_paths[sel[0]]
            self._load_photo(self.current_photo)

    def prev_photo(self):
        if not self.current_photo or not self.input_paths:
            return
        try:
            idx = self.input_paths.index(self.current_photo)
            new_idx = max(0, idx - 1)
        except ValueError:
            new_idx = 0
        self.input_list.selection_clear(0, tk.END)
        self.input_list.selection_set(new_idx)
        self.input_list.see(new_idx)
        self._save_current_custom_fields()
        self.current_photo = self.input_paths[new_idx]
        self._load_photo(self.current_photo)

    def next_photo(self):
        if not self.current_photo or not self.input_paths:
            return
        try:
            idx = self.input_paths.index(self.current_photo)
            new_idx = min(len(self.input_paths) - 1, idx + 1)
        except ValueError:
            new_idx = 0
        self.input_list.selection_clear(0, tk.END)
        self.input_list.selection_set(new_idx)
        self.input_list.see(new_idx)
        self._save_current_custom_fields()
        self.current_photo = self.input_paths[new_idx]
        self._load_photo(self.current_photo)

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

        # If not found, fall back to current photo file
        if not source:
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
            self.input_list.insert(tk.END, os.path.basename(dest))
        new_idx = self.input_paths.index(dest)
        self.input_list.selection_clear(0, tk.END)
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
        self._load_custom_fields(path)
        self.path_var.set(path)
        self._update_button_states()

    def _display_photo(self, path: str):
        if not PIL_AVAILABLE:
            self.canvas.delete('all')
            self.canvas.create_text(200, 150, text="Pillow not installed", fill='white')
            return
        try:
            img = Image.open(path)
            self.photo_dim_var.set(f"{img.width} × {img.height} px  |  {img.mode}")
            cw = max(self.canvas.winfo_width(), 100)
            ch = max(self.canvas.winfo_height(), 100)
            img.thumbnail((cw, ch), Image.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(img)
            self.canvas.delete('all')
            self.canvas.create_image(cw // 2, ch // 2,
                                     anchor=tk.CENTER,
                                     image=self.photo_image)
        except Exception as e:
            self.canvas.delete('all')
            self.canvas.create_text(10, 10, anchor=tk.NW,
                                    text=f"Cannot display image:\n{e}",
                                    fill='red', font=('TkDefaultFont', 10))

    def _on_canvas_resize(self, event):
        if self.current_photo:
            self._display_photo(self.current_photo)

    # -----------------------------------------------------------------------
    # EXIF
    # -----------------------------------------------------------------------
    def _load_exif(self, path: str):
        self.exif_tree.delete(*self.exif_tree.get_children())
        self._exif_data = {}
        self._exif_has_artist = False
        if not PIL_AVAILABLE:
            return
        try:
            img = Image.open(path)
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
                    self.exif_tree.insert('', tk.END, values=(display_tag, str(val)[:120]))
                self._exif_has_artist      = bool(self._exif_data.get('Artist'))
                self._exif_has_description = bool(self._exif_data.get('Description'))
            else:
                self._exif_has_artist      = False
                self._exif_has_description = False
                self.exif_tree.insert('', tk.END, values=('(No EXIF data)', ''))
        except Exception as e:
            self._exif_has_artist      = False
            self._exif_has_description = False
            self.exif_tree.insert('', tk.END, values=('Error reading EXIF', str(e)))

    def _on_exif_select(self, event):
        sel = self.exif_tree.selection()
        if sel:
            vals = self.exif_tree.item(sel[0], 'values')
            if vals:
                self.exif_edit_var.set(vals[1] if len(vals) > 1 else '')

    def _refresh_exif_tree(self, reselect_key: str | None = None):
        """Rebuild the EXIF treeview from self._exif_data in place."""
        self.exif_tree.delete(*self.exif_tree.get_children())
        reselect_iid = None
        for key, val in self._exif_data.items():
            iid = self.exif_tree.insert('', tk.END, values=(key, str(val)[:120]))
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

    # File info tab and separate file-info widget removed; path shown in `path_var`.

    # (File info method removed — path displayed in `path_var`)


    # -----------------------------------------------------------------------
    # Custom fields
    # -----------------------------------------------------------------------
    def _load_custom_fields(self, path: str):
        data = self.custom_data.get(path, {})
        for key, widget in self.custom_vars.items():
            val = data.get(key, '')
            if isinstance(widget, tk.Text):
                widget.delete('1.0', tk.END)
                widget.insert('1.0', val)
            else:
                widget.set(val)

    def _save_current_custom_fields(self):
        if not self.current_photo:
            return
        data = {}
        for key, widget in self.custom_vars.items():
            if isinstance(widget, tk.Text):
                data[key] = widget.get('1.0', tk.END).strip()
            else:
                data[key] = widget.get()
        self.custom_data[self.current_photo] = data

    # -----------------------------------------------------------------------
    # Processing
    # -----------------------------------------------------------------------
    # Characters forbidden in Windows filenames; also covers Linux/macOS
    _ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
    # Windows reserved names (case-insensitive, with or without extension)
    _RESERVED_NAMES = re.compile(
        r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.[^.]*)?$', re.IGNORECASE)

    def _validate_output_filename(self, name: str) -> str | None:
        """Return an error message if *name* is not a valid filename, else None."""
        if not name:
            return None  # empty means "use original filename" — always valid
        if self._ILLEGAL_FILENAME_CHARS.search(name):
            bad = ''.join(sorted(set(self._ILLEGAL_FILENAME_CHARS.findall(name))))
            return (f'The output filename contains illegal character(s): {bad}\n\n'
                    'A filename may not contain  \\ / : * ? " < > |')
        if self._RESERVED_NAMES.match(name):
            return (f'"{name}" is a reserved Windows filename and cannot be used.')
        if name != name.strip() or name.endswith('.'):
            return ('The output filename may not start or end with a space, '
                    'or end with a period.')
        return None

    def _load_file_dict(self) -> dict:
        """Load FileDict.json and return it, or an empty dict if unavailable."""
        p = Path(DownloadAlbumStructure.FILE_INDEX_FILE)
        if not p.exists():
            return {}
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _queue_for_upload(self):
        """Move the current photo to the output queue, then show the next one."""
        if not self.current_photo:
            self.set_status("No photo selected.")
            return
        self._save_current_custom_fields()
        path = self.current_photo

        # Validate output filename before queuing
        out_name = self.custom_vars['output_filename'].get().strip()
        err = self._validate_output_filename(out_name)
        if err:
            messagebox.showwarning("Invalid Output Filename", err, parent=self.root)
            self.set_status("Queue blocked: invalid output filename.")
            return

        filename = os.path.basename(path)

        # Check against the upload queue — block if already queued but not yet uploaded
        queued_match = next(
            (p for p in self.output_paths if os.path.basename(p) == filename), None)
        if queued_match:
            messagebox.showwarning(
                "File Already in Upload Queue",
                f'"{filename}" is already in the upload queue (not yet uploaded):\n\n'
                f"  {queued_match}\n\n"
                "Upload blocked to prevent duplicate uploads.",
                parent=self.root,
            )
            self.set_status(f"Blocked: {filename} already in upload queue.")
            return

        # Check against the Piwigo file dictionary — block duplicates already on server
        file_dict = self._load_file_dict()
        if filename in file_dict:
            albums = file_dict[filename]
            album_list = "\n  • ".join(albums)
            messagebox.showwarning(
                "File Already on Piwigo",
                f'"{filename}" already exists on Piwigo in:\n  • {album_list}\n\n'
                "Upload blocked to prevent overwriting.",
                parent=self.root,
            )
            self.set_status(f"Blocked: {filename} already exists on Piwigo.")
            return

        # Capture position before removing so we can select the successor
        next_idx = None
        if path in self.input_paths:
            idx = self.input_paths.index(path)
            self.input_paths.pop(idx)
            self.input_list.delete(idx)
            if self.input_paths:
                next_idx = min(idx, len(self.input_paths) - 1)

        if path not in self.output_paths:
            self.output_paths.append(path)
            self.output_list.insert(tk.END, os.path.basename(path))

        self._update_counts()
        self.set_status(f"Queued for upload: {os.path.basename(path)}")

        # Advance to the next photo, or clear the viewer if the queue is empty
        if next_idx is not None:
            self.input_list.selection_clear(0, tk.END)
            self.input_list.selection_set(next_idx)
            self.input_list.see(next_idx)
            self.current_photo = self.input_paths[next_idx]
            self._load_photo(self.current_photo)
        else:
            self.current_photo = None
            self.photo_image = None
            self.photo_label_var.set("No photo selected")
            self.photo_dim_var.set("")
            self.path_var.set("")
            self.canvas.delete('all')
            self.exif_tree.delete(*self.exif_tree.get_children())
            for widget in self.custom_vars.values():
                if isinstance(widget, tk.Text):
                    widget.delete('1.0', tk.END)
                else:
                    widget.set('')
            self._update_button_states()


    # -----------------------------------------------------------------------
    # Utilities
    # -----------------------------------------------------------------------
    def _update_button_states(self):
        self.upload_queue_btn.config(
            state=tk.NORMAL if self.output_paths else tk.DISABLED)
        self.revert_btn.config(
            state=tk.NORMAL if self.current_photo else tk.DISABLED)
        has_input_sel = bool(self.input_list.curselection())
        self.input_remove_btn.config(
            state=tk.NORMAL if has_input_sel else tk.DISABLED)

    def _update_counts(self):
        self.input_count_var.set(f"{len(self.input_paths)} item(s)")
        self.output_count_var.set(f"{len(self.output_paths)} item(s)")
        self._update_button_states()

    def set_status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()


    def _add_new_album(self):
        DownloadAlbumStructure.add_album(self.root, self.set_status)

    # -----------------------------------------------------------------------
    # Piwigo
    # -----------------------------------------------------------------------

    def _record_upload(self, path: str, album_fullname: str):
        """Call after a file has been successfully uploaded to Piwigo."""
        filename = os.path.basename(path)
        DownloadAlbumStructure.record_uploaded_file(filename, album_fullname)

    def _upload_queue(self):
        if not self.output_paths:
            self.set_status("Upload queue is empty.")
            return
        album = self.upload_album_var.get()
        if not album or album == "(none)":
            messagebox.showwarning("No album selected",
                                   "Please select an upload album first.",
                                   parent=self.root)
            return
        # TODO: implement actual upload loop here.
        # After each successful upload, call:
        #   self._record_upload(path, album)
        messagebox.showinfo("Upload Queue",
                            f"Upload of {len(self.output_paths)} photo(s) "
                            f"to '{album}' is not yet implemented.",
                            parent=self.root)

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

    def _on_close(self):
        self.root.update_idletasks()
        x, y = self.root.winfo_x(), self.root.winfo_y()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        self.state_data["geometry"] = f"{w}x{h}+{x}+{y}"
        album = self.upload_album_var.get()
        self.state_data["upload_album"] = album if album != "(none)" else ""
        save_state(self.state_data)
        self.root.destroy()

    def _bind_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self.add_photos_dialog())
        self.root.bind('<Control-Return>', lambda e: self._queue_for_upload())
        self.root.bind('<Control-Right>', lambda e: self.next_photo())
        self.root.bind('<Control-Left>', lambda e: self.prev_photo())


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
