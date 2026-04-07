"""
DownloadAlbumStructure.py
Connects to a Piwigo instance, downloads the full album hierarchy, writes
it to AlbumHierarchy.json, and shows a modal progress dialog while working.

Public API
----------
    run(parent, set_status_cb)
        Download album hierarchy → AlbumHierarchy.json.

    add_album(parent, set_status_cb)
        Create a new album on the server and refresh AlbumHierarchy.json.

    pick_album(parent, set_status_cb, on_select_cb)
        Show a tree-picker dialog; calls on_select_cb(album_id, fullname).

    download_file_index(parent, set_status_cb)
        Optional: walk every album, build filename → [album fullname, …]
        and write FileDict.json.  Useful for detecting duplicates before
        uploading.
"""

import sys
import json
import time
import threading
import warnings
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Any

try:
    import requests
    import urllib3
except ImportError:
    sys.exit(
        "ERROR: The 'requests' library is required but not installed.\n"
        "Run:  pip install requests"
    )

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------
def _center_dialog(parent: tk.Widget, dlg: tk.Toplevel):
    """Centre dlg over parent."""
    parent.update_idletasks()
    dlg.update_idletasks()
    rx, ry = parent.winfo_rootx(), parent.winfo_rooty()
    rw, rh = parent.winfo_width(), parent.winfo_height()
    dw, dh = dlg.winfo_width(), dlg.winfo_height()
    dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
def get_data_dir() -> Path:
    """Get the data directory from params.json, defaulting to '.' if not set."""
    try:
        params = load_params()
        data_path = params.get('path', '.')
        return Path(data_path)
    except Exception:
        # If params can't be loaded, default to current directory
        return Path('.')


PARAMS_FILE     = Path(".") / "PhotosUploader Params.json"
REQUIRED_PARAMS = ("url", "username", "password")


def _album_hierarchy_file() -> Path:
    return get_data_dir() / "AlbumHierarchy.json"


def _file_index_file() -> Path:
    return get_data_dir() / "FileDict.json"


# ---------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------
def load_params() -> dict[str, Any]:
    """Load Piwigo connection parameters from PhotosUploader Params.json."""
    if not PARAMS_FILE.exists():
        raise FileNotFoundError(
            f"Parameters file not found: {PARAMS_FILE}\n\n"
            "Please create PhotosUploader Params.json next to this script with:\n"
            '{\n'
            '  "url": "https://your-piwigo-site.example.com",\n'
            '  "username": "your-username-here",\n'
            '  "password": "your-password-here",\n'
            '  "verify_ssl": false\n'
            '}\n\n'
            'Optional fields:\n'
            '  "path": "/path/to/data/directory"  (defaults to ".")\n'
            '  "rate_limit_calls_per_secondond": 2.0  (API calls per second, defaults to 2.0)'
        )
    with open(PARAMS_FILE) as f:
        params = json.load(f)
    missing = [k for k in REQUIRED_PARAMS if not params.get(k)]
    if missing:
        raise ValueError(
            f"Missing required fields in PhotosUploader Params.json: "
            f"{', '.join(missing)}"
        )
    return params


# ---------------------------------------------------------------------------
# Piwigo API client
# ---------------------------------------------------------------------------
class PiwigoClient:
    def __init__(self, base_url: str, username: str, password: str,
                 verify_ssl: bool = True, rate_limit_calls_per_second: float = 2.0):
        """Initialize Piwigo client.

        Args:
            base_url: Piwigo server URL
            username: Piwigo username
            password: Piwigo password
            verify_ssl: Whether to verify HTTPS certificates
            rate_limit_calls_per_second: Maximum API calls per second (default 2.0)
        """
        url = base_url.strip().rstrip("/")
        if url.startswith("http://"):
            url = "https://" + url[7:]
        elif not url.startswith("https://"):
            url = "https://" + url
        self.base_url = url
        self.api_url  = f"{self.base_url}/ws.php?format=json"
        self.session    = requests.Session()
        self.session.verify = verify_ssl
        self._verify_ssl = verify_ssl
        self.rate_limit_calls_per_second = max(0.1, float(rate_limit_calls_per_second))  # at least 0.1 calls/sec
        self._last_api_call_time = 0

    def _apply_rate_limit(self):
        """Enforce rate limiting by sleeping if necessary."""
        if self.rate_limit_calls_per_second <= 0:
            return
        min_interval = 1.0 / self.rate_limit_calls_per_second
        elapsed = time.time() - self._last_api_call_time
        if elapsed < min_interval:
            time.sleep(min_interval - elapsed)
        self._last_api_call_time = time.time()

    def _call(self, method: str, params: dict = None) -> dict:
        """Call Piwigo API with retry logic on timeout."""
        self._apply_rate_limit()
        max_retries = 3
        retry_count = 0
        last_error = None

        while retry_count < max_retries:
            try:
                payload = {"method": method}
                if params:
                    payload.update(params)
                with warnings.catch_warnings():
                    if not self._verify_ssl:
                        warnings.simplefilter("ignore", urllib3.exceptions.InsecureRequestWarning)
                    r = self.session.post(self.api_url, data=payload, timeout=30)
                r.raise_for_status()
                try:
                    data = r.json()
                except ValueError:
                    preview = r.text[:300].strip() if r.text else "(empty)"
                    raise RuntimeError(
                        f"Server did not return valid JSON for '{method}'.\n"
                        f"URL: {self.api_url}\n"
                        f"HTTP status: {r.status_code}\n"
                        f"Response: {preview}"
                    )
                if data.get("stat") != "ok":
                    raise RuntimeError(data.get("message", "Unknown Piwigo API error"))
                return data.get("result", {})
            except (requests.Timeout, requests.ConnectTimeout, requests.ReadTimeout) as e:
                retry_count += 1
                last_error = e
                if retry_count < max_retries:
                    wait_time = 2 ** retry_count  # exponential backoff: 2s, 4s, 8s
                    logger.debug(f"API timeout (attempt {retry_count}/{max_retries}). "
                                f"Retrying in {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    raise RuntimeError(
                        f"Connection to Piwigo server timed out after {max_retries} attempts.\n\n"
                        f"Possible causes:\n"
                        f"  • Piwigo server is slow or overloaded\n"
                        f"  • Network connection is unstable\n"
                        f"  • Server is temporarily unavailable\n\n"
                        f"To fix:\n"
                        f"  1. Check your internet connection\n"
                        f"  2. Verify the Piwigo URL in PhotosUploader Params.json\n"
                        f"  3. Try again in a few moments\n"
                        f"  4. Contact your server administrator if the problem persists\n\n"
                        f"Details: {str(last_error)}"
                    )
        raise RuntimeError(f"API call '{method}' failed: {last_error}")

    def login(self, username: str, password: str):
        self._call("pwg.session.login", {
            "username": username,
            "password": password,
        })

    def logout(self):
        try:
            self._call("pwg.session.logout")
        except Exception:
            pass

    def get_albums(self) -> list:
        result = self._call("pwg.categories.getList", {
            "recursive": "true",
            "fullname":  "true",
        })
        return result.get("categories", [])

    def create_album(self, name: str, parent_id: int = None) -> int:
        """Create a new album and return its id."""
        params = {"name": name}
        if parent_id is not None:
            params["parent"] = parent_id
        result = self._call("pwg.categories.add", params)
        return int(result.get("id", 0))

    def upload_image(self, path: str, category_id: int, name: str = '',
                     author: str = '', comment: str = '',
                     tags: str = '', date_creation: str = '') -> dict:
        """Upload a single image file to Piwigo via pwg.images.addSimple.

        Returns the API result dict which contains 'image_id' on success.
        """
        self._apply_rate_limit()
        filename = path.rsplit('/', 1)[-1].rsplit('\\', 1)[-1]
        logger.info(f"[upload] uploading {filename} to category_id {category_id}")
        data = {
            'method':   'pwg.images.addSimple',
            'category': str(category_id),
            'level':    '0',   # 0 = public; omitting this can leave the image hidden
        }
        if name:
            data['name'] = name
        if author:
            data['author'] = author
        if comment:
            data['comment'] = comment
        if tags:
            data['tags'] = tags
        if date_creation:
            data['date_creation'] = date_creation

        max_retries = 2
        retry_count = 0
        last_error = None

        while retry_count < max_retries:
            try:
                with warnings.catch_warnings():
                    if not self._verify_ssl:
                        warnings.simplefilter('ignore', urllib3.exceptions.InsecureRequestWarning)
                    with open(path, 'rb') as fh:
                        r = self.session.post(
                            self.api_url,
                            data=data,
                            files={'image': (filename, fh, 'image/jpeg')},
                            timeout=120,
                        )
                r.raise_for_status()
                try:
                    resp = r.json()
                except ValueError:
                    preview = r.text[:400].strip() if r.text else "(empty)"
                    raise RuntimeError(
                        f"Server did not return valid JSON for upload of {filename}.\n"
                        f"HTTP status: {r.status_code}\n"
                        f"Response preview:\n{preview}"
                    )
                stat = resp.get('stat')
                message = resp.get('message', '')
                if message:
                    logger.info(f"[upload] server message: {message}")
                if stat != 'ok':
                    raise RuntimeError(message or 'Unknown Piwigo upload error')
                result = resp.get('result', {})
                logger.info(f"[upload] success: image_id={result.get('image_id')}, url={result.get('url', '?')}")
                return result
            except (requests.Timeout, requests.ConnectTimeout, requests.ReadTimeout) as e:
                retry_count += 1
                last_error = e
                if retry_count < max_retries:
                    wait_time = 2 ** retry_count  # exponential backoff: 2s, 4s
                    logger.debug(f"Upload timeout (attempt {retry_count}/{max_retries}). "
                                f"Retrying in {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    raise RuntimeError(
                        f"Upload of {filename} timed out after {max_retries} attempts (large file?).\n\n"
                        f"Possible causes:\n"
                        f"  • File is very large\n"
                        f"  • Network connection is unstable\n"
                        f"  • Piwigo server is overloaded\n\n"
                        f"To fix:\n"
                        f"  1. Try uploading a smaller image\n"
                        f"  2. Check your internet connection\n"
                        f"  3. Try again when the server is less busy\n"
                        f"  4. Contact your server administrator if the problem persists\n\n"
                        f"Details: {str(last_error)}"
                    )
        raise RuntimeError(f"Upload of '{filename}' failed: {last_error}")

    def sync_metadata(self, image_id: int) -> None:
        """Ask Piwigo to re-read EXIF/IPTC metadata for an image from its file."""
        logger.info(f"[sync] pwg.images.syncMetadata image_id={image_id}")
        self._call("pwg.images.syncMetadata", {"image_id": image_id})

    def refresh_representative(self, category_id: int) -> None:
        """Refresh the representative thumbnail for an album."""
        logger.info(f"[sync] pwg.categories.refreshRepresentative category_id={category_id}")
        self._call("pwg.categories.refreshRepresentative", {"cat_id": category_id})

    def get_album_images(self, cat_id: int, per_page: int = 500) -> list[dict]:
        """Return all images in a category, handling pagination automatically.

        Each dict has at least 'file' (filename) and 'id' (image id).
        """
        images = []
        page = 0
        while True:
            result = self._call("pwg.categories.getImages", {
                "cat_id":   cat_id,
                "per_page": per_page,
                "page":     page,
            })
            batch = result.get("images", [])
            images.extend(batch)
            paging = result.get("paging", {})
            total  = int(paging.get("total_count", len(images)))
            if len(images) >= total or not batch:
                break
            page += 1
        return images


# ---------------------------------------------------------------------------
# Hierarchy builder
# ---------------------------------------------------------------------------
def _build_hierarchy(flat: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Convert a flat list of Piwigo category dicts into a nested tree.

    Each node in the result has:
        id, name, fullname, nb_images, total_nb_images, children
    Root nodes (id_uppercat absent, null, or "0") appear at the top level.
    """
    by_id: dict[int, dict] = {}
    for cat in flat:
        node_id = int(cat["id"])
        node = {
            "id":              node_id,
            "name":            cat.get("name", ""),
            "nb_images":       int(cat.get("nb_images", 0)),
            "total_nb_images": int(cat.get("total_nb_images", 0)),
            "children":        [],
        }
        # When fullname=true the API puts the full breadcrumb in `name`;
        # the short name is the last segment after " / ".
        parts = node["name"].rsplit(" / ", 1)
        node["fullname"] = node["name"]
        node["name"]     = parts[-1]
        by_id[node_id] = node

    roots = []
    orphans = []
    for cat in flat:
        node      = by_id[int(cat["id"])]
        parent_id = cat.get("id_uppercat")
        if not parent_id or str(parent_id) == "0":
            roots.append(node)
        elif int(parent_id) in by_id:
            by_id[int(parent_id)]["children"].append(node)
        else:
            orphans.append(node)

    if orphans:
        orphan_container = {
            "id":              -1,
            "name":            "Orphans",
            "fullname":        "Orphans",
            "nb_images":       0,
            "total_nb_images": 0,
            "children":        orphans,
        }
        roots.append(orphan_container)

    def _sort(nodes):
        nodes.sort(key=lambda n: n["name"].lower())
        for n in nodes:
            _sort(n["children"])

    _sort(roots)
    return roots


# ---------------------------------------------------------------------------
# Shared fetch-and-save helper
# ---------------------------------------------------------------------------
def _fetch_and_save_hierarchy(client: PiwigoClient, step_cb) -> int:
    """Fetch albums from Piwigo, build the hierarchy, and write the JSON file.

    step_cb(str) is called at each stage to report progress.  It must be
    safe to call from a background thread (wrap with root.after if needed).
    Returns the total number of albums fetched.
    """
    step_cb("Fetching album list…")
    flat = client.get_albums()
    step_cb("Building hierarchy…")
    hierarchy = _build_hierarchy(flat)
    step_cb(f"Writing {_album_hierarchy_file().name}…")
    with open(_album_hierarchy_file(), "w", encoding="utf-8") as f:
        json.dump(hierarchy, f, indent=2, ensure_ascii=False)
    return len(flat)


# ---------------------------------------------------------------------------
# File-index builder
# ---------------------------------------------------------------------------
def _fetch_and_save_file_index(client: PiwigoClient, flat_albums: list,
                               progress_cb) -> dict[str, list[dict]]:
    """Walk every album, collect filenames, and write FileDict.json.

    progress_cb(done: int, total: int, album_name: str) is called after each
    album is processed.  It must be safe to call from a background thread.

    Returns the completed index dict:
        {filename: [{"fullname": str, "album_id": int, "file_id": int}, …], …}
    The fullname used for each album is the breadcrumb stored in the flat
    album list (e.g. "Fan Photos / Ackermansion").
    """
    # Build a fast id → fullname map from the flat list
    fullname_by_id: dict[int, str] = {}
    for cat in flat_albums:
        cat_id   = int(cat["id"])
        fullname = cat.get("name", "")          # fullname=true → breadcrumb
        fullname_by_id[cat_id] = fullname

    index: dict[str, list[dict]] = {}
    total = len(flat_albums)

    for done, cat in enumerate(flat_albums, 1):
        cat_id   = int(cat["id"])
        fullname = fullname_by_id[cat_id]
        # Strip the short name for progress reporting
        short    = fullname.rsplit(" / ", 1)[-1]
        progress_cb(done, total, short)

        images = client.get_album_images(cat_id)
        for img in images:
            filename = img.get("file", "").strip()
            if not filename:
                continue
            file_id = int(img.get("id", 0))
            entry = {"fullname": fullname, "album_id": cat_id, "file_id": file_id}
            entries = index.setdefault(filename, [])
            if not any(e["album_id"] == cat_id for e in entries):
                entries.append(entry)

    with open(_file_index_file(), "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False, sort_keys=True)

    return index


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------
def run(parent: tk.Widget, set_status_cb):
    """Download album hierarchy and write AlbumHierarchy.json.

    Opens a modal progress dialog on *parent*.  Calls set_status_cb(str)
    to update the caller's status bar at each step and on completion.
    """

    try:
        params = load_params()
    except (FileNotFoundError, ValueError) as exc:
        messagebox.showerror("Configuration error", str(exc), parent=parent)
        return

    # ── Progress dialog ───────────────────────────────────────────────────
    dlg = tk.Toplevel(parent)
    dlg.title("Downloading Album Hierarchy")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.geometry("360x120")
    _center_dialog(parent, dlg)

    ttk.Label(dlg, text="Downloading album hierarchy from Piwigo…",
              padding=(12, 10, 12, 4)).pack()
    step_var = tk.StringVar(value="Connecting…")
    ttk.Label(dlg, textvariable=step_var, foreground="gray",
              padding=(12, 0, 12, 6)).pack()
    bar = ttk.Progressbar(dlg, mode="indeterminate", length=320)
    bar.pack(padx=12, pady=(0, 12))
    bar.start(12)

    def set_step(msg):
        step_var.set(msg)
        set_status_cb(msg)

    def finish_ok(n_albums):
        bar.stop()
        dlg.destroy()
        msg = (f"Downloaded {n_albums} album(s). "
               f"Hierarchy written to {_album_hierarchy_file().name}")
        set_status_cb(msg)

    def finish_err(err):
        bar.stop()
        dlg.destroy()
        messagebox.showerror("Piwigo error", err, parent=parent)
        set_status_cb("Download failed.")

    # ── Background worker ─────────────────────────────────────────────────
    def worker():
        client = PiwigoClient(
            params["url"],
            params["username"],
            params["password"],
            verify_ssl=params.get("verify_ssl", True),
            rate_limit_calls_per_second=params.get("rate_limit_calls_per_second", 2.0),
        )
        try:
            parent.after(0, lambda: set_step("Logging in…"))
            client.login(params["username"], params["password"])

            def step(msg):
                parent.after(0, lambda m=msg: set_step(m))

            n = _fetch_and_save_hierarchy(client, step)

            # # TEMPORARY: also build the file index for performance testing
            # parent.after(0, lambda: set_step("Building file index…"))
            # flat = client.get_albums()
            # _fetch_and_save_file_index(client, flat, lambda d, t, nm: None)

            parent.after(0, lambda: finish_ok(n))
        except Exception as exc:
            err = str(exc)
            parent.after(0, lambda: finish_err(err))
        finally:
            client.logout()

    threading.Thread(target=worker, daemon=True).start()


def add_album(parent: tk.Widget, set_status_cb):
    """Open a dialog to create a new Piwigo album.

    The user selects an optional parent from the existing album tree (leave
    nothing selected to create a top-level album), enters a name, and clicks
    Create.  The album is added via the API and AlbumHierarchy.json is
    refreshed afterwards.
    """

    try:
        params = load_params()
    except (FileNotFoundError, ValueError) as exc:
        messagebox.showerror("Configuration error", str(exc), parent=parent)
        return

    # Load existing hierarchy for the parent picker (may be empty if not yet
    # downloaded — user can still create a top-level album).
    hierarchy = []
    if _album_hierarchy_file().exists():
        try:
            with open(_album_hierarchy_file(), encoding="utf-8") as f:
                data = json.load(f)
            # Validate that it's a list (expected hierarchy format)
            if not isinstance(data, list):
                raise ValueError(f"Expected a list, got {type(data).__name__}")
            hierarchy = data
        except json.JSONDecodeError as e:
            messagebox.showerror(
                "Invalid Album Data",
                f"The album hierarchy file is corrupted (invalid JSON):\n\n{e}\n\n"
                "To fix this, download the album hierarchy again:\n"
                "1. Click 'Download Album Hierarchy' in the main window\n"
                "2. Try adding a new album again",
                parent=parent,
            )
            return
        except (ValueError, TypeError) as e:
            messagebox.showerror(
                "Invalid Album Data",
                f"The album hierarchy file has an unexpected format:\n\n{e}\n\n"
                "To fix this, download the album hierarchy again:\n"
                "1. Click 'Download Album Hierarchy' in the main window\n"
                "2. Try adding a new album again",
                parent=parent,
            )
            return
        except Exception as e:
            messagebox.showerror(
                "Cannot Read Album Data",
                f"Error reading album hierarchy file:\n\n{e}\n\n"
                "To fix this:\n"
                "1. Check that the file exists and is readable\n"
                "2. Download the album hierarchy again:\n"
                "   - Click 'Download Album Hierarchy' in the main window\n"
                "3. Try adding a new album again",
                parent=parent,
            )
            return

    # ── Dialog ───────────────────────────────────────────────────────────────
    dlg = tk.Toplevel(parent)
    dlg.title("Add New Album")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.geometry("480x520")
    _center_dialog(parent, dlg)

    ttk.Label(dlg,
              text="Parent album  (leave unselected to create a top-level album):",
              padding=(12, 10, 12, 4)).pack(anchor="w")

    # ── Album tree ───────────────────────────────────────────────────────────
    tree_frame = ttk.Frame(dlg, padding=(12, 0, 12, 0))
    tree_frame.pack(fill='both', expand=True)

    yscroll = ttk.Scrollbar(tree_frame, orient='vertical')
    xscroll = ttk.Scrollbar(tree_frame, orient='horizontal')
    tree = ttk.Treeview(tree_frame, selectmode="browse", show="tree",
                        yscrollcommand=yscroll.set,
                        xscrollcommand=xscroll.set)
    yscroll.config(command=tree.yview)
    xscroll.config(command=tree.xview)
    yscroll.pack(side='right', fill="y")
    xscroll.pack(side='bottom', fill="x")
    tree.pack(side='left', fill='both', expand=True)

    # fullname_by_id: id → fullname, built while populating the tree
    fullname_by_id: dict[int, str] = {}
    node_by_id: dict[int, dict] = {}

    if not hierarchy:
        tree.insert("", 'end', text="(No album hierarchy loaded — "
                    "use 'Download Album Hierarchy' first)", tags=("hint",))
        tree.tag_configure("hint", foreground="gray")
    else:
        def _index_nodes(nodes):
            for node in nodes:
                fullname_by_id[node["id"]] = node.get("fullname", node["name"])
                node_by_id[node["id"]] = node
                if node.get("children"):
                    _index_nodes(node["children"])

        _index_nodes(hierarchy)

        def _populate(parent_iid, nodes):
            for node in nodes:
                iid = str(node["id"])
                tree.insert(parent_iid, 'end', iid=iid,
                            text=node["name"], open=False)
                if node.get("children"):
                    _populate(iid, node["children"])

        _populate("", hierarchy)

    # ── Name entry ───────────────────────────────────────────────────────────
    name_frame = ttk.Frame(dlg, padding=(12, 8, 12, 0))
    name_frame.pack(fill='x')
    ttk.Label(name_frame, text="New album name:").pack(side='left', padx=(0, 6))
    name_var = tk.StringVar()
    name_entry = ttk.Entry(name_frame, textvariable=name_var, width=32)
    name_entry.pack(side='left', fill='x', expand=True)
    name_entry.focus_set()

    # ── Status / buttons ─────────────────────────────────────────────────────
    status_var = tk.StringVar(value="")
    ttk.Label(dlg, textvariable=status_var, foreground="gray",
              padding=(12, 4, 12, 0)).pack(anchor='w')

    btn_frame = ttk.Frame(dlg, padding=(12, 6, 12, 12))
    btn_frame.pack(fill='x')

    def on_create():
        name = name_var.get().strip()
        if not name:
            messagebox.showwarning("Name required",
                                   "Please enter a name for the new album.",
                                   parent=dlg)
            name_entry.focus_set()
            return

        sel = tree.selection()
        # Only use the selection if it has a numeric iid (i.e. a real album)
        parent_id = None
        parent_fullname = ""
        if sel:
            try:
                parent_id = int(sel[0])
                parent_fullname = fullname_by_id.get(parent_id, "")
            except ValueError:
                parent_id = None

        new_fullname = f"{parent_fullname} / {name}" if parent_fullname else name

        # Check for a duplicate name among siblings
        if parent_id is not None and parent_id in node_by_id:
            siblings = node_by_id[parent_id]["children"]
        else:
            siblings = hierarchy
        name_lower = name.lower()
        duplicate = next((n for n in siblings if n["name"].lower() == name_lower), None)
        if duplicate:
            location = f'under "{parent_fullname}"' if parent_fullname else "at the top level"
            messagebox.showwarning(
                "Duplicate album name",
                f'An album named "{duplicate["name"]}" already exists {location}.\n\n'
                "Please choose a different name.",
                parent=dlg,
            )
            name_entry.focus_set()
            return

        create_btn.config(state=tk.DISABLED)
        cancel_btn.config(state=tk.DISABLED)
        status_var.set("Creating album…")

        def worker():
            client = PiwigoClient(
                params["url"],
                params["username"],
                params["password"],
                verify_ssl=params.get("verify_ssl", True),
            )
            try:
                client.login(params["username"], params["password"])
                new_id = client.create_album(name, parent_id)

                # Insert new node into the local hierarchy and save
                parent.after(0, lambda: status_var.set("Updating local hierarchy…"))
                new_node = {
                    "id":              new_id,
                    "name":            name,
                    "fullname":        new_fullname,
                    "nb_images":       0,
                    "total_nb_images": 0,
                    "children":        [],
                }
                if parent_id is not None and parent_id in node_by_id:
                    siblings = node_by_id[parent_id]["children"]
                else:
                    siblings = hierarchy
                siblings.append(new_node)
                siblings.sort(key=lambda n: n["name"].lower())

                with open(_album_hierarchy_file(), "w", encoding="utf-8") as f:
                    json.dump(hierarchy, f, indent=2, ensure_ascii=False)

                parent.after(0, lambda fn=new_fullname, nid=new_id: finish_ok(fn, nid))
            except Exception as exc:
                err = str(exc)
                parent.after(0, lambda: finish_err(err))
            finally:
                client.logout()

        threading.Thread(target=worker, daemon=True).start()

    def finish_ok(new_fullname, new_id):
        dlg.destroy()
        set_status_cb(
            f"Album '{new_fullname}' created (id {new_id}). "
            f"{_album_hierarchy_file().name} updated."
        )

    def finish_err(err):
        create_btn.config(state=tk.NORMAL)
        cancel_btn.config(state=tk.NORMAL)
        status_var.set("")
        messagebox.showerror("Piwigo error", err, parent=dlg)
        set_status_cb("Album creation failed.")

    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=dlg.destroy)
    cancel_btn.pack(side='right', padx=(4, 0))
    create_btn = ttk.Button(btn_frame, text="Create", command=on_create)
    create_btn.pack(side='right')

    dlg.bind("<Return>", lambda e: on_create())


# ---------------------------------------------------------------------------
# Album picker
# ---------------------------------------------------------------------------
def _hierarchy_is_fresh() -> bool:
    """Return True if AlbumHierarchy.json exists and is less than 24 hours old."""
    if not _album_hierarchy_file().exists():
        return False
    return (time.time() - _album_hierarchy_file().stat().st_mtime) < 86400


def pick_album(parent: tk.Widget, set_status_cb, on_select_cb):
    """Show an album picker, auto-refreshing the hierarchy first if stale.

    on_select_cb(album_id: int, fullname: str) is called when the user
    confirms a selection.
    """

    def open_picker():
        try:
            with open(_album_hierarchy_file(), encoding="utf-8") as f:
                hierarchy = json.load(f)
        except Exception as exc:
            messagebox.showerror("Error",
                                 f"Cannot read album hierarchy:\n{exc}",
                                 parent=parent)
            return
        _show_picker_dialog(parent, hierarchy, on_select_cb)

    if _hierarchy_is_fresh():
        open_picker()
        return

    # Hierarchy is missing or stale — refresh silently first
    try:
        params = load_params()
    except (FileNotFoundError, ValueError) as exc:
        messagebox.showerror("Configuration error", str(exc), parent=parent)
        return

    dlg = tk.Toplevel(parent)
    dlg.title("Refreshing Album Hierarchy")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.geometry("340x90")
    _center_dialog(parent, dlg)

    ttk.Label(dlg, text="Refreshing album hierarchy from Piwigo…",
              padding=(12, 10, 12, 4)).pack()
    bar = ttk.Progressbar(dlg, mode="indeterminate", length=300)
    bar.pack(padx=12, pady=(0, 12))
    bar.start(12)

    def on_done():
        bar.stop()
        dlg.destroy()
        set_status_cb("Album hierarchy refreshed.")
        open_picker()

    def on_err(err):
        bar.stop()
        dlg.destroy()
        if _album_hierarchy_file().exists():
            if messagebox.askyesno(
                "Refresh failed",
                f"Could not refresh album hierarchy:\n{err}\n\n"
                "Use the existing (possibly stale) data instead?",
                parent=parent,
            ):
                open_picker()
        else:
            messagebox.showerror("Refresh failed", err, parent=parent)

    def worker():
        client = PiwigoClient(
            params["url"],
            params["username"],
            params["password"],
            verify_ssl=params.get("verify_ssl", True),
            rate_limit_calls_per_second=params.get("rate_limit_calls_per_second", 2.0),
        )
        try:
            client.login(params["username"], params["password"])
            _fetch_and_save_hierarchy(client, lambda msg: None)
            parent.after(0, on_done)
        except Exception as exc:
            err = str(exc)
            parent.after(0, lambda: on_err(err))
        finally:
            client.logout()

    threading.Thread(target=worker, daemon=True).start()


def _show_picker_dialog(parent: tk.Widget, hierarchy: list, on_select_cb):
    """Render the album-tree picker dialog (RV Menu Tree style)."""
    dlg = tk.Toplevel(parent)
    dlg.title("Select Upload Album")
    dlg.grab_set()

    dlg.geometry("440x580")
    _center_dialog(parent, dlg)

    # ── Filter bar ───────────────────────────────────────────────────────────
    filter_frame = ttk.Frame(dlg, padding=(8, 8, 8, 4))
    filter_frame.pack(fill='x')
    ttk.Label(filter_frame, text="Filter:").pack(side='left', padx=(0, 6))
    filter_var = tk.StringVar()
    filter_entry = ttk.Entry(filter_frame, textvariable=filter_var)
    filter_entry.pack(side='left', fill='x', expand=True)
    filter_entry.focus_set()

    # ── Tree ─────────────────────────────────────────────────────────────────
    tree_frame = ttk.Frame(dlg, padding=(8, 0, 8, 4))
    tree_frame.pack(fill='both', expand=True)

    yscroll = ttk.Scrollbar(tree_frame, orient='vertical')
    xscroll = ttk.Scrollbar(tree_frame, orient='horizontal')
    tree = ttk.Treeview(tree_frame, selectmode="browse", show="tree",
                        yscrollcommand=yscroll.set,
                        xscrollcommand=xscroll.set)
    yscroll.config(command=tree.yview)
    xscroll.config(command=tree.xview)
    yscroll.pack(side='right', fill='y')
    xscroll.pack(side='bottom', fill='x')
    tree.pack(side='left', fill='both', expand=True)
    tree.column("#0", minwidth=200)

    # Populate; top-level nodes open, children collapsed
    all_items = []  # (iid, name_lower, fullname)

    def _populate(parent_iid, nodes, top_level=False):
        for node in nodes:
            iid     = str(node["id"])
            count   = node["total_nb_images"]
            text    = f"{node['name']}  (id {node['id']}, {count:,} photos)"
            tree.insert(parent_iid, 'end', iid=iid, text=text,
                        open=top_level)
            all_items.append((iid,
                               node["name"].lower(),
                               node.get("fullname", node["name"])))
            if node.get("children"):
                _populate(iid, node["children"])

    _populate("", hierarchy, top_level=True)

    # ── Filter: scroll to and select first match ──────────────────────────
    def _on_filter(*_):
        q = filter_var.get().strip().lower()
        if not q:
            tree.selection_remove(tree.selection())
            return
        for iid, name_lower, _ in all_items:
            if q in name_lower:
                tree.selection_set(iid)
                tree.see(iid)
                break

    filter_var.trace_add("write", _on_filter)

    # ── Selected-album label ─────────────────────────────────────────────────
    sel_frame = ttk.Frame(dlg, padding=(8, 0, 8, 6))
    sel_frame.pack(fill='x')
    ttk.Label(sel_frame, text="Selected:").pack(side='left', padx=(0, 4))
    sel_var = tk.StringVar(value="(none)")
    ttk.Label(sel_frame, textvariable=sel_var,
              foreground="gray", anchor='w').pack(side='left',
                                                    fill='x', expand=True)

    # Build a fast iid→fullname lookup
    fullname_by_iid = {iid: fn for iid, _, fn in all_items}

    def _on_tree_select(_event):
        sel = tree.selection()
        sel_var.set(fullname_by_iid.get(sel[0], "(none)") if sel else "(none)")

    tree.bind("<<TreeviewSelect>>", _on_tree_select)

    # ── Buttons ──────────────────────────────────────────────────────────────
    btn_frame = ttk.Frame(dlg, padding=(8, 0, 8, 10))
    btn_frame.pack(fill='x')

    def on_select():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("No album selected",
                                   "Please select an album first.",
                                   parent=dlg)
            return
        album_id  = int(sel[0])
        fullname  = fullname_by_iid.get(sel[0], "")
        dlg.destroy()
        on_select_cb(album_id, fullname)

    ttk.Button(btn_frame, text="Cancel",
               command=dlg.destroy).pack(side='right', padx=(4, 0))
    ttk.Button(btn_frame, text="Select",
               command=on_select).pack(side='right')
    dlg.bind("<Return>", lambda e: on_select())


# ---------------------------------------------------------------------------
# File-index download (optional)
# ---------------------------------------------------------------------------
def download_file_index(parent: tk.Widget, set_status_cb):
    """Walk every album on the server and build a filename → album-path index.

    The result is written to FileDict.json as:
        { "photo.jpg": ["Album / SubAlbum", "Other Album"], … }

    A determinate progress dialog shows one row per album processed.
    If AlbumHierarchy.json is missing or stale the hierarchy is refreshed
    first (re-uses the same credentials and session).
    """

    try:
        params = load_params()
    except (FileNotFoundError, ValueError) as exc:
        messagebox.showerror("Configuration error", str(exc), parent=parent)
        return

    # ── Progress dialog ───────────────────────────────────────────────────────
    dlg = tk.Toplevel(parent)
    dlg.title("Downloading File Index")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.geometry("400x130")
    _center_dialog(parent, dlg)

    ttk.Label(dlg, text="Building file index from Piwigo…",
              padding=(12, 10, 12, 2)).pack()

    step_var = tk.StringVar(value="Connecting…")
    ttk.Label(dlg, textvariable=step_var, foreground="gray",
              padding=(12, 0, 12, 4)).pack()

    bar = ttk.Progressbar(dlg, mode="determinate", length=360)
    bar.pack(padx=12, pady=(0, 12))

    def set_step(msg):
        step_var.set(msg)
        set_status_cb(msg)

    def on_progress(done, total, album_name):
        parent.after(0, lambda d=done, t=total, n=album_name: _apply_progress(d, t, n))

    def _apply_progress(done, total, album_name):
        bar["maximum"] = total
        bar["value"]   = done
        step_var.set(f"({done}/{total})  {album_name}")

    def finish_ok(n_files, n_albums):
        bar.stop()
        dlg.destroy()
        msg = (f"File index built: {n_files:,} unique file(s) across "
               f"{n_albums} album(s). Written to {_file_index_file().name}.")
        set_status_cb(msg)

    def finish_err(err):
        bar.stop()
        dlg.destroy()
        messagebox.showerror("Piwigo error", err, parent=parent)
        set_status_cb("File index download failed.")

    # ── Background worker ─────────────────────────────────────────────────────
    def worker():
        client = PiwigoClient(
            params["url"],
            params["username"],
            params["password"],
            verify_ssl=params.get("verify_ssl", True),
            rate_limit_calls_per_second=params.get("rate_limit_calls_per_second", 2.0),
        )
        try:
            parent.after(0, lambda: set_step("Logging in…"))
            client.login(params["username"], params["password"])

            # Refresh hierarchy if stale so flat_albums is up to date
            parent.after(0, lambda: set_step("Fetching album list…"))
            flat_albums = client.get_albums()

            if not _hierarchy_is_fresh():
                parent.after(0, lambda: set_step("Saving album hierarchy…"))
                hierarchy = _build_hierarchy(flat_albums)
                with open(_album_hierarchy_file(), "w", encoding="utf-8") as f:
                    json.dump(hierarchy, f, indent=2, ensure_ascii=False)

            index = _fetch_and_save_file_index(client, flat_albums, on_progress)
            parent.after(0, lambda: finish_ok(len(index), len(flat_albums)))
        except Exception as exc:
            err = str(exc)
            parent.after(0, lambda: finish_err(err))
        finally:
            client.logout()

    threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# File-index update after upload
# ---------------------------------------------------------------------------
def record_uploaded_file(filename: str, album_fullname: str,
                         album_id: int = 0, file_id: int = 0):
    """Add filename → entry to FileDict.json after a successful upload.

    Each entry is {"fullname": str, "album_id": int, "file_id": int}.
    Loads the existing index if present, appends the new entry (avoiding
    duplicates by album_id), and writes the file back.
    Safe to call from the main thread.
    """
    index: dict[str, list[dict]] = {}
    if _file_index_file().exists():
        try:
            with open(_file_index_file(), encoding="utf-8") as f:
                index = json.load(f)
        except Exception:
            pass

    entries = index.setdefault(filename, [])
    if not any(e.get("album_id") == album_id for e in entries):
        entries.append({"fullname": album_fullname,
                        "album_id": album_id,
                        "file_id":  file_id})

    with open(_file_index_file(), "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False, sort_keys=True)
