"""
DownloadAlbumStructure.py
Connects to a Piwigo instance, downloads the full album hierarchy, writes
it to AlbumHierarchy.json, and shows a modal progress dialog while working.

Public API
----------
    run(parent, set_status_cb)
        parent        – a tkinter widget used as the dialog parent
        set_status_cb – callable(str) that updates the caller's status bar
"""

import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

try:
    import requests
    import urllib3
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
PARAMS_FILE          = Path(".") / "PhotosUploader Params.json"
ALBUM_HIERARCHY_FILE = Path(".") / "AlbumHierarchy.json"
REQUIRED_PARAMS      = ("url", "username", "password")


# ---------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------
def load_params() -> dict:
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
            '}'
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
                 verify_ssl: bool = True):
        url = base_url.strip().rstrip("/")
        if url.startswith("http://"):
            url = "https://" + url[7:]
        elif not url.startswith("https://"):
            url = "https://" + url
        self.base_url = url
        self.api_url  = f"{self.base_url}/ws.php?format=json"
        self.session  = requests.Session()
        self.session.verify = verify_ssl
        if not verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def _call(self, method: str, params: dict = None) -> dict:
        payload = {"method": method}
        if params:
            payload.update(params)
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


# ---------------------------------------------------------------------------
# Hierarchy builder
# ---------------------------------------------------------------------------
def _build_hierarchy(flat: list) -> list:
    """Convert a flat list of Piwigo category dicts into a nested tree.

    Each node in the result has:
        id, name, fullname, nb_images, total_nb_images, children
    Root nodes (id_uppercat absent, null, or "0") appear at the top level.
    """
    by_id = {}
    for cat in flat:
        node = {
            "id":              int(cat["id"]),
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
        by_id[node["id"]] = node

    roots = []
    for cat in flat:
        node      = by_id[int(cat["id"])]
        parent_id = cat.get("id_uppercat")
        if parent_id and str(parent_id) != "0" and int(parent_id) in by_id:
            by_id[int(parent_id)]["children"].append(node)
        else:
            roots.append(node)

    def _sort(nodes):
        nodes.sort(key=lambda n: n["name"].lower())
        for n in nodes:
            _sort(n["children"])

    _sort(roots)
    return roots


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def run(parent: tk.Widget, set_status_cb):
    """Download album hierarchy and write AlbumHierarchy.json.

    Opens a modal progress dialog on *parent*.  Calls set_status_cb(str)
    to update the caller's status bar at each step and on completion.
    """
    if not REQUESTS_AVAILABLE:
        messagebox.showerror(
            "Missing dependency",
            "The 'requests' library is required.\nRun: pip install requests",
            parent=parent,
        )
        return

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

    parent.update_idletasks()
    rx = parent.winfo_x() + parent.winfo_width()  // 2
    ry = parent.winfo_y() + parent.winfo_height() // 2
    dlg.geometry(f"360x120+{rx - 180}+{ry - 60}")

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
               f"Hierarchy written to {ALBUM_HIERARCHY_FILE.name}")
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
        )
        try:
            parent.after(0, lambda: set_step("Logging in…"))
            client.login(params["username"], params["password"])

            parent.after(0, lambda: set_step("Fetching album list…"))
            flat = client.get_albums()

            parent.after(0, lambda: set_step("Building hierarchy…"))
            hierarchy = _build_hierarchy(flat)

            parent.after(0, lambda: set_step("Writing AlbumHierarchy.json…"))
            with open(ALBUM_HIERARCHY_FILE, "w", encoding="utf-8") as f:
                json.dump(hierarchy, f, indent=2, ensure_ascii=False)

            n = len(flat)
            parent.after(0, lambda: finish_ok(n))
        except Exception as exc:
            err = str(exc)
            parent.after(0, lambda: finish_err(err))
        finally:
            client.logout()

    threading.Thread(target=worker, daemon=True).start()
