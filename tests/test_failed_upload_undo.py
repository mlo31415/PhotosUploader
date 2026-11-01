"""A failed upload leaves things as they were before it.

With Auto Rename on, the file is renamed just before it is sent -- and when
the server refused the upload, the rename stayed: the file kept its new name,
the counter stayed used, and trying again renamed the renamed file, giving
"es-jk00121 - es-jk00120 - PICT0943.jpg".

The real app is built, against a temp folder for its state, credentials and
file index, and a stand-in Piwigo client takes the place of the network.
Every dialog is stubbed, so nothing appears on screen but the (withdrawn)
main window.

Run with:  python -m unittest discover -s tests
"""
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_APP = _HERE.parent / "PhotosUploader.py"
sys.path.insert(0, str(_APP.parent.parent / "PiwigoHelpers"))
sys.path.insert(0, str(_APP.parent))


def load_pu():
    spec = importlib.util.spec_from_file_location("pu_undo_under_test", _APP)
    module = importlib.util.module_from_spec(spec)
    module.__spec__ = spec
    spec.loader.exec_module(module)
    return module


pu = load_pu()
from CredentialStore import CredentialStore                     # noqa: E402
from PIL import Image                                           # noqa: E402

REFUSAL = ("The server refused the upload before it reached Piwigo (HTTP 403).")


class StandInClient:
    """Takes the place of AlbumHierarchy.PiwigoClient.  What upload_image
    does is set per test, through the class attribute."""
    on_upload = None
    uploads = []

    def __init__(self, *a, **k):
        pass

    def login(self, *a):
        pass

    def logout(self):
        pass

    def upload_image(self, path, album_id, **kw):
        StandInClient.uploads.append(kw.get("name"))
        return StandInClient.on_upload(path, kw)

    def sync_metadata(self, image_id):
        pass

    def refresh_representative(self, album_id):
        pass


def refuse(path, kw):
    raise RuntimeError(REFUSAL)


def accept(path, kw):
    return {"image_id": 4321}


class _Fixture:

    def setUp(self):
        self.work = Path(tempfile.mkdtemp())
        self.photos = self.work / "photos"
        self.photos.mkdir()
        (self.work / "Piwigo Credentials.json").write_text(json.dumps(
            {"url": "https://piwigo.invalid", "username": "u", "password": "p",
             "verify_ssl": False}), encoding="utf-8")

        self.saved = (pu.AlbumHierarchy.PARAMS_FILE, pu._store,
                      pu.AlbumHierarchy.PiwigoClient,
                      pu.messagebox.showerror, pu.messagebox.showinfo,
                      pu.messagebox.showwarning, pu.messagebox.askyesno,
                      pu.messagebox.askokcancel, pu.messagebox.askyesnocancel)
        pu.AlbumHierarchy.PARAMS_FILE = self.work / "PhotosUploader Params.json"
        pu._store = CredentialStore(self.work, "PhotosUploader Params.json")
        pu.AlbumHierarchy.PiwigoClient = StandInClient
        StandInClient.uploads = []
        self.errors = []
        pu.messagebox.showerror = lambda title, msg, **k: self.errors.append(msg)
        pu.messagebox.showinfo = pu.messagebox.showwarning = lambda *a, **k: None
        pu.messagebox.askyesno = pu.messagebox.askokcancel = lambda *a, **k: True
        pu.messagebox.askyesnocancel = lambda *a, **k: True

        # The same kind of root main() makes: the queue registers for drops
        self.root = pu.TkinterDnD.Tk() if pu.DND_AVAILABLE else tk.Tk()
        self.root.withdraw()
        self.app = pu.PhotosUploader(self.root)
        self.root.withdraw()
        self.app.upload_album_var.set("Test Album")
        self.app.upload_album_id = 77

    def tearDown(self):
        try:
            # The withdrawn window never gets a height, so the sash placement
            # keeps re-scheduling itself; stop it outliving the window
            for pending in self.root.tk.splitlist(self.root.tk.call('after', 'info')):
                self.root.after_cancel(pending)
            self.root.destroy()
        except tk.TclError:
            pass
        (pu.AlbumHierarchy.PARAMS_FILE, pu._store,
         pu.AlbumHierarchy.PiwigoClient,
         pu.messagebox.showerror, pu.messagebox.showinfo,
         pu.messagebox.showwarning, pu.messagebox.askyesno,
         pu.messagebox.askokcancel, pu.messagebox.askyesnocancel) = self.saved
        shutil.rmtree(self.work, ignore_errors=True)

    def queue_photo(self, name="PICT0943.jpg", counter=None):
        """Put a photo in the queue with Auto Rename on, and select it."""
        if counter is not None:
            self.app.state_data.setdefault("auto_rename_counts", {})["es-jk"] = counter
        path = str(self.photos / name)
        Image.new("RGB", (60, 40), "#3080c0").save(path)
        self.app.auto_rename_var.set("es-jk")
        self.app._add_single_image(path, {})
        self.app.input_list.selection_set(0)
        self.app._on_input_select(None)
        self.app.custom_vars['date_of_photo'].set("1962-09-01")
        self.app.custom_vars['comments'].insert('1.0', "A caption")
        self.root.update()
        return path

    def upload(self, behaviour):
        """Upload the current photo and wait for the outcome.  The worker
        reports back through root.after, which needs a running mainloop."""
        StandInClient.on_upload = behaviour
        self.app.status_var.set("")
        deadline = time.time() + 15
        said = []

        def poll():
            now = self.app.status_var.get()
            if now.startswith(("Upload failed", "Uploaded", "All ")) \
                    or time.time() > deadline:
                said.append(now)
                self.root.quit()
                return
            self.root.after(20, poll)

        def start():
            self.app._upload_current_photo()
            poll()
        self.root.after(0, start)
        self.root.mainloop()
        if time.time() > deadline:
            self.fail("the upload never finished")
        return said[0]

    def files(self):
        return sorted(os.listdir(self.photos))

    def counter_on_disk(self):
        state = json.loads((self.work / "PhotosUploader State.json")
                           .read_text(encoding="utf-8"))
        return state.get("auto_rename_counts", {}).get("es-jk")


class ARefusedUpload(_Fixture, unittest.TestCase):

    def test_the_file_gets_its_own_name_back(self):
        self.queue_photo()
        self.upload(refuse)
        self.assertEqual(self.files(), ["PICT0943.jpg"])

    def test_the_queue_names_the_file_on_disk(self):
        path = self.queue_photo()
        self.upload(refuse)
        self.assertEqual(self.app.input_paths, [path])
        self.assertEqual(self.app.current_photo, path)
        self.assertIn("PICT0943", self.app.input_list.get(0))
        self.assertEqual(self.app.input_list.curselection(), (0,))

    def test_it_is_still_waiting_to_be_uploaded(self):
        path = self.queue_photo()
        self.upload(refuse)
        self.assertIn(path, self.app._dirty_paths)
        self.assertNotIn(path, self.app.uploaded_info)

    def test_the_number_is_not_used_up(self):
        self.queue_photo(counter=119)
        self.upload(refuse)
        self.assertEqual(self.app.state_data["auto_rename_counts"]["es-jk"], 119)
        self.assertEqual(self.counter_on_disk(), 119)

    def test_a_counter_that_was_never_set_is_left_unset(self):
        self.queue_photo()
        self.upload(refuse)
        self.assertNotIn("es-jk", self.app.state_data.get("auto_rename_counts", {}))
        self.assertIsNone(self.counter_on_disk())

    def test_the_output_filename_is_what_it_was(self):
        self.queue_photo(counter=119)
        before = self.app.custom_vars['output_filename'].get()
        self.upload(refuse)
        self.assertEqual(self.app.custom_vars['output_filename'].get(), before)
        self.assertEqual(before, "es-jk00120.jpg")

    def test_the_user_is_still_told_why(self):
        self.queue_photo()
        self.upload(refuse)
        self.assertEqual(len(self.errors), 1)
        self.assertIn(REFUSAL, self.errors[0])
        self.assertNotIn("still named", self.errors[0])


class TryingAgain(_Fixture, unittest.TestCase):
    """The bug as it was met: retrying renamed the renamed file."""

    def test_a_second_failure_still_leaves_the_original(self):
        self.queue_photo(counter=119)
        self.upload(refuse)
        self.upload(refuse)
        self.assertEqual(self.files(), ["PICT0943.jpg"])
        self.assertEqual(StandInClient.uploads, ["es-jk00120.jpg", "es-jk00120.jpg"])

    def test_then_succeeding_gives_one_prefix_and_the_same_number(self):
        self.queue_photo(counter=119)
        self.upload(refuse)
        self.upload(accept)
        self.assertEqual(self.files(), ["es-jk00120 - PICT0943.jpg"])
        self.assertEqual(self.counter_on_disk(), 120)
        self.assertEqual(StandInClient.uploads[-1], "es-jk00120.jpg")


class WhenItCannotAllBeUndone(_Fixture, unittest.TestCase):

    def test_a_file_that_cannot_be_renamed_back_is_reported_and_followed(self):
        """Something has taken the original name meanwhile.  The queue must
        go on naming the file that is actually there, and the number stays
        used, since that name is now on disk."""
        path = self.queue_photo(counter=119)

        def take_the_name_then_refuse(upload_path, kw):
            Path(path).write_bytes(b"in the way")
            raise RuntimeError(REFUSAL)

        self.upload(take_the_name_then_refuse)
        renamed = str(self.photos / "es-jk00120 - PICT0943.jpg")
        self.assertTrue(os.path.exists(renamed))
        self.assertIn("still named es-jk00120 - PICT0943.jpg", self.errors[0])
        self.assertEqual(self.app.input_paths, [renamed])
        self.assertEqual(self.app.current_photo, renamed)
        self.assertEqual(self.counter_on_disk(), 120)


class AFailureBeforeTheUpload(_Fixture, unittest.TestCase):

    def test_a_client_that_cannot_be_made_is_reported_and_undone(self):
        """It used to be made outside the try: the worker died silently and
        the modal progress dialog stayed up for good."""
        self.queue_photo(counter=119)

        class Broken(StandInClient):
            def __init__(self, *a, **k):
                raise ValueError("no such server")
        pu.AlbumHierarchy.PiwigoClient = Broken
        said = self.upload(refuse)
        self.assertTrue(said.startswith("Upload failed"), said)
        self.assertIn("no such server", self.errors[0])
        self.assertEqual(self.files(), ["PICT0943.jpg"])
        dialogs = [w for w in self.root.winfo_children()
                   if isinstance(w, tk.Toplevel) and w.winfo_exists()
                   and w.title() == "Uploading…"]
        self.assertEqual(dialogs, [])


if __name__ == "__main__":
    unittest.main()
