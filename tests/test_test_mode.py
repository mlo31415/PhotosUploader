"""Test Mode: the program works as usual, and nothing lasting is changed.

Nothing is sent to Piwigo that would change the site, and no lasting file on
this computer is written: not the state file, not the shared caches, not the
photos.  Turning it off undoes what was done in it.

Built on the real app, as in test_failed_upload_undo, with every dialog
stubbed and the state in a temp folder.

Run with:  python -m unittest discover -s tests
"""
import time
import unittest
from pathlib import Path

from PIL import Image

import test_failed_upload_undo as base
import _slides

pu = base.pu
AH = pu.AlbumHierarchy


class _TestMode(base._Fixture):

    def setUp(self):
        super().setUp()
        self.state_file = self.work / "PhotosUploader State.json"
        self.app._download_and_load_uploaded_real = self.app._download_and_load_uploaded

    def tearDown(self):
        if self.app._test_mode:
            try:
                self.app._set_test_mode(False)
            except Exception:
                pass
        super().tearDown()

    def add(self, *names):
        paths = []
        for n in names:
            p = str(self.photos / n)
            Image.new("RGB", (60, 40), "#3080c0").save(p)
            self.app._add_single_image(p, {})
            paths.append(p)
        self.app._go_to_queue_index(0)
        self.root.update()
        return paths

    def ready(self):
        """Fill in what an upload needs."""
        self.app.custom_vars['date_of_photo'].set("1962-09-01")
        if not self.app.custom_vars['comments'].get('1.0', 'end').strip():
            self.app.custom_vars['comments'].insert('1.0', "A caption")
        self.root.update()

    def upload(self, behaviour=base.accept):
        """As base.upload, but Test Mode says what it did in its own words."""
        base.StandInClient.on_upload = behaviour
        self.app.status_var.set("")
        deadline = time.time() + 15
        said = []

        def poll():
            now = self.app.status_var.get()
            if now.startswith(("Upload failed", "Uploaded", "All ", "Test Mode:")) \
                    or time.time() > deadline:
                said.append(now)
                self.root.quit()
                return
            self.root.after(20, poll)

        self.root.after(0, lambda: (self.app._upload_current_photo(), poll()))
        self.root.mainloop()
        return said[0]

    def state_on_disk(self):
        return self.state_file.read_bytes() if self.state_file.exists() else None

    def on(self):
        self.app._set_test_mode(True)
        self.root.update()

    def off(self):
        self.app._set_test_mode(False)
        self.root.update()


class TheButton(_TestMode, unittest.TestCase):

    def test_ordinary_when_off(self):
        b = self.app.test_mode_btn
        self.assertEqual(b.cget("relief"), "raised")
        self.assertNotEqual(b.cget("bg"), self.app._TEST_ORANGE)

    def test_bright_orange_when_on(self):
        self.app.test_mode_btn.invoke()
        b = self.app.test_mode_btn
        self.assertTrue(self.app._test_mode)
        self.assertEqual((b.cget("bg"), b.cget("relief")), (self.app._TEST_ORANGE, "sunken"))

    def test_and_ordinary_again(self):
        before = self.app.test_mode_btn.cget("bg")
        self.app.test_mode_btn.invoke()
        self.app.test_mode_btn.invoke()
        self.assertFalse(self.app._test_mode)
        self.assertEqual(self.app.test_mode_btn.cget("bg"), before)
        self.assertEqual(self.app.test_mode_btn.cget("relief"), "raised")


class Piwigo(_TestMode, unittest.TestCase):

    def test_an_upload_sends_nothing(self):
        self.add("a.jpg")
        self.ready()
        self.on()
        said = self.upload()
        self.assertEqual(base.StandInClient.uploads, [])
        self.assertTrue(said.startswith("Test Mode:"), said)

    def test_but_otherwise_goes_as_usual(self):
        paths = self.add("a.jpg", "b.jpg")
        self.ready()
        self.on()
        self.upload()
        self.assertIn("(uploaded)", self.app.input_list.get(0))
        self.assertNotIn(paths[0], self.app._dirty_paths)
        self.assertEqual(self.app.current_photo, paths[1], "moved on to the next")

    def test_every_client_refuses_what_would_change_the_site(self):
        sent = []

        class Client:
            def _call(self, method, params=None):
                sent.append(method)
                return {}

            def upload_image(self, *a, **k):
                sent.append("upload")

        ro = pu._read_only_client(Client)()
        for method in ("pwg.session.login", "pwg.categories.getList", "pwg.images.getInfo",
                       "pwg.tags.getList", "pwg.categories.getImages", "pwg.session.logout"):
            ro._call(method)
        for method in ("pwg.categories.add", "pwg.categories.setInfo", "pwg.categories.move",
                       "pwg.categories.delete", "pwg.images.setInfo", "pwg.tags.add",
                       "pwg.images.syncMetadata", "pwg.categories.refreshRepresentative"):
            with self.assertRaises(pu.TestModeRefusal, msg=method):
                ro._call(method)
        with self.assertRaises(pu.TestModeRefusal):
            ro.upload_image("x.jpg", 77)
        self.assertEqual(sent, ["pwg.session.login", "pwg.categories.getList", "pwg.images.getInfo",
                                "pwg.tags.getList", "pwg.categories.getImages", "pwg.session.logout"])

    def test_the_shared_helpers_get_the_refusing_client_too(self):
        real = AH.PiwigoClient
        self.on()
        self.assertTrue(issubclass(AH.PiwigoClient, real))
        self.assertIsNot(AH.PiwigoClient, real)
        self.off()
        self.assertIs(AH.PiwigoClient, real)


class ThisComputer(_TestMode, unittest.TestCase):

    def test_the_state_file_is_not_written(self):
        self.add("a.jpg")
        self.ready()
        before = self.state_on_disk()
        self.on()
        self.upload()
        self.app._save_state_quietly()
        self.app._persist_state()
        self.assertEqual(self.state_on_disk(), before)

    def test_nor_on_closing(self):
        self.app._persist_state()
        before = self.state_on_disk()
        self.on()
        self.app.root.geometry("900x700+50+50")
        self.app._on_close()
        self.assertEqual(self.state_on_disk(), before)
        self.assertFalse(pu._test_mode, "the next start begins outside Test Mode")

    def test_auto_rename_leaves_the_file_alone(self):
        self.app.state_data.setdefault("auto_rename_counts", {})["es-jk"] = 119
        self.app.auto_rename_var.set("es-jk")
        self.add("PICT0943.jpg", "PICT0944.jpg")
        self.ready()
        self.on()
        said = self.upload()
        self.assertEqual(self.files(), ["PICT0943.jpg", "PICT0944.jpg"])
        self.assertIn("es-jk00120.jpg", said)
        self.off()
        self.assertEqual(self.app.state_data["auto_rename_counts"]["es-jk"], 119,
                         "the number is not used up")

    def test_the_caches_are_copies(self):
        real_index = AH._data_dir() / "FileDict.json"            # still under its old name
        real_index.write_text('{"old.jpg": []}', encoding="utf-8")
        self.on()
        copy = AH._file_index_file()
        self.assertNotEqual(copy.parent, AH._data_dir())
        self.assertEqual(copy.read_text(encoding="utf-8"), '{"old.jpg": []}')
        AH.record_uploaded_file("new.jpg", "Test Album", album_id=77, file_id=1)
        self.off()
        self.assertEqual(real_index.read_text(encoding="utf-8"), '{"old.jpg": []}')
        self.assertTrue(real_index.exists(), "not even renamed")
        self.assertFalse(copy.exists(), "the copies go when Test Mode ends")

    def test_revert_puts_its_copy_in_the_temp_folder(self):
        self.add("a.jpg")
        self.on()
        self.app._revert_photo()
        self.assertEqual(self.files(), ["a.jpg"])

    def test_an_uploaded_photo_is_not_overwritten_with_piwigos(self):
        paths = self.add("a.jpg", "b.jpg")
        self.ready()
        self.upload()                                           # really, before Test Mode
        before = Path(paths[0]).read_bytes()
        self.on()
        self.app._download_and_load_uploaded = self.app._download_and_load_uploaded_real
        self.app._go_to_queue_index(0)
        self.root.update()
        self.assertEqual(self.app.current_photo, paths[0])
        self.assertEqual(Path(paths[0]).read_bytes(), before)


class Undone(_TestMode, unittest.TestCase):

    def test_its_uploads_are_undone(self):
        paths = self.add("a.jpg", "b.jpg")
        self.ready()
        self.on()
        self.upload()
        self.off()
        self.assertNotIn(paths[0], self.app.uploaded_info)
        self.assertIn(paths[0], self.app._dirty_paths)
        self.assertNotIn("(uploaded)", self.app.input_list.get(0))

    def test_a_real_upload_before_it_stands(self):
        paths = self.add("a.jpg", "b.jpg")
        self.ready()
        self.upload()                                           # really
        self.on()
        self.app._go_to_queue_index(0)
        self.app.custom_vars['comments'].insert('end', " again")
        self.upload()                                           # in Test Mode
        self.off()
        self.assertEqual(self.app.uploaded_info[paths[0]]["image_id"], 4321)
        self.assertIn("(uploaded)", self.app.input_list.get(0))

    def test_pdf_progress_is_undone(self):
        pdf = _slides.make_slideshow(self.work / "Slide Show.pdf")
        self.app._con_list_cache = [pu.PdfPhotos.ConList.load(_slides.make_timeline(self.work / "t.txt"))]
        self.on()
        self.app._add_pdf(str(pdf), {})
        self.app._go_to_queue_index(0)
        self.app._skip_photo()
        self.assertEqual(self.app.state_data["pdf_progress"]["Slide Show.pdf"], ["p002"])
        self.off()
        self.assertNotIn("pdf_progress", self.app.state_data)
        self.assertIsNone(self.state_on_disk())

    def test_it_is_off_when_the_program_starts(self):
        self.assertFalse(self.app._test_mode)
        self.assertFalse(pu._test_mode)


if __name__ == "__main__":
    unittest.main()
