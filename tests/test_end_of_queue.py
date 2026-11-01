"""Moving past the last photo in the queue is the end of the queue.

If every photo has been uploaded, the queue is cleared, as it is after the
last upload.  If not, the user is offered the first photo that hasn't been:
Yes goes there, No clears the queue anyway, Cancel stays put.

Built on the real app, as in test_failed_upload_undo, with every dialog
stubbed and the state in a temp folder.

Run with:  python -m unittest discover -s tests
"""
import unittest

from PIL import Image

import test_failed_upload_undo as base

pu = base.pu


class _Queue(base._Fixture):

    def setUp(self):
        super().setUp()
        self.asked = []
        self.answer = True
        pu.messagebox.askyesnocancel = self.ask
        # Moving onto an uploaded photo fetches Piwigo's copy; the local one will do
        self.app._download_and_load_uploaded = self.app._load_photo

    def ask(self, title, msg, **k):
        self.asked.append(msg)
        return self.answer

    def fill(self, *names):
        """Queue photos and select the first."""
        paths = []
        for n in names:
            p = str(self.photos / n)
            Image.new("RGB", (60, 40), "#3080c0").save(p)
            self.app._add_single_image(p, {})
            paths.append(p)
        self.app.input_list.selection_set(0)
        self.app._on_input_select(None)
        self.root.update()
        return paths

    def uploaded(self, *paths):
        """Mark photos uploaded, as a successful upload does."""
        for p in paths:
            self.app.uploaded_info[p] = {"image_id": 1, "album_id": 77,
                                         "output_filename": "x.jpg"}
            self.app._dirty_paths.discard(p)

    def down(self, times=1):
        for _ in range(times):
            self.app._nav_next()
            self.root.update()


class InTheMiddle(_Queue, unittest.TestCase):
    """Up and Down used to move only the highlight: a selection made in code
    raises no <<ListboxSelect>>, so the photo shown never changed."""

    def test_down_moves_on_and_loads_the_photo(self):
        paths = self.fill("a.jpg", "b.jpg", "c.jpg")
        self.down()
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(self.app.photo_label_var.get(), "b.jpg")
        self.assertEqual(self.asked, [])
        self.assertEqual(len(self.app.input_paths), 3)

    def test_up_loads_the_photo_too(self):
        paths = self.fill("a.jpg", "b.jpg", "c.jpg")
        self.down(2)
        self.app._nav_prev()
        self.root.update()
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(self.app.photo_label_var.get(), "b.jpg")

    def test_leaving_an_edited_photo_still_asks_once(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.app._user_edited_paths.add(paths[0])
        stay = []
        pu.messagebox.askyesno = lambda *a, **k: (stay.append(1), False)[1]
        self.down()
        self.assertEqual(self.app.current_photo, paths[0])
        self.assertEqual(self.app.input_list.curselection(), (0,))
        self.assertEqual(len(stay), 1)


class AllUploaded(_Queue, unittest.TestCase):

    def test_moving_past_the_last_clears_the_queue(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.uploaded(*paths)
        self.down(2)
        self.assertEqual(self.app.input_paths, [])
        self.assertEqual(self.app.input_list.size(), 0)
        self.assertIsNone(self.app.current_photo)
        self.assertEqual(self.app.status_var.get(), "All 2 photo(s) uploaded — queue cleared.")

    def test_without_asking(self):
        paths = self.fill("a.jpg")
        self.uploaded(*paths)
        self.down()
        self.assertEqual(self.asked, [])

    def test_the_files_are_left_alone(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.uploaded(*paths)
        self.down(2)
        self.assertEqual(self.files(), ["a.jpg", "b.jpg"])


class SomeNotUploaded(_Queue, unittest.TestCase):

    def test_the_user_is_asked_and_told_which(self):
        paths = self.fill("a.jpg", "b.jpg", "c.jpg")
        self.uploaded(paths[0], paths[2])
        self.down(2)
        self.down()
        self.assertEqual(len(self.asked), 1)
        self.assertIn("1 photo(s) have not been uploaded", self.asked[0])
        self.assertIn("b.jpg", self.asked[0])

    def test_yes_goes_to_the_first_not_uploaded(self):
        paths = self.fill("a.jpg", "b.jpg", "c.jpg", "d.jpg")
        self.uploaded(paths[0], paths[3])
        self.answer = True
        self.down(4)
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(self.app.photo_label_var.get(), "b.jpg")
        self.assertEqual(self.app.input_list.curselection(), (1,))
        self.assertEqual(len(self.app.input_paths), 4)

    def test_no_clears_the_queue_anyway(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.uploaded(paths[0])
        self.answer = False
        self.down(2)
        self.assertEqual(self.app.input_paths, [])
        self.assertEqual(self.app.input_list.size(), 0)
        self.assertIsNone(self.app.current_photo)
        self.assertIn("1 photo(s) were not uploaded", self.app.status_var.get())
        self.assertEqual(self.files(), ["a.jpg", "b.jpg"])

    def test_cancel_stays_on_the_last(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.answer = None
        self.down(2)
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(self.app.input_list.curselection(), (1,))
        self.assertEqual(len(self.app.input_paths), 2)


class SkipOnTheLast(_Queue, unittest.TestCase):
    """Skip on the last photo does what Down does."""

    def skip(self):
        self.app._skip_photo()
        self.root.update()

    def test_all_uploaded_clears_the_queue(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.uploaded(*paths)
        self.skip()
        self.skip()
        self.assertEqual(self.app.input_paths, [])
        self.assertEqual(self.app.status_var.get(), "All 2 photo(s) uploaded — queue cleared.")

    def test_yes_goes_to_the_first_not_uploaded(self):
        paths = self.fill("a.jpg", "b.jpg", "c.jpg")
        self.uploaded(paths[0], paths[2])
        self.answer = True
        self.skip()
        self.skip()
        self.skip()
        self.assertEqual(len(self.asked), 1)
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(self.app.photo_label_var.get(), "b.jpg")

    def test_no_clears_the_queue_anyway(self):
        self.fill("a.jpg", "b.jpg")
        self.answer = False
        self.skip()
        self.skip()
        self.assertEqual(self.app.input_paths, [])
        self.assertIn("2 photo(s) were not uploaded", self.app.status_var.get())

    def test_cancel_stays_on_the_last(self):
        paths = self.fill("a.jpg", "b.jpg")
        self.answer = None
        self.skip()
        self.skip()
        self.assertEqual(self.app.current_photo, paths[1])
        self.assertEqual(len(self.app.input_paths), 2)


class AfterAnUploadStill(_Queue, unittest.TestCase):
    """The clearing after the last upload now shares _clear_queue."""

    def test_uploading_the_last_one_clears_the_queue(self):
        self.fill("a.jpg")
        self.app.auto_rename_var.set("")
        self.app.custom_vars['date_of_photo'].set("1962-09-01")
        self.app.custom_vars['comments'].insert('1.0', "A caption")
        said = self.upload(base.accept)
        self.assertEqual(said, "All 1 photo(s) uploaded — queue cleared.")
        self.assertEqual(self.app.input_paths, [])


if __name__ == "__main__":
    unittest.main()
