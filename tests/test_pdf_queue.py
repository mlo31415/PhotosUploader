"""Photos from a slide-show PDF, in the queue.

Opening a PDF queues its photos, named for the PDF and the slide.  Each shows
the caption, date, source and con the slide gives, and is edited and uploaded
like any photo.  What has been uploaded or skipped is remembered, by the
PDF's name, in the state file, so opening it again carries on where it left
off.  The photos themselves are not kept.

Built on the real app, as in test_failed_upload_undo, with every dialog
stubbed and the state in a temp folder.  The PDF is the made-up one of
_slides.py.

Run with:  python -m unittest discover -s tests
"""
import json
import os
import shutil
import unittest

import test_failed_upload_undo as base
import _slides

pu = base.pu
ALL = ["p002", "p003-1", "p003-2", "p004", "p006", "p007", "p008"]


class _Pdf(base._Fixture):

    def setUp(self):
        super().setUp()
        self.pdf = _slides.make_slideshow(self.work / "Slide Show.pdf")
        cons = pu.PdfPhotos.ConList.load(_slides.make_timeline(self.work / "timeline.txt"))
        self.app._con_list_cache = [cons]
        self.asked = []
        self.answer = True
        pu.messagebox.askyesno = self.ask
        # Moving onto an uploaded photo fetches Piwigo's copy; the local one will do
        self.app._download_and_load_uploaded = self.app._load_photo

    def tearDown(self):
        if self.app._pdf_temp_root:
            shutil.rmtree(self.app._pdf_temp_root, ignore_errors=True)
        super().tearDown()

    def ask(self, title, msg, **k):
        self.asked.append(title)
        return self.answer

    def open_pdf(self, pdf=None):
        added = self.app._add_pdf(str(pdf or self.pdf), {})
        self.app._update_counts()
        self.root.update()
        return added

    def keys(self):
        return [self.app._pdf_items[p]["key"] for p in self.app.input_paths]

    def show(self, idx):
        self.app._go_to_queue_index(idx)
        self.root.update()

    def fields(self):
        v = self.app.custom_vars
        return {"source": v['photo_source'].get(), "date": v['date_of_photo'].get(),
                "caption": v['comments'].get('1.0', 'end').strip(), "tags": v['tags'].get(),
                "name": v['output_filename'].get()}

    def done_on_disk(self, name="Slide Show.pdf"):
        state = json.loads((self.work / "PhotosUploader State.json").read_text(encoding="utf-8"))
        return state.get("pdf_progress", {}).get(name, [])


class Opening(_Pdf, unittest.TestCase):

    def test_every_photo_is_queued(self):
        self.assertEqual(self.open_pdf(), 7)
        self.assertEqual(self.keys(), ALL)

    def test_named_for_the_pdf_and_the_slide(self):
        self.open_pdf()
        self.assertEqual([os.path.basename(p) for p in self.app.input_paths][:3],
                         ["Slide Show p002.jpg", "Slide Show p003-1.jpg", "Slide Show p003-2.jpg"])

    def test_opening_it_twice_does_not_queue_it_twice(self):
        self.open_pdf()
        self.assertEqual(self.open_pdf(), 0)
        self.assertEqual(len(self.app.input_paths), 7)

    def test_all_are_waiting_to_be_uploaded(self):
        self.open_pdf()
        self.assertEqual(set(self.app.input_paths), self.app._dirty_paths)


class Fields(_Pdf, unittest.TestCase):

    def test_the_slide_fills_them(self):
        self.open_pdf()
        self.show(0)
        self.assertEqual(self.fields(), {
            "source": "Colin Hinz", "date": "April 4, 1987",
            "caption": "Bob Tucker & Pat Virzi at Corflu IV — Covington, KY",
            "tags": "Corflu 4", "name": "Slide Show p002.jpg"})

    def test_they_are_valid_as_they_stand(self):
        self.open_pdf()
        self.show(0)
        self.assertTrue(all(self.app._field_validity.values()), self.app._field_validity)

    def test_no_date_on_the_slide_is_no_date_not_the_files(self):
        self.open_pdf()
        self.show(ALL.index("p008"))
        self.assertEqual(self.fields()["date"], "")

    def test_what_the_user_typed_is_kept(self):
        self.open_pdf()
        self.show(0)
        self.app.custom_vars['comments'].insert('end', " (with Bob Webber)")
        self.app._nav_next()
        self.app._nav_prev()
        self.root.update()
        self.assertEqual(self.fields()["caption"],
                         "Bob Tucker & Pat Virzi at Corflu IV — Covington, KY (with Bob Webber)")

    def test_a_kept_source_gives_way_to_the_slides(self):
        """Even where the slide gives none: a kept value is whatever the
        field holds, so it would be the last slide's photographer."""
        self.app.persist_vars['photo_source'].set(True)
        self.app.custom_vars['photo_source'].set("Joe Siclari")
        self.open_pdf()
        self.show(0)
        self.assertEqual(self.fields()["source"], "Colin Hinz")
        self.show(ALL.index("p008"))                     # no credit line
        self.assertEqual(self.fields()["source"], "")

    def test_not_auto_renamed(self):
        self.app.auto_rename_var.set("es-jk")
        self.open_pdf()
        self.show(0)
        self.assertEqual(self.fields()["name"], "Slide Show p002.jpg")


class Remembering(_Pdf, unittest.TestCase):

    def test_skip_is_remembered(self):
        self.open_pdf()
        self.show(0)
        self.app._skip_photo()
        self.assertEqual(self.done_on_disk(), ["p002"])

    def test_upload_is_remembered(self):
        self.open_pdf()
        self.show(1)
        self.upload(base.accept)
        self.assertEqual(self.done_on_disk(), ["p003-1"])

    def test_uploaded_under_its_own_name_with_auto_rename_on(self):
        self.app.auto_rename_var.set("es-jk")
        self.open_pdf()
        self.show(0)
        self.upload(base.accept)
        self.assertEqual(base.StandInClient.uploads, ["Slide Show p002.jpg"])
        self.assertEqual(os.path.basename(self.app.input_paths[0]), "Slide Show p002.jpg")

    def test_moving_past_a_photo_is_not_skipping_it(self):
        self.open_pdf()
        self.show(0)
        self.app._nav_next()
        self.assertEqual(self.done_on_disk() if (self.work / "PhotosUploader State.json").exists()
                         else [], [])

    def test_opening_again_leaves_out_what_is_done(self):
        self.open_pdf()
        self.show(0)
        self.app._skip_photo()
        self.upload(base.accept)                         # p003-1, which Skip moved to
        self.app._clear_queue("")
        self.assertEqual(self.open_pdf(), 5)
        self.assertEqual(self.keys(), ["p003-2", "p004", "p006", "p007", "p008"])

    def test_by_the_pdfs_name(self):
        self.open_pdf()
        self.show(0)
        self.app._skip_photo()
        self.app._clear_queue("")
        moved = self.work / "elsewhere"
        moved.mkdir()
        shutil.copy(self.pdf, moved / "Slide Show.pdf")
        shutil.copy(self.pdf, moved / "Another Show.pdf")
        self.open_pdf(moved / "Slide Show.pdf")
        self.assertNotIn("p002", self.keys())
        self.app._clear_queue("")
        self.open_pdf(moved / "Another Show.pdf")
        self.assertIn("p002", self.keys())


class AllDone(_Pdf, unittest.TestCase):

    def finish_it(self):
        self.app.state_data["pdf_progress"] = {"Slide Show.pdf": list(ALL)}

    def test_it_asks_to_start_over(self):
        self.finish_it()
        self.answer = False
        self.assertEqual(self.open_pdf(), 0)
        self.assertEqual(self.asked, ["All Photos Done"])
        self.assertEqual(self.app.input_paths, [])

    def test_yes_starts_over(self):
        self.finish_it()
        self.answer = True
        self.assertEqual(self.open_pdf(), 7)
        self.assertEqual(self.done_on_disk(), [])


class NothingKept(_Pdf, unittest.TestCase):

    def test_the_photos_go_when_the_app_closes(self):
        self.open_pdf()
        folder = self.app._pdf_temp_root
        self.assertTrue(os.path.isdir(folder))
        self.app._on_close()
        self.assertFalse(os.path.exists(folder))


if __name__ == "__main__":
    unittest.main()
