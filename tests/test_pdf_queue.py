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
        self.app._people_cache = [pu.PdfPhotos.PeopleNames.load(
            _slides.make_people(self.work / "Peoples names.txt"))]
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

    def test_a_long_pdf_name_is_cut_in_the_photos_names(self):
        long = self.work / "Celebrating Friendship slide show_Sullivan-Virzi_LAcon V.pdf"
        shutil.copy(self.pdf, long)
        self.open_pdf(long)
        self.assertEqual(os.path.basename(self.app.input_paths[0]), "Celebrating Friendship sl p002.jpg")

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


class KeepBoxes(_Pdf, unittest.TestCase):
    """The Keep boxes for source and date mean nothing to a photo from a PDF,
    so they are greyed out while one is shown -- and back for other photos."""

    def greyed(self):
        return {k: self.app._persist_checks[k].instate(["disabled"])
                for k in ("photo_source", "date_of_photo", "tags", "comments")}

    def test_greyed_for_a_photo_from_a_pdf(self):
        self.open_pdf()
        self.show(0)
        self.assertEqual(self.greyed(), {"photo_source": True, "date_of_photo": True,
                                         "tags": False, "comments": False})

    def test_back_for_an_ordinary_photo(self):
        self.open_pdf()
        path = str(self.photos / "ordinary.jpg")
        base.Image.new("RGB", (60, 40), "#3080c0").save(path)
        self.app._add_single_image(path, {})
        self.show(0)
        self.show(len(self.app.input_paths) - 1)
        self.assertFalse(any(self.greyed().values()), self.greyed())

    def test_ticked_stays_ticked(self):
        self.app.persist_vars['date_of_photo'].set(True)
        self.open_pdf()
        self.show(0)
        self.assertTrue(self.app.persist_vars['date_of_photo'].get())

    def test_back_when_the_queue_is_cleared(self):
        self.open_pdf()
        self.show(0)
        self.app._clear_queue("")
        self.assertFalse(any(self.greyed().values()), self.greyed())


class NoConventionList(_Pdf, unittest.TestCase):
    """Without the convention list, opening a PDF says so first."""

    def setUp(self):
        super().setUp()
        self.app._con_list_cache = []
        pu._store.set_op_param("con_timeline_file", str(self.work / "not there.txt"))
        self.warned = []
        self.go_on = True
        pu.messagebox.askokcancel = lambda title, msg, **k: (self.warned.append((title, msg)),
                                                             self.go_on)[1]

    def test_it_warns_and_says_where_it_looked(self):
        self.open_pdf()
        self.assertEqual(len(self.warned), 1)
        self.assertEqual(self.warned[0][0], "Convention List Not Found")
        self.assertIn("not there.txt", self.warned[0][1])

    def test_cancel_opens_nothing(self):
        self.go_on = False
        self.assertEqual(self.open_pdf(), 0)
        self.assertEqual(self.app.input_paths, [])

    def test_ok_opens_it_without_cons(self):
        self.assertEqual(self.open_pdf(), 7)
        self.show(0)
        self.assertEqual((self.fields()["tags"], self.fields()["date"]), ("", "April 1987"))

    def test_found_the_next_time_once_it_is_there(self):
        self.open_pdf()
        self.app._clear_queue("")
        _slides.make_timeline(self.work / "not there.txt")
        self.open_pdf()
        self.assertEqual(len(self.warned), 1)

    def test_no_warning_when_it_is_found(self):
        _slides.make_timeline(self.work / "not there.txt")
        self.open_pdf()
        self.assertEqual(self.warned, [])


class Progress(_Pdf, unittest.TestCase):
    """A big PDF shows its progress while it is read; a small one only the
    wait cursor.  The test PDF is small, so what counts as big is lowered."""

    def setUp(self):
        super().setUp()
        self.seen = []           # (stage text, bar value, bar maximum) while working
        real_read, real_save = pu.PdfPhotos.read_pdf, pu.PdfPhotos.save_photos
        self.saved_fns = real_read, real_save

        def watch(real):
            def call(*a, **k):
                inner = k.get("progress")

                def progress(n, of):
                    inner(n, of)
                    dlg = self.dialog()
                    if dlg:
                        stage = [w for w in dlg.winfo_children() if isinstance(w, pu.ttk.Label)][1]
                        bar = [w for w in dlg.winfo_children() if isinstance(w, pu.ttk.Progressbar)][0]
                        self.seen.append((stage.cget("text"), int(float(bar.cget("value"))),
                                          int(float(bar.cget("maximum")))))
                k["progress"] = progress
                return real(*a, **k)
            return call
        pu.PdfPhotos.read_pdf, pu.PdfPhotos.save_photos = watch(real_read), watch(real_save)

    def tearDown(self):
        pu.PdfPhotos.read_pdf, pu.PdfPhotos.save_photos = self.saved_fns
        super().tearDown()

    def dialog(self):
        for w in self.root.winfo_children():
            if isinstance(w, pu.tk.Toplevel) and w.winfo_exists() and w.title() == "Opening PDF…":
                return w
        return None

    def test_a_big_pdf_shows_it(self):
        self.app._PDF_PROGRESS_BYTES = 0
        self.open_pdf()
        self.assertIn(("Reading slides: 3 of 8", 3, 8), self.seen)
        self.assertIn(("Taking out photos: 7 of 7", 7, 7), self.seen)

    def test_and_takes_it_down(self):
        self.app._PDF_PROGRESS_BYTES = 0
        self.open_pdf()
        self.assertIsNone(self.dialog())

    def test_also_when_all_are_done_and_it_starts_over(self):
        self.app._PDF_PROGRESS_BYTES = 0
        self.app.state_data["pdf_progress"] = {"Slide Show.pdf": list(ALL)}
        self.assertEqual(self.open_pdf(), 7)
        self.assertIn(("Taking out photos: 7 of 7", 7, 7), self.seen)
        self.assertIsNone(self.dialog())

    def test_a_small_pdf_does_not(self):
        self.open_pdf()
        self.assertEqual(self.seen, [])


class LeftToRight(_Pdf, unittest.TestCase):
    """"L-R: ??" on the caption of a photo of people that names none of them.
    The test slides are flat colour, so whether a photo has faces is set."""

    def setUp(self):
        super().setUp()
        self.faces = True
        self.looked = []
        pu.PdfPhotos.has_faces, self.saved_has_faces = (
            lambda path, model: (self.looked.append(path), self.faces)[1], pu.PdfPhotos.has_faces)

    def tearDown(self):
        pu.PdfPhotos.has_faces = self.saved_has_faces
        super().tearDown()

    def caption_of(self, key):
        self.show(ALL.index(key))
        return self.fields()["caption"]

    def test_people_but_no_names(self):
        self.open_pdf()
        self.assertEqual(self.caption_of("p004"), "Not-A-D-Con outing L-R: ??")

    def test_no_people(self):
        self.faces = False
        self.open_pdf()
        self.assertEqual(self.caption_of("p004"), "Not-A-D-Con outing")

    def test_names_already(self):
        self.open_pdf()
        self.assertEqual(self.caption_of("p002"), "Bob Tucker & Pat Virzi at Corflu IV — Covington, KY")
        self.assertEqual(self.looked, [], "no need to look for faces")

    def test_once_a_photo(self):
        self.open_pdf()
        self.caption_of("p004")
        self.app.custom_vars['comments'].delete('1.0', 'end')
        self.app.custom_vars['comments'].insert('1.0', "Not-A-D-Con outing L-R: Joe, Rob")
        self.show(0)
        self.assertEqual(self.caption_of("p004"), "Not-A-D-Con outing L-R: Joe, Rob")
        self.assertEqual(len(self.looked), 1)

    def test_the_real_detector_and_model(self):
        """Wired to the model beside the program: a flat colour has no face."""
        pu.PdfPhotos.has_faces = self.saved_has_faces
        self.open_pdf()
        self.assertEqual(self.caption_of("p004"), "Not-A-D-Con outing")
        self.assertIs(self.app._pdf_items[self.app.current_photo]["faces"], False)


class NoPeopleList(_Pdf, unittest.TestCase):
    """Without the list of people's names, opening a PDF says so first; then
    no caption gets "L-R: ??"."""

    def setUp(self):
        super().setUp()
        self.app._people_cache = []
        (self.work / "Peoples names.txt").unlink()
        pu._store.set_op_param("con_timeline_file", str(self.work / "timeline.txt"))
        self.warned = []
        self.go_on = True
        pu.messagebox.askokcancel = lambda title, msg, **k: (self.warned.append((title, msg)),
                                                             self.go_on)[1]
        pu.PdfPhotos.has_faces, self.saved_has_faces = (lambda path, model: True,
                                                        pu.PdfPhotos.has_faces)

    def tearDown(self):
        pu.PdfPhotos.has_faces = self.saved_has_faces
        super().tearDown()

    def test_it_warns_and_says_where_it_looked(self):
        self.open_pdf()
        self.assertEqual([t for t, _m in self.warned], ["People's Names Not Found"])
        self.assertIn(str(self.work / "Peoples names.txt"), self.warned[0][1])

    def test_cancel_opens_nothing(self):
        self.go_on = False
        self.assertEqual(self.open_pdf(), 0)

    def test_ok_opens_it_without(self):
        self.assertEqual(self.open_pdf(), 7)
        self.show(ALL.index("p004"))
        self.assertEqual(self.fields()["caption"], "Not-A-D-Con outing")

    def test_found_beside_the_con_list(self):
        _slides.make_people(self.work / "Peoples names.txt")
        self.open_pdf()
        self.assertEqual(self.warned, [])


class NothingKept(_Pdf, unittest.TestCase):

    def test_the_photos_go_when_the_app_closes(self):
        self.open_pdf()
        folder = self.app._pdf_temp_root
        self.assertTrue(os.path.isdir(folder))
        self.app._on_close()
        self.assertFalse(os.path.exists(folder))


if __name__ == "__main__":
    unittest.main()
