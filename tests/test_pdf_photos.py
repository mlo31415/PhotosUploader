"""Reading photos, captions, dates, sources and cons out of a slide-show PDF.

The credit lines tested are real ones, from the slide show this was written
for ("Celebrating Friendship", Geri Sullivan & Pat Virzi, LAcon V), given as
the lines they are set in on the slide.  The PDF is a made-up one in the same
form: see _slides.py.

Run with:  python -m unittest discover -s tests
"""
import shutil
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import PdfPhotos                                                # noqa: E402
import _slides                                                  # noqa: E402

credit = PdfPhotos.parse_credit


class CreditLines(unittest.TestCase):

    def test_photographer_on_the_next_line(self):
        c = credit(["1983 / photo by ", "David Dyer-Bennet  ", "St. Paul, MN /GFS "])
        self.assertEqual((c.source, c.date, c.place), ("David Dyer-Bennet", "1983", "St. Paul, MN"))

    def test_photo_and_by_on_different_lines(self):
        c = credit(["December 2003 / photo ", "by David Dyer-Bennet ", "Minneapolis, MN / GFS"])
        self.assertEqual((c.source, c.date, c.place),
                         ("David Dyer-Bennet", "December 2003", "Minneapolis, MN"))

    def test_a_name_cut_after_its_first_word(self):
        c = credit(["May 2016 / photo by Gary ", "Mattingly / Chicago, IL / PV"])
        self.assertEqual((c.source, c.place), ("Gary Mattingly", "Chicago, IL"))

    def test_a_name_ending_a_line_is_not_given_the_place(self):
        c = credit(["March 2020 / Corflu 37 ", "photo by Gary Mattingly ", "College Station, TX / PV"])
        self.assertEqual((c.source, c.place), ("Gary Mattingly", "Corflu 37 College Station, TX"))

    def test_selfies(self):
        self.assertEqual(credit(["August 2019 / group selfie by  ",
                                 "James Bacon / Dublin, Ireland / GFS"]).source, "James Bacon")
        self.assertEqual(credit(["August 2020 / selfie  ", "by Gavi Levy Haskell ",
                                 "Wales, MA / GFS"]).source, "Gavi Levy Haskell")

    def test_initials_alone(self):
        c = credit(["March 1984 ", "Minneapolis, MN / GFS"])
        self.assertEqual((c.source, c.place), ("GFS", "Minneapolis, MN"))

    def test_two_sets_of_initials(self):
        self.assertEqual(credit(["May 2018 / Corflu 35 ", "Toronto ON Canada  ", "/ GFS & PV"]).source,
                         "GFS & PV")

    def test_initials_without_a_slash(self):
        self.assertEqual(credit(["November 2021 ", "Tibbett Ave. Bronx NY GFS"]).source, "GFS")

    def test_a_state_is_not_initials(self):
        c = credit(["February 2023 ", "Ruidoso, NM "])
        self.assertEqual((c.source, c.place), ("", "Ruidoso, NM"))

    def test_a_named_collection(self):
        c = credit(["May 1985 / From the ", "Lien–Garey photo ", "collection / GFS"])
        self.assertEqual((c.source, c.place), ("Lien–Garey photo collection", ""))

    def test_a_decade_is_its_middle_year_and_its_words_are_kept(self):
        c = credit(["1980s / PV"])
        self.assertEqual((c.date, c.vague), ("1985", "1980s"))

    def test_something(self):
        c = credit(["1990–something  ", "Fairview, TX / PV"])
        self.assertEqual((c.date, c.vague, c.place), ("1995", "1990–something", "Fairview, TX"))
        self.assertEqual(credit(["1980-something / PV"]).date, "1985")

    def test_early_and_late(self):
        self.assertEqual(credit(["Early 2000s / PV"]).date, "2002")
        self.assertEqual(credit(["late 1990s / PV"]).date, "1998")

    def test_circa(self):
        c = credit(["circa 1983   ", "MippleStipple, MN ", "GFS"])
        self.assertEqual((c.date, c.vague, c.source, c.place),
                         ("1983", "circa 1983", "GFS", "MippleStipple, MN"))

    def test_a_full_date(self):
        self.assertEqual(credit(["July 23, 2026 / PV"]).date, "July 23, 2026")


class Names(unittest.TestCase):

    def test_first_mentions_are_written_out(self):
        self.assertEqual(PdfPhotos.expand_names("Kip Williams, Pat, & Cathy Doyle"),
                         "Kip Williams, Pat Virzi, & Cathy Doyle")

    def test_only_the_first(self):
        self.assertEqual(PdfPhotos.expand_names("Geri broke her ankle, but Geri stayed"),
                         "Geri Sullivan broke her ankle, but Geri stayed")

    def test_not_when_a_surname_follows(self):
        self.assertEqual(PdfPhotos.expand_names("Pat Cadigan & Geri"),
                         "Pat Cadigan & Geri Sullivan")

    def test_not_when_already_named_in_full(self):
        self.assertEqual(PdfPhotos.expand_names("Pat Virzi in Pat’s living room"),
                         "Pat Virzi in Pat’s living room")

    def test_possessives(self):
        self.assertEqual(PdfPhotos.expand_names("Pat’s Mom"), "Pat Virzi’s Mom")


class NamesSomeone(unittest.TestCase):
    """Whether a caption names anyone: the test for "L-R: ??"."""

    @classmethod
    def setUpClass(cls):
        cls.work = Path(tempfile.mkdtemp())
        cls.people = PdfPhotos.PeopleNames.load(_slides.make_people(cls.work / "Peoples names.txt"))

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.work, ignore_errors=True)

    def named(self, text):
        return self.people.names_someone(text)

    def test_a_first_name_and_surname(self):
        self.assertTrue(self.named("Sandi Cupp… and her shoes"))
        self.assertTrue(self.named("Liz Smith-Jones at the bar"), "a surname the list does not have")

    def test_a_middle_bit_in_brackets(self):
        self.assertTrue(self.named("Bob (pre-Levy) Tucker"))

    def test_a_first_name_alone_if_two_fans_have_it(self):
        self.assertTrue(self.named("Before long, Joe joined the group dance"))
        self.assertTrue(self.named("Road-tripping with Rob. Remember road maps?"),
                        "a qualified name counts")
        self.assertFalse(self.named("Sandi’s shoes"), "only one Sandi")

    def test_the_hosts_alone(self):
        self.assertTrue(self.named("Pat reacts"))
        self.assertTrue(self.named("Geri’s office"))

    def test_no_names(self):
        for text in ("Minicon Opening Ceremonies", "Not-A-D-Con outing to the Como Zoo",
                     "Toner group outing to the Chicago Hot Dogs stand",
                     "Fan lounge set-up at L.A.con III", "… with friends", ""):
            self.assertFalse(self.named(text), text)

    def test_ordinary_words_that_are_also_first_names(self):
        self.assertFalse(self.named("The Count of Monte Cristo"))
        self.assertFalse(self.named("during what would have been the Sunday finale"))


class Faces(unittest.TestCase):

    MODEL = Path(__file__).resolve().parent.parent / "face_detection_yunet_2023mar.onnx"

    def setUp(self):
        self.work = Path(tempfile.mkdtemp())
        self.photo = self.work / "plain.jpg"
        _slides.Image.new("RGB", (900, 600), "#80a0c0").save(self.photo)

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_none_in_a_plain_photo(self):
        self.assertIs(PdfPhotos.has_faces(self.photo, self.MODEL), False)

    def test_cannot_tell_without_the_model(self):
        self.assertIsNone(PdfPhotos.has_faces(self.photo, self.work / "no model.onnx"))

    def test_cannot_tell_without_a_photo(self):
        self.assertIsNone(PdfPhotos.has_faces(self.work / "not there.jpg", self.MODEL))


class Cons(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.work = Path(tempfile.mkdtemp())
        cls.cons = PdfPhotos.ConList.load(_slides.make_timeline(cls.work / "timeline.txt"))

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.work, ignore_errors=True)

    def find(self, caption, heading="", credit_text="", year=0, month=0):
        return self.cons.find(caption, heading, credit_text, year, month)

    def test_dates(self):
        c = self.cons.by_year[1985][0]
        self.assertEqual((c.start, c.end, c.middle()),
                         (date(1985, 8, 30), date(1985, 9, 2), date(1985, 8, 31)))

    def test_a_con_with_no_date_still_has_its_year(self):
        c = [c for c in self.cons.by_year[1987] if c.name == "Corflu 3"][0]
        self.assertIsNone(c.middle())

    def test_by_the_name_shown(self):
        self.assertEqual(self.find("Bob Webber at Corflu IV", year=1987)[0].name, "Corflu 4")

    def test_the_year_decides(self):
        self.assertIsNone(self.find("Bob Webber at Corflu IV", year=1988)[0])

    def test_corflu_3_is_not_inside_corflu_35(self):
        self.assertEqual(self.find("at Corflu 35", year=2018)[0].name, "Corflu 35")
        self.assertIsNone(self.find("at Corflu 35", year=1987)[0])

    def test_in_the_credit_line(self):
        self.assertEqual(self.find("Socks at the Dead Dog", credit_text="Corflu 35 Toronto",
                                   year=2018)[0].name, "Corflu 35")

    def test_a_series_in_a_heading(self):
        self.assertEqual(self.find("Al & Nate", heading="MINICON", year=1997, month=3)[0].name,
                         "Minicon 32")

    def test_a_series_where_the_photo_was_taken(self):
        self.assertEqual(self.find("Andy Hooper at Corflu Quire", year=2020)[0].name, "Corflu 37")

    def test_not_a_series_named_in_passing(self):
        self.assertIsNone(self.find("Pat reviewing a copy of The Spirit of Corflu", year=2020)[0])

    def test_a_named_con_in_another_month_is_only_a_note(self):
        con, note = self.find("Boskone 55 was Geri’s first Boskone", year=2018, month=3)
        self.assertIsNone(con)
        self.assertIn("Boskone 55", note)


class ThePdf(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.work = Path(tempfile.mkdtemp())
        cls.pdf = _slides.make_slideshow(cls.work / "Slide Show.pdf")
        cls.cons = PdfPhotos.ConList.load(_slides.make_timeline(cls.work / "timeline.txt"))
        cls.photos = {p.key: p for p in PdfPhotos.read_pdf(cls.pdf, cls.cons)}

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.work, ignore_errors=True)

    def test_which_photos(self):
        """Not the title slide, the shadows, the background or the cut-up
        graphic; the photo placed twice once."""
        self.assertEqual(sorted(self.photos), ["p002", "p003-1", "p003-2", "p004", "p006", "p007", "p008"])

    def test_one_photo(self):
        p = self.photos["p002"]
        self.assertEqual(p.caption, "Bob Tucker & Pat Virzi at Corflu IV — Covington, KY")
        self.assertEqual((p.source, p.con, p.tags), ("Colin Hinz", "Corflu 4", "Corflu 4"))
        self.assertEqual(p.date, "April 4, 1987", "the middle of the con")

    def test_two_photos_share_a_caption(self):
        for key in ("p003-1", "p003-2"):
            p = self.photos[key]
            self.assertEqual((p.caption, p.date, p.source), ("Pat Virzi and the Thames — London, UK", "1987", "PV"))

    def test_hidden_text_is_left_out(self):
        p = self.photos["p004"]
        self.assertEqual(p.caption, "Not-A-D-Con outing")
        self.assertEqual((p.source, p.date), ("Lien-Garey photo collection", "May 1985"))

    def test_a_vague_date(self):
        p = self.photos["p006"]
        self.assertEqual((p.caption, p.date), ("Liz Danforth (1980s)", "1985"))

    def test_the_con_from_a_heading_goes_into_the_caption(self):
        p = self.photos["p007"]
        self.assertEqual(p.con, "Minicon 32")
        self.assertEqual(p.caption, "Al Standish & Nate Bucklin — Bloomington, MN (Minicon 32)")
        self.assertEqual(p.date, "March 29, 1997")

    def test_no_credit_line(self):
        p = self.photos["p008"]
        self.assertEqual((p.caption, p.date, p.source), ("Thank you", "", ""))

    def test_without_a_con_list(self):
        p = {p.key: p for p in PdfPhotos.read_pdf(self.pdf)}["p007"]
        self.assertEqual((p.con, p.date), ("", "March 1997"))

    def test_whether_each_caption_names_anyone(self):
        """From the slide's caption, not the place or the con added to it."""
        people = PdfPhotos.PeopleNames.load(_slides.make_people(self.work / "Peoples names.txt"))
        named = {p.key: p.named for p in PdfPhotos.read_pdf(self.pdf, self.cons, people)}
        self.assertEqual(named, {"p002": True, "p003-1": True, "p003-2": True, "p004": False,
                                 "p006": True, "p007": True, "p008": False})

    def test_not_known_without_the_list(self):
        self.assertTrue(all(p.named is None for p in self.photos.values()))

    def test_saved_as_they_are_in_the_pdf(self):
        out = self.work / "out"
        out.mkdir()
        photos = list(self.photos.values())
        paths = PdfPhotos.save_photos(self.pdf, photos, out)
        for p, path in zip(photos, paths):
            self.assertEqual(Path(path).name, f"Slide Show {p.key}.jpg")
            self.assertEqual(Path(path).read_bytes(), _slides.PHOTOS[p.key], p.key)

    def test_which_colour_profiles_are_srgb(self):
        """The profile names in the real slide show, and PyMuPDF's own."""
        for name in ("sRGB IEC61966-2.1", "sRGB", "sRGB IEC61966-2-1 black scaled",
                     "Artifex Software sRGB ICC Profile"):
            self.assertTrue(PdfPhotos._is_srgb(name), name)
        for name in ("Display P3", "DCI-P3 D65 Gamut with sRGB Transfer", "Adobe RGB (1998)",
                     "Apple Wide Color Sharing Profile", "uRGB"):
            self.assertFalse(PdfPhotos._is_srgb(name), name)

    def test_a_file_already_there_is_kept(self):
        out = self.work / "kept"
        out.mkdir()
        (out / "Slide Show p002.jpg").write_bytes(b"edited")
        PdfPhotos.save_photos(self.pdf, [self.photos["p002"]], out)
        self.assertEqual((out / "Slide Show p002.jpg").read_bytes(), b"edited")


if __name__ == "__main__":
    unittest.main()
