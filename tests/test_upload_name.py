"""The copy sent to Piwigo carries the photo's real name.

Piwigo stores the name of the file it is given.  PhotosUploader used to send a
NamedTemporaryFile, so the photo landed on the server as "tmp2qfmcfwp.jpg" --
which is how six of the twenty-nine photos in the SlideShow log got their
names, their titles keeping the real one.

Only _prepare_upload_copy is exercised, and it writes to a temp folder, so
nothing here goes near the network or the real files.

Run with:  python -m unittest discover -s tests
"""
import importlib.util
import os
import re
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_APP = _HERE.parent / "PhotosUploader.py"
sys.path.insert(0, str(_APP.parent.parent / "PiwigoHelpers"))
sys.path.insert(0, str(_APP.parent))


def load_pu():
    spec = importlib.util.spec_from_file_location("pu_under_test", _APP)
    module = importlib.util.module_from_spec(spec)
    module.__spec__ = spec
    spec.loader.exec_module(module)
    return module


pu = load_pu()
try:
    from PIL import Image
    PIL_OK = True
except ImportError:
    PIL_OK = False

TEMPFILE_NAME = re.compile(r"^tmp[a-z0-9_]{8}\.", re.IGNORECASE)


class Bare:
    """Just enough of the app to call the one method under test."""
    _JPEG_EXTENSIONS = pu.PhotosUploader._JPEG_EXTENSIONS
    _ILLEGAL_FILENAME_CHARS = pu.PhotosUploader._ILLEGAL_FILENAME_CHARS
    _prepare_upload_copy = pu.PhotosUploader._prepare_upload_copy

    def set_status(self, msg):
        pass


class _Fixture:
    """Shared set-up.  Not a TestCase, so the classes below inherit the helpers
    without also re-running each other's tests."""

    def setUp(self):
        self.app = Bare()
        self.work = Path(tempfile.mkdtemp())
        self.made = []

    def tearDown(self):
        for d in self.made:
            shutil.rmtree(d, ignore_errors=True)
        shutil.rmtree(self.work, ignore_errors=True)

    def source(self, name, size=(60, 40), colour="#3080c0"):
        path = self.work / name
        Image.new("RGB", size, colour).save(path)
        return str(path)

    def copy(self, src, upload_name="", params=None, image=None):
        out = self.app._prepare_upload_copy(src, params or {},
                                            source_image=image,
                                            upload_name=upload_name)
        self.assertIsNotNone(out, "no copy was prepared")
        self.made.append(os.path.dirname(out))
        return out


@unittest.skipUnless(PIL_OK, "PIL is needed to write the test images")
class TheNameTheCopyCarries(_Fixture, unittest.TestCase):

    def test_a_jpeg_is_sent_under_the_output_filename(self):
        out = self.copy(self.source("mlo00195.JPG"), "mlo00195.jpg")
        self.assertEqual(os.path.basename(out), "mlo00195.jpg")

    def test_a_converted_png_is_too(self):
        out = self.copy(self.source("scan.png"), "es-eb00102.jpg")
        self.assertEqual(os.path.basename(out), "es-eb00102.jpg")

    def test_with_no_name_given_the_source_name_is_used(self):
        out = self.copy(self.source("PICT0943.JPG"))
        self.assertEqual(os.path.basename(out), "PICT0943.JPG")

    def test_never_a_tempfile_name(self):
        """The bug itself: what six photos on the server are called."""
        for name, given in (("mlo00195.JPG", "mlo00195.jpg"),
                            ("scan.png", "es-eb00102.jpg"),
                            ("PICT0943.JPG", "")):
            with self.subTest(name):
                out = self.copy(self.source(name), given)
                self.assertNotRegex(os.path.basename(out), TEMPFILE_NAME)


@unittest.skipUnless(PIL_OK, "PIL is needed to write the test images")
class NamesThatCannotBeUsed(_Fixture, unittest.TestCase):
    """The field is validated in the GUI, but the copy is made on a worker
    thread, so the name is checked again rather than trusted."""

    def test_a_name_with_no_extension_gains_one(self):
        out = self.copy(self.source("x.jpg"), "no dots here")
        self.assertEqual(os.path.basename(out), "no dots here.jpg")

    def test_an_illegal_name_falls_back_to_the_source(self):
        out = self.copy(self.source("y.jpg"), 'bad:name?.jpg')
        self.assertEqual(os.path.basename(out), "y.jpg")

    def test_a_trailing_dot_falls_back_too(self):
        out = self.copy(self.source("z.jpg"), "trailing.")
        self.assertEqual(os.path.basename(out), "z.jpg")

    def test_a_blank_name_falls_back(self):
        out = self.copy(self.source("w.jpg"), "   ")
        self.assertEqual(os.path.basename(out), "w.jpg")


@unittest.skipUnless(PIL_OK, "PIL is needed to write the test images")
class TheRestOfThePreparation(_Fixture, unittest.TestCase):
    """Giving the copy a real name must not have cost anything else."""

    def test_two_photos_of_the_same_name_do_not_collide(self):
        """Each copy has a directory to itself, which is what lets it keep a
        name that is not unique."""
        a = self.copy(self.source("one.jpg", colour="#c03030"), "same.jpg")
        b = self.copy(self.source("two.jpg", colour="#30c030"), "same.jpg")
        self.assertNotEqual(os.path.dirname(a), os.path.dirname(b))
        with Image.open(a) as ia, Image.open(b) as ib:
            self.assertNotEqual(ia.getpixel((1, 1)), ib.getpixel((1, 1)),
                                "the second copy overwrote the first")

    def test_an_oversized_photo_is_still_reduced(self):
        out = self.copy(self.source("big.jpg"), "big.jpg",
                        params={"max_upload_pixels": 400_000},
                        image=Image.new("RGB", (2000, 1500), "#802020"))
        with Image.open(out) as got:
            self.assertLessEqual(got.size[0]*got.size[1], 400_000)

    def test_the_edited_image_is_what_gets_written(self):
        """Not the file on disk: the source is blue, the edit red."""
        out = self.copy(self.source("edited.jpg", colour="#0000ff"), "edited.jpg",
                        image=Image.new("RGB", (60, 40), "#ff0000"))
        with Image.open(out) as got:
            self.assertGreater(got.getpixel((30, 20))[0], 200)

    def test_a_non_jpeg_source_comes_out_as_a_jpeg(self):
        out = self.copy(self.source("scan.png"), "scan.jpg")
        with Image.open(out) as got:
            self.assertEqual(got.format, "JPEG")

    def test_the_original_file_is_never_touched(self):
        src = self.source("keepme.jpg", size=(200, 150))
        before = os.path.getsize(src)
        self.copy(src, "keepme.jpg", params={"max_upload_pixels": 1000})
        self.assertEqual(os.path.getsize(src), before)
        with Image.open(src) as still:
            self.assertEqual(still.size, (200, 150))


if __name__ == "__main__":
    unittest.main()
