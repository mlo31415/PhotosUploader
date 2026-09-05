"""Deciding what a temp-named photo should be called again.

The repair tool restores a photo's `file` from its title.  Which titles can be
restored from, and what the restored name is, is the whole of the judgement --
and it is a pure function, so it is tested here without a server.

The cases are taken from the 2026-09-05 survey of photos.fanac.org: 1,481
photos stored under a tempfile name, 1,420 of them repairable.
"""
import importlib.util
import sys
import unittest
from pathlib import Path

_TOOL = Path(__file__).resolve().parent.parent / "tools" / "repair_tempfile_names.py"
sys.path.insert(0, str(_TOOL.parent.parent.parent / "PiwigoHelpers"))


def load_tool():
    spec = importlib.util.spec_from_file_location("repair_under_test", _TOOL)
    module = importlib.util.module_from_spec(spec)
    module.__spec__ = spec
    spec.loader.exec_module(module)
    return module


tool = load_tool()


def row(file, name):
    return {"id": 1, "file": file, "name": name, "album": "somewhere"}


class NamesThatCanBeRestored(unittest.TestCase):

    def test_the_real_cases_from_the_survey(self):
        for file, title in (("tmp3figxcg9.jpg", "es-eb00102.jpg"),
                            ("tmp_9yj5dgg.jpg", "es-p6-03.jpg"),
                            ("tmprzsvndii.jpg", "es-eb00118.jpg"),
                            ("tmp86i1nw3b.JPG", "mlo00195.jpg"),
                            ("tmp12cehr09.JPG", "PICT0943.JPG"),
                            ("tmp8xzn0oi5.JPG", "PICT8625.JPG"),
                            ("tmp_wth5jir.jpeg", "x08h001.jpeg"),
                            ("tmpsafuyfy6.JPG", "PICT1057.JPG")):
            with self.subTest(file):
                self.assertEqual(tool.repaired_name(row(file, title)), title)

    def test_a_title_with_no_extension_takes_the_current_one(self):
        """Piwigo's title is often the name without its extension."""
        self.assertEqual(tool.repaired_name(row("tmp86i1nw3b.JPG", "mlo00195")),
                         "mlo00195.JPG")

    def test_surrounding_space_is_dropped(self):
        self.assertEqual(tool.repaired_name(row("tmp12cehr09.JPG", "  PICT0943.JPG  ")),
                         "PICT0943.JPG")


class NamesThatCannotBe(unittest.TestCase):

    def test_a_title_that_is_itself_a_tempfile_name(self):
        """61 photos are like this: uploaded twice through the broken path, so
        the second upload took the first's temp name as its title.  Piwigo no
        longer holds the real name, and guessing would be worse than leaving
        it."""
        self.assertIsNone(tool.repaired_name(
            row("tmpzpkh0v71.JPG", "tmpatga9z9w.JPG")))

    def test_an_empty_or_blank_title(self):
        self.assertIsNone(tool.repaired_name(row("tmp12cehr09.JPG", "")))
        self.assertIsNone(tool.repaired_name(row("tmp12cehr09.JPG", "   ")))

    def test_a_title_that_is_not_a_usable_file_name(self):
        for bad in ('a/b.jpg', 'a:b.jpg', 'a?b.jpg', 'a|b.jpg'):
            with self.subTest(bad):
                self.assertIsNone(tool.repaired_name(row("tmp12cehr09.JPG", bad)))


class SpottingATempName(unittest.TestCase):
    """The same pattern decides who is in the list at all."""

    def test_the_names_really_seen_on_the_server(self):
        for name in ("tmp2qfmcfwp.JPG", "tmp3figxcg9.jpg", "tmp_9yj5dgg.jpg",
                     "tmp_wth5jir.jpeg", "tmpq8_1zx2o.JPG"):
            with self.subTest(name):
                self.assertRegex(name, tool.TEMPFILE)

    def test_ordinary_names_are_not_matched(self):
        for name in ("es-b00273.jpg", "PICT0943.JPG", "x15-002.jpeg",
                     "temperature.jpg", "tmp.jpg", "tmp123.jpg",
                     "tmpanything-much-longer.jpg"):
            with self.subTest(name):
                self.assertNotRegex(name, tool.TEMPFILE)


class TheToolItself(unittest.TestCase):

    def test_plan_is_what_it_does_when_told_nothing(self):
        """So that a stray run reports rather than writes."""
        source = _TOOL.read_text(encoding="utf-8")
        self.assertIn('mode = sys.argv[1] if len(sys.argv) > 1 else "plan"', source)

    def test_only_apply_writes(self):
        """setInfo is called from one place, and it is not survey or plan."""
        import ast
        tree = ast.parse(_TOOL.read_text(encoding="utf-8"))
        writers = [n.name for n in tree.body
                   if isinstance(n, ast.FunctionDef)
                   and "pwg.images.setInfo" in ast.dump(n)]
        self.assertEqual(writers, ["apply"])


if __name__ == "__main__":
    unittest.main()
