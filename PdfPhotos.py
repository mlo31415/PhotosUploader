"""
PdfPhotos.py
The photos in a slide-show PDF, with what the slides say about them.

Made for a Keynote slide show exported as a PDF: each slide a photo or two,
a caption in large type, and a credit line in small type -- "May 1985 / photo
by David Dyer-Bennet / St. Paul, MN / GFS".  For each photo it works out the
caption, the date, the source and the convention, ready for PhotosUploader's
fields.  Nothing here touches the GUI.

What a slide's text is, is told by its type size:
    caption      48 pt and up, one size throughout a line
    credit line  under 47 pt
    heading      large and small capitals mixed in one block -- a section
                 heading such as "LONESTARCON 1 1985 NASFIC"
Drop shadows are images too: each is drawn just before its photo, a little
larger all round, and is left out.
"""

import io
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import pymupdf

# ---------------------------------------------------------------------------
# Names: con names are compared in compressed form -- lower case, with . , - '
# and ’ dropped -- matched as runs of whole words, so "Corflu 3" is not found
# inside "Corflu 35".
# ---------------------------------------------------------------------------
_DROP_IN_WORDS = re.compile(r"[.\-'’]")


def _words(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", _DROP_IN_WORDS.sub("", text.lower()))


def _compress(name: str) -> str:
    return "".join(_words(name))


def _runs(words: list[str], longest: int = 6) -> set[str]:
    """Every run of up to `longest` consecutive words, run together."""
    found = set()
    for i in range(len(words)):
        s = ""
        for w in words[i:i + longest]:
            s += w
            found.add(s)
    return found


def _runs_after_at(words: list[str], longest: int = 6) -> set[str]:
    """Runs that start the text, or follow "at", "at the", "during" or
    "during the": where a caption says the photo was taken."""
    found = set()
    for i in range(len(words)):
        before = words[max(0, i - 2):i]
        if i == 0 or before[-1:] in (["at"], ["during"]) or before in (["at", "the"], ["during", "the"]):
            s = ""
            for w in words[i:i + longest]:
                s += w
                found.add(s)
    return found


_MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august",
     "september", "october", "november", "december"], 1)}
_MONTH_RE = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|Sept?(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
_MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July",
                "August", "September", "October", "November", "December"]


def _month_no(name: str) -> int:
    name = name.lower()
    for full, n in _MONTHS.items():
        if full.startswith(name[:3]):
            return n
    return 0


# ---------------------------------------------------------------------------
# Conventions, from FancyAnalyzer's "Convention timeline (Fancy).txt"
# ---------------------------------------------------------------------------
@dataclass
class Con:
    name:    str                 # the Fancyclopedia page name
    aliases: list[str]           # the names to look for
    year:    int
    start:   date | None = None
    end:     date | None = None

    def middle(self) -> date | None:
        if not self.start:
            return None
        end = self.end or self.start
        return date.fromordinal((self.start.toordinal() + end.toordinal()) // 2)

    def months(self) -> set[int]:
        if not self.start:
            return set()
        return {self.start.month, (self.end or self.start).month}


# A trailing number, Roman numeral or year: what is left is the series name
_INSTANCE_TAIL = re.compile(r"\s+(?:\d+|[IVXLC]+|'\d\d)$")
_YEAR_HEAD = re.compile(r"'''(\d{4})'''")
_LINK = re.compile(r"\[\[([^\]|#]+)(?:#[^\]|]*)?(?:\|([^\]]+))?\]\]")


def _parse_con_dates(text: str, year: int) -> tuple[date | None, date | None]:
    """'August 30-September 3, 1984', 'May 4-6, 2018', 'October 22, 1936',
    'August 2003', '2020' or ''.  Only a day gives a start; a month alone
    gives none, but the year still comes from the heading."""
    text = re.sub(r"\s+", " ", text).strip()
    m = re.fullmatch(rf"({_MONTH_RE}) (\d{{1,2}})(?:, (\d{{4}}))?(?: ?- ?(?:({_MONTH_RE}) )?(\d{{1,2}}))?, (\d{{4}})", text)
    if not m:
        return None, None
    try:
        y_end = int(m.group(6))
        y_start = int(m.group(3)) if m.group(3) else y_end
        start = date(y_start, _month_no(m.group(1)), int(m.group(2)))
        end = start
        if m.group(5):
            end_month = _month_no(m.group(4)) if m.group(4) else start.month
            end = date(y_end, end_month, int(m.group(5)))
        return start, end
    except ValueError:
        return None, None


class ConList:
    """The cons of the timeline, by year."""

    def __init__(self, cons: list[Con]):
        self.by_year: dict[int, list[Con]] = {}
        for c in cons:
            self.by_year.setdefault(c.year, []).append(c)

    @classmethod
    def load(cls, path: "str | Path") -> "ConList":
        cons, year = [], 0
        for line in Path(path).read_text(encoding="utf-8").splitlines():
            m = _YEAR_HEAD.search(line)
            if m and "colspan" in line:
                year = int(m.group(1))
                continue
            if "||" not in line or not year:
                continue
            when, rest = line.split("||", 1)
            names = []
            for lm in _LINK.finditer(rest):
                for n in (lm.group(1), lm.group(2)):
                    n = (n or "").strip()
                    if n and n not in names:
                        names.append(n)
            if not names:
                continue
            start, end = _parse_con_dates(when, year)
            cons.append(Con(names[0], names, year, start, end))
        return cls(cons)

    def find(self, caption: str, heading: str, credit: str, year: int,
             month: int = 0) -> tuple[Con | None, str]:
        """The con a photo is from, and a note for the user if one was named
        but does not fit.

        A con's full name ("Corflu 35") may be anywhere on the slide.  Its
        series name alone ("Corflu", "Chicon") counts in a section heading, or
        where the caption says the photo is *at* it ("at Corflu Quire",
        "Minicon Opening Ceremonies") -- not "The Spirit of Corflu" -- and
        only if one con of that series fits the year, and the month when the
        slide gives one."""
        cons = self.by_year.get(year, [])
        if not cons:
            return None, ""
        cap_words = _words(caption)
        all_runs = _runs(cap_words) | _runs(_words(heading)) | _runs(_words(credit))

        def month_ok(c: Con) -> bool:
            return not month or not c.start or month in c.months()

        exact = [c for c in cons
                 if any(len(_compress(a)) >= 4 and _compress(a) in all_runs for a in c.aliases)]
        if exact:
            fitting = [c for c in exact if month_ok(c)]
            if not fitting:
                c = exact[0]
                return None, (f"The slide names {c.name}, which was in "
                              f"{_MONTH_NAMES[c.start.month]} {year}; the slide says "
                              f"{_MONTH_NAMES[month]} {year}.")
            return max(fitting, key=lambda c: max(len(_compress(a)) for a in c.aliases)), ""

        head_runs = _runs(_words(heading))
        at_runs = _runs_after_at(cap_words)
        series = []
        for c in cons:
            for a in c.aliases:
                s = _INSTANCE_TAIL.sub("", a.strip())
                cs = _compress(s)
                if s != a.strip() and len(cs) >= 5 and (cs in head_runs or cs in at_runs):
                    series.append(c)
                    break
        series = [c for c in series if month_ok(c)]
        return (series[0] if len(series) == 1 else None), ""


# ---------------------------------------------------------------------------
# The credit line
# ---------------------------------------------------------------------------
# Two capitals at the end of a credit line are usually someone's initials, but
# they may be the state ("Ruidoso, NM" has no initials at all)
_PLACE_CODES = set(
    "AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO "
    "MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC "
    "AB BC MB NB NL NS NT NU ON PE QC SK YT UK USA".split())


@dataclass
class Credit:
    date:   str = ""     # for PhotosUploader's date field: "May 5, 2018", "March 1984", "1985"
    year:   int = 0
    month:  int = 0
    day:    int = 0
    vague:  str = ""     # the slide's own words when the date is a guess: "1980s", "circa 1983"
    source: str = ""     # photographer, else a named collection, else the initials
    place:  str = ""     # what is left: usually the place


_BR = " ¶ "              # a line break in the credit line, while it is being read
_BY = r"\b(?:photos?|(?:group )?selfie)(?: ¶)? by\b(?: ¶)?"


def parse_credit(lines: list[str]) -> Credit:
    """Read a slide's credit line, given as the lines it is set in."""
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in lines]
    rest = _BR.join(ln for ln in lines if ln)
    c = Credit()

    # Photographer: "photo by X", "selfie by X" -- X runs to a "/" or the end
    # of its line.  A line break may come anywhere in "photo by"; and a name
    # cut after its first word goes on with the next line's first word
    # ("photo by Gary" / "Mattingly / Chicago, IL").
    photographer = ""
    m = re.search(_BY + r"\s*([^/¶]*)", rest, re.I)
    if m:
        photographer, end = m.group(1).strip(), m.end()
        nxt = re.match(r"\s*¶\s*([A-Z][\w’'\-]+)\s*(?=/|¶|$)", rest[end:])
        if photographer and " " not in photographer and nxt:
            photographer += " " + nxt.group(1)
            end += nxt.end()
        rest = rest[:m.start()] + " " + rest[end:]
        photographer = photographer.strip(" ,;")

    collection = ""
    m = re.search(r"\bfrom the\b(.+?collection)\b", rest, re.I)
    if m:
        collection = re.sub(r"\s*¶\s*", " ", m.group(1)).strip()
        rest = rest[:m.start()] + " " + rest[m.end():]

    rest = rest.replace(_BR, " ")
    initials = ""
    m = re.search(r"(?:^|[\s/])([A-Z]{2,4}(?:\s*&\s*[A-Z]{2,4})*)\s*$", rest)
    if m and m.group(1) not in _PLACE_CODES:
        initials = re.sub(r"\s*&\s*", " & ", m.group(1))
        rest = rest[:m.start(1)]

    c.source = photographer or collection or initials

    # The date: the most precise form there is.  A decade or "1980-something"
    # becomes its middle year ("early" and "late" a little either side), and
    # the slide's own words are kept for the caption.
    m = re.search(rf"\b({_MONTH_RE})\.? (\d{{1,2}}),? ((?:19|20)\d\d)\b", rest)
    if m:
        c.month, c.day, c.year = _month_no(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(rf"\b({_MONTH_RE})\.?,? ((?:19|20)\d\d)\b", rest)
        if m:
            c.month, c.year = _month_no(m.group(1)), int(m.group(2))
        else:
            m = re.search(r"\b(?:(circa|c\.|ca\.) ?|(early|mid|late)[ \-–]?)?"
                          r"((?:19|20)\d\d)(s\b|[\-–]something\b)?", rest, re.I)
            if m:
                c.year = int(m.group(3))
                if m.group(4):
                    c.year += {"early": 2, "late": 8}.get((m.group(2) or "").lower(), 5)
                if m.group(1) or m.group(2) or m.group(4):
                    c.vague = m.group(0).strip()
    if m:
        rest = rest[:m.start()] + " " + rest[m.end():]
    c.date = _date_text(c.year, c.month, c.day)

    place = re.sub(r"(\s*/\s*)+", " / ", rest)
    c.place = re.sub(r"\s+", " ", place).strip(" /,;")
    return c


def _date_text(year: int, month: int = 0, day: int = 0) -> str:
    if not year:
        return ""
    if month and day:
        return f"{_MONTH_NAMES[month]} {day}, {year}"
    if month:
        return f"{_MONTH_NAMES[month]} {year}"
    return str(year)


def _caption_date(caption: str, year: int) -> tuple[int, int] | None:
    """A full date in the caption itself, in the credit line's year:
    "Geri & Pat at 78°13′ North -- July 23, 2026"."""
    m = re.search(rf"\b({_MONTH_RE}) (\d{{1,2}}),? ((?:19|20)\d\d)\b", caption)
    if m and int(m.group(3)) == year:
        return _month_no(m.group(1)), int(m.group(2))
    return None


# ---------------------------------------------------------------------------
# The caption
# ---------------------------------------------------------------------------
# First names alone, written out in full -- the first mention only, and not
# when a surname follows already
_FULL_NAMES = {"Pat": "Pat Virzi", "Geri": "Geri Sullivan"}


def expand_names(text: str) -> str:
    for first, full in _FULL_NAMES.items():
        if re.search(rf"\b{re.escape(full)}\b", text):
            continue                                    # already named in full
        text = re.sub(rf"\b{first}\b(?!\s+[A-Z][a-z])", full, text, count=1)
    return text


# ---------------------------------------------------------------------------
# Who is in the photo: "L-R: ??" goes on a caption that names no one, when
# the photo has faces in it
# ---------------------------------------------------------------------------
LR_PROMPT = "L-R: ??"

# Capitalised words that are some fan's first name but are much more often
# something else
_NOT_NAMES = set(
    "The A An And Mr Mrs Ms Dr Con Eve Time American Big Saint St Sir Lady "
    "May June April August Will Christmas Hope Joy "
    "Sunday Monday Tuesday Wednesday Thursday Friday Saturday".split())
_NAME_WORD = r"[^\W\d_][\w\-]+"       # two letters at least: not the L of "L.A.con"


class PeopleNames:
    """Whether a caption names anyone, from the first names in FancyAnalyzer's
    "Peoples names.txt" -- one name a line, perhaps with a "(qualifier)".

    A known first name followed by another capitalised word ("Sandi Cupp",
    "Fred (pre-Levy) Haskell") is a name; so is one standing alone ("Joe
    joined the group dance") if two or more fans have it, which leaves out
    the odd fan whose first name is an ordinary word."""

    def __init__(self, firsts: dict[str, int]):
        self.firsts = dict(firsts)
        for first in _FULL_NAMES:
            self.firsts[first] = max(self.firsts.get(first, 0), 2)

    @classmethod
    def load(cls, path: "str | Path") -> "PeopleNames":
        firsts: dict[str, int] = {}
        for line in Path(path).read_text(encoding="utf-8").splitlines():
            words = re.sub(r"\s*\([^)]*\)\s*$", "", line).split()
            if len(words) >= 2:
                firsts[words[0]] = firsts.get(words[0], 0) + 1
        return cls(firsts)

    def _first(self, word: str, at_least: int) -> bool:
        return word not in _NOT_NAMES and self.firsts.get(word, 0) >= at_least

    def names_someone(self, text: str) -> bool:
        for m in re.finditer(rf"(?<![\w\-])({_NAME_WORD})(?=\s+(?:\([^)]*\)\s+)?(\w))", text):
            if m.group(2).isupper() and self._first(m.group(1), 1):
                return True
        return any(self._first(w, 2) for w in re.findall(rf"(?<![\w\-]){_NAME_WORD}", text))


def has_faces(image_path: "str | Path", model_path: "str | Path",
              threshold: float = 0.6) -> bool | None:
    """Whether there is a face in the photo, by OpenCV's YuNet detector (the
    .onnx model SlideShow uses); None if it cannot tell -- no OpenCV, no
    model, or an image it cannot read."""
    try:
        import cv2
        import numpy as np
        from PIL import Image
    except ImportError:
        return None
    if not Path(model_path).exists():
        return None
    try:
        cv2.utils.logging.setLogLevel(cv2.utils.logging.LOG_LEVEL_ERROR)
    except AttributeError:
        pass
    try:
        with Image.open(image_path) as im:
            im = im.convert("RGB")
            # Reduced, for speed, as SlideShow does: big enough for small faces
            scale = min(1.0, 1600 / max(im.size))
            if scale < 1:
                im = im.resize((round(im.width * scale), round(im.height * scale)))
            arr = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
        detector = cv2.FaceDetectorYN.create(str(model_path), "", (arr.shape[1], arr.shape[0]),
                                             threshold)
        _, faces = detector.detect(arr)
        return faces is not None and len(faces) > 0
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Photos
# ---------------------------------------------------------------------------
@dataclass
class PdfPhoto:
    key:      str          # "p006", or "p037-2" on a slide with several
    page:     int          # 1-based
    xref:     int
    caption:  str = ""
    date:     str = ""
    source:   str = ""
    con:      str = ""     # the Fancyclopedia name, or ""
    tags:     str = ""
    notes:    list[str] = field(default_factory=list)   # for the user's attention
    named:    bool | None = None   # whether the caption names anyone; None if not known


_MIN_PIXELS = 200          # an image narrower or shorter than this is not a photo
_MAX_SLIVERS = 10          # a slide with this many tiny images is one cut-up graphic
_SHADOW_MARGIN = 40        # points: how much bigger than its photo a shadow may be


def _rect(r) -> tuple[float, float, float, float]:
    return (r.x0, r.y0, r.x1, r.y1)


def _gap(a, b) -> float:
    """Distance between two rectangles; 0 if they touch or overlap."""
    dx = max(b[0] - a[2], a[0] - b[2], 0)
    dy = max(b[1] - a[3], a[1] - b[3], 0)
    return (dx * dx + dy * dy) ** 0.5


def _photos_on(page) -> list[tuple[int, tuple]]:
    """(xref, rect) of the photos on a page, top to bottom then left to right.

    Every placement of every image, from what is actually drawn: an image may
    be placed more than once, and Keynote leaves some off the edge of the
    slide.  The drawing list does not say which image each placement is, so
    it is matched by size in pixels -- by position, the slow way, only when
    two images on the slide are the same size."""
    page_rect = _rect(page.rect)
    area = page.rect.width * page.rect.height
    by_size: dict[tuple, list[int]] = {}
    for img in page.get_images(full=True):
        by_size.setdefault((img[2], img[3]), []).append(img[0])
    placed, slivers = [], 0
    for info in page.get_image_info():
        w, h = info["width"], info["height"]
        r = tuple(info["bbox"])
        if _covered(r, page_rect) < 0.5:
            continue                                    # off the edge of the slide
        if (r[2] - r[0]) * (r[3] - r[1]) >= 0.9 * area:
            continue                                    # the slide's background
        if w < _MIN_PIXELS or h < _MIN_PIXELS:
            slivers += 1
            continue
        xrefs = by_size.get((w, h), [])
        xref = xrefs[0] if len(xrefs) == 1 else _xref_placed_at(page, xrefs, r)
        if xref and all(x != xref for x, _r in placed):
            placed.append((xref, r))
    if slivers >= _MAX_SLIVERS:
        return []

    def is_shadow(r) -> bool:
        for _x, o in placed:
            if o is r:
                continue
            margins = (o[0] - r[0], o[1] - r[1], r[2] - o[2], r[3] - o[3])
            if all(0 <= m <= _SHADOW_MARGIN for m in margins):
                return True
        return False

    photos = [(x, r) for x, r in placed if not is_shadow(r)]
    photos.sort(key=lambda p: (round(p[1][1] / 50), p[1][0]))
    return photos


def _xref_placed_at(page, xrefs: list[int], r) -> int:
    for x in xrefs:
        for rr in page.get_image_rects(x):
            if _covered(r, _rect(rr)) > 0.9 and _covered(_rect(rr), r) > 0.9:
                return x
    return 0


@dataclass
class _Block:
    kind:  str             # "caption", "credit" or "heading"
    rect:  tuple
    lines: list[str]
    block: Any = None      # the PDF text block it came from


def _covered(a, b) -> float:
    """How much of rectangle a lies inside rectangle b, from 0 to 1."""
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    area = (a[2] - a[0]) * (a[3] - a[1])
    return w * h / area if w > 0 and h > 0 and area > 0 else 0.0


def _hidden_test(page):
    """A test for text that cannot be seen: text with a picture drawn over
    it later.  Keynote keeps a text box that a photo was afterwards put on
    top of, and it is in the PDF though no one can see it.  Text drawn over
    a photo is another matter -- that is visible."""
    log = page.get_bboxlog()
    texts = [(i, tuple(r)) for i, (k, r) in enumerate(log) if "text" in k]
    images = [(i, tuple(r)) for i, (k, r) in enumerate(log) if k == "fill-image"]

    def hidden(line) -> bool:
        drawn = [i for i, r in texts if _covered(r, line) >= 0.5]
        if not drawn:
            return False
        last = max(drawn)
        return any(i > last and _covered(line, r) >= 0.8 for i, r in images)
    return hidden


def _blocks_on(page) -> list[_Block]:
    """The slide's text, by kind, leaving out what is hidden."""
    blocks = []
    hidden = _hidden_test(page)
    # Without the images: by default "dict" carries every image's pixels too,
    # which makes it some seventy times slower
    flags = pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_PRESERVE_IMAGES
    for b in page.get_text("dict", flags=flags)["blocks"]:
        if b["type"] != 0:
            continue
        # Each line by its own type size: a block may hold a caption and its
        # credit line both
        for ln in b["lines"]:
            r = tuple(ln["bbox"])
            if hidden(r):
                continue
            sizes = {round(s["size"]) for s in ln["spans"] if s["text"].strip()}
            if not sizes:
                continue
            if max(sizes) - min(sizes) >= 5:
                kind = "heading"                        # large and small capitals
            elif max(sizes) >= 47:
                kind = "caption"
            else:
                kind = "credit"
            text = "".join(s["text"] for s in ln["spans"])
            last = blocks[-1] if blocks else None
            if last and last.kind == kind and last.block is b:
                last.rect = (min(last.rect[0], r[0]), last.rect[1],
                             max(last.rect[2], r[2]), max(last.rect[3], r[3]))
                last.lines.append(text)
            else:
                blocks.append(_Block(kind, r, [text], b))
    return blocks


def _clusters(blocks: list[_Block]) -> list[_Block]:
    """Join blocks of one kind that sit one under the other: Keynote often
    splits one text box into several blocks."""
    out: list[_Block] = []
    for b in sorted(blocks, key=lambda b: b.rect[1]):
        for c in out:
            if (c.kind == b.kind and b.rect[1] - c.rect[3] <= 20
                    and min(b.rect[2], c.rect[2]) > max(b.rect[0], c.rect[0])):
                c.rect = (min(c.rect[0], b.rect[0]), c.rect[1],
                          max(c.rect[2], b.rect[2]), max(c.rect[3], b.rect[3]))
                c.lines += b.lines
                break
        else:
            out.append(_Block(b.kind, b.rect, list(b.lines)))
    return out


def _nearest(rect, clusters: list[_Block]) -> list[_Block]:
    """The cluster for this photo: the only one, or the nearest."""
    if len(clusters) <= 1:
        return clusters
    return [min(clusters, key=lambda c: _gap(rect, c.rect))]


def _text(lines: list[str]) -> str:
    return re.sub(r"\s+", " ", " ".join(lines)).strip()


def read_pdf(pdf_path: "str | Path", cons: ConList | None = None,
             people: PeopleNames | None = None) -> list[PdfPhoto]:
    """Every photo in the PDF, with its caption, date, source and con, and
    whether the caption names anyone (if people is given)."""
    photos = []
    with pymupdf.open(str(pdf_path)) as doc:
        for pno, page in enumerate(doc, 1):
            placed = _photos_on(page)
            if not placed:
                continue
            clusters = _clusters(_blocks_on(page))
            # Small type that is not a credit line is part of the caption --
            # a quotation's attribution, say.  Where some small type names a
            # source, that is what credit lines look like on this slide;
            # otherwise it takes a date or a source
            small = [(c, parse_credit(c.lines)) for c in clusters if c.kind == "credit"]
            sourced = any(cr.source for _c, cr in small)
            for c, cr in small:
                if not (cr.source if sourced else (cr.year or cr.source)):
                    c.kind = "caption"
            captions = [c for c in clusters if c.kind == "caption"]
            credits = [c for c in clusters if c.kind == "credit"]
            headings = " ".join(_text(c.lines) for c in clusters if c.kind == "heading")
            several = len(placed) > 1
            for i, (xref, rect) in enumerate(placed, 1):
                key = f"p{pno:03d}" + (f"-{i}" if several else "")
                # One photo: every caption on the slide, top to bottom.  Several:
                # each its nearest, unless the slide has only the one
                mine = _nearest(rect, captions) if several and len(captions) > 1 else captions
                caption = " ".join(_text(c.lines) for c in mine)
                # A credit line sits with its caption.  With one photo, that is
                # the way to choose when a slide has more than one credit line;
                # with several, each photo takes its nearest
                near = rect if several or not mine else mine[0].rect
                credit_lines = [ln for c in _nearest(near, credits) for ln in c.lines]
                p = _describe(key, pno, xref, caption, credit_lines, headings, cons)
                if people:
                    p.named = people.names_someone(expand_names(caption))
                if len(credits) > max(1, len(placed)):
                    p.notes.append("The slide has more than one credit line; "
                                   "check that this one is the photo's.")
                photos.append(p)
    return photos


# When the caption says the photo was taken around the con rather than at it,
# the con's dates are not the photo's
_NOT_AT_THE_CON = re.compile(r"\b(before|after|pre-?con|post-?con|on the way|road-?trip\w*)\b", re.I)


def _describe(key: str, pno: int, xref: int, caption: str, credit_lines: list[str],
              heading: str, cons: ConList | None) -> PdfPhoto:
    cr = parse_credit(credit_lines)
    p = PdfPhoto(key, pno, xref, source=cr.source)
    year, month, day = cr.year, cr.month, cr.day

    if year and not day:
        found = _caption_date(caption, year)
        if found and (not month or found[0] == month):
            month, day = found

    con, note = (cons.find(caption, heading, cr.place, year, month)
                 if cons and year else (None, ""))
    if note:
        p.notes.append(note)
    when = _date_text(year, month, day)
    if con:
        p.con = con.name
        p.tags = con.name
        mid = con.middle()
        if (not day and mid and mid.year == year and (not month or month in con.months())
                and not _NOT_AT_THE_CON.search(caption)):
            when = _date_text(mid.year, mid.month, mid.day)
    p.date = when

    text = caption
    if cr.place:
        text = f"{text} — {cr.place}" if text else cr.place
    if cr.vague:
        text = f"{text} ({cr.vague})" if text else cr.vague
    if con and not any(len(_compress(a)) >= 4 and _compress(a) in _runs(_words(text))
                       for a in con.aliases):
        text = f"{text} ({con.name})" if text else con.name
    p.caption = expand_names(text).strip()
    return p


# ---------------------------------------------------------------------------
# The image itself
# ---------------------------------------------------------------------------
def save_photos(pdf_path: "str | Path", photos: list[PdfPhoto], folder: "str | Path") -> list[str]:
    """Write each photo to folder as "<PDF name> <key>.jpg" (or .png), and
    return the paths, in the same order.  A file already there is left as it
    is: it may be one the queue already has, and the user may have edited it."""
    stem = Path(pdf_path).stem
    paths = []
    with pymupdf.open(str(pdf_path)) as doc:
        for p in photos:
            data, ext = None, ""
            for e in (".jpg", ".png", ".jpx", ".bmp", ".tif"):
                if (Path(folder) / f"{stem} {p.key}{e}").exists():
                    ext = e
                    break
            if not ext:
                data, ext = image_bytes(doc, p.xref)
            out = Path(folder) / f"{stem} {p.key}{ext}"
            if data is not None:
                out.write_bytes(data)
            paths.append(str(out))
    return paths


def _is_srgb(profile_name: str) -> bool:
    """"sRGB IEC61966-2.1", "Artifex Software sRGB ICC Profile" -- but not
    "DCI-P3 D65 Gamut with sRGB Transfer", which is P3 with sRGB's curve."""
    name = profile_name.lower()
    return bool(re.search(r"\bsrgb\b", name)) and not re.search(r"p3|gamut|adobe|wide", name)


def image_bytes(doc, xref: int) -> tuple[bytes, str]:
    """The photo's own bytes and extension (".jpg", ".png" ...).

    A JPEG is passed on exactly as it is in the PDF.  A photo in a wider colour
    space than sRGB -- Display P3, Adobe RGB -- is converted to sRGB, since the
    profile that says so lives in the PDF, not in the image, and an upload
    without it would show duller colours."""
    from PIL import Image, ImageCms

    info = doc.extract_image(xref)
    data, ext = info["image"], "." + info["ext"].replace("jpeg", "jpg")
    icc = _icc_profile(doc, xref)
    if not icc:
        return data, ext
    try:
        src = ImageCms.ImageCmsProfile(io.BytesIO(icc))
        if _is_srgb(ImageCms.getProfileDescription(src).strip()) or info.get("colorspace") != 3:
            return data, ext
        with Image.open(io.BytesIO(data)) as im:
            im = im.convert("RGB")
            out = ImageCms.profileToProfile(im, src, ImageCms.createProfile("sRGB"),
                                            outputMode="RGB")
        buf = io.BytesIO()
        if ext == ".jpg":
            out.save(buf, "JPEG", quality=95, subsampling=0)
        else:
            out.save(buf, "PNG")
        return buf.getvalue(), ext
    except Exception:
        return data, ext


def _icc_profile(doc, xref: int) -> bytes | None:
    obj = doc.xref_object(xref, compressed=True)
    m = (re.search(r"/ColorSpace\s*\[\s*/ICCBased\s+(\d+)\s+0\s+R", obj)
         or re.search(r"/ColorSpace\s+(\d+)\s+0\s+R", obj))
    if not m:
        return None
    cs = int(m.group(1))
    inner = re.search(r"/ICCBased\s+(\d+)\s+0\s+R", doc.xref_object(cs, compressed=True))
    try:
        return doc.xref_stream(int(inner.group(1)) if inner else cs)
    except Exception:
        return None
