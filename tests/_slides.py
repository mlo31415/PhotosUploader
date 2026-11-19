"""A small slide-show PDF made the way Keynote makes one, for the PDF tests.

    p1  a title slide: text, no photo
    p2  one photo with its drop shadow; a caption and a three-line credit
    p3  two photos, one caption between them
    p4  a photo over an old text box (hidden), a collection credit
    p5  a graphic cut into a dozen slivers and two strips: no photos
    p6  a photo placed twice, once off the edge of the slide; "1980s"
    p7  a section heading in large and small capitals; no con in the caption
    p8  a photo with no credit line at all

and a convention timeline in the form FancyAnalyzer writes.
"""
import io
from pathlib import Path

import pymupdf
from PIL import Image

W, H = 1920, 1080


def jpeg(colour: str, size=(600, 400)) -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", size, colour).save(buf, "JPEG", quality=90)
    return buf.getvalue()


def png(colour: str, size) -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", size, colour).save(buf, "PNG")
    return buf.getvalue()


PHOTOS = {"p002": jpeg("#c04040"), "p003-1": jpeg("#40c040"), "p003-2": jpeg("#4040c0", (500, 300)),
          "p004": jpeg("#c0c040"), "p006": jpeg("#40c0c0", (700, 500)),
          "p007": jpeg("#c040c0", (640, 480)), "p008": jpeg("#808080", (480, 640))}


def _photo(page, rect, data):
    """A photo with its shadow: a black image 20 pt bigger all round, first."""
    x0, y0, x1, y1 = rect
    shadow = png("#000000", (int(x1 - x0) + 40, int(y1 - y0) + 40))
    page.insert_image(pymupdf.Rect(x0 - 20, y0 - 20, x1 + 20, y1 + 20), stream=shadow)
    page.insert_image(pymupdf.Rect(*rect), stream=data)


def _lines(page, x, y, lines, size):
    for i, ln in enumerate(lines):
        page.insert_text((x, y + i * size * 1.15), ln, fontsize=size)


def make_slideshow(path: Path) -> Path:
    doc = pymupdf.open()
    background = jpeg("#9ab0c0", (W, H))

    def slide():
        page = doc.new_page(width=W, height=H)
        page.insert_image(page.rect, stream=background)
        return page

    p = slide()                                                         # 1
    _lines(p, 400, 400, ["Celebrating Friendship"], 94)

    p = slide()                                                         # 2
    _photo(p, (100, 100, 1100, 767), PHOTOS["p002"])
    _lines(p, 1250, 400, ["Bob Tucker & Pat at", "Corflu IV"], 56)
    _lines(p, 1250, 800, ["April 1987 / photo by", "Colin Hinz", "Covington, KY / GFS"], 44)

    p = slide()                                                         # 3
    _photo(p, (60, 60, 760, 527), PHOTOS["p003-1"])
    _photo(p, (1100, 500, 1800, 920), PHOTOS["p003-2"])
    _lines(p, 800, 300, ["Pat and the Thames"], 56)
    _lines(p, 800, 1000, ["1987 / London, UK / PV"], 44)

    p = slide()                                                         # 4
    _lines(p, 200, 300, ["An old caption no one can see"], 56)
    _photo(p, (100, 100, 1100, 767), PHOTOS["p004"])
    _lines(p, 1250, 400, ["Not-A-D-Con outing"], 56)
    _lines(p, 1250, 800, ["May 1985 / From the", "Lien-Garey photo", "collection / GFS"], 44)

    p = slide()                                                         # 5
    p.insert_image(pymupdf.Rect(100, 100, 900, 1000), stream=jpeg("#203040", (900, 1000)))
    p.insert_image(pymupdf.Rect(900, 100, 1000, 1000), stream=jpeg("#203041", (120, 1000)))
    for i in range(12):
        p.insert_image(pymupdf.Rect(1000 + i * 20, 500, 1015 + i * 20, 505),
                       stream=jpeg(f"#2030{50 + i:02x}", (30, 10)))
    _lines(p, 1300, 400, ["Book Launch Party"], 56)

    p = slide()                                                         # 6
    p.insert_image(pymupdf.Rect(2400, 100, 3100, 600), stream=PHOTOS["p006"])
    _photo(p, (100, 100, 800, 600), PHOTOS["p006"])
    _lines(p, 1000, 400, ["Liz Danforth"], 56)
    _lines(p, 1000, 800, ["1980s / PV"], 44)

    p = slide()                                                         # 7
    _photo(p, (100, 200, 740, 680), PHOTOS["p007"])
    # A heading in small capitals: one line, two sizes
    p.insert_text((1000, 150), "M", fontsize=56)
    p.insert_text((1047, 150), "INICON", fontsize=45)
    _lines(p, 1000, 400, ["Al Standish & Nate Bucklin"], 56)
    _lines(p, 1000, 800, ["March 1997", "Bloomington, MN / GFS"], 44)

    p = slide()                                                         # 8
    _photo(p, (100, 100, 580, 740), PHOTOS["p008"])
    _lines(p, 1000, 400, ["Thank you"], 56)

    doc.save(str(path))
    doc.close()
    return path


TIMELINE = """This is a chronological list of SF conventions
<tab>
colspan="2"| <big><big>'''1985'''</big></big>
August 30-September 2, 1985||[[LoneStarCon]]&nbsp;&nbsp;&nbsp;<small>(Austin, TX)</small>
colspan="2"| <big><big>'''1987'''</big></big>
April 3-5, 1987||[[Corflu 4|Corflu IV]]&nbsp;&nbsp;&nbsp;<small>(Covington, KY)</small>
 ||[[Corflu 3]]&nbsp;&nbsp;&nbsp;<small>(Somewhere)</small>
colspan="2"| <big><big>'''1997'''</big></big>
March 28-30, 1997||[[Minicon 32]]&nbsp;&nbsp;&nbsp;<small>(Bloomington, MN)</small>
colspan="2"| <big><big>'''2018'''</big></big>
February 16-18, 2018||[[Boskone 55]]&nbsp;&nbsp;&nbsp;<small>(Boston, MA)</small>
May 4-6, 2018||[[Corflu 35]]&nbsp;&nbsp;&nbsp;<small>(Toronto)</small>
colspan="2"| <big><big>'''2020'''</big></big>
March 13-15, 2020||[[Corflu 37]]&nbsp;&nbsp;&nbsp;<small>(College Station, TX)</small>
"""


def make_timeline(path: Path) -> Path:
    path.write_text(TIMELINE, encoding="utf-8")
    return path
