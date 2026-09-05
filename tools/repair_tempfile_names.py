"""Restore the real file name of photos stored on Piwigo under a tempfile's.

PhotosUploader used to build its upload copy with NamedTemporaryFile, and
Piwigo stores the name of the file it is given, so those photos are stored as
"tmp2qfmcfwp.jpg".  Their titles kept the real name, and that is what this puts
back.  The cause was fixed in PhotosUploader (and in PhotosEditor, the only
other program that uploads), so the damage is historical: it does not grow.

Only the `file` field changes.  The image itself lives at a generated path --
"upload/2026/04/17/20260417162339-741e115b.jpg" -- that has nothing to do with
this field, so nothing about the picture moves.  Verified on both a temp-named
and an ordinary photo.

Three steps, deliberately separate: the survey is slow and read-only, the plan
is there to be read before anything happens, and only "apply" writes.

    python tools/repair_tempfile_names.py survey     # walk the site (~15 min)
    python tools/repair_tempfile_names.py plan       # say what would change
    python tools/repair_tempfile_names.py apply      # change it (~25 min)

As surveyed on 2026-09-05: 12,558 photos over 595 albums, 1,481 stored under a
tempfile name, 1,420 of them repairable.  The other 61 have a title that is
itself a *different* tempfile name -- a photo uploaded twice through the broken
path, the second upload taking the first's temp name as its title.  Piwigo no
longer holds their real name; it may survive in FileDict.json or in the
uploaded_info state, both of which record output_filename.
"""
import json
import os
import re
import sys
from pathlib import Path

_TOOLS = Path(__file__).resolve().parent
_APP = _TOOLS.parent
sys.path.insert(0, str(_APP.parent / "PiwigoHelpers"))

import AlbumHierarchy                                          # noqa: E402
from CredentialStore import CredentialStore                    # noqa: E402

LISTING = _TOOLS / "tempfile-named photos.json"

# Python's tempfile names are "tmp" and eight characters from [a-z0-9_]
TEMPFILE = re.compile(r"^tmp[a-z0-9_]{8}(\.|$)", re.IGNORECASE)
ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def repaired_name(row: dict) -> "str | None":
    """The file name to put back, or None when this one cannot be repaired.

    The title is the only surviving record of the real name.  A title that is
    itself a tempfile name, or empty, or unusable as a file name, leaves
    nothing to restore from.  An extension is taken from the current name when
    the title has none.
    """
    title = (row.get("name") or "").strip()
    if not title or TEMPFILE.match(title) or ILLEGAL.search(title):
        return None
    if not os.path.splitext(title)[1]:
        title += os.path.splitext(row.get("file") or "")[1]
    return title


def connect():
    store = CredentialStore(_APP, "PhotosUploader Params.json")
    creds, params = store.load_credentials(), store.load_op_params()
    client = AlbumHierarchy.PiwigoClient(
        creds["url"], creds["username"], creds["password"],
        verify_ssl=creds.get("verify_ssl", True),
        rate_limit_calls_per_second=params.get("rate_limit_calls_per_second", 2.0))
    client.login(creds["username"], creds["password"])
    return client


def survey():
    """Walk every album and list the photos stored under a tempfile name."""
    client = connect()
    found, seen, photos = [], set(), 0
    try:
        albums = client.get_albums()
        print(f"{len(albums)} albums to walk")
        for i, album in enumerate(albums, 1):
            cid = album.get("id")
            if cid is None:
                continue
            try:
                for img in client.get_album_images(int(cid)):
                    iid = img.get("id")
                    if iid in seen:
                        continue
                    seen.add(iid)
                    photos += 1
                    if TEMPFILE.match(str(img.get("file") or "")):
                        found.append({"id": iid,
                                      "file": str(img.get("file") or ""),
                                      "name": str(img.get("name") or ""),
                                      "album": album.get("name", "")})
            except Exception as e:
                print(f"  album {cid} ({album.get('name', '')}): {e}")
            if i % 25 == 0:
                print(f"  {i}/{len(albums)} albums, {photos} photos, "
                      f"{len(found)} temp-named so far", flush=True)
    finally:
        try:
            client.logout()
        except Exception:
            pass
    LISTING.write_text(json.dumps(found, indent=2), encoding="utf-8")
    print(f"\n{photos} photos seen, {len(found)} stored under a tempfile name")
    print(f"written to {LISTING}")


def load():
    if not LISTING.is_file():
        sys.exit(f"No survey found at {LISTING}.  Run 'survey' first.")
    rows = json.loads(LISTING.read_text(encoding="utf-8"))
    doable = [(r, n) for r, n in ((r, repaired_name(r)) for r in rows) if n]
    skipped = [r for r in rows if repaired_name(r) is None]
    return rows, doable, skipped


def plan():
    rows, doable, skipped = load()
    print(f"{len(rows)} photo(s) stored under a tempfile name\n")
    for row, new in doable:
        print(f"  id {row['id']:<7} {row['file']:<24} -> {new}")
    for row in skipped:
        print(f"  id {row['id']:<7} {row['file']:<24} -- SKIPPED, "
              f"title is {row.get('name')!r}")
    print(f"\nnothing changed.  'apply' would rename {len(doable)}, "
          f"leaving {len(skipped)}.")


def apply():
    rows, doable, skipped = load()
    client = connect()
    done, failed = [], []
    try:
        print(f"renaming {len(doable)} of {len(rows)}:")
        for row, new in doable:
            try:
                # setInfo's `file` is the stored name.  single_value_mode must
                # be "replace" or Piwigo fills rather than overwrites.
                # AlbumHierarchy.set_image_info does not carry this field; it
                # is wanted for this repair alone, so the call is made here.
                client._call("pwg.images.setInfo",
                             {"image_id": row["id"], "file": new,
                              "single_value_mode": "replace"})
                got = str(client.get_image_info(row["id"]).get("file") or "")
                if got == new:
                    done.append(row["id"])
                else:
                    failed.append((row["id"], f"reads back as {got!r}"))
                    print(f"  id {row['id']:<7} FAILED: reads back as {got!r}")
            except Exception as e:
                failed.append((row["id"], str(e)))
                print(f"  id {row['id']:<7} FAILED: {e}")
            if len(done) % 50 == 0 and done:
                print(f"  {len(done)} renamed…", flush=True)
    finally:
        try:
            client.logout()
        except Exception:
            pass
    print(f"\n{len(done)} renamed, {len(failed)} failed, {len(skipped)} skipped")
    return 1 if failed else 0


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "plan"
    if mode == "survey":
        survey()
    elif mode == "plan":
        plan()
    elif mode == "apply":
        sys.exit(apply())
    else:
        sys.exit(__doc__)
