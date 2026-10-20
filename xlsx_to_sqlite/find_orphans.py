"""
find_orphans.py

Finds image and video files that no item uses any more, and (with --apply)
removes them together with their manifest entries. Used by the "Clean up
orphans" workflow; can also be run locally from the repository root.

What counts as an orphan:
  - image: assets/images/{thumbnails,200,1000,2000}/<id>.jpg and its entries in
    assets/images/manifest.json, for an id that is not an item at all. (Only
    primary_image items and secondary_image items that are not videos use
    images. Images of other existing items are reported as "unused but kept",
    never removed.)
  - video: assets/videos/<id>.mp4 and <id>.jpg and its entry in
    assets/videos/manifest.json, for an id that is not in data/videos.csv and
    not an item.
Also reported, never removed: problems such as a videos.csv line whose parent
is gone, and original uploads (video-originals/, image-originals/) whose item
no longer exists.

Before anything is removed, the database must be the committed one, pass an
integrity check, match a fresh build of data/timeline.xlsx + data/videos.csv,
and not have lost more items than the build allows; and no more than
--max-files files may be removed. The database, workbook, videos.csv and
originals are never touched.

Usage:
    python xlsx_to_sqlite/find_orphans.py [--apply] [--max-files N]
        [--summary SUMMARY_MD] [--commit-message MESSAGE_TXT]

Without --apply it is a dry run: it lists what it would remove and changes
nothing. With --apply it removes the files with `git rm` and stages the edited
manifests; committing is left to the caller.

With --removed-since PREVIOUS_DB (used by the "Build timeline database"
workflow right after it commits a new database) only the leftovers of ids
that were items in PREVIOUS_DB and are not items now are considered. If a
limit is exceeded it removes nothing and prints a warning instead of failing,
so the database rebuild still goes through; the manual workflow handles it.
"""

import argparse
import csv
import json
import re
import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path

from xlsx_to_sqlite import IMAGE_SIZES, IMAGE_TYPES, MAX_ITEMS_REMOVED

DB = "assets/timeline_data.db"
WORKBOOK = "data/timeline.xlsx"
VIDEOS_CSV = "data/videos.csv"
IMAGES_DIR = Path("assets/images")
VIDEOS_DIR = Path("assets/videos")
IMAGE_MANIFEST = IMAGES_DIR / "manifest.json"
VIDEO_MANIFEST = VIDEOS_DIR / "manifest.json"
ORIGINALS_DIRS = [Path("video-originals"), Path("image-originals")]
TOOLS = Path(__file__).resolve().parent

IMAGE_FILE = re.compile(r"^([0-9]+)\.jpg$")
VIDEO_FILE = re.compile(r"^([0-9]+)\.(mp4|jpg)$")
ORIGINAL_FILE = re.compile(r"^([0-9]+)\.[A-Za-z0-9]+$")
DEFAULT_MAX_FILES = 40


class Refused(Exception):
    """A safety check failed; nothing is removed."""


def git(*args, check=True):
    result = subprocess.run(["git", *args], capture_output=True, text=True)
    if check and result.returncode != 0:
        raise Refused(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result


def read_json_keep_style(path):
    """Returns (data, dump) where dump(data) writes it back in the file's exact style."""
    raw = Path(path).read_text(encoding="utf-8")
    data = json.loads(raw)
    styles = [dict(indent=2), dict()]
    for style in styles:
        for ensure_ascii in (True, False):
            for newline in ("", "\n"):
                if json.dumps(data, ensure_ascii=ensure_ascii, **style) + newline == raw:
                    return data, lambda d, s=style, a=ensure_ascii, n=newline: json.dumps(d, ensure_ascii=a, **s) + n
    raise Refused(f"{path}: its formatting is not one this script can rewrite exactly; edit it by hand")


def item_ids(db_path):
    with sqlite3.connect(f"file:{db_path}?mode=ro", uri=True) as con:
        return {i for (i,) in con.execute("SELECT id FROM items")}


def check_count(count, previous_count, notes, allow_shrink=False):
    if previous_count - count > MAX_ITEMS_REMOVED and allow_shrink:
        notes.append(f"Item count {count} is {previous_count - count} below the previous version ({previous_count}), "
                     f"more than the limit of {MAX_ITEMS_REMOVED}; allowed by allow_shrink.")
        return
    if previous_count - count > MAX_ITEMS_REMOVED:
        raise Refused(f"The database has {count} items, {previous_count - count} fewer than its previous version "
                      f"({previous_count}); the limit is {MAX_ITEMS_REMOVED}. If those deletions are intended, "
                      f"remove the leftovers with the manual 'Clean up orphans' workflow: dry run first, "
                      f"then apply with allow_shrink ticked.")
    notes.append(f"Item count {count} (previous version {previous_count}) is within the limit.")


def safety_checks(notes, previous_db=None, allow_shrink=False):
    # 1. The database is the committed one and is intact.
    if git("ls-files", "--error-unmatch", DB, check=False).returncode != 0:
        raise Refused(f"{DB} is not tracked by git")
    if git("diff", "--quiet", "HEAD", "--", DB, check=False).returncode != 0:
        raise Refused(f"{DB} differs from the committed version")
    try:
        with sqlite3.connect(f"file:{DB}?mode=ro", uri=True) as con:
            result = con.execute("PRAGMA integrity_check").fetchone()[0]
    except sqlite3.DatabaseError as error:
        result = str(error)
    if result != "ok":
        raise Refused(f"{DB} failed the integrity check: {result}")
    notes.append(f"`{DB}` is the committed version and passes the integrity check.")

    # 2. It matches a fresh build from the workbook (so the last build did not fail).
    with tempfile.TemporaryDirectory() as tmp:
        fresh = Path(tmp) / "fresh.db"
        build = subprocess.run([sys.executable, str(TOOLS / "xlsx_to_sqlite.py"), WORKBOOK, str(fresh),
                                "--videos", VIDEOS_CSV, "--manifest", str(VIDEO_MANIFEST),
                                "--images", str(IMAGE_MANIFEST)], capture_output=True, text=True)
        if build.returncode != 0:
            raise Refused("data/timeline.xlsx does not build (fix it first; the build workflow shows why):\n"
                          + build.stderr.strip())
        diff = subprocess.run([sys.executable, str(TOOLS / "db_diff.py"), DB, str(fresh)], capture_output=True, text=True)
        if diff.returncode != 0:
            raise Refused(f"{DB} is not up to date with data/timeline.xlsx and data/videos.csv; the last "
                          f"'Build timeline database' run may have failed. Fix that first.\n\n" + diff.stdout.strip())
    notes.append(f"`{DB}` matches a fresh build of `{WORKBOOK}` and `{VIDEOS_CSV}`.")

    # 3. The item count is plausible compared with the previous committed version.
    count = len(item_ids(DB))
    if previous_db:
        check_count(count, len(item_ids(previous_db)), notes, allow_shrink)
        return
    versions = git("log", "-2", "--format=%H", "--", DB).stdout.split()
    if len(versions) == 2:
        with tempfile.TemporaryDirectory() as tmp:
            previous = Path(tmp) / "previous.db"
            previous.write_bytes(subprocess.run(["git", "show", f"{versions[1]}:{DB}"], capture_output=True, check=True).stdout)
            previous_count = len(item_ids(previous))
        check_count(count, previous_count, notes, allow_shrink)
    else:
        notes.append(f"Item count {count} (no previous version to compare with).")


def find(report, scope=None):
    """Fills report. With scope (a set of ids), only those ids are considered."""
    in_scope = (lambda i: True) if scope is None else (lambda i: i in scope)
    with sqlite3.connect(f"file:{DB}?mode=ro", uri=True) as con:
        types = dict(con.execute("SELECT id, type FROM items"))
    with open(VIDEOS_CSV, newline="", encoding="utf-8-sig") as f:
        videos = {int(r["video_id"]): int(r["parent_id"]) for r in csv.DictReader(f)}
    needs_image = {i for i, t in types.items() if t in IMAGE_TYPES and i not in videos}

    image_manifest, _ = read_json_keep_style(IMAGE_MANIFEST)
    video_manifest, _ = read_json_keep_style(VIDEO_MANIFEST)

    # Images: every id that has a file or a manifest entry in one of the four folders.
    image_ids = {}
    for size in IMAGE_SIZES:
        for path in sorted((IMAGES_DIR / size).iterdir()) if (IMAGES_DIR / size).is_dir() else []:
            m = IMAGE_FILE.match(path.name)
            if m:
                image_ids.setdefault(int(m.group(1)), set()).add(f"{size}/{path.name}")
            else:
                report["unrecognised"].append(str(path))
    for entry in image_manifest:
        size, _, name = entry["path"].partition("/")
        m = IMAGE_FILE.match(name)
        if size in IMAGE_SIZES and m:
            image_ids.setdefault(int(m.group(1)), set()).add(entry["path"])
    for image_id in sorted(image_ids):
        if image_id in needs_image or not in_scope(image_id):
            continue
        if image_id in types or image_id in videos:
            kind = types.get(image_id, "video in data/videos.csv")
            report["kept"].append((f"image {image_id}", f"still an item ({kind}), which never shows an image"))
            continue
        files = [IMAGES_DIR / p for p in sorted(image_ids[image_id]) if (IMAGES_DIR / p).is_file()]
        entries = [e for e in image_manifest if e["path"] in image_ids[image_id]]
        report["orphans"].append(("image", image_id, files, entries))

    # Videos: every id with a file or a manifest entry.
    video_ids = {}
    for path in sorted(VIDEOS_DIR.iterdir()):
        m = VIDEO_FILE.match(path.name)
        if m:
            video_ids.setdefault(int(m.group(1)), set()).add(path.name)
        elif path.name != "manifest.json":
            report["unrecognised"].append(str(path))
    for entry in video_manifest:
        video_ids.setdefault(entry["id"], set())
    for video_id in sorted(video_ids):
        if video_id in videos or not in_scope(video_id):
            continue
        if video_id in types:
            report["kept"].append((f"video {video_id}", f"still an item ({types[video_id]}) in the database"))
            continue
        files = [VIDEOS_DIR / n for n in sorted(video_ids[video_id])]
        entries = [e for e in video_manifest if e["id"] == video_id]
        report["orphans"].append(("video", video_id, files, entries))

    # Problems to fix by hand (never removed automatically).
    for video_id, parent_id in sorted(videos.items()) if scope is None else []:
        if parent_id not in types:
            report["errors"].append(f"`{VIDEOS_CSV}` line `{video_id},{parent_id}`: item {parent_id} no longer exists. "
                                    f"Either restore item {parent_id} in `{WORKBOOK}`, or delete this line from "
                                    f"`{VIDEOS_CSV}` (video {video_id} then becomes an orphan this workflow can remove).")
    image_paths = {e["path"] for e in image_manifest}
    for item_id in sorted(needs_image) if scope is None else []:
        for size in IMAGE_SIZES:
            p = f"{size}/{item_id}.jpg"
            if not (IMAGES_DIR / p).is_file() or p not in image_paths:
                report["errors"].append(f"Item {item_id} needs `{IMAGES_DIR / p}` (file and manifest entry); "
                                        f"one of them is missing.")

    # Originals: report only.
    for folder in ORIGINALS_DIRS:
        if not folder.is_dir():
            continue
        for path in sorted(folder.iterdir()):
            m = ORIGINAL_FILE.match(path.name)
            if m and int(m.group(1)) not in types and in_scope(int(m.group(1))):
                report["originals"].append(str(path))
    return image_manifest, video_manifest


def apply(report, max_files, image_manifest, video_manifest):
    files = [f for _, _, fs, _ in report["orphans"] for f in fs]
    if len(files) > max_files:
        raise Refused(f"{len(files)} files would be removed, more than the limit of {max_files}. "
                      f"Run again with a higher max_files if that is intended.")
    allowed = re.compile(r"^assets/(images/(thumbnails|200|1000|2000)/[0-9]+\.jpg|videos/[0-9]+\.(mp4|jpg))$")
    for f in files:
        if not allowed.match(f.as_posix()):
            raise Refused(f"Refusing to remove {f}: not an image or video file")
    gone_images = {e["path"] for kind, _, _, es in report["orphans"] if kind == "image" for e in es}
    gone_videos = {e["id"] for kind, _, _, es in report["orphans"] if kind == "video" for e in es}
    for path, entries, keep in ((IMAGE_MANIFEST, image_manifest, lambda e: e["path"] not in gone_images),
                                (VIDEO_MANIFEST, video_manifest, lambda e: e["id"] not in gone_videos)):
        _, dump = read_json_keep_style(path)
        kept = [e for e in entries if keep(e)]
        if len(kept) != len(entries):
            Path(path).write_text(dump(kept), encoding="utf-8")
            git("add", "--", str(path))
    if files:
        git("rm", "-q", "--", *map(str, files))


def write_removed_summary(report, notes, removed, refused, path):
    """Summary for --removed-since: the leftovers of the items this upload deleted."""
    files = [f for _, _, fs, _ in report["orphans"] for f in fs]
    entries = [(kind, e) for kind, _, _, es in report["orphans"] for e in es]
    if refused:
        heading = "NOT REMOVED (see warning)"
    else:
        heading = "REMOVED" if report["orphans"] else "nothing to remove"
    lines = [f"## Leftover files of deleted items: {heading}", "",
             f"Ids removed by this upload (items in the previous database, not in the new one): "
             f"{', '.join(map(str, sorted(removed))) or 'none'}", ""]
    if refused:
        lines += [f"> **Warning: no files were removed.** {refused}", ">",
                  "> The database was still rebuilt. Use the manual **Clean up orphans** workflow "
                  "(dry run first) to remove these leftovers.", ""]
    lines += ["### Safety checks", ""] + [f"- {n}" for n in notes] + [""]
    verb = "Would have deleted" if refused else "Deleted"
    lines += [f"### {verb} {len(files)} file(s) and {len(entries)} manifest entr{'y' if len(entries) == 1 else 'ies'}", ""]
    if report["orphans"]:
        lines += ["| | id | files | manifest entries |", "|---|---|---|---|"]
        for kind, i, fs, es in report["orphans"]:
            shown = ", ".join(e["path"] if kind == "image" else f"id {e['id']}" for e in es) or "none"
            lines.append(f"| {kind} | {i} | {', '.join(f'`{f}`' for f in fs) or 'none'} | {shown} |")
        lines.append("")
    if report["kept"]:
        lines += ["### Kept", ""] + [f"- {what}: {why}" for what, why in report["kept"]] + [""]
    if report["originals"]:
        lines += ["### Original uploads of removed items (report only, never deleted)", ""]
        lines += [f"- `{o}`" for o in report["originals"]] + [""]
    text = "\n".join(lines)
    print(text)
    if refused:
        print(f"::warning title=Leftover files not removed::{refused.splitlines()[0]} "
              f"The database was rebuilt; run 'Clean up orphans' to remove the leftovers.")
    if path:
        Path(path).write_text(text + "\n", encoding="utf-8")


def write_summary(report, notes, mode, max_files, refused, path):
    files = [f for _, _, fs, _ in report["orphans"] for f in fs]
    entries = sum(len(es) for _, _, _, es in report["orphans"])
    verb = {"apply": "Deleted", "dry": "Would delete"}[mode]
    if mode == "apply" and not refused:
        heading = "APPLIED" if report["orphans"] else "APPLIED (nothing to remove)"
    else:
        heading = "DRY RUN (nothing changed)" if not refused else "STOPPED (nothing changed)"
    lines = [f"## Clean up orphans: {heading}", ""]
    if refused:
        lines += [f"**Stopped, nothing was changed:** {refused}", ""]
    lines += ["### Safety checks", ""] + [f"- {n}" for n in notes] + [""]
    lines += [f"### Found: {len(report['orphans'])} orphan(s), {len(files)} file(s), {entries} manifest entr{'y' if entries == 1 else 'ies'}", ""]
    if report["orphans"]:
        lines += [f"| | id | {verb.lower()} files | {verb.lower()} manifest entries |", "|---|---|---|---|"]
        for kind, i, fs, es in report["orphans"]:
            shown = ", ".join(e["path"] if kind == "image" else f"id {e['id']}" for e in es) or "none"
            lines.append(f"| {kind} | {i} | {', '.join(f'`{f}`' for f in fs) or 'none'} | {shown} |")
        lines.append("")
        if mode == "dry" and not refused:
            over = len(files) > max_files
            lines += [f"Limit per run: {max_files} files." + (" **Over the limit: an apply run would refuse; raise max_files.**" if over else ""), ""]
    else:
        lines += ["No orphans.", ""]
    lines += ["### Skipped (unused but kept)", ""]
    if report["kept"]:
        lines += [f"- {what}: {why}" for what, why in report["kept"]]
        lines += ["", "These are downloaded by every visitor (the site preloads every image in the manifest) but "
                  "never shown. To remove them, in one commit: `git rm` the four files "
                  "`assets/images/{thumbnails,200,1000,2000}/<id>.jpg` **and** delete their four entries from "
                  "`assets/images/manifest.json`. Never remove the files without the entries: the site would fail to load.", ""]
    else:
        lines += ["None.", ""]
    lines += ["### Problems to fix by hand (never changed automatically)", ""]
    lines += ([f"- {e}" for e in report["errors"]] or ["None."]) + [""]
    lines += ["### Original uploads whose item no longer exists (report only, never deleted)", ""]
    lines += ([f"- `{o}`" for o in report["originals"]] or ["None."]) + [""]
    if report["unrecognised"]:
        lines += ["### Files with unexpected names (not touched)", ""] + [f"- `{u}`" for u in report["unrecognised"]] + [""]
    text = "\n".join(lines)
    print(text)
    if path:
        Path(path).write_text(text + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Find (and with --apply remove) unused image and video files.")
    parser.add_argument("--apply", action="store_true", help="remove the orphans (default: dry run)")
    parser.add_argument("--max-files", type=int, default=DEFAULT_MAX_FILES, help="most files to remove in one run")
    parser.add_argument("--summary", help="also write the Markdown report to this file")
    parser.add_argument("--commit-message", help="write a commit message listing the removed ids to this file")
    parser.add_argument("--allow-shrink", action="store_true",
                        help="allow it even if the database lost more items than the build's limit")
    parser.add_argument("--removed-since", metavar="PREVIOUS_DB",
                        help="only the leftovers of ids that were items in PREVIOUS_DB and are not items now; "
                             "a refusal is a warning (exit 0), not a failure")
    args = parser.parse_args()

    report = {"orphans": [], "kept": [], "errors": [], "originals": [], "unrecognised": []}
    notes, refused, removed = [], None, None
    mode = "apply" if args.apply else "dry"
    try:
        if args.removed_since:
            with open(VIDEOS_CSV, newline="", encoding="utf-8-sig") as f:
                listed_videos = {int(r["video_id"]) for r in csv.DictReader(f)}
            removed = item_ids(args.removed_since) - item_ids(DB) - listed_videos
        safety_checks(notes, args.removed_since, args.allow_shrink)
        image_manifest, video_manifest = find(report, removed)
        if args.apply and report["orphans"]:
            apply(report, args.max_files, image_manifest, video_manifest)
    except Refused as error:
        refused = str(error)
        if args.removed_since and not report["orphans"]:
            try:  # still show what is left behind
                find(report, removed)
            except (Refused, OSError, sqlite3.DatabaseError):
                pass
    if args.removed_since:
        write_removed_summary(report, notes, removed or set(), refused, args.summary)
    else:
        write_summary(report, notes, mode, args.max_files, refused, args.summary)
    if args.commit_message and args.apply and not refused and report["orphans"]:
        ids = lambda kind: ", ".join(str(i) for k, i, _, _ in report["orphans"] if k == kind) or "none"
        title = "Remove leftover files of deleted items" if args.removed_since else "Clean up orphans"
        Path(args.commit_message).write_text(
            f"{title}: images {ids('image')}; videos {ids('video')}\n\n"
            + "\n".join(f"Removed {f}" for _, _, fs, _ in report["orphans"] for f in fs) + "\n", encoding="utf-8")
    if args.removed_since:
        return 0
    return 1 if refused else 0


if __name__ == "__main__":
    sys.exit(main())
