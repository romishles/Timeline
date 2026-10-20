"""
db_to_xlsx.py

Exports the timeline database to a workbook in the format xlsx_to_sqlite.py
reads, so the database can be turned back into an editable Excel file.

  - only the non-video items (the videos are the items listed with
    --videos-csv, or found automatically: secondary_image items that link
    to exactly one primary_image and are in the video manifest)
  - about_author is written as about_creator
  - text_character and text_quote are left out (they are always empty)
  - links to videos are left out (the build adds them from the videos list)
  - English era names are written as the Hebrew era names

Usage:
    python db_to_xlsx.py DATABASE OUTPUT_XLSX [--manifest MANIFEST_JSON]
        [--videos-csv VIDEOS_CSV]

Example (from the repository root):
    python xlsx_to_sqlite/db_to_xlsx.py assets/timeline_data.db timeline.xlsx \\
        --manifest assets/videos/manifest.json --videos-csv videos.csv
"""

import argparse
import csv
import json
import sqlite3
import sys
from pathlib import Path

from openpyxl import Workbook

from xlsx_to_sqlite import ALWAYS_NULL_COLUMNS, DB_COLUMNS, ERAS, SHEET_NAME, WORKBOOK_TO_DB

# Era names the database used before the workbook became the source.
ENGLISH_TO_HEBREW_ERA = dict(zip(
    [
        "Foundations & Direct Manipulation",
        "The GUI Canon",
        "Tangible & Ambient",
        "Multi-Touch & Embodied",
        "Natural & Spatial",
        "Generative & Ambient AI",
    ],
    ERAS,
))

DB_TO_WORKBOOK = {db: wb for wb, db in WORKBOOK_TO_DB.items()}


def find_videos(con, manifest_path):
    """Returns {video_id: parent_id} for every video in the manifest."""
    manifest_ids = {e["id"] for e in json.loads(Path(manifest_path).read_text(encoding="utf-8"))}
    types = dict(con.execute("SELECT id, type FROM items"))
    videos = {}
    for video_id in sorted(manifest_ids):
        parents = [t for (t,) in con.execute("SELECT target_id FROM links WHERE id = ?", (video_id,))
                   if types.get(t) == "primary_image"]
        if types.get(video_id) != "secondary_image" or len(parents) != 1:
            sys.exit(f"Video {video_id}: expected a secondary_image item linked to exactly one "
                     f"primary_image, found type {types.get(video_id)!r} and parents {parents}")
        videos[video_id] = parents[0]
    return videos


def main():
    parser = argparse.ArgumentParser(description="Export the timeline database to the timeline workbook format.")
    parser.add_argument("database", help="e.g. assets/timeline_data.db")
    parser.add_argument("output", help="the .xlsx file to write")
    parser.add_argument("--manifest", default="assets/videos/manifest.json",
                        help="the video manifest, used to recognise the video items")
    parser.add_argument("--videos-csv", help="also write the videos list (video_id,parent_id) to this file")
    args = parser.parse_args()

    con = sqlite3.connect(f"file:{args.database}?mode=ro", uri=True)
    videos = find_videos(con, args.manifest)

    columns = [c for c in DB_COLUMNS if c not in ALWAYS_NULL_COLUMNS]
    link_targets = {}
    for source, target in con.execute("SELECT id, target_id FROM links ORDER BY rowid"):
        if source not in videos and target not in videos:
            link_targets.setdefault(source, []).append(target)
    link_count = max((len(t) for t in link_targets.values()), default=0)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    sheet.append([DB_TO_WORKBOOK.get(c, c) for c in columns] + [f"link{n}" for n in range(1, link_count + 1)])
    rows = 0
    for row in con.execute(f"SELECT {', '.join(columns)} FROM items ORDER BY rowid"):
        item = dict(zip(columns, row))
        if item["id"] in videos:
            continue
        item["era"] = ENGLISH_TO_HEBREW_ERA.get(item["era"], item["era"])
        sheet.append([item[c] for c in columns] + link_targets.get(item["id"], []))
        rows += 1
    workbook.save(args.output)
    print(f"Done -> {args.output}: {rows} items, {link_count} link columns, {len(videos)} videos left out")

    if args.videos_csv:
        with open(args.videos_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, lineterminator="\n")
            writer.writerow(["video_id", "parent_id"])
            writer.writerows(sorted(videos.items()))
        print(f"Done -> {args.videos_csv}: {len(videos)} videos")


if __name__ == "__main__":
    main()
