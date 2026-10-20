Timeline database tools

The timeline database (assets/timeline_data.db) is built from two files:
  data/timeline.xlsx  - every item except the videos (Sheet1)
  data/videos.csv     - the videos: video_id,parent_id
A GitHub workflow (.github/workflows/build_timeline_db.yml) runs the build
whenever one of them changes. See data/README.md for how to edit them.

Setup:
    pip install -r xlsx_to_sqlite/requirements.txt

xlsx_to_sqlite.py - build the database from the workbook and the videos list

    python xlsx_to_sqlite/xlsx_to_sqlite.py data/timeline.xlsx new.db \
        --videos data/videos.csv \
        --manifest assets/videos/manifest.json \
        --images assets/images/manifest.json \
        --baseline assets/timeline_data.db

  Checks everything first (columns, ids, links, types, eras, paradigm
  groups, videos and their files, that every picture item has its four
  image files and manifest entries, and with --baseline that the new
  database is not much smaller). A video that is in the manifest but not in
  data/videos.csv is only a warning; Clean up orphans removes it. If
  anything is wrong it lists every problem and writes nothing. Add
  --allow-shrink to accept a much smaller database.

db_to_xlsx.py - export the database back to the workbook format

    python xlsx_to_sqlite/db_to_xlsx.py assets/timeline_data.db timeline.xlsx \
        --manifest assets/videos/manifest.json --videos-csv videos.csv

  Writes the non-video items to a workbook (and, with --videos-csv, the
  videos list). Useful to recover the workbook from the database.

db_diff.py - list the differences between two databases

    python xlsx_to_sqlite/db_diff.py assets/timeline_data.db new.db

  Prints every changed cell (item id and column), added or removed item and
  link. Exit code 0 means the same content, 1 means different.

Run the commands from the repository root.
