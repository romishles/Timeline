# Editing the timeline

The timeline's content lives in **`data/timeline.xlsx`**. When you upload a new
version of it, GitHub rebuilds the site's database automatically and the site
updates by itself. Never edit `assets/timeline_data.db` directly: the next
rebuild replaces it.

## How to edit

1. Open the latest `data/timeline.xlsx` on GitHub and click **Download raw file**.
   Always start from the latest version, so you don't undo someone else's change.
2. Edit it in Excel. One row is one item. You can change any text, add rows,
   and change the links.
3. Save it as `.xlsx`, keeping the name `timeline.xlsx` and the sheet name `Sheet1`.

### Columns

`id`, `era`, `year`, `type`, `hebrew_header`, `english_header`, `about_item`,
`about_creator`, `text_historical`, `text_term`, `paradigm`, `paradigm_main`,
`paradigm_group`, `link1`, `link2`, `link3`

- **Never rename these headers**, and don't add other columns. A renamed or
  unknown header stops the build. You may add more link columns: `link4`, `link5`, …
- `id` must be a unique whole number. Don't reuse the ids 163–187 (videos).
- `year` is a whole number. Make sure Excel doesn't turn it into a date.
- `type` is one of: `primary_image`, `secondary_image`, `historical_event`, `tech_event`.
- `era` is one of the six era names, exactly as written in the existing rows.
- Every `primary_image` needs a `paradigm_group` from the 11 existing groups.
- `link1`, `link2`, … hold the ids of the items this row points to.
- A new picture also needs its image files in `assets/images/`.

### Videos are not in the Excel file

The videos (ids 163–187) are added by the build. Adding a video needs three things:

1. Its files: `assets/videos/<video id>.mp4` and `assets/videos/<video id>.jpg`.
2. An entry in `assets/videos/manifest.json`.
3. A line `<video id>,<parent item id>` in `data/videos.csv`.

The parent must be a `primary_image`. If a link cell in the Excel file points
to a video, it is skipped with a note; that is fine.

## How to upload

1. On GitHub, open the `data` folder and choose **Add file → Upload files**.
2. Drop your `timeline.xlsx` and choose **Commit changes** (directly to `main`).

## How long it takes

About a minute for the database to rebuild, then a few more minutes for the
site to update. Reload the site afterwards; you may need a hard refresh.

## How to see if it worked

Open the repository's **Actions** tab and click the newest
**Build timeline database** run.

- **Green check:** it worked. The run page lists every changed item and column.
  If nothing changed, it says "no changes".
- **Red cross:** nothing was changed on the site. Click the run, then the
  failed step **Build and validate into a temporary database**. It lists each
  problem with the item id, for example
  `Item 12 (row 13): era 'Modern' is not one of the six era names`.
  Fix the file and upload it again.

If a run says the new database "would remove" many items or links, that is a
safety check against uploading a wrong or cut-off file. If you really meant to
delete them, open **Actions → Build timeline database → Run workflow**, tick
**allow_shrink**, and run it.

## How to roll back

1. Open `data/timeline.xlsx` on GitHub and click **History**.
2. Open the last good version, click **View file**, then **Download raw file**.
3. Upload that file again as described above. The database is rebuilt from it.
