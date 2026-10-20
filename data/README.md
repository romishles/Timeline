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

To remove a video, delete its line from `data/videos.csv`. Its files and
manifest entry stay behind until you run **Clean up orphans** (below).

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

## Cleaning up leftover files (the "Clean up orphans" button)

When you delete an item from the Excel file, or a video from `data/videos.csv`,
its picture or video files stay in the repository. Visitors still download the
leftover pictures every time the site loads. The **Clean up orphans** button
finds these leftovers and, only when you confirm, removes them. It never runs
by itself.

**What it removes:** the four picture files (and their entries in
`assets/images/manifest.json`) of an id that is no longer an item, and the
video files (and their entry in `assets/videos/manifest.json`) of a video that
is no longer in `data/videos.csv`. It never touches the Excel file, the
database, `data/videos.csv`, or the original uploads in `video-originals/`
(those are only listed).

**How to run it:**

1. Open the **Actions** tab, click **Clean up orphans** on the left, then
   **Run workflow**.
2. First run it **with "apply" unticked**. This is a dry run: nothing changes.
   Open the run and read its summary. It lists what it found, what it would
   remove, what it keeps and why, and any problems to fix by hand.
3. If the list is what you expect, run it again with **apply ticked** and type
   the word **DELETE** (capital letters) in the **confirm** box. Without the
   exact word it only does a dry run again.
4. It removes at most 40 files per run. If it says there are more, check the
   list carefully, then run it with a higher **max_files**.

It refuses to remove anything if something looks wrong, for example if the
last database rebuild failed. The summary then says why.

Pictures of items that exist but never show a picture (such as the
`tech_event` items 13 and 150) are listed as "unused but kept". The button
never removes files of an existing item.

## Bringing deleted files back

Everything the button removes is still in the repository's history.

1. Open the repository's commit list (the clock icon **Commits** on the main
   page) and find the commit **Clean up orphans: images …; videos …**. Its
   message lists every removed file.
2. The simplest way back is to undo that commit with git:
   `git revert <commit id>`, then push. This restores the files and their
   manifest entries exactly.
3. Without git: open the commit, click **Browse files** on the commit just
   before it, download each file you need, and upload it to the same folder.
   Then add the file's entry back to the manifest. The removed entries are shown
   in red in the commit's changes to `manifest.json`.
4. If the item itself was deleted by mistake, also put its row back in the Excel
   file; otherwise the next cleanup removes the files again. If you put a row
   back before its pictures, the database rebuild stops and names the missing
   pictures, so the site is never broken by it.
