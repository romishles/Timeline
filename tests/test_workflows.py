"""The two workflows, run step by step on a throwaway origin: uploads that
delete items (with and without videos), failing uploads, the limits, the
undo path, and the Clean up orphans button."""

import json
import re
import unittest

from sandbox import Sandbox

BUILD = "build_timeline_db.yml"
CLEANUP = "cleanup_orphans.yml"

REBUILD = "Rebuild timeline database from data/timeline.xlsx"


def image_files(item_id):
    return [f"assets/images/{size}/{item_id}.jpg" for size in ("thumbnails", "200", "1000", "2000")]


class BuildWorkflow(unittest.TestCase):
    def upload(self, box, change, message="Add files via upload"):
        start = box.origin_head()
        change()
        box.push(message)
        return start, box.workflow(BUILD)

    def assert_nothing_pushed(self, box, run, after):
        self.assertEqual(box.origin_subjects(after), ["Add files via upload"], run.log)

    def test_no_changes(self):
        box = Sandbox()
        start = box.origin_head()
        run = box.workflow(BUILD)
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertIn("No changes; the live database stays as it is.", run.log)
        self.assertEqual(run.outcome("Commit removed video lines"), "success")
        self.assertEqual(run.outcome("Push"), "skipped")
        self.assertEqual(box.origin_head(), start)

    # (4) deleting an item without a video works exactly as before
    def test_deleting_an_item_without_a_video(self):
        box = Sandbox()
        csv_before = box.origin_bytes("data/videos.csv")
        start, run = self.upload(box, lambda: box.delete_items([1]))
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(start), [
            "Add files via upload", REBUILD,
            "Remove leftover files of deleted items: images 1; videos none"])
        self.assertEqual(box.origin_bytes("data/videos.csv"), csv_before)
        for path in image_files(1):
            self.assertFalse(box.origin_has(path), path)
        self.assertNotIn('"1000/1.jpg"', box.origin_show("assets/images/manifest.json"))
        self.assertNotIn("Videos removed with their deleted parent", run.summary)
        self.assertIn("Ids removed by this upload (items in the previous database, not in the new one): 1\n", run.summary)
        self.assertIn("### How to undo the file removal", run.summary)
        self.assertNotIn("How to bring back a deleted item with its video", run.summary)
        self.assertEqual([t for t, _ in run.warnings], ["Pages rebuild not requested"])

    # (1) deleting item 72 removes video 187 with it
    def test_deleting_a_parent_removes_its_video(self):
        box = Sandbox()
        csv_before = box.origin_bytes("data/videos.csv")
        video_manifest = box.origin_show("assets/videos/manifest.json")
        start, run = self.upload(box, lambda: box.delete_items([72]))
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(start), [
            "Add files via upload",
            "Remove videos of deleted items: 187 (parent 72)",
            REBUILD,
            "Remove leftover files of deleted items: images 72; videos 187"])

        # the videos.csv commit changes only that line, and nothing else
        pruned = box.git("log", "-1", "--format=%H", "--grep=^Remove videos", "main", cwd=box.origin).stdout.strip()
        changed = box.git("show", "--format=", "--name-only", pruned, cwd=box.origin).stdout.split()
        self.assertEqual(changed, ["data/videos.csv"])
        self.assertEqual(box.origin_bytes("data/videos.csv"), csv_before.replace(b"187,72\n", b""))
        body = box.git("log", "-1", "--format=%b", pruned, cwd=box.origin).stdout
        self.assertIn("Removed the line '187,72' from data/videos.csv: item 72 was deleted from the workbook.", body)
        self.assertIn("by workflow run 1000.", body)

        # files and manifest entries of the video and of the parent's pictures are gone
        for path in ["assets/videos/187.mp4", "assets/videos/187.jpg"] + image_files(72):
            self.assertFalse(box.origin_has(path), path)
        manifest = json.loads(box.origin_show("assets/videos/manifest.json"))
        self.assertEqual(sorted(e["id"] for e in manifest),
                         sorted(e["id"] for e in json.loads(video_manifest) if e["id"] != 187))
        self.assertNotIn('"200/72.jpg"', box.origin_show("assets/images/manifest.json"))
        self.assertTrue(box.origin_has("video-originals/72.mp4"))  # originals are never removed

        # the summary lists everything, with the undo instructions
        s = run.summary
        self.assertIn("## Videos removed with their deleted parent", s)
        self.assertIn("| 187 | 72 | `187,72` | `assets/videos/187.mp4`, `assets/videos/187.jpg` and its entry", s)
        before, after = map(int, re.search(r"items (\d+) → (\d+)", s).groups())
        self.assertEqual(after, before - 2)
        self.assertIn("Ids removed by this upload (items in the previous database, not in the new one): 72, 187\n", s)
        self.assertIn("### Deleted 6 file(s) and 5 manifest entries", s)
        self.assertIn("- `video-originals/72.mp4`", s)
        leftovers = box.git("rev-parse", "main", cwd=box.origin).stdout.strip()
        self.assertIn(f"Run `git revert {leftovers} {pruned}`, put the item's row back", s)
        self.assertLess(s.index("Videos removed with their deleted parent"), s.index("Timeline database: changes"))

        # (6) undo: revert both cleanup commits and restore the row, in one push
        box.pull()
        box.git("revert", "--no-edit", leftovers)
        box.git("revert", "--no-edit", pruned)
        box.git("checkout", start, "--", "data/timeline.xlsx")
        box.git("commit", "-q", "-m", "Restore item 72")
        box.git("push", "-q", "origin", "HEAD:main")
        undo_start = box.git("rev-parse", "HEAD").stdout.strip()
        run = box.workflow(BUILD)
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(undo_start), [REBUILD])
        self.assertIn("Ids removed by this upload (items in the previous database, not in the new one): none", run.summary)
        for path in ["data/videos.csv", "assets/videos/manifest.json", "assets/images/manifest.json",
                     "assets/timeline_data.db", "data/timeline.xlsx"]:
            self.assertEqual(box.origin_bytes(path), box.origin_bytes(path, ref=start), path)
        for path in ["assets/videos/187.mp4", "assets/videos/187.jpg"] + image_files(72):
            self.assertTrue(box.origin_has(path), path)

    def test_restoring_the_row_without_the_reverts_is_refused(self):
        box = Sandbox()
        start, run = self.upload(box, lambda: box.delete_items([72]))
        self.assertEqual(run.conclusion, "success", run.log)
        box.pull()
        box.git("checkout", start, "--", "data/timeline.xlsx")
        after, run = self.upload(box, lambda: None, "Restore only the row")
        self.assertEqual(run.conclusion, "failure")
        self.assertIn("Item 72 (row", run.log)
        self.assertIn("a primary_image needs its images, but these are missing", run.log)
        self.assertEqual(box.origin_subjects(after), ["Restore only the row"])

    # (2) a deleted parent plus a broken workbook: nothing changes
    def test_deleted_parent_and_invalid_workbook_change_nothing(self):
        box = Sandbox()

        def change():
            box.delete_items([72])
            box.set_cell(2, "era", "Modern")
        start, run = self.upload(box, change)
        self.assertEqual(run.conclusion, "failure")
        self.assertEqual(run.outcome("Build and validate into a temporary database"), "failure")
        self.assertEqual(run.outcome("Commit removed video lines"), "skipped")
        self.assertEqual(run.outcome("Push"), "skipped")
        self.assertIn("era 'Modern' is not one of the six era names", run.log)
        self.assertIn("note: video 187 belongs to item 72, which was deleted", run.log)
        self.assert_nothing_pushed(box, run, start)
        self.assertIn(b"187,72\n", box.origin_bytes("data/videos.csv"))

    # (3) six parents with videos: the shrink limit stops it, allow_shrink lets it through
    def test_many_parents_hit_the_shrink_limit(self):
        box = Sandbox()
        parents = [17, 18, 46, 52, 55, 66]  # videos 163..168
        start, run = self.upload(box, lambda: box.delete_items(parents))
        self.assertEqual(run.conclusion, "failure")
        self.assertIn("The new database would remove 12 items (limit 5)", run.log)
        self.assert_nothing_pushed(box, run, start)

        # run by hand with allow_shrink: the lines and the database go through,
        # the leftover files stay (over the limit) with a warning
        after = box.origin_head()
        run = box.workflow(BUILD, "workflow_dispatch", {"allow_shrink": True})
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(after), [
            "Remove videos of deleted items: 163 (parent 17), 164 (parent 18), 165 (parent 46), "
            "166 (parent 52), 167 (parent 55), 168 (parent 66)", REBUILD])
        csv = box.origin_show("data/videos.csv")
        self.assertFalse(re.search(r"^16[3-8],", csv, re.M))
        self.assertIn("## Leftover files of deleted items: NOT REMOVED (see warning)", run.summary)
        self.assertIn("12 fewer than its previous version", run.summary)
        self.assertIn("### Would have deleted 36 file(s)", run.summary)
        pruned = box.git("log", "-1", "--format=%H", "--grep=^Remove videos", "main", cwd=box.origin).stdout.strip()
        self.assertIn(f"The files were not removed. Run `git revert {pruned}`", run.summary)
        self.assertIn("Leftover files not removed", [t for t, _ in run.warnings])
        self.assertTrue(box.origin_has("assets/videos/163.mp4"))

        # the button removes them: a max_files below 36 refuses, then allow_shrink + 40 works
        before = box.origin_head()
        run = box.workflow(CLEANUP, "workflow_dispatch",
                           {"apply": True, "confirm": "DELETE", "allow_shrink": True, "max_files": 30})
        self.assertEqual(run.conclusion, "failure")
        self.assertIn("36 files would be removed, more than the limit of 30", run.summary)
        self.assertEqual(box.origin_head(), before)
        run = box.workflow(CLEANUP, "workflow_dispatch", {"apply": True, "confirm": "DELETE", "allow_shrink": True})
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(before), [
            "Clean up orphans: images 17, 18, 46, 52, 55, 66; videos 163, 164, 165, 166, 167, 168"])
        self.assertFalse(box.origin_has("assets/videos/163.mp4"))

    # (5) a misspelled or blank parent in videos.csv still stops the build
    def test_misspelled_or_blank_parent_still_errors(self):
        for line, message in (("187,721", "parent 721 is not an item in the workbook, and it was not a picture item"),
                              ("187,", "video_id and parent_id must be whole numbers")):
            with self.subTest(line=line):
                box = Sandbox()

                def change():
                    path = box.path("data/videos.csv")
                    path.write_text(path.read_text().replace("187,72\n", line + "\n"))
                start, run = self.upload(box, change)
                self.assertEqual(run.conclusion, "failure")
                self.assertIn(message, run.log)
                self.assert_nothing_pushed(box, run, start)

    def test_editing_videos_csv_by_hand_still_works(self):
        # Deleting the row and the line yourself, as before: no extra commit.
        box = Sandbox()

        def change():
            box.delete_items([72])
            path = box.path("data/videos.csv")
            path.write_text(path.read_text().replace("187,72\n", ""))
        start, run = self.upload(box, change)
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(start), [
            "Add files via upload", REBUILD, "Remove leftover files of deleted items: images 72; videos 187"])
        self.assertNotIn("Videos removed with their deleted parent", run.summary)

        # Removing only a video line keeps its parent and removes just the video.
        box = Sandbox()

        def change():
            path = box.path("data/videos.csv")
            path.write_text(path.read_text().replace("163,17\n", ""))
        start, run = self.upload(box, change)
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertEqual(box.origin_subjects(start), [
            "Add files via upload", REBUILD, "Remove leftover files of deleted items: images none; videos 163"])
        self.assertTrue(box.origin_has("assets/images/200/17.jpg"))


class CleanupWorkflow(unittest.TestCase):
    def add_orphan_images(self, box, item_id):
        manifest = box.read_json("assets/images/manifest.json")
        for path in image_files(item_id):
            box.path(path).write_text("orphan")
            manifest.append({"path": path.removeprefix("assets/images/"), "width": 200, "height": 200})
        box.write_json("assets/images/manifest.json", manifest)
        box.push("Add an orphan")

    def test_dry_run_then_apply(self):
        box = Sandbox()
        self.add_orphan_images(box, 999)
        start = box.origin_head()

        run = box.workflow(CLEANUP, "workflow_dispatch")
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertIn("## Clean up orphans: DRY RUN (nothing changed)", run.summary)
        self.assertIn("apply was not ticked", run.summary)
        self.assertIn("| image | 999 |", run.summary)
        self.assertIn("- image 13: still an item (tech_event)", run.summary)
        self.assertEqual(run.outcome("Commit and push the removal"), "skipped")
        self.assertEqual(box.origin_head(), start)

        run = box.workflow(CLEANUP, "workflow_dispatch", {"apply": True, "confirm": "delete"})
        self.assertIn("apply was ticked but confirm was not exactly DELETE", run.summary)
        self.assertEqual(box.origin_head(), start)

        run = box.workflow(CLEANUP, "workflow_dispatch", {"apply": True, "confirm": "DELETE", "max_files": 3})
        self.assertEqual(run.conclusion, "failure")
        self.assertIn("4 files would be removed, more than the limit of 3", run.summary)
        self.assertEqual(box.origin_head(), start)

        run = box.workflow(CLEANUP, "workflow_dispatch", {"apply": True, "confirm": "DELETE"})
        self.assertEqual(run.conclusion, "success", run.log)
        self.assertIn("## Clean up orphans: APPLIED", run.summary)
        self.assertEqual(box.origin_subjects(start), ["Clean up orphans: images 999; videos none"])
        for path in image_files(999):
            self.assertFalse(box.origin_has(path))
        self.assertEqual(box.origin_bytes("assets/images/manifest.json"),
                         box.origin_bytes("assets/images/manifest.json", ref=f"{start}~1"))

    def test_refuses_when_the_database_is_out_of_date(self):
        box = Sandbox()
        self.add_orphan_images(box, 999)
        box.set_cell(2, "english_header", "Changed")
        box.push("Upload without a rebuild")
        start = box.origin_head()
        run = box.workflow(CLEANUP, "workflow_dispatch", {"apply": True, "confirm": "DELETE"})
        self.assertEqual(run.conclusion, "failure")
        self.assertIn("is not up to date with data/timeline.xlsx", run.summary)
        self.assertEqual(box.origin_head(), start)


if __name__ == "__main__":
    unittest.main()
