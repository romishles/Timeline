"""The converter on its own: the live files, every rejection, the shrink limit
and --prune-videos."""

import datetime
import shutil
import sqlite3
import unittest

from sandbox import Sandbox


def counts(path):
    """(items, links) in a database."""
    items, links = items_and_links(path)
    return len(items), len(links)


def items_and_links(path):
    con = sqlite3.connect(path)
    try:
        return con.execute("SELECT * FROM items").fetchall(), con.execute("SELECT * FROM links").fetchall()
    finally:
        con.close()


class LiveFiles(unittest.TestCase):
    def test_live_files_rebuild_a_byte_identical_database(self):
        box = Sandbox()
        result = box.build()
        self.assertEqual(result.returncode, 0, result.stderr)
        items, links = counts(box.path("assets/timeline_data.db"))
        self.assertIn(f"items: {items} -> {items}, links: {links} -> {links}", result.stdout)
        self.assertEqual(box.path("new.db").read_bytes(), box.path("assets/timeline_data.db").read_bytes())

    def test_prune_videos_changes_nothing_when_no_parent_was_deleted(self):
        box = Sandbox()
        before = box.path("data/videos.csv").read_bytes()
        result = box.build("--prune-videos", "--prune-summary", "s.md", "--prune-commit-message", "m.txt")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(box.path("new.db").read_bytes(), box.path("assets/timeline_data.db").read_bytes())
        self.assertEqual(box.path("data/videos.csv").read_bytes(), before)
        self.assertFalse(box.path("s.md").exists() or box.path("m.txt").exists())

    def test_prune_videos_needs_a_baseline(self):
        box = Sandbox()
        result = box.build("--prune-videos", baseline=False)
        self.assertEqual(result.returncode, 2)
        self.assertIn("--prune-videos needs --baseline", result.stderr)


class Rejections(unittest.TestCase):
    """Each broken upload is refused with a clear message and nothing is
    written, both as a plain build and with --prune-videos (as the workflow
    runs it), which must not touch the videos list either."""

    def assert_rejected(self, box, *messages, plain=None, prune=True):
        """plain: the messages expected without --prune-videos, if different.
        prune=False checks only the plain build."""
        csv_before = box.path("data/videos.csv").read_bytes()
        modes = [("plain", [], plain or messages)]
        if prune:
            modes.append(("prune", ["--prune-videos", "--prune-commit-message", "m.txt"], messages))
        for mode, extra, expected in modes:
            with self.subTest(mode=mode):
                result = box.build(*extra)
                self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
                self.assertIn("Nothing was written", result.stderr)
                for message in expected:
                    self.assertIn(message, result.stderr)
                self.assertFalse(box.path("new.db").exists())
                self.assertFalse(box.path("m.txt").exists())
                self.assertEqual(box.path("data/videos.csv").read_bytes(), csv_before)
        return result

    def test_missing_required_column(self):
        box = Sandbox()
        box.edit_workbook(lambda sheet: sheet.delete_cols(2))  # era
        self.assert_rejected(box, "Missing required column 'era'")

    def test_unknown_column(self):
        box = Sandbox()
        box.edit_workbook(lambda sheet: sheet.cell(1, sheet.max_column + 1, "colour"))
        self.assert_rejected(box, "Unknown column 'colour'")

    def test_dangling_link(self):
        box = Sandbox()
        box.set_cell(2, "link1", 9999)
        self.assert_rejected(box, "Item 2: link1 points to 9999, which is not an item in the workbook")

    def test_link_to_a_deleted_item(self):
        box = Sandbox()
        box.set_cell(2, "link1", 72)
        box.delete_items([72], clear_links=False)
        self.assert_rejected(box, "Item 2: link1 points to 72, which is not an item in the workbook")

    def test_duplicate_id(self):
        box = Sandbox()
        box.set_cell(3, "id", 2)
        self.assert_rejected(box, "Item 2: duplicate id")

    def test_unknown_era(self):
        box = Sandbox()
        box.set_cell(2, "era", "Modern")
        self.assert_rejected(box, "era 'Modern' is not one of the six era names")

    def test_unknown_type(self):
        box = Sandbox()
        box.set_cell(2, "type", "photo")
        self.assert_rejected(box, "type 'photo' is not one of")

    def test_missing_paradigm_group(self):
        box = Sandbox()
        box.set_cell(72, "paradigm_group", None)
        self.assert_rejected(box, "Item 72 (row", "primary_image needs a paradigm_group")

    def test_year_turned_into_a_date(self):
        box = Sandbox()
        box.set_cell(2, "year", datetime.datetime(1928, 1, 1))
        self.assert_rejected(box, "Excel turned it into a date")

    def test_item_id_clashes_with_a_video_id(self):
        box = Sandbox()
        box.set_cell(2, "id", 163)
        self.assert_rejected(box, "id 163 is also a video id in the videos list")

    def test_video_missing_from_the_manifest(self):
        box = Sandbox()
        box.write_json("assets/videos/manifest.json",
                       [e for e in box.read_json("assets/videos/manifest.json") if e["id"] != 187])
        self.assert_rejected(box, "Video 187 is in the videos list but not in assets/videos/manifest.json")

    def test_missing_video_file(self):
        box = Sandbox()
        box.path("assets/videos/187.mp4").unlink()
        self.assert_rejected(box, "Video 187: file assets/videos/187.mp4 is missing")

    def test_video_parent_not_a_primary_image(self):
        box = Sandbox()
        csv = box.path("data/videos.csv")
        csv.write_text(csv.read_text().replace("187,72\n", "187,13\n"))
        self.assert_rejected(box, "Video 187: parent 13 is a tech_event, not a primary_image")

    def test_video_parent_misspelled(self):
        # 721 was never an item, so it is a typo, not a deleted parent.
        box = Sandbox()
        csv = box.path("data/videos.csv")
        csv.write_text(csv.read_text().replace("187,72\n", "187,721\n"))
        self.assert_rejected(box, "Video 187: parent 721 is not an item in the workbook, and it was not a "
                                  "picture item before this upload either", "for a typo",
                             plain=["Video 187: parent 721 is not an item in the workbook"])

    def test_video_parent_missing_without_pruning(self):
        box = Sandbox()
        box.delete_items([72])
        self.assert_rejected(box, "Video 187: parent 72 is not an item in the workbook. If you deleted item 72 "
                                  "on purpose, also delete the line '187,72'", prune=False)

    def test_video_parent_blank_or_not_a_number(self):
        for bad in ("187,", "187,72a", "187, "):
            with self.subTest(line=bad):
                box = Sandbox()
                csv = box.path("data/videos.csv")
                csv.write_text(csv.read_text().replace("187,72\n", bad + "\n"))
                self.assert_rejected(box, "video_id and parent_id must be whole numbers")

    def test_missing_image_file_and_manifest_entry(self):
        box = Sandbox()
        box.path("assets/images/200/72.jpg").unlink()
        box.write_json("assets/images/manifest.json",
                       [e for e in box.read_json("assets/images/manifest.json") if e["path"] != "1000/72.jpg"])
        self.assert_rejected(box, "Item 72 (row", "200/72.jpg (file)", "1000/72.jpg (manifest entry)")

    def test_deleted_parent_with_another_problem_prunes_nothing(self):
        box = Sandbox()
        box.delete_items([72])
        box.set_cell(2, "era", "Modern")
        result = self.assert_rejected(box, "era 'Modern' is not one of the six era names",
                                      plain=["era 'Modern'", "parent 72 is not an item"])
        self.assertIn("note: video 187 belongs to item 72, which was deleted; it will be removed automatically "
                      "once the problems above are fixed", result.stderr)
        self.assertNotIn("parent 72 is not an item", result.stderr)


class ShrinkLimit(unittest.TestCase):
    def test_more_than_five_items_removed_is_refused_unless_allowed(self):
        box = Sandbox()
        box.delete_items([1, 2, 3, 4, 5, 6])
        result = box.build()
        self.assertEqual(result.returncode, 1)
        self.assertIn("The new database would remove 6 items (limit 5)", result.stderr)
        self.assertIn("--allow-shrink", result.stderr)
        result = box.build("--allow-shrink")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("allowed by --allow-shrink", result.stdout)

    def test_five_items_removed_is_within_the_limit(self):
        box = Sandbox()
        box.delete_items([1, 2, 3, 4, 5])
        result = box.build()
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_more_than_ten_percent_of_links_removed_is_refused(self):
        box = Sandbox()
        baseline = box.path("baseline.db")
        shutil.copyfile(box.path("assets/timeline_data.db"), baseline)
        with sqlite3.connect(baseline) as con:  # a baseline with 20 more links than the workbook makes
            con.executemany("INSERT INTO links VALUES (?, ?)", [(2, 3)] * 20)
        con.close()
        result = box.build("--baseline", "baseline.db", baseline=False)
        self.assertEqual(result.returncode, 1)
        limit = int(0.10 * (counts(box.path("assets/timeline_data.db"))[1] + 20))
        self.assertIn(f"would remove 20 links (limit 10% = {limit})", result.stderr)

    def test_a_deleted_parent_and_its_video_count_as_two_items(self):
        box = Sandbox()
        box.delete_items([17, 18, 46])  # videos 163, 164, 165
        result = box.build("--prune-videos")
        self.assertEqual(result.returncode, 1)
        self.assertIn("would remove 6 items (limit 5)", result.stderr)
        self.assertIn("187,72", box.path("data/videos.csv").read_text())
        self.assertIn("163,17", box.path("data/videos.csv").read_text())


class PruneVideos(unittest.TestCase):
    def test_deleted_parent_removes_only_its_line(self):
        box = Sandbox()
        before = box.path("data/videos.csv").read_bytes()
        old_items, old_links = counts(box.path("assets/timeline_data.db"))
        box.delete_items([72])
        result = box.build("--prune-videos", "--prune-summary", "s.md", "--prune-commit-message", "m.txt")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Removed video 187 from data/videos.csv: its parent 72 was deleted", result.stdout)
        self.assertNotIn("warning: video 187", result.stdout)
        self.assertEqual(box.path("data/videos.csv").read_bytes(), before.replace(b"187,72\n", b""))
        items, links = items_and_links(box.path("new.db"))
        self.assertEqual(len(items), old_items - 2)
        self.assertNotIn(72, [i[0] for i in items])
        self.assertNotIn(187, [i[0] for i in items])
        self.assertEqual(len(links), old_links - 2)
        self.assertTrue(box.path("m.txt").read_text().startswith("Remove videos of deleted items: 187 (parent 72)\n"))
        summary = box.path("s.md").read_text()
        self.assertIn("## Videos removed with their deleted parent", summary)
        self.assertIn("| 187 | 72 | `187,72` | `assets/videos/187.mp4`, `assets/videos/187.jpg`", summary)

    def test_line_endings_and_bom_are_kept(self):
        box = Sandbox()
        csv = box.path("data/videos.csv")
        csv.write_bytes(b"\xef\xbb\xbf" + csv.read_bytes().replace(b"\n", b"\r\n"))
        before = csv.read_bytes()
        box.delete_items([72, 17])
        result = box.build("--prune-videos")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(csv.read_bytes(), before.replace(b"163,17\r\n", b"").replace(b"187,72\r\n", b""))

    def test_parent_that_was_not_a_picture_is_not_pruned(self):
        # A line pointing at a former tech_event was already wrong; it stays an error.
        box = Sandbox()
        csv = box.path("data/videos.csv")
        csv.write_text(csv.read_text().replace("187,72\n", "187,13\n"))
        box.delete_items([13])
        result = box.build("--prune-videos")
        self.assertEqual(result.returncode, 1)
        self.assertIn("parent 13 is not an item in the workbook, and it was not a picture item", result.stderr)


if __name__ == "__main__":
    unittest.main()
