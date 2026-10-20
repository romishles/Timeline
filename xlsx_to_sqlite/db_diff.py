"""
db_diff.py

Compares two timeline databases and prints a Markdown report of every
difference: changed cells (item id and column), added or removed items, and
added or removed links. Used by the build workflow for its job summary.

Exit code: 0 if the databases have the same content, 1 if they differ.

Usage:
    python db_diff.py OLD_DB NEW_DB
"""

import sqlite3
import sys
from collections import Counter


def load(path):
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    schema = sorted(con.execute("SELECT type, name, sql FROM sqlite_master"))
    columns = [r[1] for r in con.execute("PRAGMA table_info(items)")]
    items = {row[0]: dict(zip(columns, row)) for row in con.execute("SELECT * FROM items ORDER BY rowid")}
    item_order = list(items)
    links = list(con.execute("SELECT id, target_id FROM links ORDER BY rowid"))
    return schema, columns, items, item_order, links


def short(value, limit=60):
    if value is None:
        return "NULL"
    text = repr(value)
    return text if len(text) <= limit else text[:limit - 3] + "..."


def cell(text):
    return str(text).replace("|", "\\|").replace("\n", "\\n")


def main():
    old_path, new_path = sys.argv[1], sys.argv[2]
    old_schema, old_columns, old_items, old_order, old_links = load(old_path)
    new_schema, new_columns, new_items, new_order, new_links = load(new_path)

    lines = []
    if old_schema != new_schema:
        lines += ["### Schema changed", "", "The table definitions differ.", ""]

    changed = []
    for item_id in sorted(set(old_items) & set(new_items)):
        for column in new_columns:
            old, new = old_items[item_id].get(column), new_items[item_id].get(column)
            if old != new:
                changed.append((item_id, column, old, new))
    added = sorted(set(new_items) - set(old_items))
    removed = sorted(set(old_items) - set(new_items))
    links_added = Counter(new_links) - Counter(old_links)
    links_removed = Counter(old_links) - Counter(new_links)
    same_order = [i for i in old_order if i in new_items] == [i for i in new_order if i in old_items]
    links_reordered = not links_added and not links_removed and old_links != new_links

    if changed:
        lines += [f"### {len(changed)} changed cell(s) in {len({c[0] for c in changed})} item(s)", "",
                  "| item id | column | before | after |", "|---|---|---|---|"]
        lines += [f"| {i} | {c} | {cell(short(o))} | {cell(short(n))} |" for i, c, o, n in changed]
        lines.append("")
    if added:
        lines += [f"### {len(added)} item(s) added", "", ", ".join(map(str, added)), ""]
    if removed:
        lines += [f"### {len(removed)} item(s) removed", "", ", ".join(map(str, removed)), ""]
    if links_added:
        lines += [f"### {sum(links_added.values())} link(s) added", "",
                  ", ".join(f"{a} → {b}" for a, b in sorted(links_added.elements())), ""]
    if links_removed:
        lines += [f"### {sum(links_removed.values())} link(s) removed", "",
                  ", ".join(f"{a} → {b}" for a, b in sorted(links_removed.elements())), ""]
    if not same_order:
        lines += ["### Item order changed", ""]
    if links_reordered:
        lines += ["### Link order changed (same links)", ""]

    differs = bool(lines)
    summary = (f"## Timeline database: {'changes' if differs else 'no changes'}\n\n"
               f"items {len(old_items)} → {len(new_items)}, links {len(old_links)} → {len(new_links)}\n")
    print(summary)
    print("\n".join(lines) if differs else "The new database has the same content as the current one.")
    return 1 if differs else 0


if __name__ == "__main__":
    sys.exit(main())
