# Tests

Run every test from the repository root with one command:

```
python3 tests/run_tests.py
```

Add `-v` to list each test. They need only what the build needs
(`pip install -r xlsx_to_sqlite/requirements.txt`), plus `git` and `bash`.
They take about a minute.

## What they do

Every test works in its own temporary copy of the repository, with a
temporary bare repository standing in for GitHub. Your files are never
changed: the script checks that the repository is the same before and
after. Large picture and video files are copied as small placeholders,
because the scripts only check that they exist.

- `test_build.py`: the converter on its own. It checks that the live files
  rebuild a byte-identical database. It checks that each kind of broken upload
  is refused with its message and that nothing is written. It also covers the
  shrink limits and the removal of videos whose parent was deleted.
- `test_workflows.py`: the two workflows in `.github/workflows/`, run step
  by step the way GitHub runs them (their `if:`, `env:` and `${{ }}`
  expressions included). It uploads changes, checks the commits pushed and
  the run summary, tests the limits and the undo path, and covers the
  **Clean up orphans** dry run and apply.
- `sandbox.py`: the temporary copies and the small workflow runner.

What they can't check: GitHub itself. That includes the checkout and setup
actions, whether a push made with `GITHUB_TOKEN` starts a new run (GitHub
says it doesn't), and GitHub Pages. Set `KEEP_SANDBOX=1` to keep the
temporary folders for a look after a run.
