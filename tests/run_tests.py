"""
Runs every test in this folder:

    python3 tests/run_tests.py          (add -v to list each test)

The tests work only in throwaway copies in a temporary folder. As a guard,
this script also checks that the repository itself is unchanged afterwards.
"""

import hashlib
import subprocess
import sys
import unittest
from pathlib import Path

sys.dont_write_bytecode = True  # no __pycache__ folders in the repository
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import sandbox  # noqa: E402


def repository_state():
    """git status plus a hash of every changed and untracked file's content."""
    status = subprocess.run(["git", "status", "--porcelain=v1", "-z", "--untracked-files=all"],
                            cwd=sandbox.REPO, capture_output=True, check=True).stdout
    digest = hashlib.sha256(status)
    for entry in filter(None, status.split(b"\0")):
        path = sandbox.REPO / entry[3:].decode()
        if path.is_file():
            digest.update(path.read_bytes())
    head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=sandbox.REPO, capture_output=True, check=True).stdout
    return head + digest.digest()


def main():
    before = repository_state()
    suite = unittest.defaultTestLoader.discover(str(HERE), pattern="test_*.py", top_level_dir=str(HERE))
    verbosity = 2 if "-v" in sys.argv[1:] else 1
    try:
        result = unittest.TextTestRunner(verbosity=verbosity).run(suite)
    finally:
        sandbox.cleanup()
    if repository_state() != before:
        print("ERROR: the repository changed while the tests ran; they must only touch temporary copies.")
        return 1
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(main())
