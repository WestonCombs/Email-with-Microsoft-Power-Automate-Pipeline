"""
Force-sync the Git repository with GitHub (or whatever ``origin`` points to).

Runs in the repository root (the folder *above* ``tools`` when this project is
the repo root). Fetches from ``origin`` and resets the current branch hard to
``origin/<branch>``, discarding local commits and uncommitted changes on tracked files.

Usage (from anywhere, if Git is on PATH):

    python tools/git/pull_latest.py

Optional: also remove untracked files and directories (like a clean clone of the branch):

    python tools/git/pull_latest.py --clean
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent


def _run_git(args: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _git_root(start: Path) -> Path:
    r = _run_git(["rev-parse", "--show-toplevel"], start)
    if r.returncode != 0:
        print(
            "ERROR: Not inside a Git repository (or git is not installed / not on PATH).",
            file=sys.stderr,
        )
        if r.stderr.strip():
            print(r.stderr.strip(), file=sys.stderr)
        sys.exit(1)
    return Path(r.stdout.strip())


def _current_branch(repo: Path) -> str:
    r = _run_git(["rev-parse", "--abbrev-ref", "HEAD"], repo)
    if r.returncode != 0:
        print("ERROR: Could not determine current branch.", file=sys.stderr)
        sys.exit(1)
    b = r.stdout.strip()
    if b == "HEAD":
        print(
            "ERROR: Detached HEAD. Checkout a branch first, then run this script.",
            file=sys.stderr,
        )
        sys.exit(1)
    return b


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch origin and hard-reset current branch to match the remote."
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="After reset, run `git clean -fd` to delete untracked files and folders.",
    )
    args = parser.parse_args()

    repo = _git_root(_SCRIPT_DIR)
    branch = _current_branch(repo)

    print(f"Repository: {repo}")
    print(f"Branch:     {branch}")
    print("Fetching from origin ...")
    fetch = subprocess.run(["git", "fetch", "origin"], cwd=str(repo))
    if fetch.returncode != 0:
        print("ERROR: git fetch failed.", file=sys.stderr)
        sys.exit(fetch.returncode)

    remote_ref = f"origin/{branch}"
    print(f"Resetting hard to {remote_ref} ...")
    reset = subprocess.run(["git", "reset", "--hard", remote_ref], cwd=str(repo))
    if reset.returncode != 0:
        print(
            f"ERROR: git reset --hard {remote_ref} failed.\n"
            "  If the branch does not exist on the remote, push it once or merge on GitHub first.",
            file=sys.stderr,
        )
        sys.exit(reset.returncode)

    if args.clean:
        print("Removing untracked files and directories (git clean -fd) ...")
        clean = subprocess.run(["git", "clean", "-fd"], cwd=str(repo))
        if clean.returncode != 0:
            print("ERROR: git clean failed.", file=sys.stderr)
            sys.exit(clean.returncode)

    print("Done. Local tree matches the remote branch (tracked files).")


if __name__ == "__main__":
    main()
