"""Folder checks under BASE_DIR. Called from runner.py — paths are relative to project root."""

from __future__ import annotations

import shutil
from pathlib import Path

from verification_common import ensure_parent_chain_exists, path_under_root, resolved_root, verification_log


def folder_verification(
    project_root: Path | str,
    relative_path: str,
    *,
    clear_if_exists: bool = False,
) -> Path:
    """Ensure a directory exists at ``BASE_DIR / relative_path``.

    Missing parent folders are created as needed.

    *clear_if_exists* (default False): if the directory already exists, only create missing
    parents — existing contents are left alone. If True, any existing directory is emptied
    (files removed, subfolders removed); use that for scratch dirs like attachments.
    """
    root = resolved_root(project_root)
    target = path_under_root(root, relative_path)

    verification_log(
        f"[FolderVerification] {relative_path!r} (clear_if_exists={clear_if_exists}) "
        f"-> {target}"
    )

    ensure_parent_chain_exists(target)

    if target.exists():
        if not target.is_dir():
            raise NotADirectoryError(f"Expected a directory, found a file: {target}")
        if clear_if_exists:
            for item in target.iterdir():
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
            verification_log(f"[FolderVerification] Cleared contents of {target}")
    else:
        target.mkdir(parents=True, exist_ok=True)
        verification_log(f"[FolderVerification] Created directory {target}")

    verification_log(f"[FolderVerification] OK {target}")
    return target
