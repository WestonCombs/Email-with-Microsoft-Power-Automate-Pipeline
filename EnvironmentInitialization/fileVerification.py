"""File checks under BASE_DIR. Called from runner.py — paths are relative to project root."""

from __future__ import annotations

from pathlib import Path

from verification_common import ensure_parent_chain_exists, path_under_root, resolved_root, verification_log


def file_verification(
    project_root: Path | str,
    relative_path: str,
    *,
    overwrite: bool = False,
) -> Path:
    """Ensure a file exists at ``BASE_DIR / relative_path``.

    Creates the full parent directory chain if any part is missing, then creates an empty
    UTF-8 file if the file is missing. If the file already exists, leaves it unchanged unless
    *overwrite* is True (then truncated to empty).
    """
    root = resolved_root(project_root)
    path = path_under_root(root, relative_path)

    verification_log(
        f"[FileVerification] {relative_path!r} (overwrite={overwrite}) -> {path}"
    )

    ensure_parent_chain_exists(path)
    verification_log(f"[FileVerification] Ensured parent directories for {path}")

    if path.exists() and path.is_dir():
        raise IsADirectoryError(f"Expected a file, found a directory: {path}")

    if overwrite or not path.exists():
        path.write_text("", encoding="utf-8")
        verification_log(f"[FileVerification] Wrote empty file {path}")
    else:
        verification_log(f"[FileVerification] File already exists, left unchanged {path}")

    verification_log(f"[FileVerification] OK {path}")
    return path


# Backwards-compatible names
def ensure_empty_file(
    project_root: Path | str,
    relative_path: str,
    *,
    overwrite: bool = False,
) -> Path:
    return file_verification(project_root, relative_path, overwrite=overwrite)


def _normalize_extension(extension: str) -> str:
    ext = extension.strip()
    if not ext:
        raise ValueError("extension must be non-empty")
    return ext if ext.startswith(".") else f".{ext}"


def ensure_empty_files_with_extension(
    project_root: Path | str,
    directory_relative: str,
    basenames: list[str],
    extension: str,
    *,
    overwrite: bool = False,
) -> list[Path]:
    """Create empty files under *directory_relative*; same overwrite rules as :func:`file_verification`."""
    ext = _normalize_extension(extension)
    out: list[Path] = []
    dr = directory_relative.strip().strip("/").strip("\\")
    for name in basenames:
        stem = name.strip()
        if not stem:
            continue
        rel = f"{dr}/{stem}{ext}" if dr else f"{stem}{ext}"
        out.append(file_verification(project_root, rel, overwrite=overwrite))
    return out
