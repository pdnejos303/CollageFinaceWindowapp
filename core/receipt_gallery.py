"""Pure-Python backend services for scanning, filtering, and managing receipt images."""

from dataclasses import dataclass
import datetime
import os
import re
import shutil
from typing import Any, Dict, List, Optional

from core.database import DatabaseManager

SUPPORTED_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")

# Regex patterns for extracting receipt numbers from exported image filenames.
_RECEIPT_PREFIX_RE = re.compile(
    r"(?:reprint_)?receipt[_\-\s]+([^\s_]+)",
    re.IGNORECASE,
)
_CODE_PREFIX_RE = re.compile(
    r"\b((?:RCP|RC|REC|INV)[-_]?\d[A-Za-z0-9\-]*)",
    re.IGNORECASE,
)
_LEADING_NUMERIC_RE = re.compile(
    r"^(\d{4,}(?:-\d+)*)(?:[_\s]|$)"
)
_INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class ReceiptImageRecord:
    """Metadata for a single receipt image file enriched with database info."""

    file_path: str
    filename: str
    folder: str
    rel_path: str
    ext: str
    size_bytes: int
    modified_ts: float
    modified_date_str: str
    receipt_no: str = ""
    student_name: str = ""
    class_group: str = ""
    year_level: str = ""
    receipt_date: str = ""
    net_total_amount: float = 0.0
    status: str = ""


def extract_receipt_no_from_filename(filename: str) -> str:
    """Extract a receipt number from common receipt image filename patterns."""
    if not filename:
        return ""
    base = os.path.basename(str(filename).strip())
    stem, _ = os.path.splitext(base)
    if not stem:
        return ""

    match = _RECEIPT_PREFIX_RE.search(stem)
    if match:
        return match.group(1).strip("-_ ")

    match = _CODE_PREFIX_RE.search(stem)
    if match:
        return match.group(1).strip("-_ ")

    match = _LEADING_NUMERIC_RE.match(stem)
    if match:
        return match.group(1).strip("-_ ")

    return ""


def _fetch_receipt_metadata_map(
    db: Optional[DatabaseManager],
    receipt_nos: List[str],
) -> Dict[str, Dict[str, Any]]:
    """Bulk-load receipt metadata from the database keyed by receipt_no."""
    if db is None or not receipt_nos:
        return {}

    unique_nos = sorted({no for no in receipt_nos if no})
    if not unique_nos:
        return {}

    metadata: Dict[str, Dict[str, Any]] = {}
    batch_size = 500
    try:
        with db._connect() as conn:
            for start in range(0, len(unique_nos), batch_size):
                batch = unique_nos[start : start + batch_size]
                placeholders = ", ".join("?" for _ in batch)
                query = (
                    "SELECT receipt_no, snapshot_student_name, snapshot_class_group, "
                    "snapshot_education_level, receipt_date, net_total_amount, status "
                    f"FROM receipts WHERE receipt_no IN ({placeholders})"
                )
                for row in conn.execute(query, tuple(batch)).fetchall():
                    rno = str(row["receipt_no"] or "")
                    if rno:
                        metadata[rno] = dict(row)
    except Exception:
        return {}

    return metadata


def _iter_image_files(root_real: str, recursive: bool) -> List[str]:
    """Safely collect supported image file paths under root_real."""
    found: List[str] = []
    if recursive:
        for dirpath, _dirnames, filenames in os.walk(root_real, onerror=lambda _err: None):
            for fname in filenames:
                if os.path.splitext(fname)[1].lower() in SUPPORTED_IMAGE_EXTS:
                    found.append(os.path.join(dirpath, fname))
        return found

    try:
        with os.scandir(root_real) as entries:
            for entry in entries:
                try:
                    if entry.is_file() and os.path.splitext(entry.name)[1].lower() in SUPPORTED_IMAGE_EXTS:
                        found.append(entry.path)
                except OSError:
                    continue
    except OSError:
        return []

    return found


def scan_receipt_images(
    folder: str,
    db: Optional[DatabaseManager] = None,
    recursive: bool = True,
) -> List[ReceiptImageRecord]:
    """Scan a folder for supported receipt images and enrich them from the database."""
    if not folder or not os.path.isdir(folder):
        return []

    root_real = os.path.realpath(folder)
    image_paths = _iter_image_files(root_real, recursive=recursive)
    if not image_paths:
        return []

    records: List[ReceiptImageRecord] = []
    for full_path in image_paths:
        try:
            st = os.stat(full_path)
        except OSError:
            continue

        filename = os.path.basename(full_path)
        ext = os.path.splitext(filename)[1].lower()
        modified_ts = float(st.st_mtime)
        modified_date_str = datetime.datetime.fromtimestamp(modified_ts).strftime("%Y-%m-%d")
        rel_path = os.path.relpath(full_path, root_real)
        receipt_no = extract_receipt_no_from_filename(filename)

        records.append(
            ReceiptImageRecord(
                file_path=full_path,
                filename=filename,
                folder=os.path.dirname(full_path),
                rel_path=rel_path,
                ext=ext,
                size_bytes=int(st.st_size),
                modified_ts=modified_ts,
                modified_date_str=modified_date_str,
                receipt_no=receipt_no,
            )
        )

    metadata_map = _fetch_receipt_metadata_map(
        db,
        [rec.receipt_no for rec in records if rec.receipt_no],
    )
    if metadata_map:
        for rec in records:
            info = metadata_map.get(rec.receipt_no)
            if not info:
                continue
            rec.student_name = str(info.get("snapshot_student_name") or "")
            rec.class_group = str(info.get("snapshot_class_group") or "")
            rec.year_level = str(info.get("snapshot_education_level") or "")
            rec.receipt_date = str(info.get("receipt_date") or "")
            rec.net_total_amount = float(info.get("net_total_amount") or 0.0)
            rec.status = str(info.get("status") or "")

    records.sort(key=lambda r: (r.modified_ts, r.filename), reverse=True)
    return records


def get_gallery_filter_options(records: List[ReceiptImageRecord]) -> Dict[str, List[str]]:
    """Return sorted unique class_groups and year_levels for gallery dropdowns."""
    class_groups = sorted({rec.class_group.strip() for rec in records if rec.class_group and rec.class_group.strip()})
    year_levels = sorted({rec.year_level.strip() for rec in records if rec.year_level and rec.year_level.strip()})
    return {
        "class_groups": class_groups,
        "year_levels": year_levels,
    }


def filter_receipt_images(
    records: List[ReceiptImageRecord],
    query: str = "",
    class_group: str = "",
    year_level: str = "",
    date_str: str = "",
) -> List[ReceiptImageRecord]:
    """Filter scanned receipt image records by text query, class, year level, and date."""
    q = (query or "").strip().lower()
    cg = (class_group or "").strip()
    yl = (year_level or "").strip()
    ds = (date_str or "").strip().lower()

    filtered: List[ReceiptImageRecord] = []
    for rec in records:
        if cg and rec.class_group.strip() != cg:
            continue
        if yl and rec.year_level.strip() != yl:
            continue
        if ds:
            r_date = (rec.receipt_date or "").lower()
            m_date = (rec.modified_date_str or "").lower()
            if not (r_date.startswith(ds) or m_date.startswith(ds)):
                continue
        if q:
            haystack = " ".join(
                [
                    rec.receipt_no,
                    rec.student_name,
                    rec.class_group,
                    rec.year_level,
                    rec.receipt_date,
                    rec.filename,
                    rec.rel_path,
                ]
            ).lower()
            if q not in haystack:
                continue
        filtered.append(rec)

    return filtered


def _resolve_within_root(
    path: str,
    allowed_root: Optional[str],
    label: str = "path",
) -> str:
    """Resolve path with realpath and enforce allowed_root containment if provided."""
    if not path or not str(path).strip():
        raise ValueError(f"Invalid {label}: path cannot be empty")

    real_path = os.path.realpath(path)
    if allowed_root is not None and str(allowed_root).strip():
        real_root = os.path.realpath(allowed_root)
        try:
            common = os.path.commonpath([real_root, real_path])
        except ValueError as exc:
            raise ValueError(
                f"Path traversal detected: {path} is outside allowed root {allowed_root}"
            ) from exc
        if common != real_root:
            raise ValueError(
                f"Path traversal detected: {path} is outside allowed root {allowed_root}"
            )

    return real_path


def _resolve_unique_dest_path(dest_path: str, overwrite: bool) -> str:
    """Return dest_path if overwrite is True or unused, otherwise append a numeric suffix."""
    if overwrite or not os.path.exists(dest_path):
        return dest_path

    stem, ext = os.path.splitext(dest_path)
    counter = 1
    while True:
        candidate = f"{stem}_{counter}{ext}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def copy_receipt_image(
    src_path: str,
    dest_path: str,
    overwrite: bool = False,
) -> str:
    """Safely copy a receipt image to dest_path, handling collisions and same-file cases."""
    if not src_path or not str(src_path).strip():
        raise ValueError("Source path cannot be empty")
    if not dest_path or not str(dest_path).strip():
        raise ValueError("Destination path cannot be empty")

    src_real = os.path.realpath(src_path)
    if not os.path.isfile(src_real):
        raise FileNotFoundError(f"Source image not found: {src_path}")

    dest_real = os.path.realpath(dest_path)
    if os.path.isdir(dest_real):
        dest_real = os.path.join(dest_real, os.path.basename(src_real))

    if src_real == dest_real:
        if overwrite:
            return src_real
        dest_real = _resolve_unique_dest_path(dest_real, overwrite=False)
    else:
        dest_real = _resolve_unique_dest_path(dest_real, overwrite=overwrite)

    parent_dir = os.path.dirname(dest_real)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)

    shutil.copy2(src_real, dest_real)
    return os.path.realpath(dest_real)


def move_receipt_image(
    src_path: str,
    dest_dir: str,
    overwrite: bool = False,
    allowed_root: Optional[str] = None,
) -> str:
    """Safely move a receipt image into dest_dir with optional allowed_root traversal guard."""
    src_real = _resolve_within_root(src_path, allowed_root, "src_path")
    if not os.path.isfile(src_real):
        raise FileNotFoundError(f"Source image not found: {src_path}")

    dest_dir_real = _resolve_within_root(dest_dir, allowed_root, "dest_dir")
    if os.path.exists(dest_dir_real) and not os.path.isdir(dest_dir_real):
        raise ValueError(f"Destination is not a directory: {dest_dir}")
    os.makedirs(dest_dir_real, exist_ok=True)

    target_path = os.path.join(dest_dir_real, os.path.basename(src_real))
    if os.path.realpath(target_path) == src_real:
        return src_real

    target_path = _resolve_unique_dest_path(target_path, overwrite=overwrite)
    if overwrite and os.path.exists(target_path):
        os.remove(target_path)

    shutil.move(src_real, target_path)
    return os.path.realpath(target_path)


def rename_receipt_image(
    src_path: str,
    new_stem: str,
    allowed_root: Optional[str] = None,
) -> str:
    """Safely rename a receipt image using a sanitized stem while keeping its extension."""
    src_real = _resolve_within_root(src_path, allowed_root, "src_path")
    if not os.path.isfile(src_real):
        raise FileNotFoundError(f"Source image not found: {src_path}")

    raw_stem = str(new_stem or "").strip()
    meaningful = re.sub(r'[<>:"/\\|?*\x00-\x1f.\s]', "", raw_stem)
    clean_stem = _INVALID_FILENAME_CHARS_RE.sub("_", raw_stem).strip(" .")
    if not meaningful or not clean_stem or clean_stem in {".", ".."}:
        raise ValueError(f"Invalid new filename stem: {new_stem!r}")

    _, ext = os.path.splitext(src_real)
    target_path = os.path.join(os.path.dirname(src_real), f"{clean_stem}{ext}")
    if os.path.realpath(target_path) == src_real:
        return src_real

    target_path = _resolve_unique_dest_path(target_path, overwrite=False)
    os.replace(src_real, target_path)
    return os.path.realpath(target_path)


def delete_receipt_image(
    file_path: str,
    allowed_root: Optional[str] = None,
) -> bool:
    """Safely delete a supported receipt image file within allowed_root."""
    real_path = _resolve_within_root(file_path, allowed_root, "file_path")
    if not os.path.exists(real_path):
        return False
    if not os.path.isfile(real_path):
        raise ValueError(f"Target path is not a file: {file_path}")

    ext = os.path.splitext(real_path)[1].lower()
    if ext not in SUPPORTED_IMAGE_EXTS:
        raise ValueError(f"Refusing to delete unsupported file type: {ext or '<none>'}")

    os.remove(real_path)
    return True


def migrate_receipt_folder(
    src_dir: str,
    dest_dir: str,
    move_files: bool = False,
) -> Dict[str, Any]:
    """Copy or move all supported receipt images from src_dir to dest_dir."""
    if not src_dir or not os.path.isdir(src_dir):
        return {
            "processed": 0,
            "skipped": 0,
            "errors": [f"Source directory does not exist: {src_dir}"],
        }
    if not dest_dir or not str(dest_dir).strip():
        return {
            "processed": 0,
            "skipped": 0,
            "errors": ["Destination directory cannot be empty"],
        }

    src_real = os.path.realpath(src_dir)
    dest_real = os.path.realpath(dest_dir)
    if src_real == dest_real:
        return {"processed": 0, "skipped": 0, "errors": []}

    try:
        os.makedirs(dest_real, exist_ok=True)
    except OSError as exc:
        return {"processed": 0, "skipped": 0, "errors": [str(exc)]}

    processed = 0
    skipped = 0
    errors: List[str] = []

    try:
        with os.scandir(src_real) as entries:
            for entry in entries:
                try:
                    if not entry.is_file():
                        skipped += 1
                        continue
                    ext = os.path.splitext(entry.name)[1].lower()
                    if ext not in SUPPORTED_IMAGE_EXTS:
                        skipped += 1
                        continue

                    if move_files:
                        move_receipt_image(entry.path, dest_real, overwrite=False)
                    else:
                        copy_receipt_image(
                            entry.path,
                            os.path.join(dest_real, entry.name),
                            overwrite=False,
                        )
                    processed += 1
                except (OSError, ValueError) as exc:
                    errors.append(f"{entry.name}: {exc}")
    except OSError as exc:
        errors.append(str(exc))

    return {
        "processed": processed,
        "skipped": skipped,
        "errors": errors,
    }
