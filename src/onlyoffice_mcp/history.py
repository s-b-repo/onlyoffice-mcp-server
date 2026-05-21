"""Edit history / version-control subsystem.

Every mutating tool can be wrapped with the :func:`record_operation`
decorator, which transparently:

1. Reads the file hash + text content BEFORE the operation
2. Runs the wrapped function
3. Reads the file hash + text content AFTER
4. Computes a unified text diff (per-format ``read`` function)
5. Appends a JSON Lines record to ``ops.jsonl``
6. Copies the file to ``snapshots/rev_NNNNNN.<ext>`` (keep last N)
7. Updates ``meta.json``

The data lives in ``~/.onlyoffice-mcp/history/<doc_id>/``. Disk usage is
capped to 20 snapshots per doc and 100 MB total (overridable via
``ONLYOFFICE_MCP_HISTORY_MAX_SNAPSHOTS`` and ``_MAX_BYTES``).
"""

from __future__ import annotations

import difflib
import functools
import json
import logging
import os
import time
from pathlib import Path
from typing import Any, Callable

from . import storage
from .errors import HistoryError, SnapshotPruned

logger = logging.getLogger("onlyoffice-mcp.history")


# --------------------------------------------------------------------------
# Per-format text extraction (used for diffs)
# --------------------------------------------------------------------------

def _text_for(path: Path) -> str:
    """Return a textual representation of a document for diffing.

    Heavy reads are imported lazily so that history.py has no hard dependency
    on python-docx / openpyxl / python-pptx at import time.
    """
    if not path.exists():
        return ""
    ext = path.suffix.lower().lstrip(".")
    try:
        if ext == "docx":
            from . import docx_ops
            return docx_ops.read(str(path), include_tables=True)
        if ext == "xlsx":
            from . import xlsx_ops
            data = xlsx_ops.read(str(path))
            lines: list[str] = []
            for sheet_name, rows in data.items():
                lines.append(f"### Sheet: {sheet_name}")
                for row in rows:
                    lines.append(" | ".join("" if c is None else str(c) for c in row))
                lines.append("")
            return "\n".join(lines)
        if ext == "pptx":
            from . import pptx_ops
            info = pptx_ops.read(str(path))
            lines = []
            for slide in info["slides"]:
                lines.append(f"# Slide {slide['index']}: {slide['title']}")
                for t in slide["text"]:
                    lines.append(f"  - {t}")
                if slide["notes"]:
                    lines.append(f"  (notes) {slide['notes']}")
                lines.append("")
            return "\n".join(lines)
    except Exception as e:  # noqa: BLE001
        logger.warning("history: text extraction failed for %s: %s", path, e)
        return ""
    return ""


# --------------------------------------------------------------------------
# Meta / log I/O
# --------------------------------------------------------------------------

def _meta_path(doc_path: str | Path) -> Path:
    return storage.doc_dir(doc_path) / "meta.json"


def _log_path(doc_path: str | Path) -> Path:
    return storage.doc_dir(doc_path) / "ops.jsonl"


def _snapshots_dir(doc_path: str | Path) -> Path:
    return storage.doc_dir(doc_path) / "snapshots"


def _load_meta(doc_path: str | Path) -> dict:
    p = _meta_path(doc_path)
    if not p.exists():
        return {
            "doc_id": storage.doc_id(doc_path),
            "abs_path": str(Path(doc_path).expanduser().resolve()),
            "format": Path(doc_path).suffix.lstrip(".").lower(),
            "created_ts": time.time(),
            "last_modified_ts": time.time(),
            "head_revision": -1,
            "cursor": {},
            "snapshot_count": 0,
            "snapshot_oldest_rev": 0,
        }
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:  # noqa: BLE001
        logger.warning("history: corrupt meta.json (%s); rebuilding", e)
        return {
            "doc_id": storage.doc_id(doc_path),
            "abs_path": str(Path(doc_path).expanduser().resolve()),
            "format": Path(doc_path).suffix.lstrip(".").lower(),
            "created_ts": time.time(),
            "last_modified_ts": time.time(),
            "head_revision": -1,
            "cursor": {},
            "snapshot_count": 0,
            "snapshot_oldest_rev": 0,
        }


def _save_meta(doc_path: str | Path, meta: dict) -> None:
    storage.atomic_write_json(_meta_path(doc_path), meta)


def _append_log(doc_path: str | Path, record: dict) -> None:
    p = _log_path(doc_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def _read_log(doc_path: str | Path) -> list[dict]:
    p = _log_path(doc_path)
    if not p.exists():
        return []
    out = []
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out


# --------------------------------------------------------------------------
# Snapshot pruning
# --------------------------------------------------------------------------

def _prune_snapshots(doc_path: str | Path) -> None:
    """Keep at most ``storage.max_snapshots()`` snapshot files per doc; obey
    the global ``max_total_bytes`` budget by dropping oldest files first
    across all tracked docs."""
    snap_dir = _snapshots_dir(doc_path)
    snapshots = sorted(snap_dir.glob("rev_*.bin")) + sorted(snap_dir.glob("rev_*.*"))
    # Deduplicate; sorting by name puts oldest first thanks to zero-padded numbers.
    snapshots = sorted({p for p in snapshots if p.is_file()}, key=lambda p: p.name)
    limit = storage.max_snapshots()
    while len(snapshots) > limit:
        oldest = snapshots.pop(0)
        try:
            oldest.unlink()
        except OSError:
            pass
    # Global byte cap: prune across ALL docs, oldest first.
    total_cap = storage.max_total_bytes()
    if total_cap == 0:
        return
    all_snapshots: list[Path] = []
    for d in storage.list_tracked_docs():
        all_snapshots.extend((d / "snapshots").glob("rev_*"))
    all_snapshots = sorted([p for p in all_snapshots if p.is_file()], key=lambda p: p.stat().st_mtime)
    total = sum(p.stat().st_size for p in all_snapshots)
    while total > total_cap and all_snapshots:
        victim = all_snapshots.pop(0)
        size = victim.stat().st_size
        try:
            victim.unlink()
            total -= size
        except OSError:
            break


# --------------------------------------------------------------------------
# Diff summary helpers
# --------------------------------------------------------------------------

def _summarise_diff(diff_lines: list[str]) -> tuple[int, int, str]:
    """Return (added, removed, short_summary) given unified-diff lines."""
    added = removed = 0
    for line in diff_lines:
        if line.startswith("+++") or line.startswith("---") or line.startswith("@@"):
            continue
        if line.startswith("+"):
            added += 1
        elif line.startswith("-"):
            removed += 1
    if added == 0 and removed == 0:
        summary = "no textual change"
    elif removed == 0:
        summary = f"+{added} line(s)"
    elif added == 0:
        summary = f"-{removed} line(s)"
    else:
        summary = f"+{added} / -{removed} line(s)"
    return added, removed, summary


def _unified_diff(before: str, after: str) -> list[str]:
    return list(
        difflib.unified_diff(
            before.splitlines(),
            after.splitlines(),
            fromfile="before",
            tofile="after",
            lineterm="",
        )
    )


# --------------------------------------------------------------------------
# record_operation decorator
# --------------------------------------------------------------------------

def _doc_path_from_args(args: tuple, kwargs: dict) -> str | None:
    """The wrapped tools all take ``path`` as the first positional or keyword
    argument."""
    if "path" in kwargs:
        return kwargs["path"]
    if args:
        first = args[0]
        if isinstance(first, (str, os.PathLike)):
            return str(first)
    return None


def record_operation(op_name: str) -> Callable:
    """Wrap a mutating tool to capture before/after state in the history log.

    The wrapped function must take ``path`` as its first positional argument
    (or as a keyword argument). The wrapper is fail-safe: any error inside the
    history machinery is logged at WARNING and the underlying op is preserved.
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            if not storage.history_enabled():
                return func(*args, **kwargs)
            doc_path = _doc_path_from_args(args, kwargs)
            if doc_path is None:
                return func(*args, **kwargs)
            doc_p = Path(doc_path).expanduser().resolve()

            before_hash = storage.content_hash(doc_p) if doc_p.exists() else ""
            before_text = _text_for(doc_p) if doc_p.exists() else ""

            result = func(*args, **kwargs)

            try:
                with storage.lock(doc_p):
                    after_hash = storage.content_hash(doc_p)
                    after_text = _text_for(doc_p)
                    if after_hash == before_hash and after_text == before_text:
                        # No-op or read-only — skip logging.
                        return result

                    meta = _load_meta(doc_p)
                    revision = meta.get("head_revision", -1) + 1

                    diff_lines = _unified_diff(before_text, after_text)
                    added, removed, diff_summary = _summarise_diff(diff_lines)

                    snapshot_saved = False
                    try:
                        snap = _snapshots_dir(doc_p) / f"rev_{revision:06d}{doc_p.suffix}"
                        if doc_p.exists():
                            storage.copy_to_snapshot(doc_p, snap)
                            snapshot_saved = True
                    except OSError as e:
                        logger.warning("history: snapshot copy failed: %s", e)

                    record = {
                        "ts": time.time(),
                        "revision": revision,
                        "tool": op_name,
                        "args_summary": _summarise_args(args, kwargs),
                        "before_hash": before_hash,
                        "after_hash": after_hash,
                        "diff_summary": diff_summary,
                        "text_diff_lines_added": added,
                        "text_diff_lines_removed": removed,
                        "text_diff": "\n".join(diff_lines),
                        "snapshot_saved": snapshot_saved,
                        "session_id": f"mcp-{os.getpid()}",
                    }
                    _append_log(doc_p, record)

                    meta["head_revision"] = revision
                    meta["last_modified_ts"] = record["ts"]
                    if "created_ts" not in meta:
                        meta["created_ts"] = record["ts"]
                    if snapshot_saved:
                        meta["snapshot_count"] = meta.get("snapshot_count", 0) + 1
                    _save_meta(doc_p, meta)

                    _prune_snapshots(doc_p)
            except Exception as e:  # noqa: BLE001
                logger.warning("history: failed to record %s on %s: %s", op_name, doc_p, e)

            return result

        return wrapper

    return decorator


def _summarise_args(args: tuple, kwargs: dict) -> dict:
    """Build a small JSON-safe summary of arguments for the ops log."""
    summary: dict[str, Any] = {}
    if args:
        summary["positional_count"] = len(args)
        if isinstance(args[0], (str, os.PathLike)):
            summary["path"] = str(args[0])
    for k, v in kwargs.items():
        if isinstance(v, (str, int, float, bool)) or v is None:
            summary[k] = v
        elif isinstance(v, (list, tuple)):
            summary[k] = f"<{type(v).__name__} len={len(v)}>"
        elif isinstance(v, dict):
            summary[k] = f"<dict keys={list(v.keys())[:8]}>"
        else:
            summary[k] = f"<{type(v).__name__}>"
    return summary


# --------------------------------------------------------------------------
# Public query API — used by MCP tools
# --------------------------------------------------------------------------

def list_history(path: str | Path, limit: int = 20) -> list[dict]:
    """Return the most recent ``limit`` ops for a document (newest first)."""
    log = _read_log(path)
    log.sort(key=lambda r: r.get("revision", 0), reverse=True)
    return [
        {
            "revision": r.get("revision"),
            "ts": r.get("ts"),
            "tool": r.get("tool"),
            "diff_summary": r.get("diff_summary"),
            "args_summary": r.get("args_summary"),
            "snapshot_saved": r.get("snapshot_saved"),
        }
        for r in log[:limit]
    ]


def show_revision(path: str | Path, revision: int) -> dict:
    log = _read_log(path)
    for r in log:
        if r.get("revision") == revision:
            return r
    raise HistoryError(f"No such revision: {revision}")


def last_edit(path: str | Path) -> dict:
    log = _read_log(path)
    if not log:
        raise HistoryError("No history recorded for this document")
    log.sort(key=lambda r: r.get("revision", 0), reverse=True)
    record = log[0]
    record_copy = dict(record)
    record_copy["age_seconds"] = time.time() - record.get("ts", time.time())
    return record_copy


def what_was_removed(path: str | Path, revision: int | None = None) -> dict:
    """Return only the ``-`` lines from a revision's diff."""
    if revision is None:
        record = last_edit(path)
        revision = record["revision"]
    else:
        record = show_revision(path, revision)
    diff = record.get("text_diff", "")
    removed = [
        line[1:]
        for line in diff.split("\n")
        if line.startswith("-") and not line.startswith("---")
    ]
    return {
        "revision": revision,
        "removed_line_count": len(removed),
        "removed_lines": removed,
    }


def diff_revisions(
    path: str | Path,
    from_rev: int | None = None,
    to_rev: int | None = None,
) -> dict:
    log = _read_log(path)
    if not log:
        return {"diff": "", "from_rev": None, "to_rev": None}
    log.sort(key=lambda r: r.get("revision", 0))
    if to_rev is None:
        to_rev = log[-1]["revision"]
    if from_rev is None:
        from_rev = max(0, to_rev - 1)
    if from_rev > to_rev:
        from_rev, to_rev = to_rev, from_rev

    # Reconstruct text by sequentially applying diffs is hard; instead, look up
    # snapshots and re-read them when available, otherwise reuse the recorded
    # ``text_diff`` from the to_rev record (the simplest unambiguous answer).
    snap_dir = _snapshots_dir(path)
    suffix = Path(path).suffix
    from_snap = snap_dir / f"rev_{from_rev:06d}{suffix}"
    to_snap = snap_dir / f"rev_{to_rev:06d}{suffix}"
    if from_snap.exists() and to_snap.exists():
        before = _text_for(from_snap)
        after = _text_for(to_snap)
        diff_lines = _unified_diff(before, after)
        return {
            "from_rev": from_rev,
            "to_rev": to_rev,
            "diff": "\n".join(diff_lines),
            "source": "snapshot",
        }

    # Fallback to the diff stored in the to_rev op record.
    rec = next((r for r in log if r["revision"] == to_rev), None)
    return {
        "from_rev": from_rev,
        "to_rev": to_rev,
        "diff": rec.get("text_diff", "") if rec else "",
        "source": "record",
    }


def revert(path: str | Path, revision: int) -> dict:
    """Restore the document from snapshot at ``revision`` and record the
    revert itself as a new edit."""
    snap_dir = _snapshots_dir(path)
    suffix = Path(path).suffix
    snap = snap_dir / f"rev_{revision:06d}{suffix}"
    if not snap.exists():
        log = _read_log(path)
        revertable = sorted(
            (r["revision"] for r in log if r.get("snapshot_saved"))
        )
        raise SnapshotPruned(
            f"Snapshot for revision {revision} was pruned. "
            f"Revertable revisions: {revertable[-storage.max_snapshots():]}"
        )
    target = Path(path).expanduser().resolve()
    with storage.lock(target):
        # Capture the "before revert" state so what_was_removed still works.
        before_text = _text_for(target) if target.exists() else ""
        storage.atomic_write_bytes(target, snap.read_bytes())
        after_text = _text_for(target)
        diff_lines = _unified_diff(before_text, after_text)
        added, removed, summary = _summarise_diff(diff_lines)
        meta = _load_meta(target)
        new_rev = meta.get("head_revision", -1) + 1
        record = {
            "ts": time.time(),
            "revision": new_rev,
            "tool": "doc_revert",
            "args_summary": {"path": str(target), "revert_to": revision},
            "before_hash": storage.content_hash(target),
            "after_hash": storage.content_hash(target),
            "diff_summary": f"reverted to rev {revision}: {summary}",
            "text_diff_lines_added": added,
            "text_diff_lines_removed": removed,
            "text_diff": "\n".join(diff_lines),
            "snapshot_saved": False,
            "session_id": f"mcp-{os.getpid()}",
        }
        _append_log(target, record)
        meta["head_revision"] = new_rev
        meta["last_modified_ts"] = record["ts"]
        _save_meta(target, meta)
    return {"new_revision": new_rev, "reverted_to": revision}


def clear_history(path: str | Path, keep_last: int = 0) -> dict:
    d = storage.doc_dir(path)
    snap_dir = d / "snapshots"
    freed = 0
    if keep_last > 0:
        log = sorted(_read_log(path), key=lambda r: r.get("revision", 0))
        keep_revs = {r["revision"] for r in log[-keep_last:]}
        for p in snap_dir.glob("rev_*"):
            try:
                rev = int(p.stem.split("_")[1])
            except (IndexError, ValueError):
                continue
            if rev not in keep_revs:
                freed += p.stat().st_size
                p.unlink()
        # Trim log file too.
        kept = [r for r in log if r["revision"] in keep_revs]
        log_path = _log_path(path)
        log_path.write_text(
            "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in kept),
            encoding="utf-8",
        )
        return {"freed_bytes": freed, "kept_revisions": sorted(keep_revs)}
    # Wipe everything for this doc.
    freed = storage.dir_size_bytes(d)
    import shutil
    shutil.rmtree(d, ignore_errors=True)
    return {"freed_bytes": freed, "kept_revisions": []}


def history_stats() -> dict:
    docs = storage.list_tracked_docs()
    total_bytes = sum(storage.dir_size_bytes(d) for d in docs)
    oldest_ts = None
    snapshot_total = 0
    for d in docs:
        snapshot_total += len(list((d / "snapshots").glob("rev_*")))
        meta_p = d / "meta.json"
        if meta_p.exists():
            try:
                ts = json.loads(meta_p.read_text("utf-8")).get("created_ts")
                if ts is not None and (oldest_ts is None or ts < oldest_ts):
                    oldest_ts = ts
            except Exception:  # noqa: BLE001
                pass
    return {
        "tracked_documents": len(docs),
        "total_disk_bytes": total_bytes,
        "total_disk_human": _human_bytes(total_bytes),
        "snapshot_count": snapshot_total,
        "oldest_tracked_ts": oldest_ts,
        "max_snapshots_per_doc": storage.max_snapshots(),
        "max_total_bytes": storage.max_total_bytes(),
        "history_enabled": storage.history_enabled(),
    }


def _human_bytes(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"
