"""Workspace, doc-ID and lock helpers used by the history subsystem.

Workspace root defaults to ``~/.onlyoffice-mcp``; overridable via the
``ONLYOFFICE_MCP_HOME`` environment variable. Inside the workspace,
``history/<doc_id>/`` holds the per-document log and snapshots, where
``doc_id`` is a stable hash of the absolute file path.
"""

from __future__ import annotations

import contextlib
import fcntl
import hashlib
import json
import os
import shutil
import tempfile
import time
from pathlib import Path
from typing import Iterator

from .errors import DocumentLocked


DEFAULT_HOME = Path.home() / ".onlyoffice-mcp"
ENV_HOME = "ONLYOFFICE_MCP_HOME"
ENV_DISABLE_HISTORY = "ONLYOFFICE_MCP_HISTORY_ENABLED"
ENV_MAX_SNAPSHOTS = "ONLYOFFICE_MCP_HISTORY_MAX_SNAPSHOTS"
ENV_MAX_BYTES = "ONLYOFFICE_MCP_HISTORY_MAX_BYTES"

DEFAULT_MAX_SNAPSHOTS = 20
DEFAULT_MAX_BYTES = 100 * 1024 * 1024  # 100 MB across all docs


def home() -> Path:
    root = Path(os.environ.get(ENV_HOME, str(DEFAULT_HOME))).expanduser()
    root.mkdir(parents=True, exist_ok=True)
    return root


def history_root() -> Path:
    p = home() / "history"
    p.mkdir(parents=True, exist_ok=True)
    return p


def history_enabled() -> bool:
    return os.environ.get(ENV_DISABLE_HISTORY, "true").lower() not in {
        "0", "false", "no", "off",
    }


def max_snapshots() -> int:
    try:
        return max(1, int(os.environ.get(ENV_MAX_SNAPSHOTS, DEFAULT_MAX_SNAPSHOTS)))
    except ValueError:
        return DEFAULT_MAX_SNAPSHOTS


def max_total_bytes() -> int:
    try:
        return max(0, int(os.environ.get(ENV_MAX_BYTES, DEFAULT_MAX_BYTES)))
    except ValueError:
        return DEFAULT_MAX_BYTES


def doc_id(path: str | Path) -> str:
    """Stable identifier: first 16 hex chars of sha256(absolute_path)."""
    abs_path = str(Path(path).expanduser().resolve())
    return hashlib.sha256(abs_path.encode("utf-8")).hexdigest()[:16]


def doc_dir(path: str | Path) -> Path:
    d = history_root() / doc_id(path)
    d.mkdir(parents=True, exist_ok=True)
    (d / "snapshots").mkdir(parents=True, exist_ok=True)
    return d


def content_hash(path: str | Path) -> str:
    """SHA-256 of file contents, or empty string if the file is missing."""
    p = Path(path).expanduser()
    if not p.exists():
        return ""
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return f"sha256:{h.hexdigest()}"


def atomic_write_bytes(path: str | Path, data: bytes) -> None:
    """Write bytes atomically (tmp file + rename)."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "wb",
        delete=False,
        dir=str(target.parent),
        prefix=target.name + ".",
        suffix=".tmp",
    ) as f:
        f.write(data)
        tmp_path = f.name
    os.replace(tmp_path, str(target))


def atomic_write_text(path: str | Path, text: str, encoding: str = "utf-8") -> None:
    atomic_write_bytes(path, text.encode(encoding))


def atomic_write_json(path: str | Path, obj: object) -> None:
    atomic_write_text(path, json.dumps(obj, indent=2, ensure_ascii=False, sort_keys=True))


@contextlib.contextmanager
def lock(path: str | Path, *, timeout: float = 5.0, poll_interval: float = 0.2) -> Iterator[None]:
    """Advisory ``flock`` on a sentinel file in the doc's history dir.

    Polls every ``poll_interval`` seconds for up to ``timeout`` seconds.
    Raises :class:`DocumentLocked` if still contended.
    """
    d = doc_dir(path)
    lock_path = d / ".lock"
    lock_path.touch(exist_ok=True)
    fd = os.open(str(lock_path), os.O_RDWR | os.O_CREAT, 0o644)
    deadline = time.monotonic() + max(0.0, timeout)
    try:
        while True:
            try:
                fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    raise DocumentLocked(
                        f"Document is locked by another operation: {path}"
                    )
                time.sleep(poll_interval)
        yield
    finally:
        try:
            fcntl.flock(fd, fcntl.LOCK_UN)
        finally:
            os.close(fd)


def dir_size_bytes(directory: Path) -> int:
    if not directory.exists():
        return 0
    total = 0
    for p in directory.rglob("*"):
        if p.is_file():
            try:
                total += p.stat().st_size
            except OSError:
                pass
    return total


def list_tracked_docs() -> list[Path]:
    root = history_root()
    if not root.exists():
        return []
    return [p for p in root.iterdir() if p.is_dir()]


def copy_to_snapshot(src: str | Path, dst: str | Path) -> int:
    src_p = Path(src)
    dst_p = Path(dst)
    dst_p.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(src_p), str(dst_p))
    return dst_p.stat().st_size
