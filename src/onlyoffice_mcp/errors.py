"""Typed exceptions for the OnlyOffice MCP server.

All exceptions inherit from :class:`OnlyOfficeMCPError` so callers can catch
the base class to handle anything thrown by this package.
"""

from __future__ import annotations


class OnlyOfficeMCPError(Exception):
    """Base class for all OnlyOffice MCP server errors."""


class HistoryError(OnlyOfficeMCPError):
    """Generic problem with the edit-history subsystem."""


class SnapshotPruned(HistoryError):
    """A revert was requested for a snapshot that has been pruned.

    Disk usage is capped, so the oldest snapshots are removed first. The
    operation log still records every edit, but the binary content can no
    longer be restored. The error message includes the oldest revertable
    revision so the caller can recover.
    """


class CursorOutOfBounds(OnlyOfficeMCPError):
    """A cursor was set or moved past the document edges.

    Implementations typically *clamp* and emit a warning rather than raising
    this — but it is available for callers that prefer hard failures.
    """


class EngineMissing(OnlyOfficeMCPError):
    """A required engine (Document Builder / LibreOffice / pyspellchecker /
    pypdf) is not available on the host. Message includes install hints.
    """


class DocumentLocked(OnlyOfficeMCPError):
    """Another process holds the per-document lock. The operation was
    refused after the configured retry window expired."""


class UnsupportedFormat(OnlyOfficeMCPError):
    """The operation is not supported for the given file extension."""


class MassDeletionBlocked(OnlyOfficeMCPError):
    """Too many files were deleted in a short time window.

    This is a safety mechanism to prevent accidental or malicious bulk
    deletion. The threshold and time window are configurable via
    environment variables.
    """


class DeletionDenied(OnlyOfficeMCPError):
    """A file deletion was rejected by the safety layer.

    Common causes: the file is outside the allowed workspace, the path
    points to a sensitive location, or the deletion would exceed the
    rate limit.
    """
