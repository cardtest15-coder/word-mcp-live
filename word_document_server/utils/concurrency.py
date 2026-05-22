"""Per-document asyncio.Lock manager for live tools."""
from __future__ import annotations

import asyncio

_locks: dict[str, asyncio.Lock] = {}


def get_doc_lock(doc_path: str) -> asyncio.Lock:
    """Return (or create) an asyncio.Lock for the given document path."""
    key = doc_path.lower()
    if key not in _locks:
        _locks[key] = asyncio.Lock()
    return _locks[key]


def prune_doc_lock(doc_path: str) -> None:
    """Remove lock entry after document is closed (optional cleanup)."""
    _locks.pop(doc_path.lower(), None)
