"""Consistent JSON error helpers for MCP tool responses."""
from __future__ import annotations

from typing import Any


def ok(**kwargs: Any) -> dict[str, Any]:
    """Return a success dict. Any extra kwargs are merged in."""
    result: dict[str, Any] = {"success": True}
    result.update(kwargs)
    return result


def err(message: str) -> dict[str, Any]:
    """Return an error dict with a consistent shape."""
    return {"success": False, "error": message}
