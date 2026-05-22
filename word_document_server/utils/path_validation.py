import os

_ALLOWED_DIRS: list[str] | None = None


def configure(raw: str | None):
    global _ALLOWED_DIRS
    if not raw:
        _ALLOWED_DIRS = None
        return
    _ALLOWED_DIRS = [os.path.realpath(p) for p in raw.split(":") if p.strip()]


def validate_path(filepath: str) -> str:
    if not filepath:
        raise ValueError("Path must not be empty")
    if ".." in filepath:
        raise ValueError("Path traversal not allowed: '..' in path")
    real = os.path.realpath(filepath)
    if _ALLOWED_DIRS is not None:
        if not any(real.startswith(d + os.sep) or real == d for d in _ALLOWED_DIRS):
            raise ValueError(f"Path '{real}' is outside allowed directories")
    return real
