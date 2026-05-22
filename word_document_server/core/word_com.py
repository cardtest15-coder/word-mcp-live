"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
import sys
import unicodedata
from contextlib import contextmanager
from typing import Optional


def get_word_app():
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object that has open documents.
    When multiple Word instances are running, iterates through all
    Running Object Table (ROT) entries to find one with documents.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Word.Application")
        if app.Documents.Count > 0:
            return app
        # GetActiveObject found an empty instance — scan ROT for others
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        # No instance has documents; return the empty one (caller may open a file)
        return app
    except Exception:
        # GetActiveObject failed entirely — try ROT scan
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        # Word not running — start it
        try:
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = True
            return app
        except Exception as e:
            raise RuntimeError(
                f"Cannot start Microsoft Word: {e}"
            )


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        # Pass 1: look for a Word.Application ROT entry with documents
        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                # Direct Application entry
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                # Remember file monikers for pass 2
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                # Also collect file monikers we couldn't QI yet
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        # Pass 2: try file monikers → Document → Application
        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_document(app, filename: Optional[str] = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        # No documents open — try to open the requested file
        if filename and os.path.isfile(filename):
            try:
                doc = app.Documents.Open(os.path.abspath(filename))
                return doc
            except Exception as e:
                raise ValueError(f"Cannot open document '{filename}': {e}")
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    target_basename = unicodedata.normalize('NFC', os.path.basename(filename)).lower()
    target_fullpath = (
        unicodedata.normalize('NFC', os.path.normpath(filename)).lower()
        if os.path.isabs(filename) else None
    )

    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if unicodedata.normalize('NFC', doc.Name).lower() == target_basename:
            return doc
        if target_fullpath and unicodedata.normalize('NFC', os.path.normpath(doc.FullName)).lower() == target_fullpath:
            return doc

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    # Not found among open docs — try opening from disk
    if filename and os.path.isfile(filename):
        try:
            doc = app.Documents.Open(os.path.abspath(filename))
            return doc
        except Exception as e:
            raise ValueError(
                f"Document '{filename}' is not open and cannot be opened: {e}. "
                f"Open documents: {open_docs}"
            )
    raise ValueError(
        f"Document '{filename}' is not open in Word. "
        f"Open documents: {open_docs}"
    )


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    rec = None
    try:
        rec = app.UndoRecord
        # Clean up stale undo record from a previous crash/interrupted session
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None  # Word 2007 or earlier — proceed without
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass


@contextmanager
def comtypes_word_app(visible: bool = False):
    """Context manager that provides a Word Application via comtypes.

    Reuses a running Word instance if available; otherwise creates a new one.
    On exit, only calls Quit() if the instance was created by us (not already running).

    Args:
        visible: Whether to make the Word window visible. Default False.

    Yields:
        Tuple of (word_app, created_new: bool).

    Usage::

        with comtypes_word_app() as (word, created_new):
            doc = word.Documents.Open(path)
            doc.SaveAs2(out, FileFormat=8)
            doc.Close()
    """
    if sys.platform != "win32":
        raise RuntimeError("comtypes Word automation is only available on Windows")

    import comtypes.client

    created_new = False
    word = None
    try:
        try:
            word = comtypes.client.GetActiveObject("Word.Application")
        except Exception:
            word = comtypes.client.CreateObject("Word.Application")
            created_new = True
        word.Visible = visible
        yield word, created_new
    finally:
        if created_new and word is not None:
            try:
                word.Quit()
            except Exception:
                pass
