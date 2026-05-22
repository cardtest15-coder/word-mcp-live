"""Tests for per-document asyncio lock manager."""
import asyncio
import pytest
from word_document_server.utils.concurrency import get_doc_lock, prune_doc_lock


def test_same_path_same_lock():
    lock1 = get_doc_lock("report.docx")
    lock2 = get_doc_lock("report.docx")
    assert lock1 is lock2


def test_different_paths_different_locks():
    lock1 = get_doc_lock("a.docx")
    lock2 = get_doc_lock("b.docx")
    assert lock1 is not lock2


def test_case_insensitive():
    lock1 = get_doc_lock("Report.DOCX")
    lock2 = get_doc_lock("report.docx")
    assert lock1 is lock2


def test_prune_removes_lock():
    get_doc_lock("temp.docx")
    prune_doc_lock("temp.docx")
    lock1 = get_doc_lock("temp.docx")
    prune_doc_lock("temp.docx")
    lock2 = get_doc_lock("temp.docx")
    assert lock1 is not lock2


@pytest.mark.asyncio
async def test_lock_is_held():
    lock = get_doc_lock("async_test.docx")
    acquired = []

    async def task1():
        async with lock:
            acquired.append("t1_start")
            await asyncio.sleep(0.1)
            acquired.append("t1_end")

    async def task2():
        async with lock:
            acquired.append("t2_start")

    await asyncio.gather(task1(), task2())
    assert acquired.index("t1_end") < acquired.index("t2_start")
