"""Tests for the declarative tool registry."""
import asyncio

import pytest

from word_document_server.tools.registry import TOOLS


def test_registry_not_empty():
    assert len(TOOLS) > 0


def test_every_entry_has_name():
    for tool in TOOLS:
        assert "name" in tool, f"Tool missing 'name': {tool}"
        assert isinstance(tool["name"], str)


def test_every_entry_has_fn():
    for tool in TOOLS:
        assert "fn" in tool, f"Tool missing 'fn': {tool['name']}"
        assert callable(tool["fn"]), f"Tool fn not callable: {tool['name']}"


def test_every_entry_has_annotations():
    for tool in TOOLS:
        assert "annotations" in tool, f"Tool missing 'annotations': {tool['name']}"
        assert isinstance(tool["annotations"], dict)


def test_no_duplicate_names():
    names = [t["name"] for t in TOOLS]
    duplicates = [n for n in names if names.count(n) > 1]
    assert not duplicates, f"Duplicate tool names: {set(duplicates)}"


def test_tool_count_matches_expected():
    assert len(TOOLS) == 195, f"Expected 195 tools, got {len(TOOLS)}"


def test_all_functions_have_docstrings():
    for tool in TOOLS:
        fn = tool["fn"]
        assert fn.__doc__, f"Tool '{tool['name']}' function {fn.__name__} missing docstring"


@pytest.mark.asyncio
async def test_register_tools_end_to_end():
    import os
    os.environ.setdefault("FASTMCP_LOG_LEVEL", "WARNING")
    from word_document_server.main import mcp, register_tools
    register_tools()
    tools = await mcp._list_tools()
    assert len(tools) == 195, f"FastMCP registered {len(tools)} tools, expected 195"
    names = {t.name for t in tools}
    assert "create_document" in names
    assert "word_screen_capture" in names
