"""Tests for error_helpers utility."""
from word_document_server.utils.error_helpers import ok, err


def test_ok_returns_success_true():
    result = ok()
    assert result["success"] is True


def test_ok_merges_kwargs():
    result = ok(count=5, name="test")
    assert result["success"] is True
    assert result["count"] == 5
    assert result["name"] == "test"


def test_ok_no_error_key():
    result = ok()
    assert "error" not in result


def test_err_returns_success_false():
    result = err("something broke")
    assert result["success"] is False


def test_err_includes_error_message():
    result = err("something broke")
    assert result["error"] == "something broke"


def test_err_only_two_keys():
    result = err("msg")
    assert set(result.keys()) == {"success", "error"}
