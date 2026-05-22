"""Tests for security hardening: path validation, URL validation, position validation, JS escaping."""
import pytest


class TestPathValidation:
    def test_reject_empty_path(self):
        from word_document_server.utils.path_validation import validate_path
        with pytest.raises(ValueError, match="empty"):
            validate_path("")

    def test_reject_traversal(self):
        from word_document_server.utils.path_validation import validate_path
        with pytest.raises(ValueError, match="[Tt]raversal"):
            validate_path("../../etc/passwd")

    def test_reject_traversal_in_middle(self):
        from word_document_server.utils.path_validation import validate_path
        with pytest.raises(ValueError):
            validate_path("foo/../../../etc/passwd")

    def test_allow_normal_path(self):
        from word_document_server.utils.path_validation import validate_path
        result = validate_path("report.docx")
        assert result

    def test_configure_allowed_dirs(self, tmp_path):
        from word_document_server.utils.path_validation import configure, validate_path
        configure(str(tmp_path))
        result = validate_path(str(tmp_path / "test.docx"))
        assert result

    def test_reject_outside_allowed_dirs(self, tmp_path):
        from word_document_server.utils.path_validation import configure, validate_path
        configure(str(tmp_path))
        with pytest.raises(ValueError, match="[Aa]llowed"):
            validate_path("/etc/passwd")

    def test_empty_config_allows_all(self):
        from word_document_server.utils.path_validation import configure, validate_path
        configure(None)
        result = validate_path("any/path.docx")
        assert result


class TestUrlValidation:
    def test_allow_https(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        _validate_url("https://example.com")

    def test_allow_http(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        _validate_url("http://example.com")

    def test_allow_mailto(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        _validate_url("mailto:user@example.com")

    def test_reject_javascript(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        with pytest.raises(ValueError, match="scheme"):
            _validate_url("javascript:alert(1)")

    def test_reject_file_scheme(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        with pytest.raises(ValueError, match="scheme"):
            _validate_url("file:///etc/passwd")

    def test_reject_smb_path(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        with pytest.raises(ValueError):
            _validate_url("//attacker/share")

    def test_reject_unc_path(self):
        from word_document_server.core.hyperlink_writer import _validate_url
        with pytest.raises(ValueError):
            _validate_url("\\\\attacker\\share")


class TestPositionValidation:
    def test_valid_integer(self):
        from word_document_server.utils.text_safety import validate_position
        assert validate_position("42") == 42

    def test_valid_zero(self):
        from word_document_server.utils.text_safety import validate_position
        assert validate_position("0") == 0

    def test_reject_negative(self):
        from word_document_server.utils.text_safety import validate_position
        with pytest.raises(ValueError, match="[Nn]egative"):
            validate_position("-1")

    def test_reject_non_integer(self):
        from word_document_server.utils.text_safety import validate_position
        with pytest.raises(ValueError):
            validate_position("abc")

    def test_reject_injection(self):
        from word_document_server.utils.text_safety import validate_position
        with pytest.raises(ValueError):
            validate_position('1"); $.system("rm -rf /")')

    def test_reject_empty(self):
        from word_document_server.utils.text_safety import validate_position
        with pytest.raises(ValueError):
            validate_position("")

    def test_reject_float(self):
        from word_document_server.utils.text_safety import validate_position
        with pytest.raises(ValueError):
            validate_position("3.14")


class TestEscapeJs:
    def test_escapes_backslash(self):
        from word_document_server.core.word_mac import _escape_js
        assert "\\\\" in _escape_js("\\")

    def test_strips_null_bytes(self):
        from word_document_server.core.word_mac import _escape_js
        assert "\x00" not in _escape_js("hello\x00world")

    def test_escapes_unicode_line_terminators(self):
        from word_document_server.core.word_mac import _escape_js
        result = _escape_js("\u2028\u2029")
        assert "\u2028" not in result
        assert "\u2029" not in result

    def test_normal_text_unchanged(self):
        from word_document_server.core.word_mac import _escape_js
        assert _escape_js("Hello World") == "Hello World"
