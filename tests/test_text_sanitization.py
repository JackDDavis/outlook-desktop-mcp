import unicodedata

import pytest

from outlook_desktop_mcp.utils.formatting import (
    format_email_full,
    format_email_summary,
    sanitize_body_text,
)
from tests.fakes import FakeMailItem, make_entry_id


def test_sanitize_body_text_removes_junk_run_and_normalizes_spaces():
    text = "offer " + ("\u034f\u2007" * 20) + "today"

    assert sanitize_body_text(text) == "offer " + (" " * 20) + "today"


def test_email_preview_collapses_junk_and_truncates_after_sanitizing():
    item = FakeMailItem(
        make_entry_id(),
        subject="\u200bSale\u034f today",
    )
    item.Body = ("A" * 299) + ("\u200b" * 100) + "B"

    summary = format_email_summary(item, include_body=True)
    full = format_email_full(item)

    assert summary["subject"] == "Sale today"
    assert summary["body_preview"] == ("A" * 299) + "B"
    assert full["body"] == ("A" * 299) + "B"


@pytest.mark.parametrize(
    "character",
    [
        "\u00ad",
        "\u200b",
        "\u200c",
        "\u200d",
        "\u200e",
        "\u200f",
        "\u2060",
        "\u2061",
        "\u2062",
        "\u2063",
        "\u2064",
        "\ufeff",
    ],
)
def test_sanitize_body_text_removes_all_format_characters(character):
    assert unicodedata.category(character) == "Cf"
    assert sanitize_body_text(f"a{character}b") == "ab"


@pytest.mark.parametrize(
    "character",
    [
        "\u00a0",
        "\u1680",
        *map(chr, range(0x2000, 0x200B)),
        "\u202f",
        "\u205f",
    ],
)
def test_sanitize_body_text_normalizes_unicode_spaces(character):
    assert sanitize_body_text(f"a{character}b") == "a b"


def test_email_preview_collapses_horizontal_space_but_preserves_newlines():
    item = FakeMailItem(make_entry_id())
    item.Body = "offer " + ("\u034f\u2007" * 20) + "today\nnext\t  line"

    summary = format_email_summary(item, include_body=True)
    full = format_email_full(item)

    assert summary["body_preview"] == "offer today\nnext line"
    assert full["body"] == "offer " + (" " * 20) + "today\nnext\t  line"


def test_sanitizer_preserves_clean_ascii_accents_and_cjk():
    text = "Clean ASCII - café - 中文"

    assert sanitize_body_text(text) == text
