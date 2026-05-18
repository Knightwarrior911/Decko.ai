from app.template_slots import (BUILTIN_SLOTS, TEMPLATE_NAMES,
                                 default_content)


def test_seven_builtins_with_authoritative_slots():
    assert set(TEMPLATE_NAMES) == {
        "title", "section", "bullets", "two_col", "comparison",
        "kpi_dashboard", "quote"}
    assert BUILTIN_SLOTS["title"] == ["title", "subtitle"]
    assert BUILTIN_SLOTS["section"] == ["section_number", "section_title"]
    assert BUILTIN_SLOTS["bullets"] == ["heading", "bullets"]
    assert BUILTIN_SLOTS["two_col"] == ["heading", "left_body",
                                        "right_body"]
    assert BUILTIN_SLOTS["comparison"] == ["heading", "left_label",
                                           "left_body", "right_label",
                                           "right_body"]
    assert BUILTIN_SLOTS["kpi_dashboard"] == ["heading", "tiles"]
    assert BUILTIN_SLOTS["quote"] == ["quote_text", "attribution"]


def test_default_content_shapes():
    c = default_content("bullets")
    assert isinstance(c["bullets"], list) and c["bullets"]
    assert isinstance(c["heading"], str) and c["heading"]
    k = default_content("kpi_dashboard")
    assert isinstance(k["tiles"], list)
    assert set(k["tiles"][0].keys()) == {"stat", "label"}
    t = default_content("title")
    assert set(t.keys()) == {"title", "subtitle"}


def test_default_content_unknown_raises():
    import pytest
    with pytest.raises(KeyError):
        default_content("nope")
