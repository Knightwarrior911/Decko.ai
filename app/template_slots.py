"""Authoritative builtin-template slot map + placeholder defaults.
Mirrors modActionsTemplate.ValidateTemplateSlots (engine FROZEN).
Pure data — no COM, no I/O."""

BUILTIN_SLOTS = {
    "title": ["title", "subtitle"],
    "section": ["section_number", "section_title"],
    "bullets": ["heading", "bullets"],
    "two_col": ["heading", "left_body", "right_body"],
    "comparison": ["heading", "left_label", "left_body",
                   "right_label", "right_body"],
    "kpi_dashboard": ["heading", "tiles"],
    "quote": ["quote_text", "attribution"],
}

TEMPLATE_NAMES = list(BUILTIN_SLOTS)

_PLACEHOLDER = {
    "title": "Title",
    "subtitle": "Subtitle",
    "section_number": "01",
    "section_title": "Section title",
    "heading": "Heading",
    "left_body": "Left content",
    "right_body": "Right content",
    "left_label": "Option A",
    "right_label": "Option B",
    "quote_text": "Quote goes here.",
    "attribution": "Attribution",
}


def default_content(template: str) -> dict:
    slots = BUILTIN_SLOTS[template]          # KeyError if unknown
    out = {}
    for s in slots:
        if s == "bullets":
            out[s] = ["First point", "Second point", "Third point"]
        elif s == "tiles":
            out[s] = [{"stat": "00", "label": "Metric one"},
                      {"stat": "00", "label": "Metric two"},
                      {"stat": "00", "label": "Metric three"}]
        else:
            out[s] = _PLACEHOLDER.get(s, s.replace("_", " ").title())
    return out
