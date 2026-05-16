"""Sanitizer robustness smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_sanitizer.py

Opens the carrier (PPT_AI_Editor.pptm), calls
  app.Run("PPT_AI_Editor!SanitizeJsonInput", raw)
for every corpus case, then asserts the cleaned output either:
  - json.loads() succeeds AND deep-equals the expected object, OR
  - for the documented "no { or [" contract: the output is returned
    byte-identical to the raw input (VERBATIM sentinel).

Prints per-case PASS/FAIL, an aggregate line, and exits non-zero unless
100% of the corpus passes.

No AI/API involved -- fully deterministic.
"""
import json
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

# Unicode noise characters
BOM = "﻿"
LDQ = "“"  # left  double curly  "
RDQ = "”"  # right double curly  "
LSQ = "‘"  # left  single curly  '
RSQ = "’"  # right single curly  '

# Canonical target object reused by most structural cases.
CANON = {"actions": [{"type": "add_text_box", "slide": 1, "text": "Hello"}]}
CANON_JSON = '{"actions":[{"type":"add_text_box","slide":1,"text":"Hello"}]}'

# Sentinel: output must be byte-identical to the raw input.
VERBATIM = object()


def _curly(s: str) -> str:
    """Rewrite a pure-ASCII JSON string using only curly double quotes
    (no ASCII double quote anywhere) to exercise the smart-quote path."""
    return s.replace('"', LDQ)


# corpus :: list[ (name, raw, expected) ]
# expected is either a Python object (deep-equal after json.loads) or VERBATIM.
CORPUS = [
    # ---- 1. BOM ----
    ("bom_utf16", BOM + CANON_JSON, CANON),
    ("bom_then_fence", BOM + "```json\n" + CANON_JSON + "\n```", CANON),

    # ---- 2. smart quotes ----
    ("smart_double_all_curly", _curly(CANON_JSON), CANON),
    ("smart_double_in_fence", "```json\n" + _curly(CANON_JSON) + "\n```", CANON),
    ("smart_single_content",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"it' + RSQ + 's fine"}]}',
     {"actions": [{"type": "add_text_box", "slide": 1, "text": "it's fine"}]}),
    ("mixed_ascii_struct_curly_content",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"He said ' + LDQ + "hi" + RDQ + '"}]}',
     {"actions": [{"type": "add_text_box", "slide": 1,
                   "text": "He said " + LDQ + "hi" + RDQ}]}),

    # ---- 3. markdown fences ----
    ("fence_json", "```json\n" + CANON_JSON + "\n```", CANON),
    ("fence_bare", "```\n" + CANON_JSON + "\n```", CANON),
    ("fence_json_uppercase", "```JSON\n" + CANON_JSON + "\n```", CANON),
    ("fence_space_lang", "``` json\n" + CANON_JSON + "\n```", CANON),
    ("fence_crlf", "```json\r\n" + CANON_JSON + "\r\n```", CANON),
    ("fence_no_newline", "```json" + CANON_JSON + "```", CANON),

    # ---- 4. prose before ----
    ("prose_before", "Sure! Here is your JSON:\n" + CANON_JSON, CANON),
    ("prose_before_long",
     "Absolutely, I can help with that. Below is the requested batch.\n\n" + CANON_JSON,
     CANON),
    ("line_comment_before", "// here is the json\n" + CANON_JSON, CANON),

    # ---- 5. prose after ----
    ("prose_after", CANON_JSON + "\nLet me know if you need changes!", CANON),
    ("prose_after_multiline",
     CANON_JSON + "\n\nThis adds one text box to slide 1.\nHope this helps.",
     CANON),

    # ---- 6. JS comments ----
    ("line_comment_outside",
     '{ "actions": [ {"type":"add_text_box","slide":1,"text":"Hello"} ] } // done',
     CANON),
    ("block_comment_leading",
     '{ /* meta */ "actions": [ {"type":"add_text_box","slide":1,"text":"Hello"} ] }',
     CANON),
    ("block_comment_between",
     '{ "actions": [ {"type":"add_text_box","slide":1,"text":"Hello"} ] /* trailing note */ }',
     CANON),
    ("comment_before_comma",
     '{"a":1 /*c*/, "b":2}', {"a": 1, "b": 2}),
    ("comment_quote_inside",
     '{ "a": 1 /* he said "hi" */ }', {"a": 1}),
    ("slashes_inside_string_preserved",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"see http://a.com/b"}]}',
     {"actions": [{"type": "add_text_box", "slide": 1,
                   "text": "see http://a.com/b"}]}),
    ("blockstars_inside_string_preserved",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"a /* not */ b"}]}',
     {"actions": [{"type": "add_text_box", "slide": 1,
                   "text": "a /* not */ b"}]}),

    # ---- 7. trailing commas ----
    ("trailing_comma_obj", '{"a":1,}', {"a": 1}),
    ("trailing_comma_arr", "[1,2,3,]", [1, 2, 3]),
    ("trailing_comma_nested",
     '{"a":[1,2,],"b":{"c":3,},}', {"a": [1, 2], "b": {"c": 3}}),
    ("trailing_comma_ws_before_close", "[1, 2 ,\n\t ]", [1, 2]),
    ("trailing_comma_after_escaped_quote",
     '{"x":"a\\"b",}', {"x": 'a"b'}),
    ("trailing_comma_canon",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"Hello",},],}', CANON),

    # ---- structural / clean ----
    ("already_clean", CANON_JSON, CANON),
    ("whitespace_padding", "   \n\t" + CANON_JSON + "\n  ", CANON),
    ("array_root_prose", "Here:\n[{\"type\":\"x\"}]\nThanks", [{"type": "x"}]),
    ("escaped_quote_clean",
     '{"actions":[{"type":"t","slide":1,"text":"a \\"b\\" c"}]}',
     {"actions": [{"type": "t", "slide": 1, "text": 'a "b" c'}]}),
    ("braces_inside_string",
     '{"actions":[{"type":"add_text_box","slide":1,"text":"render {x} and [y]"}],}',
     {"actions": [{"type": "add_text_box", "slide": 1,
                   "text": "render {x} and [y]"}]}),
    ("deep_nested",
     '{"a":{"b":{"c":[1,{"d":2,},],},},}',
     {"a": {"b": {"c": [1, {"d": 2}]}}}),

    # ---- combinations ----
    ("fence_prose_trailingcomma",
     "Sure thing!\n```json\n" + '{"actions":[{"type":"add_text_box","slide":1,"text":"Hello",},],}'
     + "\n```\nDone.",
     CANON),
    ("bom_fence_prose_trailing_mega",
     BOM + "Absolutely! Here you go:\n```json\n"
     + '{ "actions": [ { "type": "add_text_box", "slide": 1, "text": "Hello", }, ], }'
     + "\n```\nHope this helps!",
     CANON),
    ("curly_fence_trailingcomma",
     "```json\n" + _curly('{"actions":[{"type":"add_text_box","slide":1,"text":"Hello",},],}')
     + "\n```",
     CANON),
    ("prose_comment_fence",
     "Here is the result:\n```json\n"
     + '{ /* batch */ "actions": [ {"type":"add_text_box","slide":1,"text":"Hello"} ] }'
     + "\n```",
     CANON),

    # ---- adversarial ----
    ("prose_with_braces_before",
     'I first tried {"foo": 1} but the correct output is:\n' + CANON_JSON,
     CANON),
    ("prose_with_brackets_before",
     "Considered [a, b] earlier. Final answer:\n" + CANON_JSON,
     CANON),
    ("json_then_braces_in_prose_after",
     CANON_JSON + "\nNote: replace {placeholder} as needed.",
     CANON),
    ("two_jsonish_pick_real",
     'Draft: {"actions": []}\n\nFinal:\n' + CANON_JSON,
     CANON),

    # ---- no-brace contract: returned VERBATIM ----
    ("nobrace_refusal", "Sorry, I can't do that.", VERBATIM),
    ("nobrace_parens_only", "function(a, b)", VERBATIM),
    ("nobrace_empty", "", VERBATIM),
    ("nobrace_whitespace", "   ", VERBATIM),
    ("nobrace_prose", "No actions are required for this request.", VERBATIM),
]


def open_app():
    import win32com.client
    return win32com.client.DispatchEx("PowerPoint.Application")


def main():
    app = open_app()
    app.Visible = True
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)

    passed = 0
    failed = 0
    failures = []
    try:
        for name, raw, expected in CORPUS:
            try:
                out = app.Run("PPT_AI_Editor!SanitizeJsonInput", raw)
            except Exception as e:  # noqa: BLE001
                failed += 1
                failures.append(name)
                print(f"  FAIL [{name}] COM error: {e!r}")
                continue

            if expected is VERBATIM:
                if out == raw:
                    passed += 1
                    print(f"  ok   [{name}] (verbatim)")
                else:
                    failed += 1
                    failures.append(name)
                    print(f"  FAIL [{name}] expected verbatim")
                    print(f"    raw: {raw!r}")
                    print(f"    out: {out!r}")
                continue

            try:
                parsed = json.loads(out)
            except Exception as e:  # noqa: BLE001
                failed += 1
                failures.append(name)
                print(f"  FAIL [{name}] not parseable: {e}")
                print(f"    out: {out!r}")
                continue

            if parsed == expected:
                passed += 1
                print(f"  ok   [{name}]")
            else:
                failed += 1
                failures.append(name)
                print(f"  FAIL [{name}] parsed != expected")
                print(f"    expected: {expected!r}")
                print(f"    parsed:   {parsed!r}")
                print(f"    raw:      {raw!r}")
                print(f"    out:      {out!r}")
    finally:
        try:
            carrier.Saved = True
        except Exception:
            pass
        try:
            carrier.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(0.5)

    total = passed + failed
    pct = (passed / total * 100.0) if total else 0.0
    print()
    print(f"sanitizer corpus: {passed}/{total} passed ({pct:.1f}%)")
    if failures:
        print(f"failing: {', '.join(failures)}")

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
