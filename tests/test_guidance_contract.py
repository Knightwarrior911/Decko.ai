"""Static guidance<->validator contract. No COM. For every action in
modExecuteInstructions.GetActionGuidance, assert its REQUIRED list +
EXAMPLE keys cover everything its ValidateAction RequireFields requires
(shape_id<->shape_name and explicit `act.Exists(a)/Exists(b)` groups
are alternatives). Exit 1 on any drift, misleading extra, or an
unparseable-case-set change. Source of truth = the validator."""
import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
BAS = (REPO / "src" / "modExecuteInstructions.bas").read_text(
    encoding="utf-8", errors="replace")

# Cases whose required fields are NOT a literal RequireFields(Array(...))
# — validated dynamically. Each must have a one-line reason. The test
# FAILS if the actually-unparseable set differs from this allowlist.
KNOWN_UNPARSEABLE = {
    "add_connector":          "from/to/kind/slide validated via ElseIf act.Exists chain; no RequireFields",
    "apply_template":         "validated via explicit act.Exists + modActionsTemplate helper",
    "build_deck_from_spec":   "validated via explicit act.Exists + modActionsSpec helper",
    "build_image_picker_slide": "no required fields; folder optional, falls back to g_LastFetchFolder",
    "capture_template":       "validated via explicit act.Exists(\"name\") check",
    "delete_template":        "validated via explicit act.Exists(\"name\") check",
    "extract_spec":           "reads live deck; no required fields",
    "generate_variants":      "validated via complex act.Exists chain (template|templates + n)",
    "insert_picture":         "path|picture_path + pos + slide via ElseIf act.Exists chain; no RequireFields",
    "list_templates":         "no required fields",
    "open_image_picker":      "no required fields; folder optional, falls back to last fetch",
    "rename_template":        "validated via explicit act.Exists(\"from\")/act.Exists(\"to\") chain",
    "run_verification":       "no required fields; scope and max_warnings are optional",
    "scan_palette":           "no required fields; optional scope validated inline (not RequireFields)",
    "set_slide_size":         "width_pt+height_pt OR preset logic; no RequireFields call",
    "set_theme_font":         "major OR minor logic; no RequireFields call",
}


def _case_blocks(func_src: str):
    """Yield (tuple_of_action_names, block_text) for each Case in a
    Select Case body. Handles grouped + line-continued Case labels."""
    joined = re.sub(r"_\r?\n\s*", " ", func_src)
    parts = re.split(r"\n\s*Case ", joined)
    for p in parts[1:]:
        head, _, body = p.partition("\n")
        if head.strip().startswith("Else"):
            continue
        names = re.findall(r'"([^"]+)"', head)
        if names:
            yield tuple(names), body


def _slice_func(name: str) -> str:
    m = re.search(r"(?:Public |Private )?Function " + re.escape(name)
                  + r"\b.*?\nEnd Function", BAS, re.S)
    assert m, f"function {name} not found"
    return m.group(0)


_NON_FIELD_WORDS = {
    "and", "or", "not", "none", "n/a", "-", "no", "all", "at", "least",
    "one", "of", "both", "smart", "only", "note", "tip", "inner", "outer",
    "usually", "multiple", "non", "empty", "unique", "false", "true",
    "optional", "non-empty", "1-based", "0-based", "1-based)", "0-based)",
    "0-based;", "string", "int", "num", "bool", "array", "object",
}


def _clean_req_token(raw: str) -> str:
    """Strip type annotations from a REQUIRED field token.

    REQUIRED lines use inline annotation like:
      value(string), shape_id(int|ref_name), scope(""deck""|""slide:N"")
    We want just the field name (snake_case) before the first '(' or space.
    Returns "" for tokens that are clearly not field names.
    """
    t = raw.strip().strip(".,;:-")
    # Take only the part before the first '(' or whitespace
    t = re.split(r"[()\s]", t)[0].strip(".,;:-")
    # Reject tokens with special chars not valid in field names
    if re.search(r"[#{}><|/\\+=!@$%^&*\"]", t):
        return ""
    # Reject numeric/basedness tokens
    if re.match(r"^\d", t):
        return ""
    # Reject English prose words, not field names
    if t.lower() in _NON_FIELD_WORDS:
        return ""
    # Must look like a snake_case identifier
    if not re.match(r"^[a-zA-Z][a-zA-Z0-9_]*$", t):
        return ""
    return t


def _parse_required_line(line: str) -> set:
    """Extract field names from a REQUIRED: guidance line.

    Handles:
    - Type annotations:   value(string) -> value
    - Bracket groups:     pos({left,top,width,height}) -> pos  (stop at {)
    - VBA string escapes: scope(""deck""|""slide:N"") -> scope
    - At-least-one notes: shape_id, AND at least one of: x, y, z  -> {shape_id}
                          (stops before AND)
    """
    # Replace VBA double-quote escapes so they don't affect token splitting
    line = line.replace('""', "'")
    # Stop at "AND at least one of:" — everything after is optional
    line = re.split(r"\bAND\b", line, flags=re.IGNORECASE)[0]
    # Split on commas that are NOT inside brackets
    tokens = []
    depth = 0
    cur = []
    for ch in line:
        if ch in "({[":
            depth += 1
            cur.append(ch)
        elif ch in ")}]":
            depth = max(0, depth - 1)
            cur.append(ch)
        elif ch == "," and depth == 0:
            tokens.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    if cur:
        tokens.append("".join(cur))
    result = set()
    for tok in tokens:
        t = _clean_req_token(tok)
        if t:
            result.add(t)
    return result


def _guidance_map():
    """action -> (set(required_tokens), set(example_keys))."""
    out = {}
    src = "".join(_slice_func(f) for f in
                  ("GetActionGuidance", "GetActionGuidance_Part2",
                   "GetActionGuidance_Part3")
                  if re.search(r"Function " + f + r"\b", BAS))
    for names, body in _case_blocks(src):
        text = body
        req = set()
        # Capture REQUIRED line: go to end of VBA string.
        # VBA uses "" to escape a literal " inside a string literal.
        # Replace "" with a placeholder so the lone " (end of VBA string) is
        # the only remaining " to act as a stop character.
        escaped_text = text.replace('""', "\x00")
        # Now the lone " marks the end of the VBA string containing REQUIRED
        m = re.search(r"REQUIRED:\s*(.*?)(?:\"|vbCrLf|\n)", escaped_text)
        if m:
            # Remove placeholder (were VBA escaped quotes — not actual content)
            req_line = m.group(1).replace("\x00", "")
            req = _parse_required_line(req_line)
        ex = set()
        em = re.search(r"EXAMPLE:\s*(\{.*?\})", text)
        if em:
            raw = em.group(1).replace('""', '"')
            try:
                ex = set(json.loads(raw).keys())
            except Exception:
                ex = set(re.findall(r'""(\w+)""\s*:', body))
        for a in names:
            out[a] = (req, ex)
    return out


def _validator_map():
    """action -> set(required) or None (unparseable/allowlisted).

    Returns None for:
    - Cases with no RequireFields AND no parseable pairs
    - Cases explicitly listed in KNOWN_UNPARSEABLE (even if partially parsed)
    """
    out = {}
    va = _slice_func("ValidateAction")
    for names, body in _case_blocks(va):
        # Check if any of the names in this group is allowlisted
        if any(n in KNOWN_UNPARSEABLE for n in names):
            for a in names:
                out[a] = None
            continue
        block = body.split("\n        Case ")[0]
        req = None
        rf = re.search(r"RequireFields\(act,\s*Array\(([^)]*)\)\)",
                       block, re.S)
        if rf:
            req = set(re.findall(r'"([^"]+)"', rf.group(1)))
        else:
            pairs = re.findall(
                r'Not act\.Exists\("([^"]+)"\)\s+And\s+Not '
                r'act\.Exists\("([^"]+)"\)', block)
            if pairs:
                req = set()
                for a, b in pairs:
                    req.add(a + "|" + b)
        for a in names:
            out[a] = req
    return out


def _satisfied(field: str, have: set) -> bool:
    if "|" in field:
        return any(p in have for p in field.split("|"))
    if field in have:
        return True
    if "shape_ids" in field:
        return field.replace("shape_ids", "shape_names") in have
    if "shape_id" in field:
        return field.replace("shape_id", "shape_name") in have
    return False


def main() -> int:
    g = _guidance_map()
    v = _validator_map()
    fails, unparseable = [], set()
    for act, req in v.items():
        if req is None:
            unparseable.add(act)
            continue
        if act not in g:
            fails.append(f"{act}: no GetActionGuidance case")
            continue
        greq, gex = g[act]
        have = greq | {x for f in greq for x in f.split("|")}
        for f in req:
            if not _satisfied(f, have):
                fails.append(f"{act}: validator requires '{f}' but "
                             f"guidance REQUIRED={sorted(greq)}")
            if not _satisfied(f, gex):
                fails.append(f"{act}: validator requires '{f}' but "
                             f"EXAMPLE keys={sorted(gex)}")
        valt = set()
        for f in req:
            valt.add(f)
            if "shape_id" in f:
                valt.add(f.replace("shape_id", "shape_name"))
            if "shape_ids" in f:
                valt.add(f.replace("shape_ids", "shape_names"))
            valt |= set(f.split("|"))
        for t in greq:
            if t not in valt and not any(
                    t in f.split("|") or
                    t == f.replace("shape_id", "shape_name") or
                    t == f.replace("shape_ids", "shape_names")
                    for f in req):
                fails.append(f"{act}: guidance REQUIRED lists '{t}' "
                             f"which the validator does NOT require")

    expected_unp = set(KNOWN_UNPARSEABLE)
    if unparseable != expected_unp:
        fails.append(
            f"unparseable-case set changed: got {sorted(unparseable)}, "
            f"allowlist {sorted(expected_unp)} — review & update "
            f"KNOWN_UNPARSEABLE")

    if fails:
        print(f"GUIDANCE CONTRACT: {len(fails)} issue(s)")
        for f in fails:
            print("  FAIL", f)
        print("\nRESULT: FAIL")
        return 1
    print(f"GUIDANCE CONTRACT: {len(v)} actions, "
          f"{len(KNOWN_UNPARSEABLE)} allowlisted-unparseable, 0 drift")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
