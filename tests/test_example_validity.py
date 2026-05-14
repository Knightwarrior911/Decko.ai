"""
test_example_validity.py — verify every EXAMPLE block in GetActionGuidance:
  1. Parses as valid JSON
  2. Contains all params listed on the REQUIRED line
  3. Has no placeholder tokens like <value>

Run: python tests/test_example_validity.py
"""

import json
import re
import sys
from pathlib import Path

import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"


def split_top_level(s: str) -> list[str]:
    """Split on commas not inside () or {}."""
    parts, depth, current = [], 0, []
    for ch in s:
        if ch in "({":
            depth += 1
            current.append(ch)
        elif ch in ")}":
            depth -= 1
            current.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(ch)
    if current:
        parts.append("".join(current).strip())
    return parts


def extract_required_params(required_line: str) -> list[str]:
    """
    Parse REQUIRED line -> list of (param_name, [alternatives]) tuples.
    Stops at 'AND at least one of:' / 'at least one' since those are optional.
    Handles '(or X)' alternatives like 'from_shape_id (or from_shape_name)'.
    Returns flat list of param names (first alternative only — caller checks all).
    """
    after_colon = required_line.split("REQUIRED:", 1)[-1].strip()
    # Truncate at "AND at least one" / "at least one" — those params are optional
    for sentinel in ("AND at least one", "at least one"):
        idx = after_colon.lower().find(sentinel.lower())
        if idx != -1:
            after_colon = after_colon[:idx].rstrip(", ")
            break
    params = []
    for chunk in split_top_level(after_colon):
        chunk = chunk.strip()
        if not chunk:
            continue
        # Take identifier before (, space, or end
        name = re.split(r"[\s\(]", chunk)[0].strip().lower()
        if name and re.match(r"^[a-z_][a-z0-9_]*$", name):
            params.append(name)
    return params


def get_alternatives(required_line: str, param: str) -> list[str]:
    """For 'from_shape_id (or from_shape_name)' return ['from_shape_id','from_shape_name']."""
    # Find the segment containing the param
    after_colon = required_line.split("REQUIRED:", 1)[-1]
    pattern = re.compile(re.escape(param) + r"\s*\(or\s+([a-z_][a-z0-9_]*)\)", re.IGNORECASE)
    m = pattern.search(after_colon)
    if m:
        return [param, m.group(1).lower()]
    return [param]


def check_guidance(action_type: str, guidance: str) -> list[str]:
    """Return list of error strings for this action (empty = pass)."""
    errors = []
    lines = guidance.splitlines()

    required_params: list[str] = []
    example_json_str: str | None = None

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("REQUIRED:"):
            required_params = extract_required_params(stripped)
        elif stripped.startswith("EXAMPLE:"):
            example_json_str = stripped[len("EXAMPLE:"):].strip()

    if example_json_str is None:
        errors.append("missing EXAMPLE line")
        return errors

    # Check for placeholder tokens
    if "<" in example_json_str and ">" in example_json_str:
        errors.append(f"placeholder token in example: {example_json_str[:120]}")
        return errors  # can't parse JSON with placeholders

    # Try JSON parse
    try:
        parsed = json.loads(example_json_str)
    except json.JSONDecodeError as e:
        errors.append(f"invalid JSON: {e} | snippet: {example_json_str[:120]}")
        return errors

    # Check required params present in example (handle "(or X)" alternatives)
    required_line_full = next((l.strip() for l in lines if l.strip().startswith("REQUIRED:")), "")
    for param in required_params:
        alternatives = get_alternatives(required_line_full, param)
        if not any(alt in parsed for alt in alternatives):
            errors.append(f"required param '{param}' missing from example")

    return errors


def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found")
        return 1

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break
    carrier = app.Presentations.Open(str(CARRIER))

    raw = app.Run("PPT_AI_Editor.pptm!modExecuteInstructions.GetAllActionTypes")
    types = sorted({t.strip() for t in raw.split(",") if t.strip()})
    print(f"Checking {len(types)} action types...")

    failures: list[tuple[str, list[str]]] = []
    for action_type in types:
        guidance = str(app.Run(
            "PPT_AI_Editor.pptm!modExecuteInstructions.GetActionGuidance",
            action_type
        ))
        errs = check_guidance(action_type, guidance)
        if errs:
            failures.append((action_type, errs))

    carrier.Close()

    if not failures:
        print(f"OK — all {len(types)} examples parse clean with required params present.")
        return 0

    print(f"\nFAIL — {len(failures)} action(s) with invalid examples:\n")
    for action_type, errs in failures:
        for e in errs:
            print(f"  {action_type}: {e}")
    print(f"\nTotal failures: {len(failures)}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
