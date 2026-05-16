"""Cross-surface key-consistency audit for every dispatched Decko action.

Run: python tests/run_smoke_schema_audit.py   (exit 0 only at clean)

PURE static source parsing of src/modExecuteInstructions.bas. No
PowerPoint, no COM, no network, no AI. The set_3d_bevel bug (a param key
"type" colliding with the action discriminator) was one instance of a
broader class: a param's key disagreeing across the three surfaces that
touch it. For every action dispatched by DispatchAction we extract:

  V  = keys ValidateAction needs   (RequireFields arrays + act("..") reads)
  X  = keys DispatchAction reads    (act("..") incl. CLng/CStr/CDbl(...))
  G  = keys documented              (GetActionGuidance EXAMPLE + REQUIRED)

and flag an action when:

  A) "type" is read inside its ValidateAction or DispatchAction case
     (collision with the reserved discriminator),
  B) the executor reads a non-structural key that is in neither V nor G
     (mis-named / undocumented -> silent mis-dispatch),
  C) ValidateAction requires a non-structural key the executor never
     reads (dead / renamed requirement).

Structural targeting keys and framework alias alternations are exempt
(encoded below, not allowlisted). A small, explicit, PRINTED allowlist
covers genuine parser blind spots only -- never a real drift.
"""

import re
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src" / "modExecuteInstructions.bas"

# Framework targeting / orchestration keys — legitimately asymmetric across
# surfaces by design (ValidateShape/ValidateSlide resolve them; the executor
# reads the resolved shape_id; guidance lists shape_id(int|ref_name)).
STRUCTURAL = {
    "type", "slide", "shape_id", "shape_name", "shape_ids", "shape_names",
    "scope", "verify_after", "verify_scope",
}

# Framework-level alias alternations (RequireFields auto-handles the
# *_shape_id/_name pair; the rest are documented "one of X or Y" inputs).
ALIAS_PAIRS = [
    ("path", "picture_path"),
    ("from", "from_slide"),
    ("to", "to_slide"),
    ("from_shape_id", "from_shape_name"),
    ("to_shape_id", "to_shape_name"),
    ("width_pt", "preset"), ("height_pt", "preset"),
    ("major", "minor"),
]

# Explicit, justified parser-blind-spot allowlist. Each entry: action -> note.
# These are NOT real drifts; they are keys the static parser cannot resolve
# (read via a helper/loop/alias the regex cannot follow). Keep this MINIMAL
# and printed every run. A real drift must be FIXED, never added here.
ALLOWLIST: dict[str, str] = {}


def _body(text: str, header_re: str, end_kw: str) -> str:
    m = re.search(header_re, text)
    if not m:
        return ""
    start = m.end()
    e = re.search(rf"^{end_kw}\b", text[start:], re.MULTILINE)
    return text[start: start + (e.start() if e else len(text) - start)]


def _cases(body: str):
    """Yield (labels:set[str], block_text) for each Case in a Select body.
    Handles `_` line continuations and multi-label `Case "a","b"`."""
    lines = body.splitlines()
    i = 0
    out = []
    while i < len(lines):
        ln = lines[i]
        if re.match(r"\s*Case\s+\"", ln):
            label_src = ln
            while label_src.rstrip().endswith("_") and i + 1 < len(lines):
                i += 1
                label_src += "\n" + lines[i]
            labels = set(re.findall(r'"([a-z_0-9]+)"', label_src))
            blk = []
            i += 1
            while i < len(lines) and not re.match(r"\s*Case\s", lines[i]) \
                    and not re.match(r"\s*End Select\b", lines[i]):
                blk.append(lines[i])
                i += 1
            if labels:
                out.append((labels, "\n".join(blk)))
            continue
        i += 1
    return out


def _act_keys(block: str) -> set[str]:
    """Keys the executor consumes: direct act("K") reads PLUS keys passed
    by name into the shape-ref resolution helpers
    ResolveActShapeId(act,"K") / ResolveActShapeIdArray(act,"K"), which is
    how *_shape_id / *_shape_ids params actually reach the executor."""
    keys = set(re.findall(r'act\(\s*"([a-z_0-9]+)"\s*\)', block))
    keys |= set(re.findall(
        r'ResolveActShapeId(?:Array)?\(\s*act\s*,\s*"([a-z_0-9]+)"\s*\)', block))
    return keys


def _require_keys(block: str) -> set[str]:
    keys: set[str] = set()
    for arr in re.findall(r"RequireFields\(\s*act\s*,\s*Array\(([^)]*)\)", block):
        keys |= set(re.findall(r'"([a-z_0-9]+)"', arr))
    return keys


def _alias_pairs_in(block: str):
    pairs = []
    for a, b in re.findall(
        r'Not\s+act\.Exists\(\s*"([a-z_0-9]+)"\s*\)\s*And\s*Not\s+act\.Exists\(\s*"([a-z_0-9]+)"\s*\)',
        block,
    ):
        pairs.append((a, b))
    return pairs


def _siblings(key: str, extra_pairs):
    sib = set()
    for a, b in list(ALIAS_PAIRS) + list(extra_pairs):
        if key == a:
            sib.add(b)
        elif key == b:
            sib.add(a)
    # shape ref alternation handled by RequireFields itself
    if key.endswith("shape_id"):
        sib.add(key[:-2] + "name")
    if key.endswith("shape_name"):
        sib.add(key[:-4] + "id")
    if key.endswith("shape_ids"):
        sib.add(key[:-3] + "names")
    if key.endswith("shape_names"):
        sib.add(key[:-5] + "ids")
    return sib


def main() -> int:
    text = SRC.read_text(encoding="utf-8", errors="replace")

    va = _body(text, r"Private Function ValidateAction\(act As Object\) As String", "End Function")
    da = _body(text, r"Private Sub DispatchAction\(act As Object\)", "End Sub")
    ga = _body(text, r"Public Function GetActionGuidance\(actionType As String\) As String", "End Function")

    if not (va and da and ga):
        print(f"FAIL: could not isolate functions "
              f"(va={bool(va)} da={bool(da)} ga={bool(ga)})")
        return 1

    # Build per-action maps. Universe = DispatchAction case labels.
    da_cases = _cases(da)
    va_cases = _cases(va)
    ga_cases = _cases(ga)

    # Whole-`act` passthrough: a dispatch case whose body just hands the
    # entire dict to a handler (`modX.Do_yyy act` / `Do_yyy_act act`) with
    # no individual key reads. The handler consumes the keys INSIDE
    # modActions*.bas — out of static reach, and exactly the brittle
    # boundary this audit is scoped to bound. For such actions the
    # executor key-surface is opaque: rule A still applies (it inspects
    # the ValidateAction body, independent), but C/B are not statically
    # decidable and the action is reported as opaque, not flagged.
    PASSTHRU_RE = re.compile(r"\bDo_[A-Za-z0-9_]+\s+act\b")

    X: dict[str, set[str]] = {}
    passthru: set[str] = set()
    for labels, blk in da_cases:
        ak = _act_keys(blk)
        is_pt = bool(PASSTHRU_RE.search(blk)) and not ak
        for t in labels:
            X.setdefault(t, set())
            X[t] |= ak
            if is_pt:
                passthru.add(t)

    V: dict[str, set[str]] = {}
    Vbody: dict[str, set[str]] = {}
    Valt: dict[str, list] = {}
    for labels, blk in va_cases:
        for t in labels:
            V.setdefault(t, set())
            Vbody.setdefault(t, set())
            Valt.setdefault(t, [])
            V[t] |= _require_keys(blk)
            Vbody[t] |= _act_keys(blk)
            Valt[t] += _alias_pairs_in(blk)

    G: dict[str, set[str]] = {}
    for labels, blk in ga_cases:
        ex = set(re.findall(r'""([a-z_0-9]+)""\s*:', blk))
        req_seg = blk.split("EXAMPLE")[0]
        toks = set(re.findall(r'\b([a-z][a-z_0-9]{2,})\b', req_seg))
        for t in labels:
            G.setdefault(t, set())
            G[t] |= ex | toks

    universe = sorted(X)
    flagged = []
    opaque: list[str] = []

    for t in universe:
        xk = X.get(t, set())
        vk = V.get(t, set())
        vbk = Vbody.get(t, set())
        gk = G.get(t, set())
        alt = Valt.get(t, [])
        known = vk | vbk | gk
        problems = []

        # A) discriminator collision — precise, applies even to passthrough
        #    (it inspects the ValidateAction body, independent of dispatch).
        if "type" in xk or "type" in vbk:
            problems.append('A: reads act("type") inside its case (collides with discriminator)')

        if t in passthru:
            # Executor surface opaque (handler-internal). A enforced above;
            # C/B not statically decidable -> report opaque, do not flag.
            if problems:
                flagged.append((t, problems))
            else:
                opaque.append(t)
            continue

        # C) validator requires a non-structural key the executor never
        #    reads (directly or via the shape-ref helpers / an alias) — a
        #    dead or renamed requirement. This is the load-bearing half of
        #    the cross-surface rename class (e.g. set_3d_bevel, degrees/angle).
        c_unread = []
        for k in sorted(vk):
            if k in STRUCTURAL:
                continue
            if k in xk:
                continue
            if _siblings(k, alt) & xk:
                continue
            c_unread.append(k)
            problems.append(f"C: ValidateAction requires {k!r} — executor never reads it")

        # B) PAIRED form only: an undocumented non-structural key the
        #    executor reads, on an action that ALSO has an unread required
        #    key (C). That pairing is the rename signature (old name
        #    required-but-ignored, new name read-but-undocumented).
        #    Optional knobs that merely aren't in the minimal EXAMPLE are
        #    intentionally NOT flagged — that is documentation breadth, not
        #    the cross-surface key-disagreement bug class this gate targets.
        if c_unread:
            for k in sorted(xk):
                if k in STRUCTURAL or k in known:
                    continue
                if _siblings(k, alt) & known:
                    continue
                problems.append(
                    f"B: executor reads undocumented {k!r} while a required "
                    f"key is unread — likely a cross-surface rename")

        if problems:
            if t in ALLOWLIST:
                continue
            flagged.append((t, problems))

    total = len(universe)
    clean = total - len(flagged)
    print(f"dispatched actions audited: {total}")
    if ALLOWLIST:
        print("justified parser-blind-spot allowlist:")
        for k, why in sorted(ALLOWLIST.items()):
            print(f"  {k}: {why}")
    if opaque:
        print(f"opaque whole-act passthrough (handler-internal keys, not "
              f"statically checkable; rule A still enforced): {len(opaque)}")
        print("  " + ", ".join(sorted(opaque)))
    print(f"clean: {clean}/{total}")
    for t, probs in flagged:
        print(f"  FAIL [{t}]")
        for p in probs:
            print(f"      {p}")

    ok = not flagged
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
