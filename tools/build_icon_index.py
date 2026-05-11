"""Build local Fluent UI icon manifest from unpkg.

Outputs (under repo /data):
  icons_index.json     -- name -> sorted sizes; name_size -> styles list
  icons_allowed.txt    -- newline-delimited list, one entry per name+size+style
                          combination available, format "name (sizes: 16,20,24,48; styles: filled,regular)"

Run when needed (network access to unpkg required, no GitHub / no npm):
    python tools/build_icon_index.py
"""
import json
import re
import sys
import urllib.request
from collections import defaultdict
from pathlib import Path

URL = "https://unpkg.com/@fluentui/svg-icons/?meta"

REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = REPO_ROOT / "data"
JSON_OUT = OUT_DIR / "icons_index.json"
TXT_OUT = OUT_DIR / "icons_allowed.txt"

PATTERN = re.compile(r"^(.+?)_(\d+)_([a-z]+)\.svg$")


def main():
    print(f"GET {URL}")
    req = urllib.request.Request(URL, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        meta = json.load(r)

    # unpkg returns a flat {"files": [{"path": "/...", ...}, ...]} structure.
    files = [entry["path"] for entry in meta.get("files", []) if "path" in entry]
    print(f"  {len(files)} files in package (version {meta.get('version', '?')})")

    icon_files = [f.rsplit("/", 1)[-1] for f in files
                  if f.startswith("/icons/") and f.endswith(".svg")]
    print(f"  {len(icon_files)} svg icons")

    names_to_sizes = defaultdict(set)
    namesize_to_styles = defaultdict(set)
    skipped = 0
    for fname in icon_files:
        m = PATTERN.match(fname)
        if not m:
            skipped += 1
            continue
        name, size, style = m.group(1), int(m.group(2)), m.group(3)
        names_to_sizes[name].add(size)
        namesize_to_styles[f"{name}_{size}"].add(style)
    if skipped:
        print(f"  {skipped} filenames did not match pattern (ignored)")

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    index = {
        "source": URL,
        "version": meta.get("version", "unknown"),
        "icon_count": len(names_to_sizes),
        "names": {n: sorted(s) for n, s in sorted(names_to_sizes.items())},
        "styles": {k: sorted(v) for k, v in sorted(namesize_to_styles.items())},
    }
    JSON_OUT.write_text(json.dumps(index, indent=2), encoding="utf-8")
    print(f"wrote {JSON_OUT}  ({JSON_OUT.stat().st_size // 1024} KB, {len(names_to_sizes)} icons)")

    # Icons irrelevant for investment banking / professional decks.
    # Removed categories: household appliances, gaming, animals, battery/device
    # status, weather, consumer personal, media player controls, emoji/decorative,
    # accessibility UI, touch/mouse UI, MS app-specific, lab science, sparkle dupes.
    IB_BLOCKLIST_PREFIXES = (
        "battery_", "animal_", "device_meeting_room",
        "multiplier_", "skip_back_", "skip_forward_",
        "speaker_", "sound_wave_", "video_background_effect", "video_bluetooth",
        "video_usb", "number_circle_", "weather_", "temperature",
        "building_mosque", "building_townhouse",
        "phone_desktop", "phone_laptop",
        "desktop_arrow_down", "desktop_arrow_down_off", "desktop_mac", "desktop_pulse",
        "window_fingerprint", "window_new",
        "barcode_scanner",
        "calendar_sparkle", "chat_sparkle", "circle_sparkle", "clock_sparkle",
        "document_one_page_multiple_sparkle", "document_one_page_sparkle",
        "document_sparkle", "glance_horizontal_sparkle", "hexagon_sparkle",
        "info_sparkle", "mic_sparkle", "notepad_sparkle", "pen_sparkle",
        "rectangle_landscape_sparkle", "slide_text_sparkle",
        "square_hint_sparkles", "text_bullet_list_square_sparkle",
        "text_grammar_lightning", "checkmark_starburst", "add_starburst",
        "person_starburst",
    )
    IB_BLOCKLIST_EXACT = {
        # Household / personal
        "couch", "dishwasher", "door", "door_arrow_right", "drink_bottle",
        "drink_bottle_off", "elevator", "fireplace", "food", "food_chicken_leg",
        "oven", "showerhead", "spatula_spoon", "swimming_pool", "washer",
        "backpack", "beach", "glasses", "glasses_off", "hat_graduation",
        "luggage", "patient", "sticker", "road_cone",
        # Gaming
        "games", "tetris_app", "xbox_controller", "xbox_controller_error",
        # Device/UI
        "bluetooth", "brightness_high", "brightness_low", "cursor", "cursor_hover",
        "dialpad", "inking_tool", "lasso", "tap_double", "tap_single",
        "picture_in_picture", "closed_caption", "auto_fit_height", "auto_fit_width",
        "autocorrect", "incognito", "mobile_optimized", "shape_organic", "space_3d",
        # Accessibility UI
        "accessibility", "accessibility_checkmark",
        # MS app-specific
        "approvals_app", "fluent", "work_iq", "breakout_room", "tetris_app",
        # Science (non-financial)
        "beaker", "beaker_off", "molecule", "planet",
        # Decorative / consumer
        "clover", "gift_open", "puzzle_cube", "puzzle_cube_piece", "puzzle_piece",
        "rocket", "water", "emoji", "emoji_hand", "emoji_hint", "emoji_sparkle",
        "device_eq",
        # Media
        "mic_link", "mic_pulse", "mic_pulse_off", "mic_sync",
        "paint_brush", "paint_brush_subtract",
    }

    def is_ib_relevant(name):
        for prefix in IB_BLOCKLIST_PREFIXES:
            if name.startswith(prefix):
                return False
        return name not in IB_BLOCKLIST_EXACT

    # Build slim allow-list: only icons available at size=32 + style=regular,
    # filtered to IB-relevant names only.
    slim = sorted(
        name for name in names_to_sizes
        if 32 in names_to_sizes[name]
        and "regular" in namesize_to_styles.get(f"{name}_32", set())
        and is_ib_relevant(name)
    )
    lines = [
        f"# Fluent UI icons available at size=32 regular  (version {meta.get('version','?')})",
        f"# {len(slim)} names. Use EXACT spelling. Semantic substitution encouraged if",
        "# user concept has no direct match.",
        "",
    ]
    lines.extend(slim)
    TXT_OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {TXT_OUT}  ({TXT_OUT.stat().st_size} bytes, {len(slim)} names)")


if __name__ == "__main__":
    try:
        main()
    except urllib.error.URLError as e:
        print(f"FAIL: network error reaching unpkg: {e}", file=sys.stderr)
        sys.exit(1)
