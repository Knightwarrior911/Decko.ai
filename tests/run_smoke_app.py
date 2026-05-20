"""SP1 metric. Runs the four SP1 gates; exit 0 only at 100%.
  store_unit    : pytest tests/app/test_store.py
  llmclient_unit: pytest tests/app/test_llm_client.py
  core_loop     : tests/run_smoke_app_core_loop.py   (COM, stub LLM)
  packaging_smoke: tests/run_smoke_app_packaging.py  (PyInstaller + COM)
UI layer is NOT gated (manual screenshot, spec §8)."""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

GATES = [
    ("store_unit",     [sys.executable, "-m", "pytest", "-q",
                         "tests/app/test_store.py"]),
    ("llmclient_unit", [sys.executable, "-m", "pytest", "-q",
                         "tests/app/test_llm_client.py"]),
    ("core_loop",      [sys.executable, "tests/run_smoke_app_core_loop.py"]),
    ("templates",      [sys.executable,
                        "tests/run_smoke_app_templates.py"]),
    ("packaging_smoke",[sys.executable, "tests/run_smoke_app_packaging.py"]),
    ("ui_polish",      [sys.executable, "tests/run_smoke_ui_polish.py"]),
]


def main() -> int:
    failed = []
    for name, cmd in GATES:
        print(f"=== {name} ===")
        if subprocess.run(cmd, cwd=REPO).returncode != 0:
            failed.append(name)
    print("\nRESULT:", "PASS" if not failed else f"FAIL {failed}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
