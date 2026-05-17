"""core_loop: snapshot -> stub-LLM -> ExecuteFromString -> assert the
engine summary. No network. Single isolated PowerPoint run with the
project's transient-COM retry. The deck is created with one slide
(0-slide decks are guarded by DeckController.EmptyDeckError and are
out of scope for this loop). Exit 0 only on PASS."""
import os
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from app.deck_controller import DeckController          # noqa: E402
from tests.app.stub_llm import StubLLM                  # noqa: E402
from app.orchestrator import ChatOrchestrator           # noqa: E402
from app.store import Store                             # noqa: E402


def run_once() -> list[str]:
    fails: list[str] = []
    dc = DeckController()
    dc.start()
    try:
        pres = dc.app.Presentations.Add()
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        # Realistic deck: one blank slide (ppLayoutBlank = 12).
        pres.Slides.Add(1, 12)
        pres.Windows(1).Activate()
        time.sleep(1.0)
        dc.deck = pres

        st = Store(Path(os.environ["TEMP"]) / "decko_coreloop.db")
        if st.db_path.exists():
            st.db_path.unlink()
        st.init()
        orch = ChatOrchestrator(deck=dc, llm=StubLLM(), store=st)
        res = orch.run("make a title slide")

        s = res["summary"] or ""
        if "applied" not in s.lower():
            fails.append(f"summary missing 'applied': {s!r}")
        if "FAILURES" in s:
            fails.append(f"engine reported FAILURES: {s!r}")
        if pres.Slides.Count < 1:
            fails.append("no slide present after turn")
        if len(st.list_turns()) != 1:
            fails.append("turn not persisted")
        try:
            pres.Saved = True
            pres.Close()
        except Exception:
            pass
        return fails
    finally:
        dc.close(save_deck=False)


def main() -> int:
    last = None
    for attempt in range(1, 4):
        try:
            fails = run_once()
            break
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  retry transient (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    else:
        print(f"FAIL: core_loop failed after retries: {last!r}")
        return 1
    if fails:
        for f in fails:
            print(f"  FAIL [core_loop] {f}")
        print("\nRESULT: FAIL")
        return 1
    print("  core_loop: snapshot -> stub LLM -> ExecuteFromString OK")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
