"""PowerPoint COM. Two modes (spec D5):
  attach -> operate on the user's already-open ActivePresentation
  file   -> open a chosen .pptx, operate, Save
The carrier .pptm is opened (hidden) alongside so its macros
(ExecuteFromString / BuildSnapshotJson) are callable via app.Run.
Reuses the project's transient-COM retry discipline."""
import time

import pythoncom
import pywintypes
import win32com.client

from app.carrier import ensure_carrier

_TRANSIENT = (pywintypes.com_error, AttributeError)


def _open_app():
    last = None
    for _ in range(15):
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = True
            return app
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2.0)
    raise RuntimeError(f"PowerPoint COM bring-up failed: {last!r}")


class NoPowerPointError(RuntimeError):
    pass


class NoOpenDeckError(RuntimeError):
    pass


class DeckController:
    def __init__(self):
        self.app = None
        self.carrier = None
        self.deck = None

    def start(self):
        pythoncom.CoInitialize()
        try:
            self.app = _open_app()
        except RuntimeError as e:
            raise NoPowerPointError(str(e))
        self.carrier = self.app.Presentations.Open(
            str(ensure_carrier()), WithWindow=False)

    def attach_open_deck(self):
        # The active deck must be a non-carrier presentation.
        for p in self.app.Presentations:
            if p.FullName != self.carrier.FullName:
                self.deck = p
                p.Windows(1).Activate()
                return
        raise NoOpenDeckError("No deck open in PowerPoint.")

    def open_file(self, path: str):
        self.deck = self.app.Presentations.Open(path, WithWindow=True)
        self.deck.Windows(1).Activate()

    def get_snapshot(self) -> str:
        return self._run("BuildSnapshotJson")

    def run_actions(self, actions_json: str) -> str:
        # ExecuteFromString runs the verify loop by default; returns a
        # human summary incl. "FAILURES (N)" contract.
        return self._run("ExecuteFromString", actions_json)

    def _run(self, macro: str, *args, _attempts: int = 3):
        last = None
        for i in range(1, _attempts + 1):
            try:
                return self.app.Run(f"PPT_AI_Editor!{macro}", *args)
            except _TRANSIENT as e:  # noqa: PERF203
                last = e
                time.sleep(2.0 * i)
        raise RuntimeError(f"{macro} failed after retries: {last!r}")

    def close(self, save_deck: bool = False):
        try:
            if self.deck is not None and save_deck:
                self.deck.Save()
        finally:
            for p in (self.carrier,):
                try:
                    if p is not None:
                        p.Saved = True
                        p.Close()
                except Exception:
                    pass
            try:
                self.app.Quit()
            except Exception:
                pass
            time.sleep(1.0)
