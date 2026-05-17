# PyInstaller spec for Decko Desktop. Bundles the carrier + web/ assets.
from pathlib import Path

REPO = Path(SPECPATH).parent

a = Analysis(
    [str(REPO / "app" / "main.py")],
    pathex=[str(REPO)],
    binaries=[],
    datas=[
        (str(REPO / "PPT_AI_Editor.pptm"), "."),
        (str(REPO / "app" / "web"), "app/web"),
    ],
    hiddenimports=["win32com", "win32com.client", "keyring.backends.Windows"],
    hookspath=[], runtime_hooks=[], excludes=[],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
          name="Decko", console=False)
