import sys
from pathlib import Path


def pytest_configure(config):
    """Pytest hook that runs very early, before test collection."""
    _root = Path(__file__).parent.resolve()
    _root_str = str(_root)
    if _root_str not in sys.path:
        sys.path.insert(0, _root_str)
