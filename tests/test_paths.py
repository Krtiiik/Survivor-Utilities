"""survivor_app.ui.paths has no PySide6 dependency (just os/sys), so it's cheap to
test directly. Added after a real bug: app_dir() originally went up only two
directory levels, resolving to survivor_app/ instead of the repository root --
invisible to the GUI smoke tests because they monkeypatch app_dir() entirely.
This test exercises the real implementation instead.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from survivor_app.ui import paths

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_app_dir_resolves_to_repo_root_when_not_frozen():
    assert os.path.normcase(paths.app_dir()) == os.path.normcase(REPO_ROOT)


def test_config_and_counts_paths_are_next_to_app_dir():
    assert os.path.dirname(paths.config_path()) == paths.app_dir()
    assert os.path.basename(paths.config_path()) == "config.json"
    assert os.path.dirname(paths.counts_path()) == paths.app_dir()
    assert os.path.basename(paths.counts_path()) == "counts.json"
    assert os.path.dirname(paths.distributions_path()) == paths.app_dir()
    assert os.path.basename(paths.distributions_path()) == "distributions.json"


def test_bundled_example_config_exists():
    assert os.path.exists(paths.bundled_example_config_path())


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"ok  {name}")
    print("All path tests passed.")
