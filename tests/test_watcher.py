import pytest
import time
import tempfile
from pathlib import Path
import pandas as pd
from mcp_excel.watcher import FileWatcher


@pytest.fixture
def temp_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_excel(temp_dir):
    file_path = temp_dir / "test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30]
    })
    df.to_excel(file_path, index=False)
    return file_path


def test_watcher_creation(temp_dir):
    callback_called = []

    def callback():
        callback_called.append(True)

    watcher = FileWatcher(temp_dir, callback)
    assert not watcher.is_running()

    watcher.start()
    assert watcher.is_running()

    watcher.stop()
    assert not watcher.is_running()


def test_watcher_detects_file_modification(temp_dir):
    callback_count = []

    def callback():
        callback_count.append(1)

    file_path = temp_dir / "test.xlsx"
    df = pd.DataFrame({"A": [1, 2]})
    df.to_excel(file_path, index=False)

    watcher = FileWatcher(temp_dir, callback, debounce_seconds=0.5)
    watcher.start()

    time.sleep(0.2)

    df = pd.DataFrame({"A": [1, 2, 3]})
    df.to_excel(file_path, index=False)

    time.sleep(1.0)

    watcher.stop()

    assert len(callback_count) >= 1


def test_watcher_debounces_rapid_changes(temp_dir):
    callback_count = []

    def callback():
        callback_count.append(1)

    file_path = temp_dir / "test.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    watcher = FileWatcher(temp_dir, callback, debounce_seconds=0.5)
    watcher.start()

    time.sleep(0.2)

    for i in range(5):
        df = pd.DataFrame({"A": [i]})
        df.to_excel(file_path, index=False)
        time.sleep(0.1)

    time.sleep(1.0)

    watcher.stop()

    assert len(callback_count) == 1


def test_watcher_handles_new_file(temp_dir):
    callback_count = []

    def callback():
        callback_count.append(1)

    watcher = FileWatcher(temp_dir, callback, debounce_seconds=0.5)
    watcher.start()

    time.sleep(0.2)

    file_path = temp_dir / "new_file.xlsx"
    df = pd.DataFrame({"B": [10, 20]})
    df.to_excel(file_path, index=False)

    time.sleep(1.0)

    watcher.stop()

    assert len(callback_count) >= 1


def test_watcher_handles_file_deletion(temp_dir):
    callback_count = []

    def callback():
        callback_count.append(1)

    file_path = temp_dir / "test.xlsx"
    df = pd.DataFrame({"A": [1, 2]})
    df.to_excel(file_path, index=False)

    watcher = FileWatcher(temp_dir, callback, debounce_seconds=0.5)
    watcher.start()

    time.sleep(0.2)

    file_path.unlink()

    time.sleep(1.0)

    watcher.stop()

    assert len(callback_count) >= 1


def test_watcher_ignores_non_excel_files(temp_dir):
    callback_count = []

    def callback():
        callback_count.append(1)

    watcher = FileWatcher(temp_dir, callback, debounce_seconds=0.5)
    watcher.start()

    time.sleep(0.2)

    txt_file = temp_dir / "test.txt"
    txt_file.write_text("not an excel file")

    time.sleep(1.0)

    watcher.stop()

    assert len(callback_count) == 0
