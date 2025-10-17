import time
import threading
from pathlib import Path
from typing import Callable, Optional
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileSystemEvent

from . import logging as log


class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, callback: Callable[[], None], debounce_seconds: float = 1.0):
        self.callback = callback
        self.debounce_seconds = debounce_seconds
        self.timer: Optional[threading.Timer] = None
        self.lock = threading.Lock()

    def on_modified(self, event: FileSystemEvent):
        if not event.is_directory and event.src_path.endswith('.xlsx'):
            self._schedule_callback()

    def on_created(self, event: FileSystemEvent):
        if not event.is_directory and event.src_path.endswith('.xlsx'):
            self._schedule_callback()

    def on_deleted(self, event: FileSystemEvent):
        if not event.is_directory and event.src_path.endswith('.xlsx'):
            self._schedule_callback()

    def _schedule_callback(self):
        with self.lock:
            if self.timer:
                self.timer.cancel()

            self.timer = threading.Timer(self.debounce_seconds, self._execute_callback)
            self.timer.start()

    def _execute_callback(self):
        with self.lock:
            self.timer = None

        try:
            self.callback()
        except Exception as e:
            log.error("file_watch_callback_failed", error=str(e))


class FileWatcher:
    def __init__(self, path: Path, callback: Callable[[], None], debounce_seconds: float = 1.0):
        self.path = path
        self.callback = callback
        self.debounce_seconds = debounce_seconds
        self.observer: Optional[Observer] = None
        self.handler: Optional[ExcelFileHandler] = None

    def start(self):
        if self.observer:
            return

        self.handler = ExcelFileHandler(self.callback, self.debounce_seconds)
        self.observer = Observer()
        self.observer.schedule(self.handler, str(self.path), recursive=True)
        self.observer.start()
        log.info("file_watcher_started", path=str(self.path), debounce_seconds=self.debounce_seconds)

    def stop(self):
        if not self.observer:
            return

        self.observer.stop()
        self.observer.join(timeout=5)
        self.observer = None
        self.handler = None
        log.info("file_watcher_stopped", path=str(self.path))

    def is_running(self) -> bool:
        return bool(self.observer and self.observer.is_alive())
