"""The app's log: what each step of Packing Labels, the address book, Packing Specs and Stitch Labels did,
so a run that went wrong can be traced.

  terminal               INFO and up: one line per step, with counts and timings (WAREHOUSE_LOG_LEVEL=DEBUG for more)
  logs/warehouse.log     everything, DEBUG included (each courier label page, each address-book lookup), kept
                         across restarts: 2 MB a file, the last 3 files kept

Use:  log = app_log.get('stitch');  log.info("...")
      with app_log.step(log, "Open360 CSV"):  ...   -> "Open360 CSV done in 0.12s", or the error with its traceback
"""
import logging
import os
import sys
import threading
import time
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler

ROOT = 'wh'
LOG_FILE = 'warehouse.log'  # in logs/; the Logs page (log_view.py) reads it and its rotated copies
LOG_BACKUPS = 3
_TERMINAL = "%(asctime)s %(levelname)-7s %(name)-10s| %(message)s"
_FILE = "%(asctime)s %(levelname)-7s %(name)-10s %(threadName)s | %(message)s"


_errors = threading.local()


class _RememberErrors(logging.Handler):
    """Keeps the last error each thread logged (when, what), so a failed page can land on it on the Logs page."""

    def __init__(self):
        super().__init__(logging.ERROR)

    def emit(self, record):
        try:
            _errors.last = (time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(record.created)), record.getMessage().strip())
        except Exception:
            pass


def last_error():
    """(time, message) of the last error this thread logged since forget_error(), or None."""
    return getattr(_errors, 'last', None)


def forget_error():
    _errors.last = None


def log_dir(base_dir):
    """logs/ in the app folder, where the file log goes and the Logs page reads it."""
    return os.path.join(base_dir, 'logs')


def setup(base_dir):
    """Terminal and file logging, once. A copy built without a console (no sys.stdout) logs to the file only."""
    root = logging.getLogger(ROOT)
    if root.handlers:
        return root
    root.setLevel(logging.DEBUG)
    root.propagate = False
    root.addHandler(_RememberErrors())
    stream = sys.stdout
    if stream is not None:
        try:
            stream.reconfigure(errors='replace')  # a store name the Windows console can't show doesn't break the log
        except (AttributeError, ValueError):
            pass
        term = logging.StreamHandler(stream)
        term.setLevel(os.environ.get('WAREHOUSE_LOG_LEVEL', 'INFO').upper())
        term.setFormatter(logging.Formatter(_TERMINAL, "%H:%M:%S"))
        root.addHandler(term)
    try:
        folder = log_dir(base_dir)
        os.makedirs(folder, exist_ok=True)
        file = RotatingFileHandler(os.path.join(folder, LOG_FILE), maxBytes=2_000_000, backupCount=LOG_BACKUPS, encoding='utf-8')
        file.setLevel(logging.DEBUG)
        file.setFormatter(logging.Formatter(_FILE, "%Y-%m-%d %H:%M:%S"))
        root.addHandler(file)
    except OSError as e:
        root.warning("Could not open logs/warehouse.log (%s): logging to the terminal only", e)
    return root


def get(name):
    """A logger for one part of the app ('packing', 'stitch', 'book', ...), under the app's log."""
    return logging.getLogger(f"{ROOT}.{name}")


@contextmanager
def step(log, what):
    """Logs how long a step took, or that it failed (with the traceback) and re-raises."""
    started = time.perf_counter()
    log.debug("%s ...", what)
    try:
        yield
    except Exception:
        log.exception("%s FAILED after %.2fs", what, time.perf_counter() - started)
        raise
    log.info("%s done in %.2fs", what, time.perf_counter() - started)
