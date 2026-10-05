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
import time
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler

ROOT = 'wh'
_TERMINAL = "%(asctime)s %(levelname)-7s %(name)-10s| %(message)s"
_FILE = "%(asctime)s %(levelname)-7s %(name)-10s %(threadName)s | %(message)s"


def setup(base_dir):
    """Terminal and file logging, once. A copy built without a console (no sys.stdout) logs to the file only."""
    root = logging.getLogger(ROOT)
    if root.handlers:
        return root
    root.setLevel(logging.DEBUG)
    root.propagate = False
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
        folder = os.path.join(base_dir, 'logs')
        os.makedirs(folder, exist_ok=True)
        file = RotatingFileHandler(os.path.join(folder, 'warehouse.log'), maxBytes=2_000_000, backupCount=3, encoding='utf-8')
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
