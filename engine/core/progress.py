"""Live progress of a long action (Generate, Stitch, Packing Sheets …) for the loading overlay.

The page puts a random progress_id in the form it posts and, while the server works on it, asks
/api/progress/<id> twice a second: the step it's on, how far along (a percent when the step counts its work, such as
label 12 of 40), and the log lines that request has written (the same lines as logs/warehouse.log and the Logs page).

  app.py binds the request to its run (before_request) and ends it (teardown_request); Handler copies the run's log
  lines; a step that knows its size says so:

      progress.stage("Drawing the packing labels", 10, 90)   # this step fills the bar from 10 % to 90 %
      progress.tick(12, 40)                                  # 12 of 40 done: the bar is at 10 + 80 * 12/40 %

Everything is in memory and per request thread, and nothing happens without a progress_id, so it costs nothing to
code that runs outside a request (tests, scripts). A run is forgotten a few minutes after it ends.
"""
import logging
import re
import threading
import time
import traceback
from collections import deque

ID_RE = re.compile(r'^[A-Za-z0-9-]{8,64}$')
KEEP_FOR = 300   # seconds a finished run can still be read
MAX_RUNS = 50
MAX_LINES = 400  # every line of a run, details too (the overlay scrolls); the Logs page keeps them all

_runs = {}
_lock = threading.Lock()
_local = threading.local()


def start(run_id):
    """This request thread now works for run_id (from the page). Returns it, or None for an id that isn't one."""
    if not run_id or not ID_RE.match(run_id):
        return None
    now = time.time()
    with _lock:
        for k in [k for k, r in _runs.items() if r['finished'] and now - r['finished'] > KEEP_FOR]:
            del _runs[k]
        while len(_runs) >= MAX_RUNS:  # many abandoned runs: the oldest go
            del _runs[min(_runs, key=lambda k: _runs[k]['started'])]
        _runs[run_id] = {'started': now, 'finished': None, 'step': 'Starting…', 'from': 0.0, 'to': 0.0,
                         'done': 0, 'total': 0, 'percent': 0.0, 'lines': deque(maxlen=MAX_LINES)}
    _local.run = run_id
    return run_id


def current():
    return getattr(_local, 'run', None)


def detach():
    """This request thread stops working for its run without ending it, and returns its id: for work handed to a
    background thread (Finalise on the matching page), which attach()es to it and finish()es it when done."""
    run_id, _local.run = current(), None
    return run_id


def attach(run_id):
    """This (background) thread now works for a run another thread started (detach)."""
    _local.run = run_id if run_id and run_id in _runs else None


def finish():
    run_id, _local.run = current(), None
    if run_id:
        with _lock:
            run = _runs.get(run_id)
            if run:
                run['finished'] = time.time()
                run['percent'] = 100.0


def stage(step, start_pct=None, end_pct=None):
    """The run moves on to `step`; with percents, the step fills the bar from start_pct to end_pct as it ticks."""
    run_id = current()
    if not run_id:
        return
    with _lock:
        run = _runs.get(run_id)
        if not run:
            return
        run['step'], run['done'], run['total'] = step, 0, 0
        if start_pct is not None:
            run['from'] = max(run['percent'], float(start_pct))  # the bar never goes back
            run['to'] = max(run['from'], float(end_pct if end_pct is not None else start_pct))
            run['percent'] = run['from']


def tick(done, total, step=None):
    """done of total of the current step finished."""
    run_id = current()
    if not run_id:
        return
    with _lock:
        run = _runs.get(run_id)
        if not run:
            return
        if step:
            run['step'] = step
        run['done'], run['total'] = done, total
        if total:
            share = min(max(done / total, 0.0), 1.0)
            run['percent'] = max(run['percent'], run['from'] + (run['to'] - run['from']) * share)


def get(run_id):
    """What the overlay shows, or None for a run this app doesn't know (not started yet, or long gone)."""
    with _lock:
        run = _runs.get(run_id)
        if not run:
            return None
        return {'step': run['step'], 'done': run['done'], 'total': run['total'], 'percent': round(run['percent'], 1),
                'counted': run['to'] > 0, 'finished': run['finished'] is not None,
                'seconds': round((run['finished'] or time.time()) - run['started'], 1), 'lines': list(run['lines'])}


def lines_now():
    """Every line the current request's run has written so far ([] outside one): for the page that says it stopped."""
    run_id = current()
    with _lock:
        run = _runs.get(run_id) if run_id else None
        return list(run['lines']) if run else []


class Handler(logging.Handler):
    """Copies every log line a request writes into its run (the fine detail and any traceback too), for the
    overlay and for the page that shows a stopped run."""

    def __init__(self):
        super().__init__(logging.DEBUG)

    def emit(self, record):
        run_id = current()
        if not run_id:
            return
        try:
            msg = record.getMessage()
        except Exception:
            return
        when = time.localtime(record.created)
        line = {'time': time.strftime('%H:%M:%S', when), 'clock': time.strftime('%I:%M:%S %p', when).lstrip('0'),
                'level': record.levelname,
                'depth': (len(msg) - len(msg.lstrip(' '))) // 2, 'text': msg.strip(), 'more': []}
        if record.exc_info:  # where it stopped: the Logs page reads the same trace (log_view.stopped_at)
            line['more'] = ''.join(traceback.format_exception(*record.exc_info)).rstrip().splitlines()
        with _lock:
            run = _runs.get(run_id)
            if run:
                run['lines'].append(line)
