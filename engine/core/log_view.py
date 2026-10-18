"""The Logs page: the app's own log (logs/warehouse.log and its rotated copies, written by app_log.py) read
back as activities people can follow, instead of raw lines.

A line in the file:   2026-10-18 16:23:44 INFO    wh.packing MainThread | Generate done in 1.46s: Spring rollout_...
                      (the thread can have spaces, "Thread-11 (work)"; the message can have " | " in it; a traceback
                      runs on over the next lines without a date)

Activities: a line that isn't indented starts one, and the indented lines after it in the same thread are its details
("  18 cartons in 8 consignments"), so a Generate with its 40 lines is one entry with a summary. Threads the page
didn't wait for (the address check after Generate) are background work (is_background).

Running now: app_log.step writes "<what> ..." when a step starts and "<what> done in" / "<what> FAILED after" when it
ends, so a start with no end, from the last few minutes, is still running.
"""
import os
import re
import threading
from datetime import datetime, timedelta

from engine.core.app_log import LOG_FILE, LOG_BACKUPS

LINE_RE = re.compile(r'^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d) (DEBUG|INFO|WARNING|ERROR|CRITICAL)\s+(wh(?:\.[\w.]+)?)\s+(.*?) \| (.*)$')
# Which part of the app wrote it, as the nav names it
AREAS = {'packing': 'Generate Labels', 'import': 'Courier Import', 'stitch': 'Stitch Labels', 'book': 'Address Book',
         'specs': 'Packing Specs', 'sheets': 'Packing Sheets', 'subgroup': 'Sub-Group', 'shuffler': 'Label Shuffler',
         'app': 'App', 'wh': 'App'}
# Level -> (rank, word on the page)
LEVELS = {'DEBUG': (0, 'Detail'), 'INFO': (1, 'Info'), 'WARNING': (2, 'Needs a look'), 'ERROR': (3, 'Problem'),
          'CRITICAL': (3, 'Problem')}
STEP_START_RE = re.compile(r'^(.*\S) \.\.\.$')
STEP_END_RE = re.compile(r'^(.*\S) (?:done in|FAILED after) ')
END_RE = re.compile(r'^(\S+)\b.*? (?:done in|ready in|FAILED after|stopped\b)')  # group 1: the run's first word
RUNNING_FOR = timedelta(minutes=15)  # a step started longer ago with no end was cut short (the app was closed)

_cache = {}  # path -> ((mtime, size), entries): the page polls, and an unchanged file isn't read again
_lock = threading.Lock()


def log_files(log_dir):
    """The log files, oldest first (warehouse.log.3 … .1, then warehouse.log)."""
    names = [f"{LOG_FILE}.{n}" for n in range(LOG_BACKUPS, 0, -1)] + [LOG_FILE]
    return [p for p in (os.path.join(log_dir, n) for n in names) if os.path.isfile(p)]


def version(log_dir):
    """Changes whenever anything is written: lets the page ask 'anything new?' cheaply."""
    path = os.path.join(log_dir, LOG_FILE)
    try:
        st = os.stat(path)
        return f"{int(st.st_mtime * 1000)}-{st.st_size}"
    except OSError:
        return ''


def parse_lines(lines):
    """[{time, level, area, thread, depth, text, more}] in file order; lines without a date belong to the one before."""
    out = []
    for line in lines:
        line = line.rstrip('\r\n')
        m = LINE_RE.match(line)
        if m:
            when, level, name, thread, msg = m.groups()
            area = name.split('.', 1)[1].split('.')[0] if '.' in name else 'wh'
            depth = (len(msg) - len(msg.lstrip(' '))) // 2
            out.append({'time': when, 'level': level, 'area': area, 'thread': thread.strip(), 'depth': depth,
                        'text': msg.strip(), 'more': []})
        elif out and line.strip():
            out[-1]['more'].append(line)  # a traceback, or a message that ran over lines
    return out


def _read(path):
    try:
        st = os.stat(path)
    except OSError:
        return []
    key = (st.st_mtime, st.st_size)
    with _lock:
        hit = _cache.get(path)
        if hit and hit[0] == key:
            return hit[1]
    with open(path, encoding='utf-8', errors='replace') as f:
        entries = parse_lines(f)
    with _lock:
        _cache[path] = (key, entries)
    return entries


def read_entries(log_dir):
    """Every line kept, oldest first."""
    out = []
    for path in log_files(log_dir):
        out.extend(_read(path))
    return out


def _grouped(log_dir):
    """(entries, activities), worked out again only when a log file changed: the page asks every few seconds."""
    files = log_files(log_dir)
    key = tuple((p, os.path.getmtime(p), os.path.getsize(p)) for p in files)
    with _lock:
        hit = _cache.get(('grouped', log_dir))
        if hit and hit[0] == key:
            return hit[1]
    entries = read_entries(log_dir)
    result = (entries, activities(entries))
    with _lock:
        _cache[('grouped', log_dir)] = (key, result)
    return result


def is_background(thread):
    """Work the page didn't wait for (the address check after Generate, 'Thread-14 (work)'). A page's own request runs
    in 'MainThread' or, in the running app, a server thread ('Thread-19 (process_request_thread)')."""
    return thread != 'MainThread' and 'process_request_thread' not in thread


def _first_word(text):
    return text.split(' ', 1)[0].rstrip(':')


def _spans(entries):
    """{index of a run's first line: index of its last}: "Generate <file>: …" … "Generate done in 1.46s: <project>"
    (or "Preview …" … "Preview ready in", "… FAILED after", "Generate stopped: …"), same part of the app, same
    thread. Everything that thread logs in between is that run's, the address book lines it writes too."""
    spans, last_start = {}, {}
    for i, e in enumerate(entries):
        if e['depth']:
            continue
        m = END_RE.match(e['text'])
        if m:
            start = last_start.pop((e['thread'], e['area'], m.group(1).rstrip(':')), None)
            if start is not None:
                spans[start] = i
        else:
            last_start[(e['thread'], e['area'], _first_word(e['text']))] = i
    return spans


def activities(entries):
    """Lines grouped into activities, oldest first: [{time, end, area, thread, background, title, result, level,
    lines}]. A run (see _spans) is one; any other line that isn't indented starts one, with the indented lines after
    it in the same thread as its details."""
    spans = _spans(entries)
    out, last_by_thread, open_run = [], {}, {}  # open_run: thread -> (activity, index of its last line)
    for i, e in enumerate(entries):
        thread = e['thread']
        run = open_run.get(thread)
        if run and i <= run[1]:
            act = run[0]
            act['lines'].append(e)
            if i == run[1]:
                act['result'] = e['text']
                del open_run[thread]
        elif e['depth'] and thread in last_by_thread and last_by_thread[thread]['area'] == e['area']:
            act = last_by_thread[thread]  # its details (another part of the app's line is its own activity)
            act['lines'].append(e)
        else:
            act = {'time': e['time'], 'end': e['time'], 'area': e['area'], 'thread': thread, 'at': i,
                   'background': is_background(thread), 'title': e['text'], 'title_more': e['more'], 'result': '',
                   'level': e['level'], 'lines': []}
            out.append(act)
            last_by_thread[thread] = act
            if i in spans:
                open_run[thread] = (act, spans[i])
            continue
        act['end'] = e['time']
        if LEVELS[e['level']][0] > LEVELS[act['level']][0]:
            act['level'] = e['level']
    return out


def running(entries, now=None):
    """Steps started (app_log.step) with no end yet, newest first: [{'what', 'area', 'since', 'background'}]."""
    now = now or datetime.now()
    cutoff = (now - RUNNING_FOR).strftime('%Y-%m-%d %H:%M:%S')
    first = len(entries)
    while first > 0 and entries[first - 1]['time'] >= cutoff:  # only the last few minutes can still be running
        first -= 1
    started = {}
    for e in entries[first:]:
        key = (e['area'], e['thread'])
        m = STEP_START_RE.match(e['text'])
        if m and e['level'] == 'DEBUG':
            started[key + (m.group(1),)] = e
            continue
        m = STEP_END_RE.match(e['text'])
        if m:
            started.pop(key + (m.group(1),), None)
    out = []
    for (area, thread, what), e in started.items():
        since = datetime.strptime(e['time'], '%Y-%m-%d %H:%M:%S')
        if now - since <= RUNNING_FOR:
            out.append({'what': what, 'area': area, 'area_label': AREAS.get(area, area.title()), 'since': e['time'],
                        'background': is_background(thread)})
    return sorted(out, key=lambda r: r['since'], reverse=True)


def _public_line(e):
    return {'time': e['time'][11:], 'level': e['level'], 'level_label': LEVELS[e['level']][1], 'depth': e['depth'],
            'text': e['text'], 'more': e['more']}


FRAME_RE = re.compile(r'^\s*File "(.+)", line (\d+), in (.+)$')
LIBRARY_RE = re.compile(r'[\\/](?:site-packages|venv|Lib|lib[\\/]python[\d.]*)[\\/]', re.I)


def _short_path(path):
    """'C:/…/Warehouse_app/engine/courier/courier_export.py' (either slash) -> 'engine/courier/courier_export.py'."""
    parts = re.split(r'[\\/]', path)
    if 'site-packages' in parts:  # a library: its own name ('pymupdf/__init__.py')
        return '/'.join(parts[parts.index('site-packages') + 1:])
    return '/'.join(parts[parts.index('engine'):]) if 'engine' in parts else parts[-1]


def stopped_at(trace):
    """A traceback's lines read for people: {'where': the app's own code it stopped in ('app.py, line 3110, in
    create_horizontal_labels'), 'code': that line, 'inside': the library it was in when it stopped, if not the app's
    code, 'error': 'ValueError: Unknown service code ...'}. None when the lines aren't a traceback."""
    if not trace or not any(t.startswith('Traceback') for t in trace):
        return None
    frames = []
    for n, t in enumerate(trace):
        m = FRAME_RE.match(t)
        if m:
            nxt = trace[n + 1] if n + 1 < len(trace) else ''
            code = nxt.strip() if nxt.startswith((' ', '\t')) and not FRAME_RE.match(nxt) else ''  # its source line
            frames.append((m.group(1), m.group(2), m.group(3), code))
    error = next((t.strip() for t in reversed(trace) if t.strip() and not t.startswith((' ', '\t'))
                  and not t.startswith('Traceback') and not t.startswith('During handling')), '')
    own = [f for f in frames if not LIBRARY_RE.search(f[0])]
    pick = own[-1] if own else (frames[-1] if frames else None)
    out = {'error': error, 'where': '', 'code': '', 'inside': ''}
    if pick:
        out['where'] = f"{_short_path(pick[0])}, line {pick[1]}, in {pick[2]}"
        out['code'] = pick[3]
    if frames and own and frames[-1] is not own[-1]:
        last = frames[-1]
        out['inside'] = f"{_short_path(last[0])}, line {last[1]}, in {last[2]}"
    return out


def _failure(a):
    for trace in [a['title_more']] + [x['more'] for x in a['lines']]:
        found = stopped_at(trace)
        if found:
            return found
    return None


def _before(entries, a, count=6, within=120):
    """What the same thread wrote just before the activity started (the last `count` lines, from the last
    `within` seconds): what the app was doing when this happened."""
    start = datetime.strptime(a['time'], '%Y-%m-%d %H:%M:%S')
    floor = (start - timedelta(seconds=within)).strftime('%Y-%m-%d %H:%M:%S')
    out, i = [], a['at'] - 1
    while i >= 0 and len(out) < count and entries[i]['time'] >= floor:
        if entries[i]['thread'] == a['thread']:
            out.append(entries[i])
        i -= 1
    return [_public_line(x) for x in reversed(out)]


def find_activity(acts, at, text):
    """The index of the activity a failed page points at: the newest one around `at` (its time, or between its first
    and last line, a few seconds either way) that wrote `text`; else the newest one around then. None when none is."""
    try:
        when = datetime.strptime(at, '%Y-%m-%d %H:%M:%S')
    except (TypeError, ValueError):
        return None
    lo, hi = (when - timedelta(seconds=3)).strftime('%Y-%m-%d %H:%M:%S'), (when + timedelta(seconds=3)).strftime('%Y-%m-%d %H:%M:%S')
    near = [n for n in range(len(acts) - 1, -1, -1)
            if acts[n]['time'] <= hi and acts[n]['end'] >= lo]
    text = (text or '').strip().lower()[:60]
    wrote = lambda a: any(text and text in x.lower() for x in [a['title']] + [y['text'] for y in a['lines']])
    return next((n for n in near if wrote(acts[n])), near[0] if near else None)


def view(log_dir, area='', show='all', q='', day='today', limit=200, before=None, now=None, focus_at=None, focus_text=''):
    """What the page shows: the activities matching the filters, newest first (at most `limit`, older than the
    `before` index for Load more), today's counts and what's running.
    show: 'all' | 'issues' (needs a look and problems) | 'problems'. day: 'today' | 'yesterday' | 'week' | 'all'."""
    now = now or datetime.now()
    try:
        entries, acts = _grouped(log_dir)
    except OSError:  # a file rotated away while it was being read: the next ask gets it
        entries, acts = [], []
    focus = find_activity(acts, focus_at, focus_text) if focus_at else None
    if focus is not None:  # a failed page sent here: its run first, every kind and day, nothing filtered out
        area, show, q, day, before = '', 'all', '', 'all', focus + 1
    today = now.strftime('%Y-%m-%d')
    days = {'today': (today, today),
            'yesterday': ((now - timedelta(days=1)).strftime('%Y-%m-%d'),) * 2,
            'week': ((now - timedelta(days=6)).strftime('%Y-%m-%d'), today)}.get(day)
    q = q.strip().lower()
    floor = {'issues': 2, 'problems': 3}.get(show, 0)

    def keep(a):
        if days and not days[0] <= a['time'][:10] <= days[1]:
            return False
        if area and a['area'] != area:
            return False
        if LEVELS[a['level']][0] < floor:
            return False
        if q and q not in a['title'].lower() and not any(q in x['text'].lower() for x in a['lines']):
            return False
        return True

    picked = []
    start = len(acts) - 1 if before is None else min(int(before), len(acts)) - 1
    i = start
    while i >= 0 and len(picked) < limit:
        if keep(acts[i]):
            picked.append(i)
        i -= 1
    todays = [a for a in acts if a['time'][:10] == today]
    return {
        'activities': [{
            'id': n, 'time': a['time'], 'day': a['time'][:10], 'clock': a['time'][11:], 'end': a['end'][11:],
            'area': a['area'], 'area_label': AREAS.get(a['area'], a['area'].title()), 'background': a['background'],
            'title': a['title'], 'title_more': a['title_more'], 'result': a['result'], 'level': a['level'], 'level_label': LEVELS[a['level']][1],
            'details': [_public_line(x) for x in a['lines']],
            'issues': sum(1 for x in a['lines'] if LEVELS[x['level']][0] >= 2),
            'thread': a['thread'], 'failure': _failure(a), 'before': _before(entries, a),
        } for n in picked for a in [acts[n]]],
        'more': i >= 0 and any(keep(acts[j]) for j in range(i, -1, -1)),
        'oldest_shown': picked[-1] if picked else None,
        'today': {'activities': len(todays),
                  'issues': sum(1 for a in todays if LEVELS[a['level']][0] == 2),
                  'problems': sum(1 for a in todays if LEVELS[a['level']][0] >= 3)},
        'running': running(entries, now),
        'last': entries[-1]['time'] if entries else None,
        'kept_from': entries[0]['time'] if entries else None,
        'areas': sorted({a['area'] for a in acts}, key=lambda k: AREAS.get(k, k)),
        'area_labels': {k: AREAS.get(k, k.title()) for k in {a['area'] for a in acts}},
        'version': version(log_dir),
        'focus': focus,
    }
