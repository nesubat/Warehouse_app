"""What a project was made with: each tool that made something in a project folder says so in its
'project_record.json', with when and the files it made. The dashboard reads it for the card's colours and buttons
(what the project can do next), instead of guessing from file names; a project from before the record is read from
its files as before (dashboard_features in app.py).

  {"version": 1,
   "features": {"packing_sheets": {"first": "2026-10-17T15:03:00", "last": "...", "runs": 1, "files": [...]}, ...},
   "history": [{"feature": "labels", "at": "...", "files": [...], "details": {"layout": "horizontal"}}, ...]}

Features: packing_sheets, sub_group, shuffler, labels (Generate Labels, vertical or horizontal), courier_import,
stitch, courier_csv (the other courier CSV, made later from the dashboard).
"""
import json
import os
import threading
from datetime import datetime

RECORD_FILE = 'project_record.json'
FEATURES = ('packing_sheets', 'sub_group', 'shuffler', 'labels', 'courier_import', 'stitch', 'courier_csv')
_lock = threading.Lock()


def read(folder):
    """The project's record, or None for a project made before records were kept (or a damaged one)."""
    try:
        with open(os.path.join(folder, RECORD_FILE), encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data.get('features'), dict) else None
    except (OSError, ValueError, AttributeError):
        return None


def record(folder, feature, files=(), seed=None, **details):
    """Notes that `feature` made `files` in this project now. seed(): the features a project from before records shows
    in its files, called only when this is its first record, so they aren't lost. Never stops the tool calling it."""
    if feature not in FEATURES:
        raise ValueError(f"Unknown feature {feature!r}")
    try:
        with _lock:
            data = read(folder)
            if data is None:
                data = {'version': 1, 'features': {}, 'history': []}
                for f in sorted(seed() if seed else ()):  # what it was made with before records were kept
                    data['features'][f] = {'first': None, 'runs': 0, 'files': [], 'from_files': True}
            now = datetime.now().isoformat(timespec='seconds')
            files = [os.path.basename(f) for f in files if f]
            entry = data['features'].setdefault(feature, {'first': now, 'runs': 0, 'files': []})
            entry['last'] = now
            entry['runs'] = entry.get('runs', 0) + 1
            entry['files'] = list(dict.fromkeys(entry.get('files', []) + files))
            data.setdefault('history', []).append({'feature': feature, 'at': now, 'files': files,
                                                    **({'details': details} if details else {})})
            path = os.path.join(folder, RECORD_FILE)
            with open(path + '.tmp', 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            os.replace(path + '.tmp', path)
    except OSError:
        pass


def set_job(folder, job):
    """The project's job number ('J476523'), when it can't be read from its files (a Label Shuffler project made
    from a name): the dashboard shows it (app.py _project_job)."""
    try:
        with _lock:
            data = read(folder) or {'version': 1, 'features': {}, 'history': []}
            data['job'] = job
            path = os.path.join(folder, RECORD_FILE)
            with open(path + '.tmp', 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            os.replace(path + '.tmp', path)
    except OSError:
        pass


def job(folder):
    """The job number the record keeps, or ''."""
    data = read(folder)
    return str((data or {}).get('job') or '')


def features(folder):
    """The features the record names, or None when there's no record."""
    data = read(folder)
    return set(data['features']) if data else None
