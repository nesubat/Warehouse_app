"""The shared address book: this computer's address book kept in step with one in Google Sheets, through the Apps
Script web app in cloud/apps_script/Code.gs (set up as in docs/CENTRAL_ADDRESS_BOOK.md).

The app always works on its own book (address_book.py); it never waits on the internet to make labels.
  push  entries and learned spellings changed here (flagged by the book's triggers: 'dirty'), within seconds
  pull  what changed in the shared book since the last pull: on start, every 5 minutes, and before a preview if the
        last pull is over a minute old (waiting at most a few seconds)

When this computer and the shared book both changed an entry, the same rule decides on both sides: a verified entry
(confirmed by the courier portal) beats one that isn't; otherwise the latest change wins. The loser is logged in the
shared book's log. Deleting is local: an entry deleted here stays deleted here (tombstones) unless it's changed again
in the shared book after the delete, and it's never deleted from the shared book or other computers.

Settings: WAREHOUSE_SYNC_URL and WAREHOUSE_SYNC_KEY in .env (env_file.py); the name of the person using this computer
in data/sync_settings.json, typed in on the Address Book page: every change in the shared log carries it. Nothing is
pushed until a name is given.
"""
import json
import os
import socket
import sqlite3
import threading
import time
import urllib.error
import urllib.request
from contextlib import contextmanager

from engine.core import app_log

log = app_log.get('sync')

FIELDS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country', 'authority_to_leave')
KEYS = ('address_key', 'receiver_key', 'contact_key')
PUSH_BATCH = 500
PULL_PAGE = 1000
CHECK_EVERY = 15  # seconds between looks for changes to push
PULL_EVERY = 300  # seconds between pulls
TIMEOUT = 45
MAX_NAME = 60


class SyncError(Exception):
    """The shared address book couldn't be reached or refused a request (message says what to do)."""


def server_wins(incoming, local):
    """The rule both sides follow (Code.gs incomingWins): verified beats not verified; otherwise the latest change."""
    iv, lv = bool(incoming.get('verified_at')), bool(local['verified_at'])
    if iv != lv:
        return iv
    return float(incoming.get('changed_at') or 0) >= float(local['changed_at'] or 0)


class CloudSync:
    def __init__(self, book, url, key, settings_path, pc=None, post=None):
        """book: the AddressBook, or a function returning it (app.py swaps it in tests). post(payload) -> reply dict
        replaces the web request (tests)."""
        self._book = book if callable(book) else (lambda: book)
        self.url, self.key = (url or '').strip(), (key or '').strip()
        self.settings_path = settings_path
        self.pc = (pc or socket.gethostname() or 'this computer')[:MAX_NAME]
        self._post = post or self._http
        self._lock = threading.Lock()  # one sync at a time
        self._thread = None
        self._stop = threading.Event()
        self.last_error = ''
        self.last_error_at = None
        self.counts = {'pushed': 0, 'pulled': 0, 'kept': 0}

    # ---------------------------------------------------------------- settings

    @property
    def configured(self):
        return bool(self.url and self.key)

    def user_name(self):
        try:
            with open(self.settings_path, encoding='utf-8') as f:
                return str(json.load(f).get('user_name') or '').strip()[:MAX_NAME]
        except (OSError, ValueError):
            return ''

    def set_user_name(self, name):
        name = " ".join(str(name or '').split())[:MAX_NAME]
        if not name:
            raise ValueError("Type your name.")
        os.makedirs(os.path.dirname(self.settings_path) or '.', exist_ok=True)
        tmp = self.settings_path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'user_name': name}, f)
        os.replace(tmp, self.settings_path)
        log.info("Shared address book: this computer's user is %s", name)
        return name

    # ---------------------------------------------------------------- state kept in the book

    def _state(self, db, key, default=None):
        r = db.execute("SELECT value FROM sync_state WHERE key = ?", (key,)).fetchone()
        return r[0] if r else default

    def _set_state(self, db, key, value):
        db.execute("INSERT INTO sync_state (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                   (key, str(value)))

    def status(self):
        """What the Address Book page shows: set up or not, the name, what's waiting, the last push / pull, the last
        problem."""
        db = self._book()._connect()
        waiting = db.execute("SELECT COUNT(*) FROM addresses WHERE dirty = 1").fetchone()[0]
        spellings = db.execute("SELECT COUNT(*) FROM aliases WHERE dirty = 1").fetchone()[0]
        num = lambda k: float(self._state(db, k) or 0) or None  # noqa: E731
        return {'configured': self.configured, 'user_name': self.user_name(), 'pc': self.pc, 'waiting': waiting,
                'waiting_spellings': spellings, 'last_push': num('last_push'), 'last_pull': num('last_pull'),
                'rev': int(float(self._state(db, 'rev') or 0)), 'error': self.last_error, 'error_at': self.last_error_at,
                **self.counts}

    @contextmanager
    def _quiet(self, db):
        """Writes from the shared book: the triggers don't take them for changes made here."""
        db.execute("BEGIN IMMEDIATE")
        db.execute("INSERT INTO sync_quiet (on_) VALUES (1)")
        try:
            yield
            db.execute("DELETE FROM sync_quiet")
            db.commit()
        except BaseException:
            db.rollback()
            raise

    # ---------------------------------------------------------------- the web request

    def _http(self, payload):
        body = json.dumps({**payload, 'key': self.key}).encode('utf-8')
        req = urllib.request.Request(self.url, data=body, headers={'Content-Type': 'application/json'}, method='POST')
        try:
            # Apps Script answers a POST with a redirect to the reply, which urllib follows as a GET (as it should)
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                text = r.read().decode('utf-8')
        except urllib.error.HTTPError as e:
            raise SyncError(f"The shared address book answered {e.code}. Check WAREHOUSE_SYNC_URL in .env.") from e
        except (urllib.error.URLError, OSError) as e:
            raise SyncError(f"Can't reach the shared address book ({getattr(e, 'reason', e)}). Working on this "
                            f"computer's copy; it syncs when the connection is back.") from e
        try:
            return json.loads(text)
        except ValueError as e:
            raise SyncError("The shared address book's reply couldn't be read: check WAREHOUSE_SYNC_URL in .env is the "
                            "web app's address (ending in /exec).") from e

    def _call(self, payload):
        reply = self._post(payload)
        if not isinstance(reply, dict):
            raise SyncError("The shared address book's reply couldn't be read.")
        if reply.get('error'):
            raise SyncError(reply['error'])
        return reply

    # ---------------------------------------------------------------- push

    def push(self):
        """Sends what changed here, in batches, then the learned spellings. Returns how many entries went up."""
        name = self.user_name()
        if not name:
            raise SyncError("Type your name on the Address Book page: changes are shared with it.")
        db = self._book()._connect()
        sent = 0
        for _ in range(200):  # at most 100,000 entries in one go; the rest next time
            rows = db.execute(f"SELECT id, gid, {', '.join(FIELDS + KEYS)}, verified_at, changed_at FROM addresses "
                              f"WHERE dirty = 1 ORDER BY changed_at LIMIT ?", (PUSH_BATCH,)).fetchall()
            if not rows:
                break
            payload = [{**{k: r[k] for k in FIELDS + KEYS}, 'gid': r['gid'], 'verified_at': r['verified_at'],
                        'changed_at': r['changed_at'], 'authority_to_leave': 1 if r['authority_to_leave'] else 0}
                       for r in rows]
            reply = self._call({'action': 'push', 'user': name, 'pc': self.pc, 'rows': payload, 'spellings': []})
            pushed = {r['gid']: r for r in rows}
            with self._quiet(db):
                for res in reply.get('results') or []:
                    mine = pushed.get(res.get('gid'))
                    if not mine:
                        continue
                    gid = res.get('canonical') or mine['gid']
                    if gid != mine['gid']:
                        # The shared book had this entry already (saved on another computer first): its id from now on
                        try:
                            db.execute("UPDATE addresses SET gid = ? WHERE id = ?", (gid, mine['id']))
                        except sqlite3.IntegrityError:
                            log.warning("Shared address book: %s is here twice; kept as it is", gid)
                            continue
                        db.execute("DELETE FROM tombstones WHERE gid = ?", (gid,))
                    if res.get('status') == 'kept' and res.get('row'):
                        # The shared copy won (verified, or newer): it replaces this one, unless edited again since
                        self.counts['kept'] += 1
                        self._apply_row(db, res['row'], only_if_changed_at=mine['changed_at'], row_id=mine['id'])
                    db.execute("UPDATE addresses SET dirty = 0 WHERE id = ? AND changed_at IS ?", (mine['id'], mine['changed_at']))
            sent += len(rows)
            self.counts['pushed'] += len(rows)
            left = db.execute(f"SELECT COUNT(*) FROM addresses WHERE dirty = 1 AND id IN ({','.join('?' * len(rows))})",
                              [r['id'] for r in rows]).fetchone()[0]
            if len(rows) < PUSH_BATCH or left == len(rows):  # all sent, or none taken (tried again next time)
                break
        self._push_spellings(db, name)
        with db:
            self._set_state(db, 'last_push', time.time())
        if sent:
            log.info("Shared address book: %d entr%s pushed", sent, 'y' if sent == 1 else 'ies')
        return sent

    def _push_spellings(self, db, name):
        for _ in range(100):
            rows = db.execute("SELECT s.source_key, s.address_id, s.changed_at, a.gid FROM aliases s JOIN addresses a "
                              "ON a.id = s.address_id WHERE s.dirty = 1 AND a.dirty = 0 LIMIT ?", (PUSH_BATCH * 2,)).fetchall()
            if not rows:
                return
            reply = self._call({'action': 'push', 'user': name, 'pc': self.pc, 'rows': [],
                                'spellings': [{'source_key': r['source_key'], 'gid': r['gid'], 'changed_at': r['changed_at']}
                                              for r in rows]})
            results = reply.get('spellings') or {}
            done = 0
            with self._quiet(db):
                for r in rows:
                    res = results.get(r['source_key']) or {}
                    if res.get('status') == 'kept' and res.get('gid'):
                        # Another computer learned this spelling for another entry later: theirs
                        target = db.execute("SELECT id FROM addresses WHERE gid = ?", (res['gid'],)).fetchone()
                        if target:
                            db.execute("UPDATE aliases SET address_id = ? WHERE source_key = ?", (target[0], r['source_key']))
                    if res.get('status') in ('saved', 'kept'):
                        db.execute("UPDATE aliases SET dirty = 0 WHERE source_key = ? AND changed_at IS ?",
                                   (r['source_key'], r['changed_at']))
                        done += 1
            if done < len(rows) or len(rows) < PUSH_BATCH * 2:
                return  # some the shared book can't place yet: next time

    # ---------------------------------------------------------------- pull

    def pull(self):
        """Fetches what changed in the shared book since the last pull. Returns how many entries came in."""
        db = self._book()._connect()
        got = 0
        for _ in range(500):
            since = int(float(self._state(db, 'rev') or 0))
            reply = self._call({'action': 'pull', 'since': since, 'limit': PULL_PAGE})
            with self._quiet(db):
                for row in reply.get('rows') or []:
                    got += self._apply_row(db, row)
                for s in reply.get('spellings') or []:
                    self._apply_spelling(db, s)
                self._set_state(db, 'rev', int(reply.get('rev') or since))
            if not reply.get('more'):
                break
        with db:
            self._set_state(db, 'last_pull', time.time())
        self.counts['pulled'] += got
        if got:
            log.info("Shared address book: %d entr%s pulled", got, 'y' if got == 1 else 'ies')
        return got

    def _apply_row(self, db, row, only_if_changed_at=None, row_id=None):
        """Puts a shared entry into this book (inside _quiet). Returns 1 when it changed something here."""
        gid = str(row.get('gid') or '')
        if not gid:
            return 0
        changed_at = float(row.get('changed_at') or 0)
        values = {k: ('' if row.get(k) is None else str(row.get(k))) for k in FIELDS + KEYS}
        values['authority_to_leave'] = 1 if row.get('authority_to_leave') else 0
        verified = row.get('verified_at') or None
        if row_id is None:
            # Deleted here: stays deleted, unless changed in the shared book after the delete
            dead = db.execute("SELECT deleted_at FROM tombstones WHERE gid = ?", (gid,)).fetchone()
            if dead and changed_at <= dead[0]:
                return 0
            if dead:
                db.execute("DELETE FROM tombstones WHERE gid = ?", (gid,))
        local = (db.execute("SELECT * FROM addresses WHERE id = ?", (row_id,)).fetchone() if row_id is not None else
                 db.execute("SELECT * FROM addresses WHERE gid = ?", (gid,)).fetchone())
        if local is None:
            # Not here by its id: the same receiver at the same address saved here on its own (before syncing, or offline)
            local = db.execute("SELECT * FROM addresses WHERE address_key = ? AND receiver_key = ? AND contact_key = ?",
                               tuple(values[k] for k in KEYS)).fetchone()
            if local is not None:
                db.execute("UPDATE addresses SET gid = ? WHERE id = ?", (gid, local['id']))
        if local is not None:
            if only_if_changed_at is not None and local['changed_at'] != only_if_changed_at:
                return 0  # edited here again while the push was on its way: that edit goes up next
            if local['dirty'] and only_if_changed_at is None and not server_wins(row, local):
                return 0  # changed here, and this copy wins: it's pushed next
            same = all(str(local[k] if local[k] is not None else '') == str(values[k]) for k in FIELDS + KEYS) \
                and bool(local['verified_at']) == bool(verified)
            try:
                db.execute(f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS + KEYS)}, verified_at = ?, "
                           f"changed_at = ?, dirty = 0, updated_at = ? WHERE id = ?",
                           [values[k] for k in FIELDS + KEYS] + [verified, changed_at, time.time(), local['id']])
            except sqlite3.IntegrityError:
                log.warning("Shared address book: %s %s clashes with another entry here; kept this computer's",
                            values['receiver'], gid)
                return 0
            return 0 if same else 1
        now = time.time()
        try:
            db.execute(f"INSERT INTO addresses ({', '.join(FIELDS + KEYS)}, verified_at, gid, changed_at, dirty, use_count, "
                       f"created_at, updated_at) VALUES ({', '.join('?' * (len(FIELDS) + len(KEYS) + 7))})",
                       [values[k] for k in FIELDS + KEYS] + [verified, gid, changed_at, 0, 0, now, now])
        except sqlite3.IntegrityError:
            log.warning("Shared address book: %s %s couldn't be added here", values['receiver'], gid)
            return 0
        return 1

    def _apply_spelling(self, db, s):
        target = db.execute("SELECT id FROM addresses WHERE gid = ?", (str(s.get('gid') or ''),)).fetchone()
        key = str(s.get('source_key') or '')
        if not target or not key:
            return
        mine = db.execute("SELECT dirty, changed_at FROM aliases WHERE source_key = ?", (key,)).fetchone()
        if mine and mine['dirty'] and float(mine['changed_at'] or 0) > float(s.get('changed_at') or 0):
            return  # learned here later: pushed next
        db.execute("INSERT INTO aliases (source_key, address_id, dirty, changed_at) VALUES (?, ?, 0, ?) "
                   "ON CONFLICT(source_key) DO UPDATE SET address_id = excluded.address_id, dirty = 0, "
                   "changed_at = excluded.changed_at", (key, target[0], float(s.get('changed_at') or 0)))

    # ---------------------------------------------------------------- when

    def sync_now(self):
        """Push, then pull. Returns {'pushed', 'pulled'}; raises SyncError when it couldn't."""
        if not self.configured:
            raise SyncError("The shared address book isn't set up on this computer: WAREHOUSE_SYNC_URL and "
                            "WAREHOUSE_SYNC_KEY in .env.")
        with self._lock:
            try:
                pushed = self.push() if self.user_name() else 0
                pulled = self.pull()
                self.last_error, self.last_error_at = '', None
                if not self.user_name():
                    self.last_error = "Type your name below to share this computer's changes."
                return {'pushed': pushed, 'pulled': pulled}
            except SyncError as e:
                self._failed(str(e))
                raise
            except Exception as e:
                log.exception("Shared address book sync failed")
                self._failed(f"Sync failed: {e}")
                raise SyncError(str(e)) from e

    def _failed(self, message):
        if message != self.last_error:  # said once in the log, not every 15 seconds
            log.warning("Shared address book: %s", message)
        self.last_error, self.last_error_at = message, time.time()

    def pull_if_stale(self, max_age=60, wait=3.0):
        """Before a preview: a pull if the last one is over max_age seconds old, waiting at most `wait` seconds (the
        preview goes on with this computer's copy if the shared book is slow)."""
        if not self.configured:
            return
        try:
            last = float(self._state(self._book()._connect(), 'last_pull') or 0)
        except sqlite3.Error:
            return
        if time.time() - last < max_age:
            return
        t = threading.Thread(target=self._quietly, args=(self.pull_locked,), daemon=True)
        t.start()
        t.join(wait)

    def pull_locked(self):
        with self._lock:
            return self.pull()

    def _quietly(self, fn):
        try:
            fn()
        except SyncError as e:
            self._failed(str(e))
        except Exception as e:
            log.exception("Shared address book: %s", e)
            self._failed(str(e))

    def start(self):
        """The background sync: pushes within CHECK_EVERY seconds of a change, pulls every PULL_EVERY seconds. Waits
        longer after a failure (up to 5 minutes), so an offline computer doesn't keep knocking."""
        if not self.configured or self._thread:
            return
        log.info("Shared address book: on (%s, as %s)", self.pc, self.user_name() or 'no name yet')

        def loop():
            wait, last_pull = 0, 0
            while not self._stop.wait(wait):
                try:
                    db = self._book()._connect()
                    due_pull = time.time() - last_pull >= PULL_EVERY
                    waiting = db.execute("SELECT 1 FROM addresses WHERE dirty = 1 UNION ALL "
                                         "SELECT 1 FROM aliases WHERE dirty = 1 LIMIT 1").fetchone()
                    if due_pull or (waiting and self.user_name()):
                        self.sync_now()
                        last_pull = time.time()
                    wait = CHECK_EVERY
                except SyncError:
                    wait = min(max(wait * 2, 30), PULL_EVERY)
                except Exception:
                    log.exception("Shared address book: background sync")
                    wait = PULL_EVERY
        self._thread = threading.Thread(target=loop, name='shared-address-book', daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
