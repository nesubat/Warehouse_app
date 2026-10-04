"""Local address book: clean delivery addresses, learned from every Generate and editable by the user.

Stored in one SQLite file (data/address_book.db):
  addresses      one row per receiver + address, with how often and how recently it was used
  addresses_fts  full-text index over the address fields (kept in sync by triggers) for instant search
  aliases        raw address spellings seen in Excel -> the address they were saved as, so the next file
                 with the same messy spelling gets the clean (possibly hand-corrected) address

SQLite was chosen because it needs no server, ships with Python, keeps memory low (only queried rows are
loaded) and stays fast at hundreds of thousands of rows. Everything goes through the AddressBook class, so
moving to a cloud database later means replacing this one file.
"""
import os
import sqlite3
import threading
import time

from packing_label_generator import _address_key, _norm

FIELDS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country', 'authority_to_leave')
TEXT_FIELDS = FIELDS[:-1]
MAX_LENGTH = 200

SCHEMA = """
CREATE TABLE IF NOT EXISTS addresses (
    id                 INTEGER PRIMARY KEY,
    receiver           TEXT NOT NULL DEFAULT '',
    contact            TEXT NOT NULL DEFAULT '',
    line1              TEXT NOT NULL DEFAULT '',
    line2              TEXT NOT NULL DEFAULT '',
    suburb             TEXT NOT NULL DEFAULT '',
    state              TEXT NOT NULL DEFAULT '',
    postcode           TEXT NOT NULL DEFAULT '',
    country            TEXT NOT NULL DEFAULT 'AU',
    authority_to_leave INTEGER NOT NULL DEFAULT 0,
    address_key        TEXT NOT NULL,
    receiver_key       TEXT NOT NULL,
    use_count          INTEGER NOT NULL DEFAULT 0,
    created_at         REAL NOT NULL,
    updated_at         REAL NOT NULL,
    last_used_at       REAL,
    UNIQUE (address_key, receiver_key)
);
CREATE INDEX IF NOT EXISTS addresses_recent ON addresses (last_used_at DESC, updated_at DESC);

CREATE VIRTUAL TABLE IF NOT EXISTS addresses_fts USING fts5(
    receiver, contact, line1, line2, suburb, state, postcode,
    content='addresses', content_rowid='id',
    tokenize="unicode61 remove_diacritics 2", prefix='2 3'
);
CREATE TRIGGER IF NOT EXISTS addresses_ai AFTER INSERT ON addresses BEGIN
    INSERT INTO addresses_fts(rowid, receiver, contact, line1, line2, suburb, state, postcode)
    VALUES (new.id, new.receiver, new.contact, new.line1, new.line2, new.suburb, new.state, new.postcode);
END;
CREATE TRIGGER IF NOT EXISTS addresses_ad AFTER DELETE ON addresses BEGIN
    INSERT INTO addresses_fts(addresses_fts, rowid, receiver, contact, line1, line2, suburb, state, postcode)
    VALUES ('delete', old.id, old.receiver, old.contact, old.line1, old.line2, old.suburb, old.state, old.postcode);
END;
CREATE TRIGGER IF NOT EXISTS addresses_au AFTER UPDATE ON addresses BEGIN
    INSERT INTO addresses_fts(addresses_fts, rowid, receiver, contact, line1, line2, suburb, state, postcode)
    VALUES ('delete', old.id, old.receiver, old.contact, old.line1, old.line2, old.suburb, old.state, old.postcode);
    INSERT INTO addresses_fts(rowid, receiver, contact, line1, line2, suburb, state, postcode)
    VALUES (new.id, new.receiver, new.contact, new.line1, new.line2, new.suburb, new.state, new.postcode);
END;

CREATE TABLE IF NOT EXISTS aliases (
    source_key TEXT PRIMARY KEY,
    address_id INTEGER NOT NULL REFERENCES addresses(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS aliases_address ON aliases (address_id);
"""


class AddressBookError(ValueError):
    """A change the user asked for can't be made (missing fields, duplicate address...)."""


def one_line(a):
    return ", ".join(x for x in (a.get('line1'), a.get('line2'), a.get('suburb'), a.get('state'), a.get('postcode')) if x)


def _clean(fields):
    out = {}
    for k in TEXT_FIELDS:
        v = " ".join(str(fields.get(k) or '').split())[:MAX_LENGTH]
        out[k] = v.upper() if k in ('state', 'country') else v
    out['country'] = out['country'] or 'AU'
    out['authority_to_leave'] = 1 if fields.get('authority_to_leave') in (True, 1, '1', 'true', 'on', 'Y', 'y') else 0
    return out


def _keys(a):
    return _address_key(one_line(a)), _norm(a['receiver'])


def _match_query(text):
    """Search-as-you-type: finished words must match whole words, the last (still being typed) word
    matches as a prefix. '49 church st' -> '"49" AND "church" AND "st"*', so 49 doesn't also find 499."""
    words = [w for w in _norm(text).split() if w][:8]
    if not words:
        return ''
    finished = text.rstrip() != text  # a trailing space means the last word is complete too
    parts = [f'"{w}"' for w in words[:-1]] + [f'"{words[-1]}"' if finished else f'"{words[-1]}"*']
    return " AND ".join(parts)


# Column weights for ranking: receiver and street lines count most, then suburb/postcode.
RANK = "bm25(addresses_fts, 4.0, 1.0, 3.0, 1.0, 2.0, 0.5, 2.0)"


class AddressBook:
    def __init__(self, path):
        self.path = path
        self._local = threading.local()
        os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
        with self._connect() as db:
            db.executescript(SCHEMA)

    # One connection per thread: Flask serves requests on several threads, and WAL mode lets
    # searches run while a Generate is writing.
    def _connect(self):
        db = getattr(self._local, 'db', None)
        if db is None:
            db = sqlite3.connect(self.path, timeout=10)
            db.row_factory = sqlite3.Row
            db.execute("PRAGMA journal_mode=WAL")
            db.execute("PRAGMA synchronous=NORMAL")
            db.execute("PRAGMA foreign_keys=ON")
            self._local.db = db
        return db

    @staticmethod
    def _row(r):
        d = {k: r[k] for k in ('id',) + TEXT_FIELDS + ('use_count', 'last_used_at', 'updated_at')}
        d['authority_to_leave'] = bool(r['authority_to_leave'])
        d['address'] = one_line(d)
        return d

    # ---------- reading ----------

    def search(self, query='', limit=50, offset=0):
        """Best matches first (closeness, then most used). An empty query lists the most recently used.
        Returns (rows, has_more) without counting every match, so typing stays fast at any size."""
        limit, offset = max(1, min(int(limit), 200)), max(0, int(offset))
        db = self._connect()
        match = _match_query(query)
        if match and len(_norm(query)) < 2:
            # One character matches a large share of the book; ranking all of it costs ~90 ms at 250k,
            # so take the first matches straight from the index. The next keystroke gets full ranking.
            sql = ("SELECT a.* FROM addresses_fts JOIN addresses a ON a.id = addresses_fts.rowid "
                   "WHERE addresses_fts MATCH ? LIMIT ? OFFSET ?")
            rows = db.execute(sql, (match, limit + 1, offset)).fetchall()
        elif match:
            sql = ("SELECT a.* FROM addresses_fts JOIN addresses a ON a.id = addresses_fts.rowid "
                   f"WHERE addresses_fts MATCH ? ORDER BY {RANK}, a.use_count DESC LIMIT ? OFFSET ?")
            rows = db.execute(sql, (match, limit + 1, offset)).fetchall()
        else:
            sql = "SELECT * FROM addresses ORDER BY last_used_at DESC, updated_at DESC LIMIT ? OFFSET ?"
            rows = db.execute(sql, (limit + 1, offset)).fetchall()
        return [self._row(r) for r in rows[:limit]], len(rows) > limit

    def get(self, address_id):
        r = self._connect().execute("SELECT * FROM addresses WHERE id = ?", (address_id,)).fetchone()
        return self._row(r) if r else None

    def count(self):
        return self._connect().execute("SELECT COUNT(*) FROM addresses").fetchone()[0]

    def lookup_aliases(self, source_keys):
        """{raw Excel spelling key: saved address} for the keys the book has learned."""
        keys = [k for k in dict.fromkeys(source_keys) if k]
        found = {}
        db = self._connect()
        for i in range(0, len(keys), 500):  # stay under SQLite's parameter limit
            chunk = keys[i:i + 500]
            sql = (f"SELECT al.source_key, a.* FROM aliases al JOIN addresses a ON a.id = al.address_id "
                   f"WHERE al.source_key IN ({','.join('?' * len(chunk))})")
            for r in db.execute(sql, chunk):
                found[r['source_key']] = self._row(r)
        return found

    # ---------- writing ----------

    def create(self, fields):
        a = _clean(fields)
        if not (a['receiver'] and a['line1'] and a['postcode']):
            raise AddressBookError("Receiver, Address Line 1 and Postcode are required.")
        address_key, receiver_key = _keys(a)
        now = time.time()
        db = self._connect()
        try:
            with db:
                cur = db.execute(
                    f"INSERT INTO addresses ({', '.join(FIELDS)}, address_key, receiver_key, created_at, updated_at) "
                    f"VALUES ({', '.join('?' * (len(FIELDS) + 4))})",
                    [a[k] for k in FIELDS] + [address_key, receiver_key, now, now])
        except sqlite3.IntegrityError:
            raise AddressBookError(f"{a['receiver']} at {one_line(a)} is already in the address book.")
        return self.get(cur.lastrowid)

    def update(self, address_id, fields):
        current = self.get(address_id)
        if not current:
            raise AddressBookError("That address no longer exists.")
        a = _clean({**current, **fields})
        if not (a['receiver'] and a['line1'] and a['postcode']):
            raise AddressBookError("Receiver, Address Line 1 and Postcode are required.")
        address_key, receiver_key = _keys(a)
        db = self._connect()
        try:
            with db:
                db.execute(
                    f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS)}, address_key = ?, receiver_key = ?, "
                    f"updated_at = ? WHERE id = ?",
                    [a[k] for k in FIELDS] + [address_key, receiver_key, time.time(), address_id])
        except sqlite3.IntegrityError:
            raise AddressBookError(f"{a['receiver']} at {one_line(a)} is already in the address book.")
        return self.get(address_id)

    def delete(self, address_id):
        db = self._connect()
        with db:
            return db.execute("DELETE FROM addresses WHERE id = ?", (address_id,)).rowcount > 0

    def record_used(self, entries):
        """After a Generate: save or refresh each delivery address and remember the raw spellings that led to it.

        entries: [(destination dict, [raw source keys]), ...]. All in one transaction, so a Generate of
        hundreds of consignments is a single fast write."""
        now = time.time()
        db = self._connect()
        with db:
            for dest, source_keys in entries:
                a = _clean(dest)
                if not (a['receiver'] and a['postcode']):
                    continue
                address_key, receiver_key = _keys(a)
                row = db.execute("SELECT id FROM addresses WHERE address_key = ? AND receiver_key = ?",
                                 (address_key, receiver_key)).fetchone()
                if row:
                    address_id = row['id']
                    # Refresh to the latest spelling/contact/ATL actually sent
                    db.execute(f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS)}, use_count = use_count + 1, "
                               f"last_used_at = ?, updated_at = ? WHERE id = ?",
                               [a[k] for k in FIELDS] + [now, now, address_id])
                else:
                    address_id = db.execute(
                        f"INSERT INTO addresses ({', '.join(FIELDS)}, address_key, receiver_key, use_count, created_at, "
                        f"updated_at, last_used_at) VALUES ({', '.join('?' * (len(FIELDS) + 6))})",
                        [a[k] for k in FIELDS] + [address_key, receiver_key, 1, now, now, now]).lastrowid
                for key in dict.fromkeys(source_keys):
                    if key:
                        db.execute("INSERT INTO aliases (source_key, address_id) VALUES (?, ?) "
                                   "ON CONFLICT(source_key) DO UPDATE SET address_id = excluded.address_id",
                                   (key, address_id))
