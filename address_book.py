"""Local address book: clean delivery addresses, learned from every Generate and editable by the user.

Stored in one SQLite file (data/address_book.db):
  addresses      one row per receiver + Attn + address, with how often and how recently it was used. One address
                 can have several receivers and Attn names; each combination sent in a Generate is kept
  addresses_fts  full-text index over the address fields (kept in sync by triggers) for instant search
  aliases        raw address spellings seen in Excel -> the address they were saved as, so the next file
                 with the same messy spelling gets the clean (possibly hand-corrected) address

SQLite was chosen because it needs no server, ships with Python, keeps memory low (only queried rows are
loaded) and stays fast at hundreds of thousands of rows. Everything goes through the AddressBook class, so
moving to a cloud database later means replacing this one file.
"""
import heapq
import os
import sqlite3
import threading
import time

from packing_label_generator import _address_key, _norm, normalize_postcode

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
    contact_key        TEXT NOT NULL DEFAULT '',
    use_count          INTEGER NOT NULL DEFAULT 0,
    created_at         REAL NOT NULL,
    updated_at         REAL NOT NULL,
    last_used_at       REAL,
    verified_at        REAL,
    UNIQUE (address_key, receiver_key, contact_key)
);
CREATE INDEX IF NOT EXISTS addresses_recent ON addresses (last_used_at DESC, updated_at DESC);
CREATE INDEX IF NOT EXISTS addresses_receiver ON addresses (receiver_key);
CREATE INDEX IF NOT EXISTS addresses_postcode ON addresses (postcode);

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
        out[k] = v.upper() if k in ('suburb', 'state', 'country') else v  # suburbs always in capitals
    out['country'] = out['country'] or 'AU'
    out['postcode'] = normalize_postcode(out['postcode'], out['country'])
    out['authority_to_leave'] = 1 if fields.get('authority_to_leave') in (True, 1, '1', 'true', 'on', 'Y', 'y') else 0
    return out


def _keys(a):
    """(address, receiver, Attn) comparison keys: together they identify an entry."""
    return _address_key(one_line(a)), _norm(a['receiver']), _norm(a['contact'])


def _who(a):
    return a['receiver'] + (f" (Attn {a['contact']})" if a['contact'] else '')


COLUMNS = ('id',) + FIELDS + ('address_key', 'receiver_key', 'contact_key', 'use_count', 'created_at', 'updated_at',
                               'last_used_at', 'verified_at')


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
CROWDED_SHORTLIST = 30  # in a postcode with more entries than near()'s limit, how many to score in full
RANK = "bm25(addresses_fts, 4.0, 1.0, 3.0, 1.0, 2.0, 0.5, 2.0)"


class AddressBook:
    def __init__(self, path):
        self.path = path
        self._local = threading.local()
        os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
        with self._connect() as db:
            db.executescript(SCHEMA)
            # Added after the first release: when an address was confirmed by the courier portal
            if 'verified_at' not in {c[1] for c in db.execute("PRAGMA table_info(addresses)")}:
                db.execute("ALTER TABLE addresses ADD COLUMN verified_at REAL")
        self._add_contact_key()
        with self._connect() as db:
            self._pad_short_postcodes(db)
            # Suburbs are kept in capitals; entries saved before that are upper-cased (matching ignores case,
            # so keys and learned spellings are unchanged)
            db.execute("UPDATE addresses SET suburb = UPPER(suburb) WHERE suburb != UPPER(suburb)")

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

    def _add_contact_key(self):
        """One-off upgrade of books made before the Attn was part of an entry's identity (an address could hold
        only one Attn per receiver). SQLite can't change a UNIQUE rule in place, so the table is rebuilt with the
        same ids (aliases and the search index stay valid). A copy of the file is kept first."""
        db = self._connect()
        if 'contact_key' in {c[1] for c in db.execute("PRAGMA table_info(addresses)")}:
            return
        backup = sqlite3.connect(f"{os.path.splitext(self.path)[0]}.before-attn-upgrade.db")
        db.backup(backup)
        backup.close()
        db.create_function('norm_key', 1, _norm)
        old = [c for c in COLUMNS if c != 'contact_key']
        create = SCHEMA.split(';')[0].replace('CREATE TABLE IF NOT EXISTS addresses', 'CREATE TABLE addresses_new')
        db.execute("PRAGMA foreign_keys=OFF")
        try:
            with db:
                db.execute(create)
                db.execute(f"INSERT OR IGNORE INTO addresses_new ({', '.join(old)}, contact_key) "
                           f"SELECT {', '.join(old)}, norm_key(contact) FROM addresses")
                for trigger in ('addresses_ai', 'addresses_ad', 'addresses_au'):
                    db.execute(f"DROP TRIGGER IF EXISTS {trigger}")
                db.execute("DROP TABLE addresses")
                db.execute("ALTER TABLE addresses_new RENAME TO addresses")
            db.executescript(SCHEMA)  # indexes and triggers for the new table
            with db:
                db.execute("INSERT INTO addresses_fts(addresses_fts) VALUES ('rebuild')")
        finally:
            db.execute("PRAGMA foreign_keys=ON")

    @staticmethod
    def _pad_short_postcodes(db):
        """One-off fix for entries saved before postcodes were padded (803 -> 0803)."""
        rows = db.execute("SELECT * FROM addresses WHERE length(postcode) = 3 AND postcode GLOB '[0-9][0-9][0-9]' "
                          "AND country IN ('AU', 'NZ')").fetchall()
        for r in rows:
            fixed = {k: r[k] for k in TEXT_FIELDS}
            fixed['postcode'] = normalize_postcode(r['postcode'], r['country'])
            try:
                db.execute("UPDATE addresses SET postcode = ?, address_key = ? WHERE id = ?",
                           (fixed['postcode'], _address_key(one_line(fixed)), r['id']))
            except sqlite3.IntegrityError:
                pass  # the padded address is already saved separately; leave this one for the user to tidy

    @staticmethod
    def _row(r):
        d = {k: r[k] for k in ('id',) + TEXT_FIELDS + ('use_count', 'last_used_at', 'updated_at', 'verified_at')}
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
        keys = _keys(a)
        now = time.time()
        db = self._connect()
        try:
            with db:
                cur = db.execute(
                    f"INSERT INTO addresses ({', '.join(FIELDS)}, address_key, receiver_key, contact_key, created_at, updated_at) "
                    f"VALUES ({', '.join('?' * (len(FIELDS) + 5))})",
                    [a[k] for k in FIELDS] + list(keys) + [now, now])
        except sqlite3.IntegrityError:
            raise AddressBookError(f"{_who(a)} at {one_line(a)} is already in the address book.")
        return self.get(cur.lastrowid)

    def update(self, address_id, fields):
        current = self.get(address_id)
        if not current:
            raise AddressBookError("That address no longer exists.")
        a = _clean({**current, **fields})
        if not (a['receiver'] and a['line1'] and a['postcode']):
            raise AddressBookError("Receiver, Address Line 1 and Postcode are required.")
        keys = _keys(a)
        db = self._connect()
        try:
            with db:
                db.execute(
                    f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS)}, address_key = ?, receiver_key = ?, "
                    f"contact_key = ?, updated_at = ? WHERE id = ?",
                    [a[k] for k in FIELDS] + list(keys) + [time.time(), address_id])
        except sqlite3.IntegrityError:
            raise AddressBookError(f"{_who(a)} at {one_line(a)} is already in the address book.")
        return self.get(address_id)

    def find(self, address):
        """The saved entry for this receiver + address (the one with the same Attn if there is one), or None."""
        address_key, receiver_key, contact_key = _keys(_clean(address))
        r = self._connect().execute("SELECT * FROM addresses WHERE address_key = ? AND receiver_key = ? "
                                    "ORDER BY contact_key = ? DESC, use_count DESC LIMIT 1",
                                    (address_key, receiver_key, contact_key)).fetchone()
        return self._row(r) if r else None

    def find_by_receiver(self, receiver):
        """Entries whose receiver name matches (ignoring capitals and punctuation), most used first."""
        rows = self._connect().execute("SELECT * FROM addresses WHERE receiver_key = ? ORDER BY use_count DESC LIMIT 50",
                                       (_norm(receiver),)).fetchall()
        return [self._row(r) for r in rows]

    def same_identity(self, entry):
        """The receiver's other entries with the same Attn (other addresses, or stale copies of this one)."""
        rows = self._connect().execute("SELECT * FROM addresses WHERE receiver_key = ? AND contact_key = ? AND id != ? LIMIT 50",
                                       (_norm(entry.get('receiver')), _norm(entry.get('contact')), entry['id'])).fetchall()
        return [self._row(r) for r in rows]

    def merge_into(self, target_id, other_id):
        """Folds a stale copy into the target entry: its learned spellings and use count move over, then it's
        removed. Only for two entries of the same receiver + Attn. Returns False if there was nothing to merge."""
        db = self._connect()
        with db:
            pair = db.execute("SELECT id, receiver_key, contact_key, use_count, last_used_at FROM addresses WHERE id IN (?, ?)",
                              (target_id, other_id)).fetchall()
            if target_id == other_id or len(pair) != 2:
                return False
            target, other = sorted(pair, key=lambda r: r['id'] != target_id)
            if (target['receiver_key'], target['contact_key']) != (other['receiver_key'], other['contact_key']):
                return False
            db.execute("UPDATE aliases SET address_id = ? WHERE address_id = ?", (target_id, other_id))
            db.execute("UPDATE addresses SET use_count = use_count + ?, "
                       "last_used_at = COALESCE(MAX(last_used_at, ?), last_used_at, ?) WHERE id = ?",
                       (other['use_count'], other['last_used_at'], other['last_used_at'], target_id))
            db.execute("DELETE FROM addresses WHERE id = ?", (other_id,))
        return True

    def mark_verified(self, address_id):
        """The courier portal confirmed this entry exactly as saved."""
        db = self._connect()
        with db:
            return db.execute("UPDATE addresses SET verified_at = ? WHERE id = ?", (time.time(), address_id)).rowcount > 0

    def at_address(self, address):
        """Every entry saved at exactly this address, whatever the receiver (an address can have several
        receivers / Attn names). Uses the (address_key, receiver_key) index."""
        rows = self._connect().execute("SELECT * FROM addresses WHERE address_key = ? ORDER BY use_count DESC LIMIT 20",
                                       (_keys(_clean(address))[0],)).fetchall()
        return [self._row(r) for r in rows]

    def candidates(self, address, cache=None):
        """Entries worth comparing with an address the book doesn't hold: same postcode, same receiver name,
        and full-text matches for the receiver (in case the postcode itself is wrong)."""
        found = {e['id']: e for e in self.near(address, cache=cache)}
        receiver = address.get('receiver')
        if receiver:
            for e in self.find_by_receiver(receiver) + self.search(receiver, 10)[0]:
                found.setdefault(e['id'], e)
        # The street as typed so far ('27 Siri'), in case neither postcode nor name lead to it
        street = " ".join(str(address.get(k) or '') for k in ('line1', 'suburb')).strip()
        if len(street) >= 3:
            for e in self.search(street, 10)[0]:
                found.setdefault(e['id'], e)
        return list(found.values())

    def near(self, address, limit=100, cache=None):
        """Short-list for fuzzy matching: entries with the same postcode (indexed). A crowded postcode
        (a CBD with hundreds of stores) is narrowed to the entries sharing the most receiver/street words,
        so the right one is never cut off by the limit. No postcode: full-text matches for the receiver.
        Pass the same `cache` dict for every row of one import so a crowded postcode is only read once."""
        db = self._connect()
        postcode = normalize_postcode(address.get('postcode'), address.get('country'))
        if not postcode:
            return self.search(address.get('receiver') or '', 20)[0] if address.get('receiver') else []
        cache = {} if cache is None else cache
        if postcode not in cache:
            rows = db.execute("SELECT * FROM addresses WHERE postcode = ? LIMIT ?", (postcode, limit + 1)).fetchall()
            if len(rows) <= limit:
                cache[postcode] = [self._row(r) for r in rows]
            else:
                # Crowded: keep just the words of each entry's name and street (indexed read, cheap)
                cache[postcode] = [(r[0], frozenset(_norm(f"{r[1]} {r[2]} {r[3]}").split())) for r in db.execute(
                    "SELECT id, receiver, line1, line2 FROM addresses WHERE postcode = ?", (postcode,))]
        entries = cache[postcode]
        if not entries or isinstance(entries[0], dict):
            return entries
        # Load in full only the entries sharing the most words with this address
        words = {w for k in ('receiver', 'line1', 'line2') for w in _norm(address.get(k)).split()}
        ids = [i for _, i in heapq.nlargest(CROWDED_SHORTLIST, ((len(words & ws), i) for i, ws in entries))]
        rows = db.execute(f"SELECT * FROM addresses WHERE id IN ({','.join('?' * len(ids))})", ids).fetchall()
        return [self._row(r) for r in rows]

    def apply_verified(self, address_id, fields):
        """Replaces an entry with the address the courier portal confirmed, and marks it verified.

        If the confirmed address is one the book already holds under another entry, the two are merged:
        the other entry is kept (with the confirmed details), it takes over this entry's use count and
        learned Excel spellings, and this entry is removed. Returns the entry that now holds the address."""
        current = self.get(address_id)
        if not current:
            raise AddressBookError("That address no longer exists.")
        a = _clean({**current, **fields})
        if not (a['receiver'] and a['postcode']):
            raise AddressBookError("A verified address needs at least a receiver and a postcode.")
        keys = _keys(a)
        now = time.time()
        db = self._connect()
        with db:
            other = db.execute("SELECT id, use_count FROM addresses WHERE address_key = ? AND receiver_key = ? "
                               "AND contact_key = ? AND id != ?", (*keys, address_id)).fetchone()
            target = other['id'] if other else address_id
            if other:
                db.execute("UPDATE aliases SET address_id = ? WHERE address_id = ?", (target, address_id))
                db.execute("UPDATE addresses SET use_count = use_count + ? WHERE id = ?", (current['use_count'], target))
                db.execute("DELETE FROM addresses WHERE id = ?", (address_id,))
            db.execute(f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS)}, address_key = ?, receiver_key = ?, "
                       f"contact_key = ?, updated_at = ?, verified_at = ? WHERE id = ?",
                       [a[k] for k in FIELDS] + list(keys) + [now, now, target])
        return self.get(target)

    def delete(self, address_id):
        db = self._connect()
        with db:
            return db.execute("DELETE FROM addresses WHERE id = ?", (address_id,)).rowcount > 0

    def record_used(self, entries):
        """After a Generate: save or refresh each delivery address and remember the raw spellings that led to it.

        A receiver or Attn not yet saved at that address is added as its own entry (one address can have several);
        no Attn means the receiver's existing entry there. A new entry at an address the courier portal has
        verified is verified too. entries: [(destination dict, [raw source keys]), ...]. All in one transaction,
        so a Generate of hundreds of consignments is a single fast write."""
        now = time.time()
        db = self._connect()
        with db:
            for dest, source_keys in entries:
                a = _clean(dest)
                if not (a['receiver'] and a['postcode']):
                    continue
                address_key, receiver_key, contact_key = _keys(a)
                row = db.execute("SELECT id, contact FROM addresses WHERE address_key = ? AND receiver_key = ? AND contact_key = ?",
                                 (address_key, receiver_key, contact_key)).fetchone()
                if row is None and not contact_key:
                    row = db.execute("SELECT id, contact FROM addresses WHERE address_key = ? AND receiver_key = ? "
                                     "ORDER BY use_count DESC LIMIT 1", (address_key, receiver_key)).fetchone()
                if row:
                    address_id = row['id']
                    # Refresh to the latest spelling/ATL actually sent (a blank Attn leaves the saved one)
                    fields = {**a, 'contact': a['contact'] or row['contact']}
                    db.execute(f"UPDATE addresses SET {', '.join(f'{k} = ?' for k in FIELDS)}, use_count = use_count + 1, "
                               f"last_used_at = ?, updated_at = ? WHERE id = ?",
                               [fields[k] for k in FIELDS] + [now, now, address_id])
                else:
                    verified = db.execute("SELECT MAX(verified_at) FROM addresses WHERE address_key = ?",
                                          (address_key,)).fetchone()[0]
                    address_id = db.execute(
                        f"INSERT INTO addresses ({', '.join(FIELDS)}, address_key, receiver_key, contact_key, use_count, "
                        f"created_at, updated_at, last_used_at, verified_at) VALUES ({', '.join('?' * (len(FIELDS) + 8))})",
                        [a[k] for k in FIELDS] + [address_key, receiver_key, contact_key, 1, now, now, now, verified]).lastrowid
                for key in dict.fromkeys(source_keys):
                    if key:
                        db.execute("INSERT INTO aliases (source_key, address_id) VALUES (?, ?) "
                                   "ON CONFLICT(source_key) DO UPDATE SET address_id = excluded.address_id",
                                   (key, address_id))
