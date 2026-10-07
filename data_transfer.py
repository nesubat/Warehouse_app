"""Back up or move the app's own data: any of the address book, packing specs and settings, as one zip.

The data folder is left out of git (it's yours, not the code's), so a fresh copy of the app starts empty. Export makes
'warehouse-data-<date>.zip' with the parts chosen (at least one); Import on the other copy shows what the zip holds and
takes the parts chosen there:

  address_book           data/address_book.db: addresses, verified marks, Keep-as-is choices, learned Excel spellings
                         Merge (adds what's missing, takes verified marks over) or Replace (the zip's book only)
  packing_specs          data/packing_specs.json      Merge (adds the named specs missing here) or Replace
  export_details         data/export_details.json     Replace
  service_usage          data/service_code_usage.json Replace
  address_data_settings  data/address_reference.json  (the LINZ API key) Replace

The official address data (G-NAF, LINZ) isn't included: it's large and the Address Data page loads it again. Before an
import changes anything, the current files are copied to data/backups/before-import-<time>/.
"""
import io
import json
import os
import shutil
import sqlite3
import tempfile
import time
import zipfile
from datetime import datetime

FORMAT = 1
ITEMS = {  # key: (label, what it holds, file in data/, import modes)
    'address_book': ('Address book', 'Saved delivery addresses with their verified marks, Keep-as-is choices and the Excel '
                     'spellings learned for them', 'address_book.db', ('merge', 'replace')),
    'packing_specs': ('Packing specs', 'Sizes, weights and item types for each Packing Spec', 'packing_specs.json', ('merge', 'replace')),
    'export_details': ('Export details', 'Defaults for deliveries outside Australia in the OpenFreight CSV', 'export_details.json', ('replace',)),
    'service_usage': ('Service code order', 'How often each courier service was used (the order of the service list)',
                      'service_code_usage.json', ('replace',)),
    'address_data_settings': ('Address data settings', 'Your LINZ API key and when the address data was last checked '
                              '(not the address data itself)', 'address_reference.json', ('replace',)),
}


class TransferError(ValueError):
    """The zip can't be used (not an export of this app, damaged, nothing chosen)."""


def _read_json(path, default=None):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def _write_json(path, data):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def _summary(key, path=None, book=None, data=None):
    """One line about what an item holds ('32 addresses, 30 verified')."""
    try:
        if key == 'address_book':
            db = book._connect() if book is not None else sqlite3.connect(f"file:{path}?mode=ro", uri=True)
            n = db.execute("SELECT COUNT(*) FROM addresses").fetchone()[0]
            v = db.execute("SELECT COUNT(*) FROM addresses WHERE verified_at IS NOT NULL").fetchone()[0]
            a = db.execute("SELECT COUNT(*) FROM aliases").fetchone()[0]
            if book is None:
                db.close()
            return f"{n:,} address{'es' if n != 1 else ''}, {v:,} verified, {a:,} learned spelling{'s' if a != 1 else ''}"
        data = data if data is not None else _read_json(path, {})
        if key == 'packing_specs':
            return f"{len(data.get('named') or [])} named specs + the OB / CS / Pallet / FP formulas"
        if key == 'export_details':
            return ", ".join(f"{k}: {v}" for k, v in list(data.items())[:3]) + (' …' if len(data) > 3 else '')
        if key == 'service_usage':
            top = sorted(data.items(), key=lambda kv: -kv[1])[:3]
            return f"{len(data)} services used" + (f" (most: {', '.join(k for k, _ in top)})" if top else '')
        if key == 'address_data_settings':
            return ("LINZ API key saved" if data.get('linz_key') else "no LINZ API key") + \
                   (f", last checked {data['last_checked'][:10]}" if data.get('last_checked') else '')
    except Exception:
        return ''
    return ''


def what_is_here(data_dir, book):
    """The items this copy of the app has, for the Export list: [{key, label, about, summary, modes, present}]."""
    out = []
    for key, (label, about, name, modes) in ITEMS.items():
        path = os.path.join(data_dir, name)
        present = book.count() > 0 if key == 'address_book' else os.path.exists(path)
        out.append({'key': key, 'label': label, 'about': about, 'modes': modes, 'present': present,
                    'summary': _summary(key, path, book if key == 'address_book' else None) if present else 'nothing saved yet'})
    return out


def export_zip(data_dir, book, chosen):
    """The chosen items as one zip (bytes), with a manifest describing them."""
    chosen = [k for k in ITEMS if k in set(chosen or [])]
    if not chosen:
        raise TransferError("Choose at least one thing to export.")
    buf = io.BytesIO()
    manifest = {'app': 'Warehouse app', 'format': FORMAT, 'exported_at': datetime.now().isoformat(timespec='seconds'),
                'items': {}}
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for key in chosen:
            label, _, name, _ = ITEMS[key]
            path = os.path.join(data_dir, name)
            if key == 'address_book':
                # A consistent copy even while the app is using the book (SQLite's backup)
                with tempfile.TemporaryDirectory() as tmp:
                    copy = os.path.join(tmp, name)
                    dest = sqlite3.connect(copy)
                    book._connect().backup(dest)
                    dest.close()
                    z.write(copy, name)
            elif os.path.exists(path):
                z.write(path, name)
            else:
                continue
            manifest['items'][key] = {'label': label, 'file': name, 'summary': _summary(key, path, book if key == 'address_book' else None)}
        z.writestr('manifest.json', json.dumps(manifest, indent=2, ensure_ascii=False))
    if not manifest['items']:
        raise TransferError("Nothing to export: none of the chosen things is saved yet.")
    return buf.getvalue(), manifest


def read_zip(data):
    """(manifest, {key: bytes}) of an exported zip, or TransferError."""
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        manifest = json.loads(z.read('manifest.json').decode('utf-8'))
    except (zipfile.BadZipFile, KeyError, ValueError):
        raise TransferError("That isn't a data export of this app (no manifest.json inside).")
    if manifest.get('app') != 'Warehouse app' or int(manifest.get('format', 0)) > FORMAT:
        raise TransferError("That zip was made by another app or a newer version of this one.")
    files = {}
    for key, info in (manifest.get('items') or {}).items():
        if key in ITEMS and ITEMS[key][2] in z.namelist():
            files[key] = z.read(ITEMS[key][2])
    return manifest, files


def describe_zip(data):
    """What an exported zip holds, for the Import list: (manifest, [{key, label, about, summary, modes}])."""
    manifest, files = read_zip(data)
    out = []
    for key, raw in files.items():
        label, about, name, modes = ITEMS[key]
        summary = manifest['items'].get(key, {}).get('summary', '')
        out.append({'key': key, 'label': label, 'about': about, 'modes': modes, 'summary': summary})
    if not out:
        raise TransferError("That export holds nothing this app can take.")
    return manifest, out


def _backup(data_dir, book, keys):
    """Copies of everything an import is about to change, in data/backups/before-import-<time>/."""
    folder = os.path.join(data_dir, 'backups', f"before-import-{datetime.now():%Y%m%d-%H%M%S}")
    os.makedirs(folder, exist_ok=True)
    for key in keys:
        name = ITEMS[key][2]
        if key == 'address_book':
            dest = sqlite3.connect(os.path.join(folder, name))
            book._connect().backup(dest)
            dest.close()
        elif os.path.exists(os.path.join(data_dir, name)):
            shutil.copy2(os.path.join(data_dir, name), os.path.join(folder, name))
    return folder


ENTRY_COLUMNS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country', 'authority_to_leave',
                 'address_key', 'receiver_key', 'contact_key', 'use_count', 'created_at', 'updated_at', 'last_used_at',
                 'verified_at', 'ref_kept')


def _import_book(book, raw, mode):
    """Merges the exported address book into this one (or replaces it). The export is opened through AddressBook
    first, so a book from an older version of the app is brought up to date before it's read."""
    from address_book import AddressBook
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'import.db')
        with open(path, 'wb') as f:
            f.write(raw)
        other = AddressBook(path)
        src = other._connect()
        rows = src.execute(f"SELECT id, {', '.join(ENTRY_COLUMNS)} FROM addresses").fetchall()
        aliases = src.execute("SELECT source_key, address_id FROM aliases").fetchall()
        src.close()
        other._local.db = None
    db = book._connect()
    added = updated = learned = 0
    with db:
        if mode == 'replace':
            db.execute("DELETE FROM aliases")
            db.execute("DELETE FROM addresses")
        id_of = {}
        for r in rows:
            values = [r[c] for c in ENTRY_COLUMNS]
            here = db.execute("SELECT id, verified_at, use_count FROM addresses WHERE address_key = ? AND receiver_key = ? "
                              "AND contact_key = ?", (r['address_key'], r['receiver_key'], r['contact_key'])).fetchone()
            if here:
                # Already here: the courier's verified mark and the higher use count carry over
                db.execute("UPDATE addresses SET verified_at = COALESCE(verified_at, ?), use_count = MAX(use_count, ?), "
                           "ref_kept = CASE WHEN ref_kept = '' THEN ? ELSE ref_kept END WHERE id = ?",
                           (r['verified_at'], r['use_count'], r['ref_kept'] or '', here['id']))
                id_of[r['id']] = here['id']
                updated += 1
            else:
                cur = db.execute(f"INSERT INTO addresses ({', '.join(ENTRY_COLUMNS)}) VALUES ({', '.join('?' * len(ENTRY_COLUMNS))})", values)
                id_of[r['id']] = cur.lastrowid
                added += 1
        for a in aliases:
            target = id_of.get(a['address_id'])
            if target and db.execute("INSERT OR IGNORE INTO aliases (source_key, address_id) VALUES (?, ?)",
                                     (a['source_key'], target)).rowcount:
                learned += 1
    return (f"{added:,} added, {updated:,} already here (verified marks kept), {learned:,} learned spellings"
            if mode == 'merge' else f"replaced: {added:,} addresses, {learned:,} learned spellings")


def _import_specs(spec_store, raw, mode):
    incoming = json.loads(raw.decode('utf-8'))
    if mode == 'replace':
        spec_store._write({'formula': incoming.get('formula') or {}, 'named': incoming.get('named') or []})
        return f"replaced: {len(incoming.get('named') or [])} named specs and the formulas"
    current = spec_store.load()
    have = {str(s.get('name', '')).strip().lower() for s in current['named']}
    new = [s for s in incoming.get('named') or [] if str(s.get('name', '')).strip().lower() not in have]
    spec_store._write({'formula': current['formula'], 'named': current['named'] + new})
    return f"{len(new)} named spec{'s' if len(new) != 1 else ''} added (the ones here kept as they are)"


def import_zip(data, choices, data_dir, book, spec_store):
    """Takes the chosen parts of an exported zip. choices: {key: 'merge' | 'replace'}. Returns
    {'backup': folder, 'done': [(label, what happened)]}."""
    manifest, files = read_zip(data)
    chosen = {k: m for k, m in (choices or {}).items() if k in files and m in ITEMS[k][3]}
    if not chosen:
        raise TransferError("Choose at least one thing to import.")
    backup = _backup(data_dir, book, chosen)
    done = []
    for key, mode in chosen.items():
        label, _, name, _ = ITEMS[key]
        raw = files[key]
        if key == 'address_book':
            done.append((label, _import_book(book, raw, mode)))
        elif key == 'packing_specs':
            done.append((label, _import_specs(spec_store, raw, mode)))
        else:
            data_ = json.loads(raw.decode('utf-8'))
            _write_json(os.path.join(data_dir, name), data_)
            done.append((label, f"replaced: {_summary(key, data=data_)}"))
    return {'backup': backup, 'done': done, 'exported_at': manifest.get('exported_at', '')}
