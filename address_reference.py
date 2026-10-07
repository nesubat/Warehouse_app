"""Official address data for checking delivery addresses: G-NAF for Australia and LINZ NZ Addresses for New Zealand,
downloaded and kept locally, updated when the user chooses (the Address Data page reminds after a month).

  Australia   G-NAF (Geoscape, data.gov.au, CC BY 4.0), released quarterly: every locality (suburb) with its state,
              every postcode used in it and its other names; every street in each locality; every street number.
  New Zealand LINZ NZ Addresses (data.linz.govt.nz, CC BY 4.0), updated weekly; needs a free LINZ API key. Suburb or
              locality, town or city, street and number. It has no postcodes (NZ Post's), so an NZ postcode isn't checked.

Each country is its own SQLite file (data/address_ref_au.db, data/address_ref_nz.db), built under a temporary name and
swapped in only when complete: lookups carry on during an update, and a failed update leaves the old data.

check_address(address) says whether suburb + postcode + state belong together, whether the street is in that suburb and
whether the number is on it, with suggestions and the time it took.
"""
import csv
import difflib
import io
import json
import os
import re
import sqlite3
import threading
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
DOWNLOADS = os.path.join(DATA_DIR, 'reference_downloads')
DB_PATHS = {'AU': os.path.join(DATA_DIR, 'address_ref_au.db'), 'NZ': os.path.join(DATA_DIR, 'address_ref_nz.db')}
SETTINGS_FILE = os.path.join(DATA_DIR, 'address_reference.json')  # LINZ key, what's loaded, last checked
REMIND_DAYS = 30
AU_STATES = {'VIC', 'NSW', 'QLD', 'SA', 'WA', 'TAS', 'NT', 'ACT', 'OT'}

GNAF_PACKAGE = 'https://data.gov.au/data/api/3/action/package_show?id=geocoded-national-address-file-g-naf'
LINZ_LAYER = 105689  # NZ Addresses
LINZ_WFS = 'https://data.linz.govt.nz/services;key={key}/wfs'
LINZ_PAGE = 100000
USER_AGENT = 'Mozilla/5.0 (Warehouse app address data)'


def _key(text):
    """Comparison key: capitals, letters and digits only, accents folded ('Mount  Waverley.' -> 'MOUNT WAVERLEY',
    'Tākaka' -> 'TAKAKA')."""
    import unicodedata
    plain = unicodedata.normalize('NFKD', str(text or '')).encode('ascii', 'ignore').decode()
    return " ".join(re.sub(r"[^0-9A-Z]+", " ", plain.upper()).split())


# ---------- settings ----------

def load_settings():
    try:
        with open(SETTINGS_FILE, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(changes):
    data = {**load_settings(), **changes}
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = SETTINGS_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, SETTINGS_FILE)
    return data


def loaded_info(country):
    """What's loaded for a country: {'version', 'loaded_at', 'counts', 'size_mb'} or None."""
    path = DB_PATHS[country]
    if not os.path.exists(path):
        return None
    try:
        with sqlite3.connect(f"file:{path}?mode=ro", uri=True) as db:
            meta = dict(db.execute("SELECT key, value FROM meta").fetchall())
    except sqlite3.Error:
        return None
    return {'version': meta.get('version', ''), 'loaded_at': meta.get('loaded_at', ''),
            'counts': json.loads(meta.get('counts', '{}')), 'source': meta.get('source', ''),
            'size_mb': round(os.path.getsize(path) / 1e6)}


def reminder():
    """Countries whose data is more than REMIND_DAYS old (or never loaded), for the page's update reminder."""
    due = []
    for country in DB_PATHS:
        info = loaded_info(country)
        if not info:
            due.append((country, None))
            continue
        try:
            age = (datetime.now() - datetime.fromisoformat(info['loaded_at'])).days
        except ValueError:
            age = REMIND_DAYS + 1
        if age > REMIND_DAYS:
            due.append((country, age))
    return due


# ---------- background jobs ----------

JOB = {'state': 'idle'}  # one download / import at a time: {'state', 'country', 'step', 'done', 'total', 'unit', 'message'}
_JOB_LOCK = threading.Lock()


def job_status():
    return dict(JOB)


def _progress(**values):
    JOB.update(values)


def start(country, log=None, on_done=None):
    """Download and load a country's data in the background. Returns an error text, or None when started.
    on_done() runs after a successful load (the app re-checks the address book against the new data)."""
    with _JOB_LOCK:
        if JOB.get('state') == 'running':
            return f"Already loading {JOB.get('country')}: wait for it to finish."
        if country == 'NZ' and not load_settings().get('linz_key') and not linz_export():
            return ("Download NZ Addresses from data.linz.govt.nz (Shapefile) into data/reference_downloads, "
                    "or enter your LINZ API key first.")
        JOB.clear()
        JOB.update(state='running', country=country, step='Starting', done=0, total=0, unit='', message='',
                   started=time.time())

    def work():
        try:
            (_load_au if country == 'AU' else _load_nz)(log)
            _progress(state='done', step='Done', message=f"{country} address data loaded.")
            if on_done:
                on_done()
        except Exception as e:  # shown on the page; the old data stays in place
            if log:
                log.exception("Address data %s failed", country)
            _progress(state='failed', message=str(e))
    threading.Thread(target=work, daemon=True).start()
    return None


# ---------- downloading ----------

def _open(url, headers=None, timeout=60):
    return urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': USER_AGENT, **(headers or {})}),
                                  timeout=timeout)


def _get_json(url, tries=6):
    for attempt in range(tries):
        try:
            with _open(url) as r:
                return json.load(r)
        except (OSError, ValueError) as e:
            if attempt == tries - 1:
                raise RuntimeError(f"Couldn't reach {urllib.parse.urlsplit(url).netloc}: {e}") from e
            time.sleep(3 + 3 * attempt)


def _download(url, path, size=None, tries=40):
    """A big file, resumed after every dropped connection (data.gov.au cuts some off). Kept as <path>.part until whole."""
    part = path + '.part'
    for attempt in range(tries):
        have = os.path.getsize(part) if os.path.exists(part) else 0
        if size and have >= size:
            break
        try:
            with _open(url, {'Range': f'bytes={have}-'} if have else None, timeout=120) as r:
                if have and r.status != 206:
                    have = 0  # the server ignored the resume: start again
                total = size or (have + int(r.headers.get('Content-Length') or 0))
                with open(part, 'ab' if have else 'wb') as f:
                    while True:
                        chunk = r.read(1 << 20)
                        if not chunk:
                            break
                        f.write(chunk)
                        have += len(chunk)
                        _progress(step='Downloading', done=round(have / 1e6), total=round(total / 1e6), unit='MB')
            if not size or have >= size:
                break
        except OSError as e:
            _progress(message=f"Connection dropped ({e.__class__.__name__}), resuming… (try {attempt + 1})")
            time.sleep(min(5 + 5 * attempt, 30))
    else:
        raise RuntimeError("The download kept failing. Try again later.")
    os.replace(part, path)
    return path


def gnaf_latest():
    """(release name, zip URL, size in bytes) of the newest G-NAF (GDA2020, pipe-separated) on data.gov.au."""
    res = _get_json(GNAF_PACKAGE)['result']['resources']
    zips = [r for r in res if str(r.get('format', '')).upper() == 'ZIP' and 'psv' in r['url'].lower()]
    zips.sort(key=lambda r: ('gda2020' in r['url'].lower(), r.get('last_modified') or ''), reverse=True)
    if not zips:
        raise RuntimeError("No G-NAF download found on data.gov.au.")
    best = zips[0]
    name = re.sub(r'\.zip$', '', os.path.basename(urllib.parse.urlsplit(best['url']).path))
    return name, best['url'], int(best.get('size') or 0) or None


def check_updates():
    """{country: {'loaded', 'latest', 'newer'}}: what's on the source now against what's loaded. Saves when checked."""
    out = {}
    au = loaded_info('AU')
    try:
        name, _, size = gnaf_latest()
        out['AU'] = {'loaded': au and au['version'], 'latest': name, 'size_mb': round((size or 0) / 1e6),
                     'newer': not au or au['version'] != name}
    except Exception as e:
        out['AU'] = {'loaded': au and au['version'], 'error': str(e)}
    nz = loaded_info('NZ')
    # LINZ changes its addresses every week and has no release names: a month-old copy is worth refreshing
    out['NZ'] = {'loaded': nz and nz['version'], 'latest': 'updated weekly by LINZ',
                 'newer': not nz or (datetime.now() - datetime.fromisoformat(nz['loaded_at'])).days > 7}
    save_settings({'last_checked': datetime.now().isoformat(timespec='seconds')})
    return out


# ---------- building a country's database ----------

SCHEMA = """
CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE locality (id INTEGER PRIMARY KEY, name TEXT, key TEXT, state TEXT, town TEXT);
CREATE TABLE locality_name (key TEXT, locality INTEGER);           -- its name and its other names
CREATE TABLE locality_postcode (postcode TEXT, locality INTEGER, PRIMARY KEY (postcode, locality)) WITHOUT ROWID;
CREATE TABLE street (id INTEGER PRIMARY KEY, locality INTEGER, name TEXT, key TEXT);
CREATE TABLE number (street INTEGER, first INTEGER, last INTEGER, PRIMARY KEY (street, first, last)) WITHOUT ROWID;
CREATE TABLE street_type (word TEXT PRIMARY KEY, full TEXT);       -- 'RD' -> 'ROAD', 'ROAD' -> 'ROAD'
"""
INDEXES = """
CREATE INDEX locality_name_key ON locality_name (key);
CREATE INDEX locality_postcode_locality ON locality_postcode (locality);
CREATE INDEX street_locality ON street (locality, key);
CREATE INDEX street_key ON street (key);
"""


def _new_db(country):
    path = DB_PATHS[country] + '.building'
    if os.path.exists(path):
        os.remove(path)
    db = sqlite3.connect(path)
    db.executescript("PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF;" + SCHEMA)
    return db, path


def _finish_db(db, path, country, version, source):
    _progress(step='Indexing', done=0, total=0, unit='')
    db.executescript(INDEXES)
    counts = {t: db.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in ('locality', 'street', 'number')}
    counts['postcodes'] = db.execute("SELECT COUNT(DISTINCT postcode) FROM locality_postcode").fetchone()[0]
    db.executemany("INSERT INTO meta VALUES (?, ?)", [
        ('version', version), ('source', source), ('loaded_at', datetime.now().isoformat(timespec='seconds')),
        ('counts', json.dumps(counts))])
    db.commit()
    db.execute("VACUUM")
    db.close()
    os.replace(path, DB_PATHS[country])  # the new data takes over in one step
    return counts


def _psv(zf, member):
    """Rows of a pipe-separated G-NAF file inside the zip, as dicts by column name."""
    with zf.open(member) as raw:
        yield from csv.DictReader(io.TextIOWrapper(raw, encoding='utf-8-sig', newline=''), delimiter='|')


def _load_au(log=None):
    name, url, size = gnaf_latest()
    os.makedirs(DOWNLOADS, exist_ok=True)
    path = os.path.join(DOWNLOADS, name + '.zip')
    if not (os.path.exists(path) and (not size or os.path.getsize(path) == size)):
        _download(url, path, size)
    _progress(step='Reading G-NAF', done=0, total=0, unit='', message='')
    with zipfile.ZipFile(path) as zf:
        members = zf.namelist()
        find = lambda table: sorted(m for m in members if re.search(rf'/Standard/[A-Z]+_{table}_psv\.psv$', m))
        db, building = _new_db('AU')

        # Street types: 'ROAD' is written 'RD' too
        types = next((m for m in members if m.endswith('STREET_TYPE_AUT_psv.psv')), None)
        if types:
            pairs = set()
            for r in _psv(zf, types):
                full = _key(r.get('CODE')); abbr = _key(r.get('NAME'))
                if full:
                    pairs.add((full, full))
                    if abbr:
                        pairs.add((abbr, full))
            db.executemany("INSERT OR IGNORE INTO street_type VALUES (?, ?)", sorted(pairs))

        states = {}
        for m in find('STATE'):
            for r in _psv(zf, m):
                states[r['STATE_PID']] = r['STATE_ABBREVIATION']
        loc_id = {}
        for m in find('LOCALITY'):
            for r in _psv(zf, m):
                if r.get('DATE_RETIRED'):
                    continue
                i = len(loc_id) + 1
                loc_id[r['LOCALITY_PID']] = i
                db.execute("INSERT INTO locality VALUES (?, ?, ?, ?, '')",
                           (i, r['LOCALITY_NAME'], _key(r['LOCALITY_NAME']), states.get(r['STATE_PID'], '')))
                db.execute("INSERT INTO locality_name VALUES (?, ?)", (_key(r['LOCALITY_NAME']), i))
                if r.get('PRIMARY_POSTCODE'):
                    db.execute("INSERT OR IGNORE INTO locality_postcode VALUES (?, ?)", (r['PRIMARY_POSTCODE'], i))
        for m in find('LOCALITY_ALIAS'):
            for r in _psv(zf, m):
                i = loc_id.get(r['LOCALITY_PID'])
                if i and not r.get('DATE_RETIRED') and r.get('NAME'):
                    db.execute("INSERT INTO locality_name VALUES (?, ?)", (_key(r['NAME']), i))
                    if r.get('POSTCODE'):
                        db.execute("INSERT OR IGNORE INTO locality_postcode VALUES (?, ?)", (r['POSTCODE'], i))
        street_id = {}
        for m in find('STREET_LOCALITY'):
            rows = []
            for r in _psv(zf, m):
                i = loc_id.get(r['LOCALITY_PID'])
                if not i or r.get('DATE_RETIRED'):
                    continue
                sid = len(street_id) + 1
                street_id[r['STREET_LOCALITY_PID']] = sid
                full = " ".join(x for x in (r.get('STREET_NAME'), r.get('STREET_TYPE_CODE'), r.get('STREET_SUFFIX_CODE')) if x)
                rows.append((sid, i, full, _key(full)))
            db.executemany("INSERT INTO street VALUES (?, ?, ?, ?)", rows)
        db.commit()

        details = find('ADDRESS_DETAIL')
        read = 0
        for n, m in enumerate(details, start=1):
            _progress(step=f"Reading addresses: {re.search(r'/([A-Z]+)_ADDRESS', m).group(1)}", done=n, total=len(details),
                      unit='states', message=f"{read:,} addresses read")
            numbers, postcodes = set(), set()
            for r in _psv(zf, m):
                if r.get('DATE_RETIRED'):
                    continue
                read += 1
                i = loc_id.get(r['LOCALITY_PID'])
                if i and r.get('POSTCODE'):
                    postcodes.add((r['POSTCODE'], i))
                sid = street_id.get(r.get('STREET_LOCALITY_PID'))
                first = r.get('NUMBER_FIRST')
                if sid and first and first.isdigit():
                    last = r.get('NUMBER_LAST')
                    numbers.add((sid, int(first), int(last) if last and last.isdigit() else int(first)))
            db.executemany("INSERT OR IGNORE INTO locality_postcode VALUES (?, ?)", postcodes)
            db.executemany("INSERT OR IGNORE INTO number VALUES (?, ?, ?)", numbers)
            db.commit()
        _progress(message=f"{read:,} addresses read")
    counts = _finish_db(db, building, 'AU', name, 'G-NAF (Geoscape Australia, data.gov.au), CC BY 4.0')
    if log:
        log.info("Address data AU loaded: %s %s", name, counts)
    # The zip is only needed to load again: dropped once a newer one is in
    for f in os.listdir(DOWNLOADS):
        if f.startswith('g-naf') and f != name + '.zip':
            os.remove(os.path.join(DOWNLOADS, f))


def _linz_pages(key):
    """NZ Addresses from the LINZ Data Service as CSV pages (WFS), joined into one file in reference_downloads."""
    os.makedirs(DOWNLOADS, exist_ok=True)
    out = os.path.join(DOWNLOADS, f"linz-nz-addresses-{datetime.now():%Y%m%d}.csv")
    base = LINZ_WFS.format(key=urllib.parse.quote(key.strip()))
    start, header, rows = 0, None, 0
    with open(out + '.part', 'w', encoding='utf-8', newline='') as f:
        while True:
            q = urllib.parse.urlencode({'service': 'WFS', 'version': '2.0.0', 'request': 'GetFeature',
                                        'typeNames': f'data.linz.govt.nz:layer-{LINZ_LAYER}', 'outputFormat': 'csv',
                                        'count': LINZ_PAGE, 'startIndex': start, 'sortBy': 'address_id'})
            for attempt in range(8):
                try:
                    with _open(f"{base}?{q}", timeout=600) as r:
                        text = r.read().decode('utf-8-sig')
                    break
                except urllib.error.HTTPError as e:
                    if 400 <= e.code < 500:  # LINZ said no (a wrong key, a key without access): no point retrying
                        said = re.sub(r'\s+', ' ', re.sub('<[^>]+>', ' ', e.read().decode('utf-8', 'replace'))).strip()
                        raise RuntimeError(f"LINZ refused the request ({e.code}): {said[:250] or e.reason}. Check the API key "
                                           f"on data.linz.govt.nz (My account → API keys).")
                    if attempt == 7:
                        raise
                except OSError:
                    if attempt == 7:
                        raise
                time.sleep(5 + 5 * attempt)
            if text.lstrip().startswith('<'):  # an XML error instead of CSV
                raise RuntimeError("LINZ answered with an error: " + re.sub(r'\s+', ' ', re.sub('<[^>]+>', ' ', text))[:300])
            page = list(csv.reader(io.StringIO(text)))
            if not page:
                break
            if header is None:
                header = page[0]
                csv.writer(f).writerow(header)
            body = page[1:]
            csv.writer(f).writerows(body)
            rows += len(body)
            start += len(body)
            _progress(step='Downloading from LINZ', done=rows, total=0, unit='addresses')
            if len(body) < LINZ_PAGE:
                break
    os.replace(out + '.part', out)
    return out


def linz_export():
    """The newest NZ Addresses export downloaded from the LINZ website (a shapefile zip, 'lds-nz-addresses-SHP.zip') in
    reference_downloads, or None. Loading from it needs no API key."""
    if not os.path.isdir(DOWNLOADS):
        return None
    found = [os.path.join(DOWNLOADS, f) for f in os.listdir(DOWNLOADS)
             if f.lower().startswith('lds-nz-addresses') and f.lower().endswith('.zip')]
    return max(found, key=os.path.getmtime) if found else None


def _dbf_rows(zip_path):
    """The attribute table (.dbf) of a shapefile zip, row by row as {column: text}, without the shapes. dBase names are
    cut to 10 characters ('suburb_loc' = suburb_locality)."""
    import struct
    with zipfile.ZipFile(zip_path) as zf:
        name = next((n for n in zf.namelist() if n.lower().endswith('.dbf')), None)
        if not name:
            raise RuntimeError(f"{os.path.basename(zip_path)} has no .dbf table: export NZ Addresses as a Shapefile.")
        cpg = next((n for n in zf.namelist() if n.lower().endswith('.cpg')), None)
        encoding = (zf.read(cpg).decode('ascii', 'ignore').strip() if cpg else '') or 'utf-8'
        with zf.open(name) as f:
            count, header_len, record_len = struct.unpack('<IHH', f.read(12)[4:12])
            f.read(20)
            fields = []
            for _ in range((header_len - 33) // 32):
                b = f.read(32)
                fields.append((b[:11].split(b'\0')[0].decode('ascii', 'ignore').lower(), b[16]))
            f.read(header_len - 32 - 32 * len(fields))  # the 0x0D end of the header
            for n in range(count):
                rec = f.read(record_len)
                if len(rec) < record_len:
                    break
                if rec[:1] == b'*':  # deleted
                    continue
                row, pos = {}, 1
                for col, width in fields:
                    row[col] = rec[pos:pos + width].decode(encoding, 'replace').strip()
                    pos += width
                yield row
                if n % 100000 == 0:
                    _progress(done=n, total=count, unit='addresses')


def _nz_rows(log=None):
    """NZ addresses as {suburb, suburb_ascii, town, town_ascii, ta, road, road_ascii, number, high}: from a LINZ website
    export in reference_downloads when there's one, else downloaded with the LINZ API key."""
    export = linz_export()
    if export:
        _progress(step=f"Reading {os.path.basename(export)}", done=0, total=0, unit='')
        names = {'suburb': ('suburb_loc',), 'suburb_ascii': ('suburb_l_1',), 'town': ('town_city',), 'town_ascii': ('town_city_',),
                 'ta': ('territoria',), 'road': ('full_road_',), 'road_ascii': ('full_roa_1',), 'number': ('address_nu',),
                 'high': ('address__2',), 'lifecycle': ('address_li',)}
        rows, source, kept = _dbf_rows(export), f"LINZ NZ Addresses (website export {datetime.fromtimestamp(os.path.getmtime(export)):%d %b %Y})", export
    else:
        path = _linz_pages(load_settings().get('linz_key', ''))
        _progress(step='Reading NZ addresses', done=0, total=0, unit='')
        names = {'suburb': ('suburb_locality',), 'suburb_ascii': ('suburb_locality_ascii',), 'town': ('town_city',),
                 'town_ascii': ('town_city_ascii',), 'ta': ('territorial_authority',), 'road': ('full_road_name',),
                 'road_ascii': ('full_road_name_ascii',), 'number': ('address_number',), 'high': ('address_number_high',),
                 'lifecycle': ('address_lifecycle',)}
        f = open(path, encoding='utf-8', newline='')
        rows = ({k.lower(): v for k, v in r.items()} for r in csv.DictReader(f))
        source, kept = f"LINZ NZ Addresses (downloaded {datetime.now():%d %b %Y})", path
    def standard():
        for r in rows:
            yield {k: next((r.get(c, '') for c in cols if c in r), '') for k, cols in names.items()}
    return standard(), source, kept


def _load_nz(log=None):
    rows, version, kept = _nz_rows(log)
    db, building = _new_db('NZ')
    loc_id, street_id, numbers, read = {}, {}, set(), 0
    for r in rows:
        if r['lifecycle'] and r['lifecycle'].lower() not in ('current', ''):
            continue  # proposed or retired
        read += 1
        sub, town = r['suburb'].strip(), r['town'].strip()
        lk = (_key(sub), _key(town))
        if lk not in loc_id:
            i = len(loc_id) + 1
            loc_id[lk] = i
            db.execute("INSERT INTO locality VALUES (?, ?, ?, ?, ?)", (i, sub or town, lk[0] or lk[1], r['ta'].strip(), town))
            # Its suburb and its town both find it, with or without macrons ('Tākaka' / 'Takaka')
            for name in {lk[0], lk[1], _key(r['suburb_ascii']), _key(r['town_ascii'])} - {''}:
                db.execute("INSERT INTO locality_name VALUES (?, ?)", (name, i))
        i = loc_id[lk]
        road = r['road'].strip()
        sk = (i, _key(road))
        if road and sk not in street_id:
            street_id[sk] = len(street_id) + 1
            db.execute("INSERT INTO street VALUES (?, ?, ?, ?)", (street_id[sk], i, road, sk[1]))
        num = r['number'].strip()
        if road and num.isdigit():
            high = r['high'].strip()
            numbers.add((street_id[sk], int(num), int(high) if high.isdigit() and int(high) else int(num)))
        if read % 200000 == 0:
            _progress(message=f"{read:,} addresses read")
    if not read:
        raise RuntimeError("No addresses found in the LINZ file.")
    db.executemany("INSERT OR IGNORE INTO number VALUES (?, ?, ?)", numbers)
    _street_types_from_au(db)
    counts = _finish_db(db, building, 'NZ', version, 'LINZ Data Service, CC BY 4.0')
    if log:
        log.info("Address data NZ loaded: %s %s", version, counts)
    for f in os.listdir(DOWNLOADS):  # keep only the file just loaded
        p = os.path.join(DOWNLOADS, f)
        if (f.startswith('linz-nz-addresses-') or f.lower().startswith('lds-nz-addresses')) and os.path.abspath(p) != os.path.abspath(kept):
            os.remove(p)


def _street_types_from_au(db):
    """NZ street names are written out in full; the AU street-type list ('RD' = 'ROAD') reads typed abbreviations."""
    if os.path.exists(DB_PATHS['AU']):
        with sqlite3.connect(f"file:{DB_PATHS['AU']}?mode=ro", uri=True) as au:
            db.executemany("INSERT OR IGNORE INTO street_type VALUES (?, ?)", au.execute("SELECT word, full FROM street_type"))


# ---------- lookups ----------

_LOCAL = threading.local()
_CHECKED_INDEXES = set()


def _ensure_indexes(path):
    """Indexes added after a database was built (the street-name index for typing suggestions): made once, the first
    time the file is opened after an upgrade."""
    if path in _CHECKED_INDEXES:
        return
    _CHECKED_INDEXES.add(path)
    try:
        with sqlite3.connect(path) as db:
            if not db.execute("SELECT 1 FROM sqlite_master WHERE name = 'street_key'").fetchone():
                db.execute("CREATE INDEX street_key ON street (key)")
    except sqlite3.Error:
        pass  # read-only or busy: suggestions by street name alone are just slower


def _db(country):
    """A read-only connection per thread, reopened when the data was replaced by an update."""
    path = DB_PATHS[country]
    if not os.path.exists(path):
        return None
    _ensure_indexes(path)
    stamp = os.path.getmtime(path)
    cache = getattr(_LOCAL, 'dbs', None)
    if cache is None:
        cache = _LOCAL.dbs = {}
    held = cache.get(country)
    if held and held[1] == stamp:
        return held[0]
    if held:
        held[0].close()
    db = sqlite3.connect(f"file:{path}?mode=ro", uri=True, check_same_thread=False)
    cache[country] = (db, stamp)
    return db


def close_all():
    """Closes this thread's connections (an update can't replace a file that's open on Windows)."""
    for db, _ in (getattr(_LOCAL, 'dbs', None) or {}).values():
        db.close()
    _LOCAL.dbs = {}


_NAMES_CACHE = {}


def _all_names(country, db):
    """Every locality name of a country, for close spellings (cached until the data changes)."""
    stamp = os.path.getmtime(DB_PATHS[country])
    held = _NAMES_CACHE.get(country)
    if not held or held[0] != stamp:
        held = (stamp, sorted({k for (k,) in db.execute("SELECT key FROM locality_name")}))
        _NAMES_CACHE[country] = held
    return held[1]


def _localities(db, ids):
    out = []
    for i in ids:
        r = db.execute("SELECT id, name, state, town FROM locality WHERE id = ?", (i,)).fetchone()
        if r:
            pcs = [p for (p,) in db.execute("SELECT postcode FROM locality_postcode WHERE locality = ? ORDER BY postcode", (i,))]
            out.append({'id': r[0], 'name': r[1], 'state': r[2], 'town': r[3], 'postcodes': pcs})
    return out


# Common ways of writing a street type that G-NAF's own list ('AV', 'BVD' …) doesn't have
EXTRA_TYPES = {'AVE': 'AVENUE', 'AVNU': 'AVENUE', 'BLVD': 'BOULEVARD', 'BOULEVARDE': 'BOULEVARD', 'CRES': 'CRESCENT',
               'CRESC': 'CRESCENT', 'CRT': 'COURT', 'CT': 'COURT', 'DR': 'DRIVE', 'DRV': 'DRIVE', 'HWY': 'HIGHWAY',
               'HWAY': 'HIGHWAY', 'LN': 'LANE', 'PDE': 'PARADE', 'PL': 'PLACE', 'RD': 'ROAD', 'ST': 'STREET',
               'STR': 'STREET', 'TCE': 'TERRACE', 'TERR': 'TERRACE', 'SQ': 'SQUARE', 'CL': 'CLOSE', 'GR': 'GROVE',
               'GRV': 'GROVE', 'CCT': 'CIRCUIT', 'ESP': 'ESPLANADE', 'WY': 'WAY', 'MWY': 'MOTORWAY', 'FWY': 'FREEWAY',
               'PKWY': 'PARKWAY'}
_TYPES_CACHE = {}


def _types(db):
    held = _TYPES_CACHE.get(id(db))
    if held is None:
        held = _TYPES_CACHE[id(db)] = {**EXTRA_TYPES, **dict(db.execute("SELECT word, full FROM street_type").fetchall())}
    return held


def _street_key(db, text):
    """'68 Ricketts Rd.' -> ['68', 'RICKETTS', 'ROAD']: the words, street types written out in full. Stored street names
    go through it too, so 'St Kilda Rd' and 'ST KILDA ROAD' read alike (both 'STREET KILDA ROAD')."""
    types = _types(db)
    return [types.get(w, w) for w in _key(text).split()]


def _find_street(db, locality_ids, lines, number=None):
    """The street of these localities named in the address lines: (street id, street name, locality name, closest
    names). Several localities can have a street of that name (Queen Street in many Auckland suburbs): the one where
    the number exists wins."""
    if not locality_ids:
        return None, None, None, []
    streets = []
    for i in locality_ids:
        streets += [(sid, name, " ".join(_street_key(db, key)), i)
                    for sid, name, key in db.execute("SELECT id, name, key FROM street WHERE locality = ?", (i,))]
    best, found = 0, []
    for line in lines:
        for segment in re.split(r'[,;/]|\bcnr\b|\bcorner\b', line or '', flags=re.I):
            words = _street_key(db, segment)
            if not words:
                continue
            text = " ".join(words)
            for sid, name, key, loc in streets:
                # As written, or ignoring spaces ('Glenhuntly Rd' = 'GLEN HUNTLY ROAD')
                if key and (f" {key} " in f" {text} " or key.replace(' ', '') in text.replace(' ', '')):
                    if len(key) > best:
                        best, found = len(key), []
                    if len(key) == best:
                        found.append((sid, name, loc))
    if found:
        num = _number_in(lines, found[0][1])
        pick = next((f for f in found if num is not None and db.execute(
            "SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1", (f[0], num, num)).fetchone()), found[0])
        place = db.execute("SELECT name FROM locality WHERE id = ?", (pick[2],)).fetchone()[0]
        return pick[0], pick[1], place, []
    # Not found as written: the closest street names of this locality, for a typo
    text = " ".join(" ".join(_street_key(db, l)) for l in lines if l)
    words = [w for w in re.findall(r'[A-Z]+', text) if len(w) > 2]
    names = {key.split()[0]: name for _, name, key, _ in streets if key}  # 'RICKETTS' -> 'RICKETTS ROAD'
    close = [names[m] for w in words for m in difflib.get_close_matches(w, list(names), n=2, cutoff=0.8)]
    return None, None, None, list(dict.fromkeys(close))[:5]


def check_address(address):
    """Checks an address against the official data. address: {line1, line2, suburb, state, postcode, country}.
    Returns {'country', 'status', 'checks': [{'what', 'ok', 'text'}], 'suggestions': [...], 'ms'}.
    status: 'verified' (suburb + postcode + state, street and number all found), 'locality' (the suburb, postcode
    and state agree; the street or number wasn't found), 'mismatch', 'unknown', or 'no_data'."""
    started = time.perf_counter()
    country = (address.get('country') or '').strip().upper()
    country = {'AUSTRALIA': 'AU', 'NEW ZEALAND': 'NZ'}.get(country, country)
    suburb, state = _key(address.get('suburb')), _key(address.get('state'))
    postcode = str(address.get('postcode') or '').strip()
    lines = [address.get('line1') or '', address.get('line2') or '']
    # No country given: an Australian state means Australia (NZ has a Richmond too, and can't check postcodes)
    tried = [country] if country in DB_PATHS else (['AU'] if state in AU_STATES else ['AU', 'NZ'])
    result = None
    for c in tried:
        db = _db(c)
        if db is None:
            continue
        r = _check_in(c, db, suburb, state, postcode, lines)
        if result is None or r['rank'] > result['rank']:
            result = r
        if r['rank'] >= 2:
            break
    if result is None:
        result = {'country': country or '-', 'status': 'no_data', 'rank': -1,
                  'checks': [{'what': 'Data', 'ok': False, 'text': f"No address data loaded for {country or 'AU / NZ'}."}],
                  'suggestions': []}
    result['ms'] = round((time.perf_counter() - started) * 1000, 1)
    return result


def _check_in(country, db, suburb, state, postcode, lines):
    checks, suggestions = [], []
    # The suburb by its own name; its other names ('Richmond' for parts of Abbotsford) only when no suburb is named so
    # NZ: a town finds all its suburbs too (Queenstown -> Frankton, Arthurs Point …)
    ids = [i for (i,) in db.execute("SELECT id FROM locality WHERE key = ?", (suburb,))] if suburb and country == 'AU' else []
    if suburb and not ids:
        ids = [i for (i,) in db.execute("SELECT DISTINCT locality FROM locality_name WHERE key = ?", (suburb,))]
    found = _localities(db, ids)
    if suburb and not found and postcode:
        # A misspelt suburb ('Mount Waverly') whose postcode has a suburb of nearly that name: carry on with that one
        here = _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_postcode WHERE postcode = ?", (postcode,))])
        near = [l for l in here if difflib.SequenceMatcher(None, suburb, _key(l['name'])).ratio() >= 0.85]
        if near:
            checks.append({'what': 'Suburb', 'ok': False, 'text': f"{suburb.title()} is spelt {near[0]['name'].title()}"})
            suggestions.append({'suburb': near[0]['name'], 'state': near[0]['state'], 'postcode': postcode})
            sid, street, place, close_streets = _find_street(db, [l['id'] for l in near], lines)
            if sid:
                num = _number_in(lines, street)
                hit = num is not None and db.execute("SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1",
                                                     (sid, num, num)).fetchone()
                checks.append({'what': 'Street', 'ok': True, 'text': f"{street} is in {near[0]['name']}"
                               + (f"; {num} {'found' if hit else 'not found'}" if num is not None else '')})
            elif any(l.strip() for l in lines):
                checks.append({'what': 'Street', 'ok': False, 'text': f"No street of {near[0]['name']} found in the address lines"
                               + (f"; closest: {', '.join(close_streets)}" if close_streets else '')})
            return {'country': country, 'status': 'mismatch', 'rank': 1, 'checks': checks, 'suggestions': suggestions}
    if country == 'AU':
        same_state = [l for l in found if not state or l['state'] == state] or found
        with_pc = [l for l in same_state if postcode in l['postcodes']]
        if with_pc:
            matched = with_pc
            checks.append({'what': 'Suburb + postcode + state', 'ok': True,
                           'text': f"{with_pc[0]['name']} {with_pc[0]['state']} {postcode}"})
            rank = 2
        elif found:
            matched = same_state
            checks.append({'what': 'Suburb + postcode + state', 'ok': False, 'text': "; ".join(
                f"{l['name']} {l['state']} is postcode {', '.join(l['postcodes']) or '?'}" for l in same_state[:4])})
            suggestions += [{'suburb': l['name'], 'state': l['state'], 'postcode': p} for l in same_state[:3] for p in l['postcodes'][:3]]
            rank = 1
        else:
            matched = []
            rank = 0
        if postcode and not with_pc:
            here = _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_postcode WHERE postcode = ?", (postcode,))])
            if here:
                names = [l['name'] for l in here]
                close = difflib.get_close_matches(suburb, [_key(n) for n in names], n=3, cutoff=0.6)
                checks.append({'what': f'Suburbs with postcode {postcode}', 'ok': None,
                               'text': ", ".join(f"{l['name']} {l['state']}" for l in here[:12]) + (' …' if len(here) > 12 else '')})
                for l in here:
                    if _key(l['name']) in close:
                        suggestions.insert(0, {'suburb': l['name'], 'state': l['state'], 'postcode': postcode})
            elif postcode:
                checks.append({'what': 'Postcode', 'ok': False, 'text': f"{postcode} isn't an Australian postcode in G-NAF."})
    else:
        matched = found
        rank = 2 if found else 0
        if found:
            checks.append({'what': 'Suburb / town', 'ok': True, 'text': ", ".join(
                " · ".join(x for x in (l['name'], l['town'] if l['town'] != l['name'] else '') if x) for l in found[:4])})
        checks.append({'what': 'Postcode', 'ok': None, 'text': "Not checked: LINZ has no postcodes."})
    if suburb and not found:
        close = difflib.get_close_matches(suburb, _all_names(country, db), n=5, cutoff=0.75)
        checks.append({'what': 'Suburb', 'ok': False, 'text': f"{suburb.title()} isn't a {country} suburb"
                       + (f"; closest: {', '.join(c.title() for c in close)}" if close else '')})
        for c in close:
            for l in _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_name WHERE key = ?", (c,))])[:2]:
                suggestions.append({'suburb': l['name'], 'state': l['state'], 'postcode': (l['postcodes'] or [''])[0]})
    # The street, in the matched suburb(s)
    sid, street, place, close_streets = _find_street(db, [l['id'] for l in matched], lines)
    if sid:
        checks.append({'what': 'Street', 'ok': True, 'text': f"{street} is in {place}"})
        num = _number_in(lines, street)
        if num is not None:
            hit = db.execute("SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1", (sid, num, num)).fetchone()
            checks.append({'what': 'Number', 'ok': bool(hit), 'text': f"{num} {street}" + ('' if hit else ' not found (new or unlisted address?)')})
            if rank == 2 and hit:
                rank = 3
    elif matched and any(l.strip() for l in lines):
        checks.append({'what': 'Street', 'ok': False, 'text': f"No street of {matched[0]['name']} found in the address lines"
                       + (f"; closest: {', '.join(close_streets)}" if close_streets else '')})
    status = {3: 'verified', 2: 'locality', 1: 'mismatch', 0: 'unknown'}[rank]
    return {'country': country, 'status': status, 'rank': rank, 'checks': checks,
            'suggestions': [dict(t) for t in dict.fromkeys(tuple(s.items()) for s in suggestions)][:6]}


def _number_in(lines, street):
    """The street number written just before the street's name in the address lines ('68 Ricketts Rd' -> 68; a range
    '268-274 Springvale Rd' -> 268; '604 Glenhuntly Rd' for GLEN HUNTLY ROAD -> 604)."""
    first = _key(street).split()[0]
    for line in lines:
        # Only the number right in front of the street counts (not 'Shop 1', 'B0014/'); a hyphen joins a range
        text = re.sub(r'[^0-9A-Z\-]+', ' ', str(line or '').upper())
        m = re.search(rf'(?<![\dA-Z])(\d+)[A-Z]?(?:\s*-\s*\d+[A-Z]?)?\s+{re.escape(first)}', text)
        if m:
            return int(m.group(1))
    return None


# ---------- field by field: the consignment preview and the address book ----------

ASSESS_FIELDS = ('line1', 'suburb', 'state', 'postcode', 'country')
_ASSESS_CACHE = {}
_ASSESS_CACHE_MAX = 50000


def data_version():
    """Changes whenever a country's data is loaded again: cached verdicts and the address book's checks go stale."""
    return "|".join(f"{c}:{int(os.path.getmtime(p)) if os.path.exists(p) else 0}" for c, p in sorted(DB_PATHS.items()))


def available():
    return any(os.path.exists(p) for p in DB_PATHS.values())


def assess(address):
    """How an address compares with the official data, field by field, for the preview and the address book:
      {'level':   'ok' (agrees), 'note' (only the street isn't on record), 'mismatch' (suburb, state, postcode or
                  country disagree), 'none' (no data for it: another country, nothing loaded, no address),
       'country', 'message': one line saying what's wrong,
       'fields':  {field: {'level': 'bad' | 'note' | 'info', 'text'}} for the fields with something to say,
       'options': [{'label': 'RICHMOND VIC 3121', 'fields': {...}}] one-click fixes, likeliest first,
       'version': the data it was checked against}
    Cached per address and data version: a big job is checked once, and again only after an update."""
    a = {k: str(address.get(k) or '').strip() for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')}
    version = data_version()
    key = (version,) + tuple(_key(a[k]) for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country'))
    hit = _ASSESS_CACHE.get(key)
    if hit is not None:
        return hit
    result = _assess(a)
    result['version'] = version
    if len(_ASSESS_CACHE) > _ASSESS_CACHE_MAX:
        _ASSESS_CACHE.clear()
    _ASSESS_CACHE[key] = result
    return result


def _none(country, message=''):
    return {'level': 'none', 'country': country, 'message': message, 'fields': {}, 'options': []}


def _assess(a):
    country = {'AUSTRALIA': 'AU', 'NEW ZEALAND': 'NZ', '': 'AU'}.get(_key(a['country']), _key(a['country']))
    if not (a['suburb'] or a['postcode']):
        return _none(country)
    if country not in DB_PATHS:
        return _none(country, f"No official data for {country}.")
    db = _db(country)
    if db is None:
        return _none(country, f"{country} address data isn't loaded.")
    result = _assess_in(country, db, a)
    # Clearly the other country: its suburb and street are on record there, and this country doesn't know the suburb
    other = 'NZ' if country == 'AU' else 'AU'
    if result['level'] == 'mismatch' and (result.get('suburb_unknown') or result.get('foreign_postcode')) and _db(other) is not None:
        there = _assess_in(other, _db(other), a)
        if there['level'] in ('ok',) and there.get('street_found'):
            name = {'AU': 'an Australian', 'NZ': 'a New Zealand'}[other]
            fix = {'country': other}
            if other == 'NZ' and a['state']:
                fix['state'] = ''
            return {'level': 'mismatch', 'country': country,
                    'message': f"This looks like {name} address: {there.get('place', '')} is on record in {other}.",
                    'fields': {'country': {'level': 'bad', 'text': f"Looks like {name} address ({there.get('place', '')})"}},
                    'options': [{'label': f"Country {other}", 'fields': fix}]}
    return result


def _option(l, postcode=None, country='AU'):
    pc = postcode if postcode is not None else (l['postcodes'][0] if l['postcodes'] else '')
    if country == 'NZ':
        return {'label': " · ".join(x for x in (l['name'], l['town'] if l['town'] and l['town'] != l['name'] else '') if x),
                'fields': {'suburb': l['name'].upper()}}
    return {'label': f"{l['name']} {l['state']} {pc}".strip(), 'fields': {'suburb': l['name'], 'state': l['state'], 'postcode': pc}}


def _street_in(db, localities, lines):
    """(street id, street name, place, closest names) of the street named in the lines, among these localities."""
    if not localities or not any(x.strip() for x in lines):
        return None, None, None, []
    return _find_street(db, [l['id'] for l in localities], lines)


def _assess_in(country, db, a):
    suburb, state, postcode = _key(a['suburb']), _key(a['state']), a['postcode'].strip()
    lines = [a['line1'], a['line2']]
    fields, options, message = {}, [], ''
    named = []
    if suburb:
        ids = [i for (i,) in db.execute("SELECT id FROM locality WHERE key = ?", (suburb,))] if country == 'AU' else []
        if not ids:
            ids = [i for (i,) in db.execute("SELECT DISTINCT locality FROM locality_name WHERE key = ?", (suburb,))]
        named = _localities(db, ids)
    at_postcode = _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_postcode WHERE postcode = ?", (postcode,))]) \
        if postcode and country == 'AU' else []
    matched, suburb_unknown = [], False

    if country == 'NZ':
        # LINZ has suburbs and towns but no postcodes or regions: only the suburb / town is checked
        if named:
            matched = named
        elif suburb:
            suburb_unknown = True
            close = difflib.get_close_matches(suburb, _all_names(country, db), n=3, cutoff=0.8)
            fields['suburb'] = {'level': 'bad', 'text': f"{a['suburb']} isn't a New Zealand suburb or town on record"
                                + (f" (closest: {', '.join(c.title() for c in close)})" if close else '')}
            for c in close:
                for l in _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_name WHERE key = ?", (c,))])[:1]:
                    options.append(_option(l, country='NZ'))
            message = fields['suburb']['text']
    else:
        in_state = [l for l in named if not state or l['state'] == state]
        exact = [l for l in in_state if postcode in l['postcodes']]
        if exact:
            matched = exact
        elif named and not in_state:
            # The suburb exists, but not in this state
            elsewhere = [l for l in named if postcode in l['postcodes']] or named
            matched = elsewhere[:1]
            fields['state'] = {'level': 'bad', 'text': f"{named[0]['name']} isn't in {state}: it's in "
                               + ", ".join(sorted({l['state'] for l in named}))}
            options += [_option(l, postcode if postcode in l['postcodes'] else None) for l in elsewhere[:3]]
            message = fields['state']['text']
        elif named:
            # The suburb exists in this state with another postcode: the postcode, or the suburb, is wrong.
            # The street decides when it can (12 Smith St is in Richmond, not in 3122's Hawthorn)
            by_name = _street_in(db, in_state, lines)[0]
            by_postcode = [l for l in at_postcode if _street_in(db, [l], lines)[0]]
            pcs = sorted({p for l in in_state for p in l['postcodes']})
            here = ", ".join(sorted({l['name'] for l in at_postcode})) if at_postcode else ''
            if by_postcode and not by_name:
                matched = by_postcode
                fields['suburb'] = {'level': 'bad', 'text': f"The street is in {by_postcode[0]['name']} {postcode}, not {named[0]['name']}"}
                options += [_option(l, postcode) for l in by_postcode[:2]]
                options += [_option(in_state[0], p) for p in pcs[:1]]
                message = fields['suburb']['text']
            else:
                matched = in_state
                places = "; ".join(f"{l['name']} {l['state']} is {', '.join(l['postcodes'][:3])}" for l in in_state[:3])
                fields['postcode'] = {'level': 'bad', 'text': f"{postcode or 'No postcode'}: {places}" + (f"; {postcode} is {here}" if here else '')}
                options += [_option(in_state[0], p) for p in pcs[:2]]
                if not by_name:
                    options += [_option(l, postcode) for l in at_postcode[:1]]
                message = fields['postcode']['text']
        elif suburb:
            suburb_unknown = True
            near = [l for l in at_postcode if difflib.SequenceMatcher(None, suburb, _key(l['name'])).ratio() >= 0.85]
            with_street = [l for l in at_postcode if _street_in(db, [l], lines)[0]]
            if near:
                matched = near[:1]
                fields['suburb'] = {'level': 'bad', 'text': f"{a['suburb']} is spelt {near[0]['name']}"}
                options.append(_option(near[0], postcode))
            else:
                close = difflib.get_close_matches(suburb, _all_names(country, db), n=3, cutoff=0.8)
                fields['suburb'] = {'level': 'bad', 'text': f"{a['suburb']} isn't an Australian suburb on record"
                                    + (f" (closest: {', '.join(c.title() for c in close)})" if close else '')}
                matched = with_street[:1]
                options += [_option(l, postcode) for l in with_street[:2]]
                for c in close:
                    for l in _localities(db, [i for (i,) in db.execute("SELECT locality FROM locality_name WHERE key = ?", (c,))]):
                        if not state or l['state'] == state:
                            options.append(_option(l))
                            break
            message = fields['suburb']['text']
        elif at_postcode:
            matched = at_postcode
        if postcode and not at_postcode and 'postcode' not in fields:
            fields['postcode'] = {'level': 'bad', 'text': f"{postcode} isn't an Australian postcode on record"}
            message = message or fields['postcode']['text']
        if state and state not in AU_STATES and 'state' not in fields:
            fields['state'] = {'level': 'bad', 'text': f"{a['state']} isn't an Australian state"}
            options += [_option(l) for l in matched[:1]]
            message = message or fields['state']['text']

    # The street: soft, the official data lags new streets and centres. A number not on record is only mentioned.
    street_found, place = False, (matched[0]['name'] if matched else '')
    sid, street, where, close_streets = _street_in(db, matched, lines)
    if sid:
        street_found, place = True, where
        num = _number_in(lines, street)
        if num is not None and not db.execute("SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1",
                                              (sid, num, num)).fetchone():
            fields['line1'] = {'level': 'info', 'text': f"{num} {street.title()} isn't on record (new or unlisted?)"}
    elif matched and any(x.strip() for x in lines):
        fields['line1'] = {'level': 'note', 'text': f"Street not on record in {matched[0]['name']}"
                           + (f" (closest: {', '.join(n.title() for n in close_streets[:3])})" if close_streets else '')}

    levels = {f['level'] for f in fields.values()}
    level = 'mismatch' if 'bad' in levels else ('note' if 'note' in levels else 'ok')
    if level == 'note':
        message = fields['line1']['text']
    seen, unique = set(), []
    for o in options:
        k = tuple(sorted(o['fields'].items()))
        if k not in seen and any(_key(v) != _key(a.get(f, '')) for f, v in o['fields'].items()):
            seen.add(k)
            unique.append(o)
    return {'level': level, 'country': country, 'message': message, 'fields': fields, 'options': unique[:3],
            'suburb_unknown': suburb_unknown, 'street_found': street_found, 'place': place,
            'foreign_postcode': bool(country == 'AU' and postcode and not at_postcode)}


# ---------- typing suggestions, field by field ----------

def _title(text):
    return " ".join(w if any(c.isdigit() for c in w) else w.capitalize() for w in str(text).split())


def _prefix_range(key):
    return key, key + '￿'


def _context_localities(db, country, ctx):
    suburb = _key(ctx.get('suburb'))
    if not suburb:
        return []
    ids = [i for (i,) in db.execute("SELECT DISTINCT locality FROM locality_name WHERE key = ?", (suburb,))]
    found = _localities(db, ids)
    state, postcode = _key(ctx.get('state')), str(ctx.get('postcode') or '').strip()
    if country == 'AU':
        found = [l for l in found if (not state or l['state'] == state)] or found
        found = [l for l in found if not postcode or postcode in l['postcodes']] or found
    return found


def complete(field, value, ctx, limit=6):
    """Official suggestions for one field as it's typed: [{'label', 'detail', 'fields'}].
      suburb    'RICHMOND VIC 3121' fills suburb, state and postcode (NZ: the suburb, its town shown)
      postcode  '3121 · RICHMOND VIC' fills postcode, suburb and state
      state     the states the typed suburb is in
      line1/2   the streets of the typed suburb that start like this ('68 Rick' -> '68 Ricketts Road'); with no
                suburb yet, streets anywhere, each with its suburb, as a complete address
      country   AU / NZ"""
    value = str(value or '').strip()
    ctx = ctx or {}
    want = {'AUSTRALIA': 'AU', 'NEW ZEALAND': 'NZ'}.get(_key(ctx.get('country')), _key(ctx.get('country')))
    if want not in DB_PATHS and _key(ctx.get('state')) in AU_STATES:
        want = 'AU'  # an Australian state typed: Australia
    countries = [want] if want in DB_PATHS else [c for c in ('AU', 'NZ') if _db(c) is not None]
    out = []
    if field == 'country':
        return [{'label': c, 'detail': {'AU': 'Australia', 'NZ': 'New Zealand'}[c], 'fields': {'country': c}}
                for c in ('AU', 'NZ') if c.startswith(_key(value))]
    for country in countries:
        db = _db(country)
        if db is None:
            continue
        if field == 'suburb' and len(_key(value)) >= 2:
            lo, hi = _prefix_range(_key(value))
            ids = [i for (i,) in db.execute("SELECT DISTINCT locality FROM locality_name WHERE key >= ? AND key < ? LIMIT 60", (lo, hi))]
            state, postcode = _key(ctx.get('state')), str(ctx.get('postcode') or '').strip()
            locs = _localities(db, ids)
            typed = _key(value)
            # Its own name first ('RICHMOND' before Burnley, which G-NAF also calls Richmond), the exact name first of all
            locs.sort(key=lambda l: (not _key(l['name']).startswith(typed), _key(l['name']) != typed, bool(state) and l['state'] != state,
                                     bool(postcode) and postcode not in l['postcodes'], len(l['name']), l['name']))
            for l in locs[:limit]:
                if country == 'NZ':
                    o = _option(l, country='NZ')
                    out.append({'label': l['name'], 'detail': f"{l['town'] or ''} · NZ".strip(' ·'), 'fields': {**o['fields'], 'country': 'NZ'}})
                    continue
                pcs = [postcode] if postcode in l['postcodes'] else l['postcodes'][:2]
                for pc in pcs or ['']:
                    out.append({'label': f"{l['name']} {l['state']} {pc}".strip(), 'detail': 'AU',
                                'fields': {'suburb': l['name'], 'state': l['state'], 'postcode': pc, 'country': 'AU'}})
        elif field == 'postcode' and country == 'AU' and re.fullmatch(r'\d{2,4}', value):
            lo, hi = _prefix_range(value)
            rows = db.execute("SELECT postcode, locality FROM locality_postcode WHERE postcode >= ? AND postcode < ? LIMIT 80", (lo, hi)).fetchall()
            suburb = _key(ctx.get('suburb'))
            locs = [(pc, l) for pc, i in rows for l in _localities(db, [i])]
            locs.sort(key=lambda t: (bool(suburb) and _key(t[1]['name']) != suburb, t[0], t[1]['name']))
            for pc, l in locs[:limit]:
                out.append({'label': f"{pc} · {l['name']} {l['state']}", 'detail': 'AU',
                            'fields': {'postcode': pc, 'suburb': l['name'], 'state': l['state'], 'country': 'AU'}})
        elif field == 'state' and country == 'AU':
            states = sorted({l['state'] for l in _context_localities(db, country, ctx)}) or sorted(AU_STATES - {'OT'})
            out += [{'label': s, 'detail': 'AU', 'fields': {'state': s}} for s in states if s.startswith(_key(value))]
        elif field in ('line1', 'line2'):
            out += _street_options(db, country, field, value, ctx, limit)
    return out[:limit * 2]


_LINE = re.compile(r'^(?P<pre>.*?)(?P<num>\d+[A-Za-z]?(?:\s*-\s*\d+[A-Za-z]?)?)?\s+(?P<street>[A-Za-z][^,]*)$')


def _street_options(db, country, field, value, ctx, limit):
    """'68 Rick' -> '68 Ricketts Road' in the typed suburb; with no suburb yet, a complete address per street found."""
    m = _LINE.match(' ' + value)
    if not m or len(_key(m.group('street'))) < 3:
        return []
    pre, num, typed = (m.group('pre') or '').strip(), (m.group('num') or '').replace(' ', ''), m.group('street')
    words = _street_key(db, typed)
    # The last word may still be being typed: match it as a prefix ('RICKETTS RO' -> 'RICKETTS ROAD')
    lead = " ".join(words)
    first_num = int(re.match(r'\d+', num).group()) if num else None
    head = (f"{pre} " if pre else '') + (f"{num} " if num else '')
    locs = _context_localities(db, country, ctx)
    if _key(ctx.get('suburb')) and not locs:
        return []  # the suburb typed isn't in this country: no streets from here
    found = []
    if locs:
        for l in locs:
            for sid, name, key in db.execute("SELECT id, name, key FROM street WHERE locality = ?", (l['id'],)):
                if " ".join(_street_key(db, key)).startswith(lead):
                    found.append((sid, name, l))
    elif len(lead) >= 4:
        # No suburb yet: streets anywhere starting like this (by the raw name, as stored), each with its suburb
        raw = _key(typed)
        lo, hi = _prefix_range(raw.split()[0])
        for sid, name, key, loc in db.execute("SELECT id, name, key, locality FROM street WHERE key >= ? AND key < ? LIMIT 400", (lo, hi)):
            if " ".join(_street_key(db, key)).startswith(lead):
                found.append((sid, name, loc))
        found = [(sid, name, l) for sid, name, loc in found for l in _localities(db, [loc])]
    if first_num is not None:
        has = lambda sid: bool(db.execute("SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1",
                                           (sid, first_num, first_num)).fetchone())
        found.sort(key=lambda t: not has(t[0]))
    out = []
    for sid, name, l in found[:limit]:
        line = f"{head}{_title(name)}"
        on_record = first_num is not None and bool(db.execute(
            "SELECT 1 FROM number WHERE street = ? AND first <= ? AND last >= ? LIMIT 1", (sid, first_num, first_num)).fetchone())
        if locs:
            out.append({'label': line, 'detail': ('✓ number on record' if on_record else l['name']), 'fields': {field: line}})
        else:
            o = _option(l, country=country)
            full = {field: line, **o['fields'], 'country': country}
            out.append({'label': f"{line}, {o['label']}", 'detail': ('✓ on record' if on_record else country), 'fields': full})
    return out
