"""Courier services: the service codes offered in the consignment preview and written in the courier CSV, grouped so
one logo serves every service of a carrier. Stitch Labels puts the logo of each consignment's service on its labels.

Kept in data/courier_services.json; a logo dropped for a group is data/courier_logos/logo-<content hash>.png. A logo
belongs to a group (a service on its own is a group of one), and a service always shows the logo of the group it's in:
move it to another group and it takes that group's logo. Files are named by their content, so groups with the same
logo share one file, and a file no group uses any more (replaced, removed, its group deleted or emptied) is deleted
after every change. Until a logo is dropped,
the group uses the logo that comes with the app (engine/courier/default_logos/<group id>.png, made from the logo sheet
by engine/courier/logo_sheet.py), unless its logo was removed on purpose ("logo_removed"):

  {"groups": [{"id": "dhl", "name": "DHL", "logo": "dhl.png",
               "services": [{"code": "DHLWPXOD", "name": "DHL DDP EXPRESS PARCELS>30kG"}, ...]}, ...]}

A service on its own is a group of one, named after the service. Codes are unique across every group: the code is
what the courier portal reads in the CSV, so two services with one code would book the wrong one.
"""
import hashlib
import io
import json
import os
import re
import threading

from PIL import Image

# The list the app starts with (and falls back to when the file is missing or damaged)
DEFAULT_GROUPS = [
    ('DHL', [('DHLWPXOD', 'DHL DDP EXPRESS PARCELS>30kG'), ('DHLWPXUD', 'DHL DDP EXPRESS PARCELS<30KG')]),
    ('FEDEX International', [('FXIEOIE', 'FEDEX INTERNATIONAL ECONOMY ECONOMY'), ('FXPOIP', 'FEDEX INTERNATIONAL PRIORITY PRIORITY')]),
    ('Flash Couriers', [('FLASHFW', 'FLASH COURIERS EXP WAGON'), ('FLASHVIP', 'FLASH COURIERS VIP CAR'),
                        ('FLASHFV', 'FLASH COURIERS EXP VAN'), ('FLASHVW', 'FLASH COURIERS VIP WAGON'),
                        ('FLASHVV', 'FLASH COURIERS VIP VAN'), ('FLASHS', 'FLASH COURIERS STD CAR'),
                        ('FLASHSF1', 'FLASH COURIERS STD FLATTOP'), ('FLASHSF2', 'FLASH COURIERS STD FLATTOP 2SP'),
                        ('FLASHVF1', 'FLASH COURIERS VIP FLATTOP'), ('FLASHVF2', 'FLASH COURIERS VIP FLATT2SP'),
                        ('FLASHSV', 'FLASH COURIERS STD VAN'), ('FLASHFX', 'FLASH COURIERS EXP CAR'),
                        ('FLASHSW', 'FLASH COURIERS STD WAGON')]),
    ('Startrack Premium/Overnight', [('STEPRM', 'STARTRACK PREMIUM'), ('STEFPP', 'STARTRACK PREMIUM (SATCHEL)')]),
    ('Startrack Road Express', [('STEROAD', 'STARTRACK ROAD EXPRESS'), ('STERE2', 'STARTRACK ROAD EXPRESS 2 MAN SERVICE'),
                                ('STERET', 'STARTRACK ROAD EXPRESS TAILGATE')]),
    ('TNT OVERNIGHT', [('TNTNONFC', 'TNT (NSW) OVERNIGHT EXPRESS'), ('TNTN9AM', 'TNT (NSW) OVERNIGHT 9AM')]),
    ('TNT International', [('TNTIXP', 'TNT INTERNATIONAL EXPRESS PARCELS')]),
    ('Toll Priority', [('20', 'TOLL PRIORITY 1KG SATCHEL'), ('21', 'TOLL PRIORITY 3KG SATCHEL'),
                       ('22', 'TOLL PRIORITY 5KG SATCHEL'), ('2', 'TOLL PRIORITY OVERNIGHT')]),
    # On their own: each its own group, named after the service
    ('TOLL IPEC ROAD EXPRESS', [('IPECX', 'TOLL IPEC ROAD EXPRESS')]),
    ('TNT (NSW) ROAD EXPRESS', [('TNTN', 'TNT (NSW) ROAD EXPRESS')]),
    ('BORDER EXPRESS PARCEL', [('BORDERP', 'BORDER EXPRESS PARCEL')]),
    ('BORDER EXPRESS BULK', [('BORDERB', 'BORDER EXPRESS BULK')]),
    ('DETRACK LOCAL FREIGHT', [('DTRLOF', 'DETRACK LOCAL FREIGHT')]),
    ('STARTRACK NATIONAL PREMIUM', [('STNPRM', 'STARTRACK NATIONAL PREMIUM')]),
    ('STARTRACK NATIONAL PREMIUM(SATCHEL)', [('STNFPP', 'STARTRACK NATIONAL PREMIUM(SATCHEL)')]),
    ('STARTRACK NATIONAL ROAD EXPRESS', [('STNROAD', 'STARTRACK NATIONAL ROAD EXPRESS')]),
    ('TNT (NSW) SPECIALISED (TAILGATE)', [('TNTNSWSPEC', 'TNT (NSW) SPECIALISED (TAILGATE)')]),
]

OWN_GROUP = '__own__'  # update_service: take the service out of its group, on its own
CODE_RE = re.compile(r'^[A-Z0-9][A-Z0-9_-]{0,19}$')
LOGO_MAX_BYTES = 5 * 1024 * 1024
DEFAULT_LOGO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'default_logos')
LOGO_MAX_SIDE = 2400  # px: a 1000 px wide strip is plenty for the 43 mm slot; bigger ones are scaled down once


class ServiceError(ValueError):
    """A change that can't be saved (a code already used, an empty name, an image that isn't one)."""


def clean_code(code):
    return re.sub(r'\s+', '', str(code or '')).upper()


def _slug(name, taken):
    base = re.sub(r'[^a-z0-9]+', '-', str(name).lower()).strip('-')[:40] or 'group'
    slug, n = base, 2
    while slug in taken:
        slug, n = f"{base}-{n}", n + 1
    return slug


def _defaults():
    groups, taken = [], set()
    for name, services in DEFAULT_GROUPS:
        gid = _slug(name, taken)
        taken.add(gid)
        groups.append({'id': gid, 'name': name, 'logo': '', 'services': [{'code': c, 'name': n} for c, n in services]})
    return {'groups': groups}


class ServiceStore:
    def __init__(self, path, logo_dir, default_dir=DEFAULT_LOGO_DIR):
        self.path = path
        self.logo_dir = logo_dir
        self.default_dir = default_dir
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ reading
    def load(self):
        """{'groups': [...]}: the saved list, or the starting list when nothing usable is saved."""
        try:
            with open(self.path, encoding='utf-8') as f:
                data = json.load(f)
            groups = [g for g in data.get('groups') or [] if isinstance(g, dict) and g.get('id')]
            for g in groups:
                g['services'] = [{'code': clean_code(s.get('code')), 'name': str(s.get('name') or '').strip()}
                                 for s in g.get('services') or [] if clean_code(s.get('code'))]
                g['name'] = str(g.get('name') or g['id'])
                g['logo'] = g.get('logo') or ''
                g['logo_removed'] = bool(g.get('logo_removed'))
            return {'groups': groups}
        except (OSError, ValueError, AttributeError, TypeError):
            return _defaults()

    def services(self):
        """Every service in list order: [{'code', 'name', 'group', 'group_name'}]."""
        return [{'code': s['code'], 'name': s['name'], 'group': g['id'], 'group_name': g['name']}
                for g in self.load()['groups'] for s in g['services']]

    def codes(self):
        return {s['code'] for s in self.services()}

    def group_of(self, code):
        """The group a service code belongs to, or None."""
        code = clean_code(code)
        return next((g for g in self.load()['groups'] if any(s['code'] == code for s in g['services'])), None)

    def logo_path(self, group):
        """The group's logo file: the one dropped for it, else the one that comes with the app, else None."""
        if not group:
            return None
        if group.get('logo'):
            path = os.path.join(self.logo_dir, os.path.basename(group['logo']))
            if os.path.exists(path):
                return path
        if not group.get('logo_removed') and self.default_dir:
            path = os.path.join(self.default_dir, f"{os.path.basename(group['id'])}.png")
            if os.path.exists(path):
                return path
        return None

    def logo_for(self, code):
        """The logo file for a service code (its group's), or None."""
        return self.logo_path(self.group_of(code))

    # ------------------------------------------------------------------ writing
    def _write(self, data):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        tmp = f"{self.path}.tmp"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def _change(self, fn):
        with self._lock:
            data = self.load()
            result = fn(data)
            self._write(data)
            self._prune_logos(data)
            return result

    def _prune_logos(self, data):
        """Deletes the logo files no group uses any more (a replaced or removed logo, a deleted or emptied group).
        The logos that come with the app are never touched."""
        used = {os.path.basename(g['logo']) for g in data['groups'] if g.get('logo')}
        try:
            names = os.listdir(self.logo_dir)
        except OSError:
            return
        for name in names:
            if name.lower().endswith('.png') and name not in used:
                try:
                    os.remove(os.path.join(self.logo_dir, name))
                except OSError:
                    pass

    @staticmethod
    def _check(data, code, name, except_code=None):
        code, name = clean_code(code), re.sub(r'\s+', ' ', str(name or '')).strip()
        if not CODE_RE.match(code):
            raise ServiceError("A service code is 1 to 20 letters or digits (the code the courier portal reads), e.g. STEPRM or 20.")
        if not name:
            raise ServiceError("Give the service a name, e.g. STARTRACK PREMIUM.")
        for g in data['groups']:
            for s in g['services']:
                if s['code'] == code and code != except_code:
                    raise ServiceError(f"{code} is already in the list ({s['name']}, group {g['name']}).")
        return code, name

    def _group(self, data, group_id=None, new_group=''):
        """The group a service goes in: an existing one, a new one by name, or (neither) a new group of its own."""
        if group_id:
            g = next((g for g in data['groups'] if g['id'] == group_id), None)
            if g is None:
                raise ServiceError("That group isn't in the list any more: reload the page.")
            return g
        name = re.sub(r'\s+', ' ', str(new_group or '')).strip()
        if name:
            here = next((g for g in data['groups'] if g['name'].lower() == name.lower()), None)
            if here:
                return here
        return None

    def add_service(self, code, name, group_id=None, new_group='', move=False):
        """Adds a service to a group (or a new group of its own). move: a code that's already in another group is
        moved into this one instead (keeping its name unless a new one is given), like editing it to that group."""
        if move and group_id:
            code_c = clean_code(code)
            here = self.group_of(code_c)
            if here is not None:
                if here['id'] == group_id:
                    raise ServiceError(f"{code_c} is already in this group.")
                old = next(s for s in here['services'] if s['code'] == code_c)
                moved = self.update_service(code_c, code_c, str(name or '').strip() or old['name'], group_id)
                return {**moved, 'moved_from': here['name']}

        def go(data):
            c, n = self._check(data, code, name)
            g = self._group(data, group_id, new_group)
            if g is None:
                g = {'id': _slug(new_group or n, {x['id'] for x in data['groups']}), 'name': (new_group or n).strip(),
                     'logo': '', 'services': []}
                data['groups'].append(g)
            g['services'].append({'code': c, 'name': n})
            return {'code': c, 'name': n, 'group': g['id'], 'group_name': g['name']}
        return self._change(go)

    def update_service(self, old_code, code, name, group_id=None, new_group=''):
        old = clean_code(old_code)

        def go(data):
            c, n = self._check(data, code, name, except_code=old)
            src = next((g for g in data['groups'] if any(s['code'] == old for s in g['services'])), None)
            if src is None:
                raise ServiceError(f"{old} isn't in the list any more: reload the page.")
            new_name = new_group
            if group_id == OWN_GROUP:  # out of its group, on its own: a group of one named after it, no logo yet
                dest = src if len(src['services']) == 1 else None
                new_name = n
            else:
                dest = self._group(data, group_id, new_group) if (group_id or new_group) else src
            if dest is None:  # a new group's name that isn't there yet
                dest = {'id': _slug(new_name, {x['id'] for x in data['groups']}), 'name': new_name.strip(), 'logo': '', 'services': []}
                data['groups'].append(dest)
            if dest is src:  # same group: changed where it is, so the list keeps its order
                src['services'] = [{'code': c, 'name': n} if s['code'] == old else s for s in src['services']]
            else:
                src['services'] = [s for s in src['services'] if s['code'] != old]
                dest['services'].append({'code': c, 'name': n})
            self._drop_empty(data, keep=dest['id'])
            return {'code': c, 'name': n, 'group': dest['id'], 'group_name': dest['name']}
        return self._change(go)

    def delete_service(self, code):
        code = clean_code(code)

        def go(data):
            found = False
            for g in data['groups']:
                if any(s['code'] == code for s in g['services']):
                    g['services'] = [s for s in g['services'] if s['code'] != code]
                    found = True
            if not found:
                raise ServiceError(f"{code} isn't in the list any more.")
            self._drop_empty(data)
        return self._change(go)

    def _drop_empty(self, data, keep=None):
        """A group left with no services goes, and its logo with it."""
        for g in [g for g in data['groups'] if not g['services'] and g['id'] != keep]:
            self._remove_logo_file(g)
        data['groups'] = [g for g in data['groups'] if g['services'] or g['id'] == keep]

    def add_group(self, name):
        name = re.sub(r'\s+', ' ', str(name or '')).strip()
        if not name:
            raise ServiceError("Give the group a name, e.g. Toll Priority.")

        def go(data):
            if any(g['name'].lower() == name.lower() for g in data['groups']):
                raise ServiceError(f"There's already a group called {name}.")
            g = {'id': _slug(name, {x['id'] for x in data['groups']}), 'name': name, 'logo': '', 'services': []}
            data['groups'].append(g)
            return g
        return self._change(go)

    def rename_group(self, group_id, name):
        name = re.sub(r'\s+', ' ', str(name or '')).strip()
        if not name:
            raise ServiceError("Give the group a name.")

        def go(data):
            g = self._group(data, group_id)
            if any(x['name'].lower() == name.lower() and x['id'] != group_id for x in data['groups']):
                raise ServiceError(f"There's already a group called {name}.")
            g['name'] = name
            return g
        return self._change(go)

    def delete_group(self, group_id):
        """Deletes a group and its logo; its services stay in the list, each on its own."""
        def go(data):
            g = self._group(data, group_id)
            self._remove_logo_file(g)
            data['groups'] = [x for x in data['groups'] if x['id'] != group_id]
            taken = {x['id'] for x in data['groups']}
            for s in g['services']:
                gid = _slug(s['name'], taken)
                taken.add(gid)
                data['groups'].append({'id': gid, 'name': s['name'], 'logo': '', 'services': [s]})
        return self._change(go)

    # ------------------------------------------------------------------ logos
    def set_logo(self, group_id, raw):
        """Saves an image as the group's logo (always written as PNG, transparency kept)."""
        if not raw or len(raw) > LOGO_MAX_BYTES:
            raise ServiceError("Drop a PNG or JPG logo under 5 MB.")
        try:
            img = Image.open(io.BytesIO(raw))
            img.load()
        except Exception:
            raise ServiceError("That file isn't an image: drop a PNG (best, with a transparent background) or a JPG.")
        if img.format not in ('PNG', 'JPEG', 'WEBP'):
            raise ServiceError("Use a PNG or JPG logo.")
        img = img.convert('RGBA')
        if max(img.size) > LOGO_MAX_SIDE:
            img.thumbnail((LOGO_MAX_SIDE, LOGO_MAX_SIDE))

        buf = io.BytesIO()
        img.save(buf, 'PNG', optimize=True)
        png = buf.getvalue()

        def go(data):
            g = self._group(data, group_id)
            g['logo'] = self._store_png(png)  # the group's earlier logo goes once nothing uses it (_prune_logos)
            g.pop('logo_removed', None)
            return {'width': img.width, 'height': img.height}
        return self._change(go)

    def _store_png(self, png):
        """Writes a logo under the name of its content and returns that name (an identical logo is one file)."""
        name = f"logo-{hashlib.sha1(png).hexdigest()[:16]}.png"
        path = os.path.join(self.logo_dir, name)
        if not os.path.exists(path):
            os.makedirs(self.logo_dir, exist_ok=True)
            with open(path + '.tmp', 'wb') as f:
                f.write(png)
            os.replace(path + '.tmp', path)
        return name

    def _remove_logo_file(self, g):
        """The group no longer has its own logo; the file goes once no other group uses it (_prune_logos)."""
        g['logo'] = ''

    def delete_logo(self, group_id):
        """No logo for the group any more, not even the one that comes with the app (drop one to have it again)."""
        def go(data):
            g = self._group(data, group_id)
            self._remove_logo_file(g)
            g['logo_removed'] = True
        return self._change(go)

    # ------------------------------------------------------------------ export / import (Back up & move)
    def summary(self):
        groups = self.load()['groups']
        n = sum(len(g['services']) for g in groups)
        logos = sum(1 for g in groups if self.logo_path(g))
        return f"{n} services in {len(groups)} groups, {logos} logo{'s' if logos != 1 else ''}"

    def export_files(self):
        """{name in the zip: bytes}: the list, and only the logos a group (or a service on its own) uses, the app's
        own ones too, each file once however many groups share it, named by its content."""
        data = self.load()
        files = {}
        for g in data['groups']:
            path = self.logo_path(g)
            g['logo'] = ''
            if path:
                with open(path, 'rb') as f:
                    raw = f.read()
                g['logo'] = f"logo-{hashlib.sha1(raw).hexdigest()[:16]}.png"
                files[f"courier_logos/{g['logo']}"] = raw
            g.pop('logo_removed', None)
        files['courier_services.json'] = json.dumps(data, indent=2, ensure_ascii=False).encode('utf-8')
        return files

    def import_files(self, files, mode):
        """Takes an exported list and its logos. merge: adds the services and groups missing here, and logos for
        groups that have none; replace: the zip's list and logos only."""
        incoming = json.loads(files['courier_services.json'].decode('utf-8'))
        logos = {os.path.basename(k): v for k, v in files.items() if k.startswith('courier_logos/') and k.endswith('.png')}
        groups = [g for g in incoming.get('groups') or [] if isinstance(g, dict) and g.get('id')]
        if mode == 'replace':  # the logos here go as the list empties (_prune_logos); the zip's come in after
            with self._lock:
                self._write({'groups': []})
            added = 0
            for g in groups:
                for s in g.get('services') or []:
                    try:
                        self.add_service(s.get('code'), s.get('name'), new_group=g.get('name') or s.get('name'))
                        added += 1
                    except ServiceError:
                        pass
            self._import_logos(groups, logos, only_missing=False)
            return f"replaced: {self.summary()}"
        have = self.codes()
        added = 0
        for g in groups:
            for s in g.get('services') or []:
                if clean_code(s.get('code')) in have:
                    continue
                try:
                    self.add_service(s.get('code'), s.get('name'), new_group=g.get('name') or s.get('name'))
                    added += 1
                except ServiceError:
                    pass
        logos_added = self._import_logos(groups, logos, only_missing=True)
        return (f"{added} service{'s' if added != 1 else ''} added, {logos_added} logo{'s' if logos_added != 1 else ''} "
                f"added (the ones here kept as they are)")

    def _import_logos(self, groups, logos, only_missing):
        here = {g['name'].lower(): g for g in self.load()['groups']}
        n = 0
        for g in groups:
            raw = logos.get(os.path.basename(g.get('logo') or ''))
            target = here.get(str(g.get('name') or '').lower())
            if not raw or not target or (only_missing and self.logo_path(target)):
                continue
            try:
                self.set_logo(target['id'], raw)
                n += 1
            except ServiceError:
                pass
        return n
