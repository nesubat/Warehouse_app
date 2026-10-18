"""Settings kept in the .env file in the app folder (next to the .exe in a built copy), read once when the app starts.

  WAREHOUSE_HELP_URL=https://...     the ❓ Help link at the right of the nav bar (opens in a new tab)

One KEY=value per line; blank lines and lines starting with # are skipped, quotes around a value are dropped and
'export ' in front is allowed. A variable already set in the environment wins over the file. The file stays out of
git (.gitignore), so each computer keeps its own.
"""
import os


def load(path):
    """Puts the file's settings into os.environ (not over ones already set). Returns the keys it read."""
    try:
        with open(path, encoding='utf-8-sig') as f:
            lines = f.read().splitlines()
    except OSError:  # no .env: nothing to read
        return []
    keys = []
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, _, value = line.removeprefix('export ').partition('=')
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in '"\'':
            value = value[1:-1]
        if key:
            os.environ.setdefault(key, value)
            keys.append(key)
    return keys
