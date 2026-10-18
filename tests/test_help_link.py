"""The nav's Help link: WAREHOUSE_HELP_URL from the .env file (engine/core/env_file.py), web addresses only."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.core import env_file  # noqa: E402


def test_env_file_lines_are_read_and_the_environment_wins(tmp_path, monkeypatch):
    path = tmp_path / ".env"
    path.write_text('# settings\n\nWAREHOUSE_HELP_URL="https://help.example.com/warehouse"\nexport A_SET=from file\n'
                    "QUOTED='x y'\nnot a setting\n", encoding='utf-8')
    monkeypatch.delenv('WAREHOUSE_HELP_URL', raising=False)
    monkeypatch.delenv('QUOTED', raising=False)
    monkeypatch.setenv('A_SET', 'already set')
    assert env_file.load(str(path)) == ['WAREHOUSE_HELP_URL', 'A_SET', 'QUOTED']
    assert os.environ['WAREHOUSE_HELP_URL'] == 'https://help.example.com/warehouse'
    assert os.environ['A_SET'] == 'already set' and os.environ['QUOTED'] == 'x y'
    assert env_file.load(str(tmp_path / "missing.env")) == []


def test_help_link_in_the_nav_only_for_a_web_address(monkeypatch):
    import app as app_module
    assert app_module._help_url('https://help.example.com/a?b=1') == 'https://help.example.com/a?b=1'
    for bad in ('javascript:alert(1)', 'file:///C:/x', 'help.example.com', '', 'https://x" onclick="y'):
        assert app_module._help_url(bad) == '', bad
    client = app_module.app.test_client()
    monkeypatch.setattr(app_module, 'HELP_URL', 'https://help.example.com/warehouse')
    page = client.get('/packing-specs').data.decode()
    assert 'href="https://help.example.com/warehouse" class="nav-help" target="_blank" rel="noopener noreferrer"' in page
    assert 'aria-label="Help (opens in a new tab)"' in page and 'class="nav-help-icon"' in page
    monkeypatch.setattr(app_module, 'HELP_URL', '')
    assert 'nav-help' not in client.get('/packing-specs').data.decode()
