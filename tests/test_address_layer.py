"""NZ Addresses' LINZ layer number: set by hand on the Address Data page, else looked up in LINZ's catalogue, else
the app's default; and "Check for updates" says when LINZ has moved it. No network: the catalogue is stubbed."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import address_reference as ar  # noqa: E402


@pytest.fixture()
def settings(tmp_path, monkeypatch):
    """A settings file of the test's own, never the real data/address_reference.json."""
    monkeypatch.setattr(ar, "DATA_DIR", str(tmp_path))
    monkeypatch.setattr(ar, "SETTINGS_FILE", str(tmp_path / "address_reference.json"))
    return tmp_path


def catalogue(monkeypatch, layer):
    """LINZ's catalogue answering with NZ Addresses as `layer` (None: catalogue not reachable)."""
    def fake(url, tries=6):
        if layer is None:
            raise RuntimeError("Couldn't reach data.linz.govt.nz")
        return [{"id": 999, "type": "layer", "title": "NZ Addresses (old pilot)"},
                {"id": layer, "type": "layer", "title": "NZ Addresses"}]
    monkeypatch.setattr(ar, "_get_json", fake)


def test_layer_from_catalogue_then_by_hand_then_default(settings, monkeypatch):
    catalogue(monkeypatch, 200001)
    assert ar._linz_layer() == (200001, "found in LINZ catalogue")  # the exact title, not the pilot
    ar.save_settings({"linz_layer": "300002"})
    assert ar._linz_layer() == (300002, "set by hand")
    ar.save_settings({"linz_layer": ""})
    catalogue(monkeypatch, None)
    layer, how = ar._linz_layer()
    assert layer == ar.LINZ_LAYER and "not reachable" in how


def test_check_updates_says_when_linz_moved_the_layer(settings, monkeypatch):
    monkeypatch.setattr(ar, "gnaf_latest", lambda: ("G-NAF NOV26", "", 0))
    monkeypatch.setattr(ar, "loaded_info", lambda c: {"version": "v", "loaded_at": "2099-01-01T00:00:00"})
    ar.save_settings({"linz_layer_used": 123113, "linz_layer_how": "found in LINZ catalogue"})
    catalogue(monkeypatch, 123113)
    nz = ar.check_updates()["NZ"]
    assert not nz["layer"]["moved"] and not nz["newer"]
    catalogue(monkeypatch, 456789)
    nz = ar.check_updates()["NZ"]
    assert nz["layer"]["moved"] and nz["layer"]["catalogue"] == 456789 and nz["newer"]


def test_page_saves_the_layer_and_refuses_anything_but_digits(settings, monkeypatch):
    import app as app_module
    app_module.app.config["TESTING"] = True
    c = app_module.app.test_client()
    r = c.post("/api/address-data/layer", json={"linz_layer": "12a"})
    assert r.status_code == 400 and "digits" in r.get_json()["error"]
    assert c.post("/api/address-data/layer", json={"linz_layer": " 456789 "}).get_json() == {"layer": 456789}
    assert ar.load_settings()["linz_layer"] == "456789"
    page = c.get("/address-data").data.decode()
    assert 'id="ad-linz-layer" value="456789"' in page and "set by hand" in page
    assert c.post("/api/address-data/layer", json={"linz_layer": ""}).get_json() == {"layer": None}
    assert ar.load_settings()["linz_layer"] == ""
