# SPDX-License-Identifier: Apache-2.0
import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))  # tests/onenote_fakes.py


def _load():
    name = "hermes_office_desktop"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, ROOT / "__init__.py",
                                                  submodule_search_locations=[str(ROOT)])
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def plugin():
    return _load()


@pytest.fixture(scope="session")
def bridge_mod(plugin):
    return sys.modules["hermes_office_desktop.bridge_server"]


@pytest.fixture(scope="session")
def tls_data(tmp_path_factory):
    """One local CA for the whole run (never a real profile's plugin data; generating one takes a moment)."""
    return tmp_path_factory.mktemp("plugin-data")


@pytest.fixture(autouse=True)
def bridge_tls_home(tls_data, monkeypatch):
    """The bridge's TLS identity (per-install CA) goes to the session's tmp folder, never the real Hermes home."""
    import https_client
    https_client.DATA = tls_data
    mod = _load()
    if hasattr(mod, "_tls_dir"):
        monkeypatch.setattr(mod, "_tls_dir", lambda ctx=None: tls_data)
    return tls_data


@pytest.fixture(scope="session")
def tls(plugin, tls_data):
    import https_client
    https_client.DATA = tls_data
    t = https_client.local_tls()
    t.ensure()
    return t


@pytest.fixture(autouse=True)
def manifest_home(tmp_path, monkeypatch):
    """office_connect writes the add-in manifests to ~/Documents/Omnia Works/add-in manifests by default (F2);
    tests never touch the real home: the default folder points into this test's tmp_path."""
    folder = tmp_path / "Documents" / "Omnia Works" / "add-in manifests"
    mod = _load()
    if hasattr(mod, "_default_manifest_dir"):
        monkeypatch.setattr(mod, "_default_manifest_dir", lambda: folder)
    return folder
