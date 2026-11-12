# SPDX-License-Identifier: Apache-2.0
"""office_insert_image sources: path allowlist, SSRF-guarded fetch, HTML sanitizer, the tool handler and receipts.

No real network: every fetch goes to a fixture server on 127.0.0.1 through the images.CONNECT seam, and host
names resolve through a stubbed images.RESOLVE. The address policy is exercised on the stubbed answers.
"""
import base64
import http.client
import json
import os
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pytest

from test_plugin import FakeCtx, ORIGIN  # noqa: E402  (tests/ is on sys.path via conftest)

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==")
JPEG = b"\xff\xd8\xff\xe0" + b"\x00\x10JFIF" + b"\x00" * 64 + b"\xff\xd9"
GIF = b"GIF89a" + b"\x00" * 32


@pytest.fixture(scope="module")
def images(plugin):
    return sys.modules["hermes_office_desktop.images"]


# ---- fixture server ---------------------------------------------------------------------------------
ROUTES = {}


class _H(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        return

    def do_GET(self):
        route = ROUTES.get(self.path)
        if route is None:
            self.send_response(404)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        route(self)


def _send(h, status=200, ctype="image/png", body=b"", headers=None, length=True):
    h.send_response(status)
    if ctype:
        h.send_header("Content-Type", ctype)
    for k, v in (headers or {}).items():
        h.send_header(k, v)
    if length:
        h.send_header("Content-Length", str(len(body)))
    h.send_header("Connection", "close")
    h.end_headers()
    h.wfile.write(body)


@pytest.fixture(scope="module")
def server():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), _H)
    srv.daemon_threads = True
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    yield srv.server_address[1]
    srv.shutdown()


DNS = {"img.example": ["93.184.216.34"], "cdn.example": ["2606:4700::6810:84e5"], "hop.example": ["93.184.216.35"],
       "internal.example": ["10.0.0.5"], "mixed.example": ["93.184.216.36", "192.168.1.9"],
       "meta.example": ["169.254.169.254"], "cgnat.example": ["100.100.100.100"], "mcast.example": ["224.0.0.1"],
       "mapped.example": ["::ffff:127.0.0.1"], "v6local.example": ["fe80::1"], "loop6.example": ["::1"],
       "sixtofour.example": ["2002:0a00:0001::1"], "teredo.example": ["2001:0:4136:e378:8000:63bf:3fff:fdd2"],
       "zero.example": ["0.0.0.0"]}


@pytest.fixture()
def net(images, server, monkeypatch):
    """Stub DNS + route every pinned connection to the 127.0.0.1 fixture server (plain HTTP)."""
    seen = {"connect": [], "resolve": []}

    def resolve(host, port):
        seen["resolve"].append(host)
        if host not in DNS:
            raise OSError("NXDOMAIN")
        return DNS[host]

    def connect(host, address, port, timeout):
        seen["connect"].append((host, address, port))
        return http.client.HTTPConnection("127.0.0.1", server, timeout=timeout)
    monkeypatch.setattr(images, "RESOLVE", resolve)
    monkeypatch.setattr(images, "CONNECT", connect)
    ROUTES.clear()
    return seen


# ---- address policy -------------------------------------------------------------------------------
@pytest.mark.parametrize("addr", ["127.0.0.1", "10.1.2.3", "172.16.0.1", "192.168.0.1", "169.254.169.254", "100.64.0.1",
                                  "100.127.255.254", "224.0.0.251", "239.255.255.250", "0.0.0.0", "255.255.255.255",
                                  "240.0.0.1", "::1", "::", "fe80::1", "fc00::1", "ff02::1", "::ffff:10.0.0.1",
                                  "::ffff:127.0.0.1", "2002:7f00:0001::1", "2001:0:4136:e378:8000:63bf:3fff:fdd2",
                                  "198.18.0.1", "192.0.2.1", "not-an-ip"])
def test_non_public_addresses_are_refused(images, addr):
    assert images.address_allowed(addr) is False


@pytest.mark.parametrize("addr", ["93.184.216.34", "8.8.8.8", "2606:4700::6810:84e5", "::ffff:93.184.216.34"])
def test_public_addresses_are_allowed(images, addr):
    assert images.address_allowed(addr) is True


@pytest.mark.parametrize("url,cat", [("http://img.example/a.png", "Blocked"), ("ftp://img.example/a.png", "Blocked"),
                                     ("file:///etc/hosts", "Blocked"), ("https://user:pw@img.example/a.png", "Blocked"),
                                     ("https://img.example:8443/a.png", "Blocked"), ("https:///a.png", "Validation"),
                                     ("https://img.example/a b.png", "Validation")])
def test_url_shape_is_checked_before_any_lookup(images, net, url, cat):
    with pytest.raises(images.ImageError) as e:
        images.fetch_url(url)
    assert e.value.category == cat
    assert net["resolve"] == [] and net["connect"] == []


# ---- fetch ------------------------------------------------------------------------------------------
def test_fetch_png_pins_the_checked_address(images, net):
    ROUTES["/a.png"] = lambda h: _send(h, body=PNG)
    img = images.fetch_url("https://img.example/a.png")
    assert img.kind == "png" and img.data == PNG and img.source == {"type": "url", "url": "https://img.example/a.png",
                                                                   "redirects": 0}
    assert net["connect"] == [("img.example", "93.184.216.34", 443)]


def test_fetch_jpeg_over_ipv6_public(images, net):
    ROUTES["/b.jpg"] = lambda h: _send(h, ctype="image/jpeg; charset=binary", body=JPEG)
    assert images.fetch_url("https://cdn.example/b.jpg").kind == "jpeg"


@pytest.mark.parametrize("host", ["internal.example", "mixed.example", "meta.example", "cgnat.example", "mcast.example",
                                  "mapped.example", "v6local.example", "loop6.example", "sixtofour.example",
                                  "teredo.example", "zero.example"])
def test_private_resolution_is_refused_without_connecting(images, net, host):
    with pytest.raises(images.ImageError) as e:
        images.fetch_url(f"https://{host}/a.png")
    assert e.value.category == "Blocked" and net["connect"] == []


@pytest.mark.parametrize("literal", ["https://127.0.0.1/a.png", "https://[::1]/a.png", "https://10.0.0.2/a.png",
                                     "https://169.254.169.254/latest/meta-data", "https://[::ffff:192.168.0.1]/x.png"])
def test_private_ip_literals_are_refused(images, net, literal):
    with pytest.raises(images.ImageError) as e:
        images.fetch_url(literal)
    assert e.value.category == "Blocked" and net["connect"] == [] and net["resolve"] == []


def test_redirects_are_rechecked_each_hop(images, net):
    ROUTES["/r1"] = lambda h: _send(h, 302, None, headers={"Location": "https://hop.example/r2"})
    ROUTES["/r2"] = lambda h: _send(h, 301, None, headers={"Location": "/final.png"})
    ROUTES["/final.png"] = lambda h: _send(h, body=PNG)
    img = images.fetch_url("https://img.example/r1")
    assert img.source["redirects"] == 2 and img.source["finalUrl"] == "https://hop.example/final.png"
    assert [c[1] for c in net["connect"]] == ["93.184.216.34", "93.184.216.35", "93.184.216.35"]


@pytest.mark.parametrize("location,cat", [("https://internal.example/a.png", "Blocked"), ("http://img.example/a.png", "Blocked"),
                                          ("https://127.0.0.1/a.png", "Blocked"), ("https://meta.example/", "Blocked")])
def test_redirect_to_non_public_or_plain_http_is_refused(images, net, location, cat):
    ROUTES["/r"] = lambda h: _send(h, 307, None, headers={"Location": location})
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/r")
    assert e.value.category == cat
    assert len(net["connect"]) == 1  # the redirect target was never contacted


def test_more_than_three_redirects_is_refused(images, net):
    for i in range(5):
        ROUTES[f"/l{i}"] = (lambda n: lambda h: _send(h, 302, None, headers={"Location": f"/l{n + 1}"}))(i)
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/l0")
    assert e.value.category == "Blocked" and "3 redirects" in str(e.value) and len(net["connect"]) == 4


@pytest.mark.parametrize("ctype,body,cat", [("text/html", b"<html>", "WrongType"), ("image/gif", GIF, "WrongType"),
                                            ("image/svg+xml", b"<svg/>", "WrongType"), (None, PNG, "WrongType"),
                                            ("image/png", JPEG, "WrongType"), ("image/jpeg", PNG, "WrongType"),
                                            ("image/png", b"", "WrongType")])
def test_content_type_and_magic_bytes_must_agree(images, net, ctype, body, cat):
    ROUTES["/x"] = lambda h: _send(h, ctype=ctype, body=body)
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/x")
    assert e.value.category == cat


def test_compressed_responses_are_refused(images, net):
    ROUTES["/z"] = lambda h: _send(h, body=PNG, headers={"Content-Encoding": "gzip"})
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/z")
    assert e.value.category == "WrongType"


def test_oversize_by_content_length_is_refused_before_reading(images, net):
    def big(h):
        h.send_response(200)
        h.send_header("Content-Type", "image/png")
        h.send_header("Content-Length", str(images.MAX_IMAGE_BYTES + 1))
        h.send_header("Connection", "close")
        h.end_headers()
        h.wfile.write(PNG)  # far fewer bytes than claimed: the client must not wait for them
    ROUTES["/big"] = big
    t0 = time.monotonic()
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/big")
    assert e.value.category == "Capacity" and time.monotonic() - t0 < 5


def test_oversize_stream_without_length_hits_the_hard_cap(images, net):
    def stream(h):
        h.send_response(200)
        h.send_header("Content-Type", "image/png")
        h.send_header("Connection", "close")
        h.end_headers()
        h.wfile.write(PNG)
        chunk = b"\0" * 65536
        try:
            for _ in range(80):  # 5 MiB, more than the cap
                h.wfile.write(chunk)
        except OSError:
            pass
    ROUTES["/stream"] = stream
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/stream")
    assert e.value.category == "Capacity"


def test_slow_server_hits_the_deadline(images, net, monkeypatch):
    monkeypatch.setattr(images, "FETCH_TIMEOUT", 0.6)

    def slow(h):
        h.send_response(200)
        h.send_header("Content-Type", "image/png")
        h.send_header("Content-Length", "100000")
        h.end_headers()
        try:
            for _ in range(20):
                h.wfile.write(PNG[:4])
                h.wfile.flush()
                time.sleep(0.2)
        except OSError:
            pass
    ROUTES["/slow"] = slow
    t0 = time.monotonic()
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/slow")
    assert e.value.category == "Network" and time.monotonic() - t0 < 3


def test_unresolvable_host_and_http_errors(images, net):
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://nowhere.example/a.png")
    assert e.value.category == "Network"
    with pytest.raises(images.ImageError) as e:
        images.fetch_url("https://img.example/missing.png")
    assert e.value.category == "Network" and "404" in str(e.value)


# ---- local paths ------------------------------------------------------------------------------------
@pytest.fixture()
def home(tmp_path, monkeypatch):
    h = tmp_path / "hermes"
    for d in ("cache/generated/images", "cache/images", "cache/screenshots"):
        (h / d).mkdir(parents=True)
    monkeypatch.setenv("HERMES_HOME", str(h))
    return h


def test_generated_image_path_is_allowed(images, home):
    f = home / "cache" / "generated" / "images" / "image_1.png"
    f.write_bytes(PNG)
    img = images.load_path(str(f))
    assert img.kind == "png" and img.source == {"type": "path", "name": "image_1.png"}
    assert (home / "cache" / "generated" / "images").resolve() in images.allowed_dirs()


def test_paths_outside_the_allowed_dirs_are_refused(images, home, tmp_path):
    outside = tmp_path / "elsewhere.png"
    outside.write_bytes(PNG)
    for p in (str(outside), str(home / "config.yaml"), "/etc/hosts",
              str(home / "cache" / "images" / ".." / ".." / ".." / "elsewhere.png")):
        with pytest.raises(images.ImageError) as e:
            images.load_path(p)
        assert e.value.category in ("Blocked", "Validation"), p
    (home / "config.yaml").write_text("x")
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(home / "config.yaml"))
    assert e.value.category == "Blocked"


def test_symlink_out_of_an_allowed_dir_is_refused(images, home, tmp_path):
    secret = tmp_path / "secret.png"
    secret.write_bytes(PNG)
    link = home / "cache" / "images" / "innocent.png"
    link.symlink_to(secret)
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(link))
    assert e.value.category == "Blocked"


def test_symlinked_default_root_does_not_widen_the_allowlist(images, home, tmp_path):
    """A6 #3: a default root that resolves outside the Hermes cache layout (or to '/') is not an allowed root."""
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "secret.png").write_bytes(PNG)
    root = home / "cache" / "images"
    root.rmdir()
    root.symlink_to(outside)
    assert outside.resolve() not in images.allowed_dirs()
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(root / "secret.png"))
    assert e.value.category == "Blocked"
    shots = home / "cache" / "screenshots"
    shots.rmdir()
    shots.symlink_to("/")
    dirs = images.allowed_dirs()
    assert images.Path("/") not in dirs and images.Path.home().resolve() not in dirs
    for d in dirs:
        assert d == home.resolve() / d.relative_to(home.resolve()), d
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(outside / "secret.png"))
    assert e.value.category == "Blocked"


def test_dangerous_hermes_home_yields_no_default_roots(images, monkeypatch):
    for bad in ("/", os.path.expanduser("~")):
        monkeypatch.setenv("HERMES_HOME", bad)
        assert images.allowed_dirs(home=images.Path(bad)) == [], bad


def test_swap_to_symlink_after_the_check_is_refused(images, home, tmp_path, monkeypatch):
    """A6 #3: the file is replaced by a symlink between resolve() and open(); the read must not follow it."""
    secret = tmp_path / "secret.png"
    secret.write_bytes(JPEG)
    f = home / "cache" / "images" / "ok.png"
    f.write_bytes(PNG)
    real_resolve = images.Path.resolve
    swapped = []

    def resolve(self, strict=False):
        out = real_resolve(self, strict=strict)
        if out == real_resolve(f) and not swapped:
            swapped.append(1)
            f.unlink()
            f.symlink_to(secret)
        return out
    monkeypatch.setattr(images.Path, "resolve", resolve)
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(f))
    assert swapped and e.value.category == "Blocked"


def test_swap_to_a_fifo_or_directory_after_the_check_is_refused(images, home, monkeypatch):
    f = home / "cache" / "images" / "ok.png"
    f.write_bytes(PNG)
    real_resolve = images.Path.resolve

    def resolve(self, strict=False):
        out = real_resolve(self, strict=strict)
        if out == real_resolve(f) and f.is_file():
            f.unlink()
            os.mkfifo(f)
        return out
    monkeypatch.setattr(images.Path, "resolve", resolve)
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(f))
    assert e.value.category == "Validation"


def test_user_image_dirs_setting(images, home, tmp_path):
    mine = tmp_path / "Pictures"
    mine.mkdir()
    (mine / "a.jpg").write_bytes(JPEG)
    with pytest.raises(images.ImageError):
        images.load_path(str(mine / "a.jpg"))
    assert images.load_path(str(mine / "a.jpg"), [str(mine)]).kind == "jpeg"
    # relative, missing, '/' and the home folder itself are ignored
    assert images.configured_dirs(["Pictures", str(tmp_path / "nope"), "/", str(os.path.expanduser("~")), 7]) == []


# ---- A27: a bare picture name ("team.png") is looked up in the allowed folders only ---------------------------
def test_bare_name_is_found_in_a_subfolder_of_the_images_folder(images, home):
    f = home / "cache" / "images" / "demo" / "team.png"
    f.parent.mkdir()
    f.write_bytes(PNG)
    for name in ("team.png", "TEAM.PNG", "Team.Png"):
        img = images.load_path(name)
        assert img.kind == "png" and img.data == PNG
        assert img.source["type"] == "path" and img.source["name"] == "team.png"
        assert img.source["foundByName"] == name and img.source["path"] == str(f.resolve())
        assert img.source["folder"] == "cache/images"
    for near in ("team", "team.png.bak", "eam.png", " team.png"):
        with pytest.raises(images.ImageError) as e:
            images.load_path(near)
        assert e.value.category == "Validation", near


def test_bare_name_found_in_an_image_dirs_folder(images, home, tmp_path):
    mine = tmp_path / "Pictures"
    mine.mkdir()
    (mine / "logo.jpg").write_bytes(JPEG)
    with pytest.raises(images.ImageError) as e:
        images.load_path("logo.jpg")
    assert "Searched:" in str(e.value) and str((home / "cache" / "images").resolve()) in str(e.value)
    img = images.load_path("logo.jpg", [str(mine)])
    assert img.kind == "jpeg" and img.source["folder"] == "image_dirs[0]"


def test_several_matches_are_refused_with_relative_candidates_only(images, home, tmp_path):
    a = home / "cache" / "images" / "a" / "team.png"
    b = home / "cache" / "generated" / "images" / "TEAM.png"
    for f in (a, b):
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_bytes(PNG)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    msg = str(e.value)
    assert e.value.category == "Validation" and msg.startswith("2 pictures are named team.png")
    assert "a/team.png (in cache/images)" in msg and "TEAM.png (in cache/generated/images)" in msg
    assert str(home) not in msg and str(tmp_path) not in msg  # no absolute path leaks in the candidate list


def test_none_found_names_the_searched_folders(images, home):
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    msg = str(e.value)
    assert e.value.category == "Validation" and "No picture named team.png" in msg
    for d in images.allowed_dirs():
        assert str(d) in msg


@pytest.mark.parametrize("name", ["../team.png", "../../outside/team.png", "..", ".", "~", "~/team.png",
                                  "cache/images/team.png", "..\\team.png", "C:team.png", "team.png\x00.jpg"])
def test_traversal_through_the_name_is_refused(images, home, tmp_path, name):
    (tmp_path / "team.png").write_bytes(PNG)
    (tmp_path / "outside").mkdir()
    (tmp_path / "outside" / "team.png").write_bytes(PNG)
    with pytest.raises(images.ImageError) as e:
        images.load_path(name)
    assert e.value.category in ("Validation", "Blocked"), name


def test_a_symlink_named_like_the_picture_is_never_followed(images, home, tmp_path):
    secret = tmp_path / "secret.png"
    secret.write_bytes(PNG)
    (home / "cache" / "images" / "team.png").symlink_to(secret)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    assert "No picture named team.png" in str(e.value)


def test_a_symlinked_folder_is_not_walked(images, home, tmp_path):
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "team.png").write_bytes(PNG)
    (home / "cache" / "images" / "linked").symlink_to(outside)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    assert "No picture named team.png" in str(e.value)


def test_a_symlinked_default_root_is_not_searched(images, home, tmp_path):
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "team.png").write_bytes(PNG)
    root = home / "cache" / "images"
    root.rmdir()
    root.symlink_to(outside)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    assert "No picture named team.png" in str(e.value) and str(outside) not in str(e.value)


def test_swap_to_symlink_after_the_name_lookup_is_refused(images, home, tmp_path, monkeypatch):
    secret = tmp_path / "secret.png"
    secret.write_bytes(JPEG)
    f = home / "cache" / "images" / "team.png"
    f.write_bytes(PNG)
    real_find = images.find_by_name

    def find(*a, **k):
        out = real_find(*a, **k)
        f.unlink()
        f.symlink_to(secret)  # swapped after the lookup, before the open
        return out
    monkeypatch.setattr(images, "find_by_name", find)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    assert e.value.category == "Blocked"


def test_the_lookup_is_bounded(images, home, monkeypatch):
    deep = home / "cache" / "images"
    for i in range(images.MAX_NAME_DEPTH + 1):
        deep = deep / f"d{i}"
    deep.mkdir(parents=True)
    (deep / "team.png").write_bytes(PNG)
    with pytest.raises(images.ImageError) as e:
        images.load_path("team.png")
    assert "stopped early" in str(e.value)
    monkeypatch.setattr(images, "MAX_NAME_SCAN", 3)
    for i in range(5):
        (home / "cache" / "images" / f"x{i}.png").write_bytes(PNG)
    with pytest.raises(images.ImageError) as e:
        images.load_path("x4.png")
    assert "stopped early" in str(e.value)


def test_path_oversize_and_wrong_type(images, home):
    d = home / "cache" / "images"
    (d / "big.png").write_bytes(PNG + b"\0" * images.MAX_IMAGE_BYTES)
    (d / "anim.png").write_bytes(GIF)
    (d / "page.png").write_text("<html>")
    with pytest.raises(images.ImageError) as e:
        images.load_path(str(d / "big.png"))
    assert e.value.category == "Capacity"
    for name in ("anim.png", "page.png"):
        with pytest.raises(images.ImageError) as e:
            images.load_path(str(d / name))
        assert e.value.category == "WrongType"
    with pytest.raises(images.ImageError) as e:
        images.load_path("cache/images/anim.png")
    assert e.value.category == "Validation"


# ---- HTML sanitizer ----------------------------------------------------------------------------------
def _fetch_ok(url):
    images = sys.modules["hermes_office_desktop.images"]
    if "bad" in url:
        raise images.ImageError("Blocked", "private")
    return images.Image(PNG, "png", {"type": "url", "url": url})


def test_sanitizer_strips_active_content(images):
    raw = ('<p onclick="x()">Hi <b>there</b><script>alert(1)</script></p><style>p{background:url(https://t.example/p)}</style>'
           '<iframe src="https://evil.example"></iframe><object data="x"></object><embed src="x"><form action="https://e">'
           '<input name=a></form><a href="javascript:alert(1)">js</a><a href="https://ok.example/">ok</a>'
           '<span style="color:red">red</span><span style="background:url(https://t.example/q)">bg</span>'
           '<link rel=stylesheet href="https://t.example/s.css"><table><tr><td colspan="2">c</td></tr></table>'
           '<!-- comment --><div background="https://t.example/b.png">d</div><img src="https://img.example/a.png">')
    c = images.sanitize_html(raw, "word", "alt", _fetch_ok)
    h = c.html
    for bad in ("script", "alert", "onclick", "iframe", "object", "embed", "form", "input", "javascript", "url(",
                "t.example", "comment", "<link"):
        assert bad not in h, bad
    for good in ("<b>there</b>", '<a href="https://ok.example/">ok</a>', '<span style="color:red">red</span>',
                 '<td colspan="2">c</td>', ">js<", ">bg<", ">d<"):
        assert good in h, good
    assert "style with url/expression" in c.dropped and "on* handler" in c.dropped


def test_sanitizer_inlines_https_and_data_images_and_drops_the_rest(images):
    data = "data:image/png;base64," + base64.b64encode(PNG).decode()
    raw = (f'<p><img src="https://img.example/a.png" alt="A"></p><img src="{data}"><img src="http://plain.example/x.png">'
           '<img src="https://bad.example/x.png"><img src="data:image/svg+xml;base64,PHN2Zz4="><img src="/rel.png">'
           '<img srcset="https://img.example/2x.png 2x">')
    c = images.sanitize_html(raw, "onenote", "Default alt", _fetch_ok)
    assert c.inlined == 2 and c.html.count("<img ") == 2
    assert 'alt="A"' in c.html and 'alt="Default alt"' in c.html
    assert "plain.example" not in c.html and "bad.example" not in c.html and "svg" not in c.html
    assert [i["status"] for i in c.images] == ["inlined", "inlined", "dropped", "dropped", "dropped", "dropped", "dropped"]
    assert "<style" not in images.sanitize_html("<style>p{color:red}</style><p>x</p><img src=\"https://img.example/a.png\">",
                                                "onenote", "a", _fetch_ok).html
    assert "<style" not in images.sanitize_html("<style>p{color:red}</style><p>x</p><img src=\"https://img.example/a.png\">",
                                                "word", "a", _fetch_ok).html


WORD_CSS_PAYLOADS = [
    # A6 #4: URL-valued CSS without url(
    "<style>@font-face{font-family:x;src:https://attacker.invalid/font.woff2}</style><p>ok</p>",
    '<style>p{background-image:image-set("https://attacker.invalid/p.png" 1x)}</style><p>ok</p>',
    # encoded / escaped variants
    "<style>p{background:u\\72l(https://attacker.invalid/e)}</style><p>ok</p>",
    "<style>@\\69mport 'https://attacker.invalid/i.css';</style><p>ok</p>",
    "<style>p{color:red}</style><p>ok</p>",
    "<STYLE type=text/css>p{color:red}</STYLE><p>ok</p>",
]
STYLE_ATTR_PAYLOADS = [
    'background-image:image-set("https://attacker.invalid/p.png" 1x)',
    "background:image-set(https://attacker.invalid/p.png 1x)",
    "src:https://attacker.invalid/f.woff2",
    "background:u\\72l(https://attacker.invalid/e)",
    "background:&#117;rl(https://attacker.invalid/ent)",
    "background:&#x75;&#x72;&#x6c;(https://attacker.invalid/hex)",
    "width:expression(alert(1))",
    "background:url (https://attacker.invalid/sp)",
    "list-style:image(https://attacker.invalid/i.png)",
    "background:cross-fade(https://attacker.invalid/a.png, 50%)",
    "color:red;/**/background:url(https://attacker.invalid/c)",
    "--x:https://attacker.invalid/v;background:var(--x)",
    "behavior:url(x.htc)",
    "content:'https://attacker.invalid/q'",
]


@pytest.mark.parametrize("target", ["word", "onenote"])
@pytest.mark.parametrize("raw", WORD_CSS_PAYLOADS)
def test_no_stylesheet_survives_the_untrusted_html_channel(images, raw, target):
    out = images.sanitize_html(raw + '<img src="https://img.example/a.png">', target, "a", _fetch_ok).html
    assert "<style" not in out.lower() and "attacker" not in out and "@" not in out, out
    assert "<p>ok</p>" in out


@pytest.mark.parametrize("css", STYLE_ATTR_PAYLOADS)
def test_style_attributes_carry_no_url_bearing_css(images, css):
    raw = f'<p><span style="{css}">t</span><img src="https://img.example/a.png"></p>'
    out = images.sanitize_html(raw, "word", "a", _fetch_ok).html
    assert "attacker" not in out and "url" not in out.lower() and "expression" not in out and "image-set" not in out, out
    assert "style=" not in out, out  # fail closed: the whole attribute goes, not just the bad declaration
    assert ">t<" in out


def test_style_attribute_allowlist_keeps_simple_formatting(images):
    raw = ('<p><span style="color: red; font-weight: bold; position: fixed; text-align:center">a</span>'
           '<span style="background-color:#ffeeaa">b</span><img src="https://img.example/a.png"></p>')
    out = images.sanitize_html(raw, "word", "a", _fetch_ok).html
    assert '<span style="color:red;font-weight:bold;text-align:center">a</span>' in out, out
    assert '<span style="background-color:#ffeeaa">b</span>' in out and "position" not in out


@pytest.mark.parametrize("target", ["word", "onenote"])
def test_text_only_html_is_refused_for_the_picture_tool(images, target):
    """A6 #5: office_insert_image inserts a picture; HTML with no retained <img> is refused."""
    for raw in ("<p>plain text</p>", '<p>x<img src="http://plain.example/a.png"></p>', '<p>x<img src="https://bad.example/a.png"></p>'):
        with pytest.raises(images.ImageError) as e:
            images.sanitize_html(raw, target, "a", _fetch_ok)
        assert e.value.category == "Validation" and "picture" in str(e.value), raw


def test_sanitizer_bounds(images):
    with pytest.raises(images.ImageError) as e:
        images.sanitize_html("x" * (images.MAX_HTML_CHARS + 1), "word", "a")
    assert e.value.category == "Capacity"
    with pytest.raises(images.ImageError) as e:
        images.sanitize_html("<script>only</script>", "word", "a")
    assert e.value.category == "Validation"
    many = "".join(f'<img src="https://img.example/{i}.png">' for i in range(8))
    c = images.sanitize_html(many, "word", "a", _fetch_ok)
    assert c.inlined == images.MAX_HTML_IMAGES


# ---- the tool -----------------------------------------------------------------------------------------
def _fresh(plugin, port, config=None):
    if plugin._server:
        plugin._server.stop()
    plugin._server = None
    ctx = FakeCtx({"allowed_origin": ORIGIN, **(config or {})})
    ctx._test_ports = tuple(range(port, port + 5))
    plugin.register(ctx)
    return ctx


def _call(ctx, args, sid="chat-1"):
    return json.loads(ctx.tools["office_insert_image"]["handler"](args, session_id=sid))


class Pane:
    """A paired fake pane answering insert_image / image_status / undo_last on a thread."""

    def __init__(self, plugin, ctx, app, answers):
        self.server = plugin.get_server(ctx)
        self.server.ensure_started()
        code, _ = self.server.registry.mint_code("chat-1", app)
        doc = {"name": "HOL images", **({"id": "nb-1"} if app == "onenote" else {})}
        self.b = self.server.registry.pair(code, app, doc, {}, [])
        self.answers, self.seen, self.stop = answers, [], False
        self.t = threading.Thread(target=self.run, daemon=True)
        self.t.start()

    def run(self):
        reg = self.server.registry
        while not self.stop:
            cmd = reg.next_command(self.b, 0.2)
            if cmd is None:
                continue
            self.seen.append((cmd.tool, dict(cmd.arguments)))
            fn = self.answers.get(cmd.tool)
            res = fn(cmd.arguments) if callable(fn) else (fn or {"ok": False, "error": {"category": "Unsupported", "message": "no"}})
            reg.complete(self.b, cmd.request_id, res)

    def close(self):
        self.stop = True
        self.t.join(2)


@pytest.fixture()
def tool(plugin, home):
    made = []

    def make(app, answers, port=19900, config=None):
        ctx = _fresh(plugin, port, config)
        pane = Pane(plugin, ctx, app, answers) if answers is not None else None
        made.append(pane)
        return ctx, pane
    yield make
    for p in made:
        if p:
            p.close()
    if plugin._server:
        plugin._server.stop()
    plugin._server = None


def _gen(home, name="image_x.png", data=PNG):
    f = home / "cache" / "generated" / "images" / name
    f.write_bytes(data)
    return str(f)


def test_tool_is_registered_with_one_source_schema(plugin, tool):
    ctx, _ = tool("excel", None)
    t = ctx.tools["office_insert_image"]
    assert t["toolset"] == "office-desktop"
    props = t["schema"]["parameters"]["properties"]
    assert {"app", "path", "url", "html", "alt", "credit", "prompt"} <= set(props)
    for bad in ({"app": "excel"}, {"app": "excel", "path": "/a", "url": "https://x/y"},
                {"app": "excel", "url": "https://x/y", "html": "<p/>"}, {"app": "excel", "path": "/a", "cell": "A1:B2"},
                {"app": "word", "path": "/a", "slideIndex": 1}, {"app": "excel", "path": "/a", "bridgeSessionId": "x"}):
        out = _call(ctx, bad)
        assert out["ok"] is False and out["error"]["category"] == "Validation", bad


def test_excel_path_insert_is_verified_by_readback(plugin, tool, home):
    answers = {"insert_image": lambda a: {"ok": True, "data": {"sheet": "Sheet1", "shapeId": "S1", "cell": "B2",
                                                                "format": "png", "altTextApplied": True,
                                                                "undo": {"available": True, "depth": 1}}},
               "image_status": lambda a: {"ok": True, "data": {"exists": a.get("shapeId") == "S1", "type": "Image"}}}
    ctx, pane = tool("excel", answers)
    out = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "Q4 chart", "sheet": "Sheet1", "cell": "B2", "width": 200})
    assert out["ok"] is True, out
    d = out["data"]
    assert d["status"] == "verified" and d["shapeId"] == "S1" and d["undo"]["available"] is True
    assert d["readback"] == {"exists": True, "type": "Image"}
    tool_name, args = pane.seen[0]
    assert tool_name == "insert_image" and base64.b64decode(args["base64"]) == PNG
    assert args["alt"] == "Q4 chart" and args["sheet"] == "Sheet1" and args["cell"] == "B2" and args["width"] == 200
    assert pane.seen[1] == ("image_status", {"sheet": "Sheet1", "shapeId": "S1"})


def test_readback_missing_is_uncertain_and_unreadable_is_applied(plugin, tool, home):
    ack = {"ok": True, "data": {"sheet": "S", "shapeId": "S9", "format": "png", "altTextApplied": True}}
    ctx, _ = tool("excel", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": False}}})
    d = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})["data"]
    assert d["status"] == "uncertain" and "do not retry" in d["next"].lower()
    ctx, _ = tool("excel", {"insert_image": ack, "image_status": {"ok": False, "error": {"category": "Office", "message": "x"}}},
                  port=19910)
    d = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})["data"]
    assert d["status"] == "applied"


def test_a_bare_picture_name_is_inserted_and_the_receipt_names_the_file(plugin, tool, home):
    """A27 cold story-bare: "put team.png from my Hermes images folder on the title slide" — the tool finds it."""
    f = home / "cache" / "images" / "demo" / "team.png"
    f.parent.mkdir()
    f.write_bytes(PNG)
    ack = {"ok": True, "data": {"slideIndex": 0, "slideId": "256#", "shapeId": "7", "format": "png", "altTextApplied": True}}
    ctx, pane = tool("powerpoint", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True}}})
    desc = ctx.tools["office_insert_image"]["schema"]["description"]
    assert "path may be a file name in your Hermes images folder" in desc
    out = _call(ctx, {"app": "powerpoint", "path": "team.png", "alt": "The team", "slideIndex": 0})
    assert out["ok"] is True and out["data"]["status"] == "verified", out
    src = out["data"]["source"]
    assert src["name"] == "team.png" and src["foundByName"] == "team.png" and src["path"] == str(f.resolve())
    assert base64.b64decode(pane.seen[0][1]["base64"]) == PNG
    out = _call(ctx, {"app": "powerpoint", "path": "nobody.png", "alt": "x", "slideIndex": 0})
    assert out["ok"] is False and out["error"]["category"] == "Validation"
    assert "No picture named nobody.png" in out["error"]["message"] and "Searched:" in out["error"]["message"]
    assert len(pane.seen) == 2  # the miss never reached the pane


def test_alt_text_not_applied_is_partial(plugin, tool, home):
    ack = {"ok": True, "data": {"slideIndex": 1, "slideId": "256#", "shapeId": "7", "format": "png", "altTextApplied": False}}
    ctx, pane = tool("powerpoint", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True}}})
    d = _call(ctx, {"app": "powerpoint", "path": _gen(home), "alt": "logo", "slideIndex": 1})["data"]
    assert d["status"] == "partial" and "alt text" in d["next"]
    assert pane.seen[1] == ("image_status", {"slideId": "256#", "shapeId": "7"})


def test_alt_defaults_from_prompt_or_credit_and_is_required(plugin, tool, home):
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "altTextApplied": True}}
    ctx, pane = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True, "pictures": 1}}})
    out = _call(ctx, {"app": "word", "path": _gen(home)})
    assert out["ok"] is False and out["error"]["category"] == "Validation" and "alt" in out["error"]["message"]
    assert pane.seen == []
    assert _call(ctx, {"app": "word", "path": _gen(home), "prompt": "a red fox in snow"})["ok"]
    assert pane.seen[0][1]["alt"] == "a red fox in snow"


def test_url_source_records_url_and_credit(plugin, tool, net):
    ROUTES["/fox.jpg"] = lambda h: _send(h, ctype="image/jpeg", body=JPEG)
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "altTextApplied": True}}
    ctx, pane = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True, "pictures": 1}}})
    out = _call(ctx, {"app": "word", "url": "https://img.example/fox.jpg", "alt": "A fox", "credit": "Photo: J. Doe, CC BY 4.0"})
    assert out["ok"], out
    args = pane.seen[0][1]
    assert args["alt"] == "A fox — Photo: J. Doe, CC BY 4.0 — source: https://img.example/fox.jpg"
    assert args["location"] == "end" and args["tag"].startswith("hermes-image-")
    assert out["data"]["source"] == {"type": "url", "url": "https://img.example/fox.jpg", "redirects": 0,
                                     "credit": "Photo: J. Doe, CC BY 4.0"}
    assert pane.seen[1] == ("image_status", {"tag": args["tag"]})


def test_blocked_url_never_reaches_the_pane(plugin, tool, net):
    ctx, pane = tool("excel", {})
    out = _call(ctx, {"app": "excel", "url": "https://internal.example/x.png", "alt": "a"})
    assert out["ok"] is False and out["error"]["category"] == "Blocked" and out["error"]["hint"]
    assert pane.seen == [] and net["connect"] == []


def test_disallowed_path_never_reaches_the_pane(plugin, tool, tmp_path):
    f = tmp_path / "x.png"
    f.write_bytes(PNG)
    ctx, pane = tool("excel", {})
    out = _call(ctx, {"app": "excel", "path": str(f), "alt": "a"})
    assert out["error"]["category"] == "Blocked" and pane.seen == []


def test_image_dirs_setting_reaches_the_tool(plugin, tool, tmp_path):
    mine = tmp_path / "Pictures"
    mine.mkdir()
    (mine / "p.png").write_bytes(PNG)
    ack = {"ok": True, "data": {"sheet": "S", "shapeId": "1", "format": "png", "altTextApplied": True}}
    ctx, pane = tool("excel", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True}}},
                     config={"image_dirs": [str(mine)]})
    assert _call(ctx, {"app": "excel", "path": str(mine / "p.png"), "alt": "a"})["data"]["status"] == "verified"


@pytest.mark.parametrize("app", ["excel", "powerpoint"])
def test_html_for_excel_and_powerpoint_needs_raster(plugin, tool, app):
    ctx, pane = tool(app, {})
    out = _call(ctx, {"app": app, "html": "<h1>Q4</h1>", "alt": "Q4 summary card"})
    assert out["ok"] is False and out["error"]["category"] == "NeedsRaster"
    nr = out["needsRaster"]
    assert nr["app"] == app and nr["then"]["tool"] == "office_insert_image" and nr["then"]["args"]["app"] == app
    assert "path" in nr["then"]["args"] and any("screenshot" in s.lower() for s in nr["steps"])
    assert any(d.endswith(os.path.join("cache", "screenshots")) for d in nr["saveTo"])  # native paths
    assert pane.seen == []
    gate = ctx.hooks["pre_tool_call"]
    assert gate(tool_name="office_insert_image", args={"app": app, "html": "<p/>"}, session_id="chat-1") is None


def test_word_html_is_sanitized_before_it_reaches_the_pane(plugin, tool, net):
    ROUTES["/a.png"] = lambda h: _send(h, body=PNG)
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "kind": "html"}}
    ctx, pane = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True, "pictures": 1}}})
    out = _call(ctx, {"app": "word", "html": '<p onclick="x">Hi<script>bad()</script><img src="https://img.example/a.png"></p>'
                                              '<img src="https://internal.example/b.png">', "alt": "card"})
    assert out["ok"], out
    sent = pane.seen[0][1]["html"]
    assert "script" not in sent and "onclick" not in sent and "data:image/png;base64," in sent and "internal" not in sent
    assert out["data"]["status"] == "verified" and out["data"]["html"]["imagesInlined"] == 1
    assert out["data"]["html"]["imagesDropped"] == 1


def test_word_html_fewer_pictures_than_inlined_is_uncertain(plugin, tool, net):
    ROUTES["/a.png"] = lambda h: _send(h, body=PNG)
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "kind": "html"}}
    ctx, _ = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True, "pictures": 0}}})
    d = _call(ctx, {"app": "word", "html": '<p><img src="https://img.example/a.png"></p>', "alt": "c"})["data"]
    assert d["status"] == "uncertain"


@pytest.mark.parametrize("app", ["word", "onenote"])
def test_text_only_html_never_reaches_the_pane(plugin, tool, home, tmp_path, app):
    ctx, pane = tool(app, {})
    if app == "onenote":
        import sqlite3
        brainmod = sys.modules["hermes_office_desktop.brain"]
        ctx._test_brain = brainmod.Brain(lambda: sqlite3.connect(str(tmp_path / "brain.db"), check_same_thread=False))
    out = _call(ctx, {"app": app, "html": "<p>plain text</p>", "alt": "card"})
    assert out["ok"] is False and out["error"]["category"] == "Validation" and "picture" in out["error"]["message"]
    assert pane.seen == []


@pytest.mark.parametrize("readback", [{"exists": True, "pictures": 0}, {"exists": True}])
def test_word_html_readback_without_a_picture_is_not_verified(plugin, tool, net, readback):
    """A6 #5: an existing control is not enough; at least one actual picture must be read back."""
    ROUTES["/a.png"] = lambda h: _send(h, body=PNG)
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "kind": "html"}}
    ctx, _ = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": readback}})
    d = _call(ctx, {"app": "word", "html": '<p><img src="https://img.example/a.png"></p>', "alt": "c"})["data"]
    assert d["status"] == "uncertain"


def test_word_path_readback_with_zero_pictures_is_not_verified(plugin, tool, home):
    ack = {"ok": True, "data": {"tag": "t", "contentControlId": 1, "kind": "png", "altTextApplied": True}}
    ctx, _ = tool("word", {"insert_image": ack, "image_status": {"ok": True, "data": {"exists": True, "pictures": 0}}})
    assert _call(ctx, {"app": "word", "path": _gen(home), "alt": "a"})["data"]["status"] == "uncertain"


def test_pane_failure_and_uncertain_transport(plugin, tool, home):
    ctx, _ = tool("excel", {"insert_image": {"ok": False, "error": {"category": "Unsupported", "message": "needs ExcelApi 1.9"}}})
    out = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})
    assert out["ok"] is False and out["error"]["category"] == "Unsupported" and out["data"]["status"] == "failed"
    ctx, _ = tool("excel", {"insert_image": {"ok": False, "error": {"category": "ExecutionUncertain", "message": "lost"}}},
                  port=19910)
    out = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})
    assert out["ok"] is False and out["data"]["status"] == "uncertain" and "never retried" in out["data"]["next"]


def test_timeout_means_nothing_changed(plugin, tool, home, monkeypatch, bridge_mod):
    ctx, pane = tool("excel", None)
    server = plugin.get_server(ctx)
    server.ensure_started()
    code, _ = server.registry.mint_code("chat-1", "excel")
    server.registry.pair(code, "excel", {"name": "d"}, {}, [])  # paired, but no pane ever polls
    real = server.registry.call
    monkeypatch.setattr(server.registry, "call", lambda *a, **k: real(*a, **{**k, "timeout": 0.3}))
    out = _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})
    assert out["ok"] is False and out["error"]["category"] == "Timeout" and out["data"]["status"] == "failed"


def test_not_paired_is_refused_before_any_fetch(plugin, tool, net):
    ctx, _ = tool("excel", None)
    out = _call(ctx, {"app": "excel", "url": "https://img.example/a.png", "alt": "a"})
    assert out["error"]["category"] == "NotPaired" and net["resolve"] == []


def test_onenote_insert_goes_through_the_ledger_and_undo_removes_it(plugin, tool, home, tmp_path):
    import sqlite3
    brainmod = sys.modules["hermes_office_desktop.brain"]
    outline = {"id": "ol-9", "paragraphs": [{"id": "ol-9-0", "type": "Image"}, {"id": "ol-9-1", "type": "RichText",
                                                                                "text": "Photo by Ana"}]}
    fp = brainmod.outline_fingerprint(outline["paragraphs"])
    answers = {"insert_image": lambda a: {"ok": True, "data": {"pageId": "pg-1", "outlineId": "ol-9", "images": 1,
                                                                "altTextApplied": True,
                                                                "created": {"outlineIds": ["ol-9"], "fingerprints": {"ol-9": fp}}}},
               "image_status": lambda a: {"ok": True, "data": {"exists": True, "images": 1}},
               "get_context": {"ok": True, "data": {"page": {"id": "pg-1"}}},
               "get_page": {"ok": True, "data": {"pageId": "pg-1", "title": "Meeting", "outlines": [outline]}},
               "delete_owned": lambda a: {"ok": True, "data": {"pageId": "pg-1", "missing": [], "notOwned": [],
                                                               "deleted": [i for i in a["outlineIds"]
                                                                           if a["expect"].get(i) == fp]}}}
    ctx, pane = tool("onenote", answers)
    ctx._test_brain = brainmod.Brain(lambda: sqlite3.connect(str(tmp_path / "brain.db"), check_same_thread=False))
    out = _call(ctx, {"app": "onenote", "path": _gen(home), "alt": "Whiteboard photo", "credit": "Photo by Ana"})
    assert out["ok"], out
    d = out["data"]
    assert d["status"] == "verified" and d["outlineId"] == "ol-9" and d["undo"]["via"] == "office_undo app='onenote'"
    assert [t for t, _ in pane.seen][:2] == ["get_context", "insert_image"]  # A30: the open page is recorded first
    tool_name, args = pane.seen[1]
    assert tool_name == "insert_image" and args["expectNotebookId"] == "nb-1" and args["opId"].startswith("op-")
    assert "<img src=\"data:image/png;base64," in args["html"] and "Photo by Ana" in args["html"]
    assert args["alt"].startswith("Whiteboard photo")
    undo = json.loads(ctx.tools["office_undo"]["handler"]({"app": "onenote"}, session_id="chat-1"))
    assert undo["ok"] and undo["data"]["undone"] == "office_insert_image" and undo["data"]["deleted"] == ["ol-9"]


def test_approval_prompt_style(plugin, tool):
    ctx, _ = tool("powerpoint", None)
    gate = ctx.hooks["pre_tool_call"]
    server = plugin.get_server(ctx)
    code, _ = server.registry.mint_code("chat-1", "powerpoint")
    server.registry.pair(code, "powerpoint", {"name": "Q4 deck"}, {}, [])
    d = gate(tool_name="office_insert_image", args={"app": "powerpoint", "path": "/x.png", "slideIndex": 1}, session_id="chat-1")
    assert d["message"] == "PowerPoint · Q4 deck · insert image on slide 2 (undoable) · Workbench"
    assert d["rule_key"] == plugin.APPROVAL_RULE_KEY
    msg = lambda a: gate(tool_name="office_insert_image", args=a)["message"]  # noqa: E731
    assert msg({"app": "excel", "url": "https://x", "sheet": "Data", "cell": "B2"}) == \
        "Excel · insert image at Data!B2 (undoable) · Workbench"
    assert msg({"app": "word", "path": "/x"}) == "Word · insert image at the end (undoable) · Workbench"
    assert msg({"app": "word", "path": "/x", "location": "selection"}) == \
        "Word · insert image at the cursor (undoable) · Workbench"
    assert msg({"app": "onenote", "html": "<p/>"}) == "OneNote · insert image on the open page (undoable) · Workbench"


def test_unattended_insert_is_blocked(plugin, tool, monkeypatch, home):
    ctx, _ = tool("excel", None)
    monkeypatch.setenv("HERMES_CRON_SESSION", "1")
    assert ctx.hooks["pre_tool_call"](tool_name="office_insert_image", args={"app": "excel", "path": "/x"})["action"] == "block"
    assert _call(ctx, {"app": "excel", "path": _gen(home), "alt": "a"})["error"]["category"] == "Unattended"


def test_a_word_insert_error_after_the_picture_landed_is_read_back_not_called_nothing_changed(plugin, tool, home):
    """A19 (seen live, Word run 3 on 3076fb1): Word for the web raised "Cannot read properties of null (reading
    'message')" after the picture was in the document; the receipt said "Nothing was changed". The tag is read back."""
    boom = {"ok": False, "error": {"category": "Office", "message": "Cannot read properties of null (reading 'message')"}}
    ctx, pane = tool("word", {"insert_image": boom, "image_status": {"ok": True, "data": {"exists": True, "pictures": 1}}})
    out = _call(ctx, {"app": "word", "path": _gen(home), "alt": "a"})
    d = out["data"]
    assert out["ok"] is False and d["status"] == "uncertain" and "readback found the picture" in d["next"]
    assert "Nothing was changed" not in d["next"] and d["readback"]["exists"] is True
    ctx, _ = tool("word", {"insert_image": boom, "image_status": {"ok": True, "data": {"exists": False, "controls": 0}}},
                  port=19912)
    d = _call(ctx, {"app": "word", "path": _gen(home), "alt": "a"})["data"]
    assert d["status"] == "failed" and "readback found no picture" in d["next"]


def test_a28_onenote_insert_lost_mid_write_is_uncertain_and_ledgered_uncertain(plugin, tool, home, tmp_path):
    import sqlite3
    brainmod = sys.modules["hermes_office_desktop.brain"]
    msg = ("OneNote stopped answering this add-in: even reading which notebook is open fails (GeneralException), so "
           "every OneNote request from this pane will fail until it is reloaded. This change may or may not have been "
           "made: after the reload, read the page before trying again. Ask the user to reload the OneNote browser tab "
           "(or close and reopen the Workbench pane); the pane reconnects to this chat by itself. Then try again once.")
    lost = {"ok": False, "error": {"category": "PaneReload", "code": "OfficeContextLost", "message": msg}}
    ctx, pane = tool("onenote", {"insert_image": lost})
    ctx._test_brain = brainmod.Brain(lambda: sqlite3.connect(str(tmp_path / "brain.db"), check_same_thread=False))
    out = _call(ctx, {"app": "onenote", "path": _gen(home), "alt": "Whiteboard photo"})
    assert out["ok"] is False and out["data"]["status"] == "uncertain", out
    assert "Nothing was changed" not in json.dumps(out) and "reload the OneNote browser tab" in out["data"]["next"]
    assert [r["status"] for r in ctx._test_brain._q("select status from writes")] == ["uncertain"]


# ---- A30 P0: an uncertain OneNote image insert is released only by a successful read of that page -----------------
A30_LOST = {"ok": False, "error": {"category": "PaneReload", "code": "OfficeContextLost",
                                   "message": "OneNote stopped answering this add-in. This change may or may not have been made."}}


class _ImagePage:
    """One OneNote page for the image pane: insert_image is lost mid-write (landing or not), journalled 'failed'."""

    def __init__(self, land):
        self.land, self.lose, self.read_ok, self.journal = land, True, True, {}
        self.outlines = [{"id": "o-text", "paragraphs": [{"id": "o-text-0", "type": "RichText", "text": "Meeting notes"}]}]

    def insert(self, a):
        oid = "ol-%d" % (len(self.outlines) + 1)
        para = [{"id": oid + "-0", "type": "Image"}]
        if self.lose:
            if self.land:
                self.outlines.append({"id": oid, "paragraphs": para})
            self.journal[a["opId"]] = {"state": "failed", "result": A30_LOST}
            return A30_LOST
        self.outlines.append({"id": oid, "paragraphs": para})
        fp = sys.modules["hermes_office_desktop.brain"].outline_fingerprint(para)
        self.journal[a["opId"]] = {"state": "done", "result": {"ok": True, "data": {}}}
        return {"ok": True, "data": {"pageId": "pg-1", "outlineId": oid, "images": 1,
                                     "created": {"outlineIds": [oid], "fingerprints": {oid: fp}}}}

    def answers(self):
        return {"insert_image": self.insert,
                "op_status": lambda a: {"ok": True, "data": {"opId": a["opId"], **self.journal.get(a["opId"], {"state": "unknown"})}},
                "get_context": {"ok": True, "data": {"page": {"id": "pg-1"}}},
                "get_page": lambda a: {"ok": True, "data": {"pageId": "pg-1", "title": "Meeting", "outlines": self.outlines,
                                                            "text": "Meeting notes", "truncated": False}}
                if self.read_ok else {"ok": False, "error": {"category": "PaneBusy", "message": "no answer"}},
                "image_status": lambda a: {"ok": True, "data": {"exists": True, "images": 1}}}


def _a30_image(tool, tmp_path, land):
    import sqlite3
    brainmod = sys.modules["hermes_office_desktop.brain"]
    page = _ImagePage(land)
    ctx, pane = tool("onenote", page.answers())
    ctx._test_brain = brainmod.Brain(lambda: sqlite3.connect(str(tmp_path / "brain.db"), check_same_thread=False))
    return ctx, pane, page


def _inserts(pane):
    return sum(1 for t, _ in pane.seen if t == "insert_image")


def _read(ctx):
    return json.loads(ctx.tools["onenote_get_page"]["handler"]({"pageId": "pg-1"}, session_id="chat-1"))


@pytest.mark.parametrize("land", [True, False])
def test_a30_image_no_read_or_a_failed_read_never_inserts_twice(plugin, tool, home, tmp_path, land):
    ctx, pane, page = _a30_image(tool, tmp_path, land)
    path = _gen(home)
    args = {"app": "onenote", "path": path, "alt": "Whiteboard photo"}
    assert _call(ctx, args)["data"]["status"] == "uncertain"
    page.lose = False  # reloaded
    for _ in range(4):
        out = _call(ctx, args)
        assert out["ok"] is False and out["data"]["status"] == "uncertain", out
        assert "Nothing was changed" not in json.dumps(out)
    page.read_ok = False
    assert _read(ctx)["ok"] is False
    assert _call(ctx, args)["data"]["status"] == "uncertain"
    assert _inserts(pane) == 1
    assert [r["reconciled"] for r in ctx._test_brain._q("select reconciled from writes")] == [None]


def test_a30_image_read_showing_the_picture_refuses_the_retry(plugin, tool, home, tmp_path):
    ctx, pane, page = _a30_image(tool, tmp_path, True)
    args = {"app": "onenote", "path": _gen(home), "alt": "Whiteboard photo"}
    _call(ctx, args)
    page.lose = False
    assert _read(ctx)["ok"] is True
    out = _call(ctx, args)
    assert out["ok"] is False and out["error"]["category"] == "AlreadyApplied", out
    assert out["data"]["status"] == "applied" and "not inserted again" in out["data"]["next"]
    assert "Nothing was changed" not in json.dumps(out) and _inserts(pane) == 1


def test_a30_image_read_without_the_picture_inserts_it_once(plugin, tool, home, tmp_path):
    ctx, pane, page = _a30_image(tool, tmp_path, False)
    args = {"app": "onenote", "path": _gen(home), "alt": "Whiteboard photo"}
    _call(ctx, args)
    page.lose = False
    assert _read(ctx)["ok"] is True
    out = _call(ctx, args)
    assert out["ok"] is True and out["data"]["status"] == "verified", out
    assert _inserts(pane) == 2
    assert [r["status"] for r in ctx._test_brain._q("select status from writes order by id")] == ["abandoned", "done"]
