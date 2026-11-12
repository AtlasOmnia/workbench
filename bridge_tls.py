# SPDX-License-Identifier: Apache-2.0
"""Per-install localhost TLS identity for the Workbench bridge. Never modifies the OS trust store.

Office for the web loads add-in pages over HTTPS only, so the bridge serves the bundled add-in (and its own API) on
https://localhost:<port>. The first start creates a local CA and a localhost leaf certificate in this Hermes
profile's plugin data folder (``<Hermes home>/plugin-data/hermes-office-desktop/tls``, owner-only files). The CA is
name-constrained to ``localhost`` and ``127.0.0.1``. Trusting it is a manual, one-time step the user performs with
the command ``trust_instructions()`` prints; the plugin never runs it.

Adapted from AtlasOmnia/sidecar ``desk_tls.py`` (commit 74a6c41).
"""
from __future__ import annotations

import _ssl
import ipaddress
import os
import shlex
import shutil
import ssl
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

CA_NAME = "Workbench local CA"
_PERMITTED = ("localhost", "127.0.0.1/32")


def _stdlib_ssl_context_class() -> type:
    """The standard library's SSLContext, even after ``truststore.inject_into_ssl()`` replaced ssl.SSLContext."""
    if ssl.SSLContext.__module__ == "ssl":
        return ssl.SSLContext
    for cls in _ssl._SSLContext.__subclasses__():
        if cls.__module__ == "ssl" and cls.__name__ == "SSLContext":
            return cls
    raise RuntimeError("standard library SSLContext is unavailable")


def _context(protocol: int) -> ssl.SSLContext:
    # Hermes injects truststore process-wide (ssl.SSLContext becomes truststore's class), which verifies against the
    # OS store only. Build from the stdlib class; its property setters recurse while truststore is injected, so the
    # TLS floor is set through the C-level descriptor.
    context = _stdlib_ssl_context_class()(protocol)
    _ssl._SSLContext.minimum_version.__set__(context, ssl.TLSVersion.TLSv1_2)
    return context


class LocalTLS:
    def __init__(self, data_dir: Path) -> None:
        self.directory = Path(data_dir) / "tls"
        self.ca_path = self.directory / "ca.pem"
        self.cert_path = self.directory / "localhost.pem"
        self.key_path = self.directory / "localhost-key.pem"
        self.ca_key_path = self.directory / "ca-key.pem"
        self._unsafe_ca_reported = False

    @staticmethod
    def _write(path: Path, content: bytes) -> None:
        # Restrict access at creation, not after writing key material.
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "wb") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())

    @staticmethod
    def _constraints():
        from cryptography import x509
        return x509.NameConstraints([x509.DNSName(_PERMITTED[0]), x509.IPAddress(ipaddress.ip_network(_PERMITTED[1]))],
                                    None)

    def ensure(self) -> ssl.SSLContext:
        """Create the identity on first use (atomically), check it, renew the leaf when due; the server context."""
        # Lazy import: plugin validation/registration never generates a CA.
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509.oid import NameOID

        now = datetime.now(timezone.utc)
        if self.directory.is_symlink():
            raise ValueError("TLS directory must not be a symlink")
        if not self.directory.exists():
            self.directory.parent.mkdir(parents=True, exist_ok=True)
            staging = Path(tempfile.mkdtemp(prefix=".tls-", dir=self.directory.parent))
            try:
                ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
                leaf_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
                ca_name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, f"{CA_NAME} {os.urandom(8).hex()}")])
                ca = (x509.CertificateBuilder().subject_name(ca_name).issuer_name(ca_name)
                      .public_key(ca_key.public_key()).serial_number(x509.random_serial_number())
                      .not_valid_before(now - timedelta(minutes=5))
                      .not_valid_after(now - timedelta(minutes=5) + timedelta(days=825))
                      .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
                      .add_extension(x509.KeyUsage(False, False, False, False, False, True, True, False, False),
                                     critical=True)
                      .add_extension(self._constraints(), critical=True)
                      .add_extension(x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()), critical=False)
                      .sign(ca_key, hashes.SHA256()))
                leaf = self._leaf(ca, ca_key, leaf_key, now)
                for name, key in (("ca-key.pem", ca_key), ("localhost-key.pem", leaf_key)):
                    self._write(staging / name, key.private_bytes(serialization.Encoding.PEM,
                                serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
                self._write(staging / "ca.pem", ca.public_bytes(serialization.Encoding.PEM))
                self._write(staging / "localhost.pem", leaf.public_bytes(serialization.Encoding.PEM))
                try:
                    # Publish all four files together. A simultaneous first-run winner is reused.
                    os.rename(staging, self.directory)
                except OSError:
                    if not self.directory.is_dir():
                        raise
            finally:
                if staging.exists():
                    shutil.rmtree(staging)

        if self.directory.is_symlink():
            raise ValueError("TLS directory must not be a symlink")
        if os.name != "nt":
            os.chmod(self.directory, 0o700)
        for path in (self.ca_path, self.cert_path, self.key_path, self.ca_key_path):
            if path.is_symlink() or not path.is_file():
                raise ValueError("TLS identity is incomplete or unsafe; restore it or explicitly reset local trust")
            if os.name != "nt":
                os.chmod(path, 0o600)
        ca = x509.load_pem_x509_certificate(self.ca_path.read_bytes())
        try:
            constraints = ca.extensions.get_extension_for_class(x509.NameConstraints)
            basic = ca.extensions.get_extension_for_class(x509.BasicConstraints)
            usage = ca.extensions.get_extension_for_class(x509.KeyUsage)
            safe = (constraints.critical and constraints.value == self._constraints()
                    and basic.critical and basic.value == x509.BasicConstraints(True, 0)
                    and usage.critical
                    and usage.value == x509.KeyUsage(False, False, False, False, False, True, True, False, False)
                    and ca.not_valid_after_utc - ca.not_valid_before_utc <= timedelta(days=825))
        except x509.ExtensionNotFound:
            safe = False
        if not safe:
            if not self._unsafe_ca_reported:
                print("UNSAFE LOCAL CA: refusing to serve; no identity has been replaced. Remove the old Workbench "
                      "CA's trust from your login keychain (Keychain Access) or Current User Trusted Root "
                      "Certification Authorities (Windows certificate manager), then delete only this plugin's TLS "
                      f"directory: {self.directory}. Start again and review the newly printed trust command.")
                self._unsafe_ca_reported = True
            raise ValueError("Local CA policy is unsafe; remove old trust and reset TLS directory")
        ca_key = serialization.load_pem_private_key(self.ca_key_path.read_bytes(), password=None)
        leaf_key = serialization.load_pem_private_key(self.key_path.read_bytes(), password=None)
        leaf = x509.load_pem_x509_certificate(self.cert_path.read_bytes())
        spki = (serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
        if (ca.public_key().public_bytes(*spki) != ca_key.public_key().public_bytes(*spki)
                or not ca.not_valid_before_utc <= now < ca.not_valid_after_utc):
            raise ValueError("Local CA is invalid or expired; explicitly reset local trust")
        ca.verify_directly_issued_by(ca)
        leaf.verify_directly_issued_by(ca)
        if (leaf.public_key().public_bytes(*spki) != leaf_key.public_key().public_bytes(*spki)
                or now < leaf.not_valid_before_utc
                or leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
                .get_values_for_type(x509.DNSName) != ["localhost"]):
            raise ValueError("Localhost identity is invalid")
        if now + timedelta(days=30) >= leaf.not_valid_after_utc:
            # Renew the leaf under the same CA: no new trust step. Atomic certificate replacement.
            renewed = self._leaf(ca, ca_key, leaf_key, now)
            temp = self.directory / (".leaf-" + os.urandom(8).hex())
            try:
                self._write(temp, renewed.public_bytes(serialization.Encoding.PEM))
                os.replace(temp, self.cert_path)
            finally:
                temp.unlink(missing_ok=True)
        context = _context(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(str(self.cert_path), str(self.key_path))
        return context

    @staticmethod
    def _leaf(ca, ca_key, key, now):
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes
        from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
        return (x509.CertificateBuilder()
                .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")]))
                .issuer_name(ca.subject).public_key(key.public_key()).serial_number(x509.random_serial_number())
                .not_valid_before(now - timedelta(minutes=5))
                .not_valid_after(min(now + timedelta(days=365), ca.not_valid_after_utc))
                .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
                .add_extension(x509.SubjectAlternativeName([x509.DNSName("localhost")]), critical=False)
                .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
                .add_extension(x509.KeyUsage(True, False, True, False, False, False, False, False, False), critical=True)
                .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False)
                .sign(ca_key, hashes.SHA256()))

    def client_context(self) -> ssl.SSLContext:
        """Trust this install's CA only (never the OS store); hostname and certificate checks stay on."""
        context = _context(ssl.PROTOCOL_TLS_CLIENT)
        context.load_verify_locations(cafile=str(self.ca_path))
        return context

    def trust_command(self, platform: str | None = None) -> str:
        platform = platform or sys.platform
        if platform == "win32":
            # PowerShell single-quoted literal; do not interpolate an unescaped install path.
            return "certutil -user -addstore Root '" + str(self.ca_path).replace("'", "''") + "'"
        if platform == "darwin":
            return ('security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '
                    + shlex.quote(str(self.ca_path)))
        return f"(trust {self.ca_path} as a certificate authority in your user trust store)"

    def trust_instructions(self, platform: str | None = None) -> str:
        platform = platform or sys.platform
        where = " in PowerShell" if platform == "win32" else " in Terminal" if platform == "darwin" else ""
        return ("Manual step, once per Hermes profile (the plugin never changes your certificate trust): Office for "
                "the web loads the Workbench pane from https://localhost, signed by this profile's own local "
                f"certificate authority {self.ca_path}. It can vouch only for localhost and 127.0.0.1. Review and run "
                f"this command yourself{where}:\n{self.trust_command(platform)}\n"
                "Keep the plugin data folder private, and remove this trust (Keychain Access, or the Windows "
                "certificate manager) when you stop using Workbench.")
