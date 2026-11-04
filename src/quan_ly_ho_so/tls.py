"""HTTPS for LAN mode: a local certificate authority (CA) signs a certificate for this computer's LAN address.

The certificate also covers this computer's mDNS name (name.local), which stays the same when the Wi-Fi hands out
a new address. Phones show a warning until they trust ca.pem; trusting it once keeps working when the address or
name changes, because only the server certificate is reissued then, not the CA.
"""

import ipaddress
import os
import ssl
from datetime import datetime, timedelta, timezone

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

# The CA may only vouch for private addresses, localhost and mDNS names (*.local), so even a leaked CA key cannot
# impersonate real websites on the phones that trust it.
PERMITTED_NETWORKS = ["10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16", "127.0.0.0/8", "fc00::/7", "::1/128"]
# iOS rejects server certificates valid for more than 825 days.
SERVER_CERT_DAYS = 800
CA_DAYS = 3650


def lan_certificate(directory, address, hostname):
    """(cert_path, key_path) covering localhost, 127.0.0.1, `address` and `hostname`. Creates the CA on first use
    and reissues the server certificate when the address or hostname changes or it nears expiry."""
    directory.mkdir(exist_ok=True)
    ca_cert_path, ca_key_path = directory / "ca.pem", directory / "ca-key.pem"
    cert_path, key_path = directory / "cert.pem", directory / "key.pem"
    ips = [ipaddress.ip_address("127.0.0.1")]
    try:
        ips.append(ipaddress.ip_address(address))
    except ValueError:
        pass  # skip: lan_address() could not find the LAN address and returned its placeholder text
    names = ["localhost", hostname]

    if ca_cert_path.is_file() and ca_key_path.is_file():
        ca_cert = x509.load_pem_x509_certificate(ca_cert_path.read_bytes())
        ca_key = serialization.load_pem_private_key(ca_key_path.read_bytes(), password=None)
        renew = not server_cert_valid(cert_path, key_path, ca_cert, ips, names)
    else:
        ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        ca_cert = build_ca(ca_key)
        write_private_key(ca_key_path, ca_key)
        ca_cert_path.write_bytes(ca_cert.public_bytes(serialization.Encoding.PEM))
        renew = True

    if renew:
        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        write_private_key(key_path, key)
        cert_path.write_bytes(build_server_cert(key, ca_cert, ca_key, ips, names).public_bytes(serialization.Encoding.PEM))
    return cert_path, key_path


class LazyHandshakeContext(ssl.SSLContext):
    """Werkzeug wraps the listening socket, so by default each TLS handshake runs inside accept() on the one
    serving thread: a client that connects and stays silent (Chrome preconnects) freezes the whole app.
    Deferring the handshake moves it to the first read, in that connection's own request thread."""

    def wrap_socket(self, sock, *args, **kwargs):
        kwargs["do_handshake_on_connect"] = False
        return super().wrap_socket(sock, *args, **kwargs)


def lan_ssl_context(cert_path, key_path):
    context = LazyHandshakeContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(cert_path, key_path)
    return context


def server_cert_valid(cert_path, key_path, ca_cert, ips, names):
    if not cert_path.is_file() or not key_path.is_file():
        return False
    cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
    san = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
    covered_ips = san.get_values_for_type(x509.IPAddress)
    covered_names = [name.lower() for name in san.get_values_for_type(x509.DNSName)]
    return (
        cert.issuer == ca_cert.subject
        and all(ip in covered_ips for ip in ips)
        and all(name.lower() in covered_names for name in names)
        and cert.not_valid_after_utc > datetime.now(timezone.utc) + timedelta(days=7)
    )


def build_ca(key):
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Quan Ly Ho So - CA noi bo")])
    now = datetime.now(timezone.utc)
    return (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1))
        .not_valid_after(now + timedelta(days=CA_DAYS))
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=False, content_commitment=False, key_encipherment=False, data_encipherment=False,
                key_agreement=False, key_cert_sign=True, crl_sign=True, encipher_only=False, decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.NameConstraints(
                permitted_subtrees=[x509.IPAddress(ipaddress.ip_network(net)) for net in PERMITTED_NETWORKS]
                + [x509.DNSName("localhost"), x509.DNSName("local")],
                excluded_subtrees=None,
            ),
            critical=True,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
        .sign(key, hashes.SHA256())
    )


def build_server_cert(key, ca_cert, ca_key, ips, names):
    now = datetime.now(timezone.utc)
    return (
        x509.CertificateBuilder()
        .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, names[-1])]))
        .issuer_name(ca_cert.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1))
        .not_valid_after(now + timedelta(days=SERVER_CERT_DAYS))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True, content_commitment=False, key_encipherment=True, data_encipherment=False,
                key_agreement=False, key_cert_sign=False, crl_sign=False, encipher_only=False, decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .add_extension(
            x509.SubjectAlternativeName([x509.DNSName(name) for name in names] + [x509.IPAddress(ip) for ip in ips]),
            critical=False,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
        .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )


def write_private_key(path, key):
    pem = key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption())
    # Owner-only from the start, so the key is never briefly readable by other accounts on this computer.
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as file:
        file.write(pem)
