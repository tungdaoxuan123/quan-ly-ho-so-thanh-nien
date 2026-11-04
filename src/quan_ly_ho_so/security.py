"""CSRF tokens for form-submitting routes, and the checks that keep the app on the local network."""

import hmac
import ipaddress
import os
import secrets

from flask import request, session


def csrf_token():
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def valid_csrf(value):
    return bool(value) and hmac.compare_digest(value, session.get("csrf_token", ""))


def is_local_network(address):
    """True for this computer and private LAN addresses, so the app is never reachable from the internet."""
    try:
        ip = ipaddress.ip_address(address)
    except ValueError:
        return False
    ip = getattr(ip, "ipv4_mapped", None) or ip
    return ip.is_loopback or ip.is_private or ip.is_link_local


def online_mode():
    """True when the app is published to the internet through a tunnel (QLHS_ONLINE=1)."""
    return os.environ.get("QLHS_ONLINE", "0").lower() in ("1", "true", "yes")


def client_ip():
    """The caller's address. Behind the tunnel every request comes from this computer, so read the
    address the tunnel forwards; the header is only trusted when the request really came from loopback."""
    address = request.remote_addr or ""
    if online_mode() and is_loopback(address):
        forwarded = request.headers.get("CF-Connecting-IP") or request.headers.get("X-Forwarded-For", "")
        forwarded = forwarded.split(",")[0].strip()
        if forwarded:
            return forwarded
    return address


def is_loopback(address):
    try:
        ip = ipaddress.ip_address(address)
    except ValueError:
        return False
    return (getattr(ip, "ipv4_mapped", None) or ip).is_loopback
