"""CSRF tokens for form-submitting routes, and the checks that keep the app on the local network."""

import hmac
import ipaddress
import secrets

from flask import session


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
