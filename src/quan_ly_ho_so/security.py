"""CSRF token issuance and verification for form-submitting routes."""

import hmac
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
