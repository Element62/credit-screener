from __future__ import annotations

import base64
import hashlib
import hmac
import time
from dataclasses import dataclass


SESSION_COOKIE = "credit_screenr_session"
SESSION_TTL_SECONDS = 60 * 60 * 12


@dataclass(frozen=True)
class SessionUser:
    username: str


def _sign(secret: str, payload: str) -> str:
    return hmac.new(secret.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()


def create_session_token(secret: str, username: str) -> str:
    issued_at = str(int(time.time()))
    payload = f"{username}|{issued_at}"
    signature = _sign(secret, payload)
    token = f"{payload}|{signature}"
    return base64.urlsafe_b64encode(token.encode("utf-8")).decode("utf-8")


def decode_session_token(secret: str, token: str | None) -> SessionUser | None:
    if not token:
        return None
    try:
        raw = base64.urlsafe_b64decode(token.encode("utf-8")).decode("utf-8")
        username, issued_at_str, signature = raw.split("|", 2)
    except Exception:
        return None

    payload = f"{username}|{issued_at_str}"
    expected = _sign(secret, payload)
    if not hmac.compare_digest(signature, expected):
        return None

    try:
        issued_at = int(issued_at_str)
    except ValueError:
        return None

    if time.time() - issued_at > SESSION_TTL_SECONDS:
        return None

    return SessionUser(username=username)
