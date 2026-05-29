import hashlib
import json
import secrets
import time
from datetime import datetime
from typing import Optional

from fastapi import HTTPException, Request

from app.config import USERS_FILE, SESSION_COOKIE, load_config

PRESET_USERS = [
    {"username": "kpage@phillipcapital.com", "password": "Welcome2"},
    {"username": "haotian@phillipcapital.com", "password": "Welcome3"},
    {"username": "hcurtis@phillipcapital.com", "password": "Welcome 1"},
]

SESSIONS: dict[str, dict] = {}


def hash_password(password: str, salt: Optional[str] = None) -> tuple[str, str]:
    if salt is None:
        salt = secrets.token_hex(16)
    derived = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), 100_000
    )
    return salt, derived.hex()


def verify_password(password: str, salt: str, hashed: str) -> bool:
    try:
        _, candidate = hash_password(password, salt)
        return secrets.compare_digest(candidate, hashed)
    except Exception:
        return False


def load_users() -> dict:
    if USERS_FILE.exists():
        try:
            return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_users(users: dict):
    USERS_FILE.write_text(json.dumps(users, indent=2), encoding="utf-8")


def seed_preset_users():
    users = load_users()
    if users:
        return
    for preset in PRESET_USERS:
        salt, hashed = hash_password(preset["password"])
        users[preset["username"].lower()] = {
            "username": preset["username"],
            "salt": salt,
            "hash": hashed,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "created_by": "system",
        }
    save_users(users)


def add_user(username: str, password: str, created_by: str) -> tuple[bool, str]:
    username = username.strip()
    if not username:
        return False, "Username cannot be empty."
    if not password:
        return False, "Password cannot be empty."
    users = load_users()
    key = username.lower()
    if key in users:
        return False, f"User '{username}' already exists."
    salt, hashed = hash_password(password)
    users[key] = {
        "username": username,
        "salt": salt,
        "hash": hashed,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "created_by": created_by,
    }
    save_users(users)
    return True, f"User '{username}' added."


def delete_user(username: str, current_user: str) -> tuple[bool, str]:
    users = load_users()
    key = username.strip().lower()
    if key not in users:
        return False, "User not found."
    if key == current_user.strip().lower():
        return False, "You cannot delete the account you are currently signed in with."
    if len(users) <= 1:
        return False, "Cannot delete the last remaining user."
    removed = users.pop(key)
    save_users(users)
    return True, f"User '{removed.get('username', username)}' removed."


def update_password(username: str, new_password: str, actor: str) -> tuple[bool, str]:
    if not new_password:
        return False, "Password cannot be empty."
    users = load_users()
    key = username.strip().lower()
    if key not in users:
        return False, "User not found."
    record = users[key]
    salt, hashed = hash_password(new_password)
    record["salt"] = salt
    record["hash"] = hashed
    record["password_updated_at"] = datetime.now().isoformat(timespec="seconds")
    record["password_updated_by"] = actor
    users[key] = record
    save_users(users)
    actor_key = actor.strip().lower()
    if key != actor_key:
        stale_tokens = [
            tok for tok, sess in SESSIONS.items()
            if sess.get("username", "").strip().lower() == key
        ]
        for tok in stale_tokens:
            SESSIONS.pop(tok, None)
    return True, f"Password updated for '{record.get('username', username)}'."


def authenticate(username: str, password: str) -> Optional[str]:
    users = load_users()
    key = username.strip().lower()
    record = users.get(key)
    if not record:
        return None
    if verify_password(password, record.get("salt", ""), record.get("hash", "")):
        return record.get("username", username)
    return None


def create_session(username: str) -> str:
    token = secrets.token_urlsafe(32)
    now = time.time()
    SESSIONS[token] = {
        "username": username,
        "created": now,
        "last_active": now,
    }
    return token


def validate_session(token: Optional[str]) -> Optional[str]:
    if not token:
        return None
    session = SESSIONS.get(token)
    if not session:
        return None
    timeout_seconds = load_config()["session_timeout_minutes"] * 60
    if time.time() - session["last_active"] > timeout_seconds:
        SESSIONS.pop(token, None)
        return None
    session["last_active"] = time.time()
    return session["username"]


def destroy_session(token: Optional[str]):
    if token:
        SESSIONS.pop(token, None)


def current_user_or_none(request: Request) -> Optional[str]:
    token = request.cookies.get(SESSION_COOKIE)
    return validate_session(token)


def require_api_user(request: Request) -> str:
    user = current_user_or_none(request)
    if not user:
        raise HTTPException(status_code=401, detail="Login required.")
    return user
