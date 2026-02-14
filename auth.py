import hashlib
import html
import json
import os
import re
import secrets
import time
from datetime import datetime

import streamlit as st

USERS_DB_PATH = os.path.join(os.getcwd(), ".users.json")
AUTH_SESSIONS_DB_PATH = os.path.join(os.getcwd(), ".auth_sessions.json")
DEFAULT_TEMP_PASSWORD = os.getenv("DEFAULT_TEMP_PASSWORD", "Temp@123")
PASSWORD_MIN_LENGTH = 8
SYSTEM_ADMIN_USERNAME = "admin"
SYSTEM_ADMIN_PASSWORD = "System@123"
SESSION_IDLE_TIMEOUT_SECONDS = int(os.getenv("SESSION_IDLE_TIMEOUT_SECONDS", "1800"))


def _utc_now_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


class AuthManager:
    """Authentication manager with file-based users and password reset flow."""

    def __init__(self):
        self._init_session_state()
        self._ensure_user_store()

    def _init_session_state(self):
        if "authenticated" not in st.session_state:
            st.session_state.authenticated = False
        if "user_role" not in st.session_state:
            st.session_state.user_role = None
        if "username" not in st.session_state:
            st.session_state.username = None
        if "user_key" not in st.session_state:
            st.session_state.user_key = None
        if "login_attempts" not in st.session_state:
            st.session_state.login_attempts = 0
        if "last_attempt_time" not in st.session_state:
            st.session_state.last_attempt_time = 0.0
        if "session_start_time" not in st.session_state:
            st.session_state.session_start_time = None
        if "last_activity_time" not in st.session_state:
            st.session_state.last_activity_time = None
        if "auth_mode" not in st.session_state:
            st.session_state.auth_mode = "Login"
        if "_switch_to_login" not in st.session_state:
            st.session_state._switch_to_login = False
        if "show_forgot_password" not in st.session_state:
            st.session_state.show_forgot_password = False
        if "password_reset_required" not in st.session_state:
            st.session_state.password_reset_required = False
        if "auth_feedback" not in st.session_state:
            st.session_state.auth_feedback = None
        if "show_user_menu_change_password" not in st.session_state:
            st.session_state.show_user_menu_change_password = False
        if "show_user_menu_delete_confirm" not in st.session_state:
            st.session_state.show_user_menu_delete_confirm = False
        if "show_user_main_menu" not in st.session_state:
            st.session_state.show_user_main_menu = False
        if "show_user_profile_panel" not in st.session_state:
            st.session_state.show_user_profile_panel = False
        if "profile_menu_open" not in st.session_state:
            st.session_state.profile_menu_open = False
        if "user_theme" not in st.session_state:
            st.session_state.user_theme = "light"
        if "auth_token" not in st.session_state:
            st.session_state.auth_token = None

    def hash_password(self, password: str, salt: str = None) -> tuple[str, str]:
        if salt is None:
            salt = secrets.token_hex(16)
        password_hash = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100000
        )
        return password_hash.hex(), salt

    def verify_password(self, password: str, stored_hash: str, salt: str) -> bool:
        password_hash, _ = self.hash_password(password, salt)
        return password_hash == stored_hash

    def _load_user_store(self) -> dict:
        if not os.path.exists(USERS_DB_PATH):
            return {"users": {}}
        try:
            with open(USERS_DB_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {"users": {}}
            if not isinstance(data.get("users"), dict):
                data["users"] = {}
            return data
        except Exception:
            return {"users": {}}

    def _save_user_store(self, data: dict) -> bool:
        try:
            with open(USERS_DB_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def _load_auth_sessions(self) -> dict:
        if not os.path.exists(AUTH_SESSIONS_DB_PATH):
            return {"sessions": {}}
        try:
            with open(AUTH_SESSIONS_DB_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {"sessions": {}}
            if not isinstance(data.get("sessions"), dict):
                data["sessions"] = {}
            return data
        except Exception:
            return {"sessions": {}}

    def _save_auth_sessions(self, data: dict) -> bool:
        try:
            with open(AUTH_SESSIONS_DB_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def _get_query_session_token(self) -> str:
        try:
            raw = st.query_params.get("session")
            if isinstance(raw, list):
                return str(raw[0]).strip() if raw else ""
            return str(raw).strip() if raw else ""
        except Exception:
            return ""

    def _set_query_session_token(self, token: str):
        try:
            st.query_params["session"] = token
        except Exception:
            pass

    def _clear_query_session_token(self):
        try:
            if "session" in st.query_params:
                del st.query_params["session"]
        except Exception:
            pass

    def _remove_user_sessions(self, user_key: str):
        normalized = (user_key or "").strip().lower()
        if not normalized:
            return
        store = self._load_auth_sessions()
        sessions = store.setdefault("sessions", {})
        changed = False
        for token, entry in list(sessions.items()):
            if (entry.get("user_key") or "").strip().lower() == normalized:
                del sessions[token]
                changed = True
        if changed:
            self._save_auth_sessions(store)

    def _create_persistent_session(self, user_key: str, user_entry: dict):
        token = secrets.token_urlsafe(32)
        now = time.time()
        store = self._load_auth_sessions()
        sessions = store.setdefault("sessions", {})
        sessions[token] = {
            "user_key": (user_key or "").strip().lower(),
            "username": user_entry.get("display_name") or user_entry.get("username") or user_key,
            "role": user_entry.get("role", "user"),
            "created_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "created_ts": now,
            "last_activity_ts": now,
        }
        if self._save_auth_sessions(store):
            st.session_state.auth_token = token
            self._set_query_session_token(token)

    def _clear_persistent_session(self, token: str = ""):
        resolved_token = (token or st.session_state.get("auth_token") or self._get_query_session_token() or "").strip()
        if resolved_token:
            store = self._load_auth_sessions()
            sessions = store.setdefault("sessions", {})
            if resolved_token in sessions:
                del sessions[resolved_token]
                self._save_auth_sessions(store)
        st.session_state.auth_token = None
        self._clear_query_session_token()

    def _touch_persistent_session(self):
        token = (st.session_state.get("auth_token") or self._get_query_session_token() or "").strip()
        if not token:
            return
        store = self._load_auth_sessions()
        sessions = store.setdefault("sessions", {})
        entry = sessions.get(token)
        if not isinstance(entry, dict):
            return
        entry["last_activity_ts"] = time.time()
        entry["updated_at"] = _utc_now_iso()
        sessions[token] = entry
        self._save_auth_sessions(store)

    def _restore_persistent_session(self) -> bool:
        token = self._get_query_session_token()
        if not token:
            return False

        store = self._load_auth_sessions()
        sessions = store.setdefault("sessions", {})
        session_entry = sessions.get(token)
        if not isinstance(session_entry, dict):
            self._clear_query_session_token()
            return False

        last_activity = float(session_entry.get("last_activity_ts", 0) or 0)
        if (time.time() - last_activity) > SESSION_IDLE_TIMEOUT_SECONDS:
            del sessions[token]
            self._save_auth_sessions(store)
            self._clear_query_session_token()
            st.session_state.auth_feedback = {"level": "warning", "text": "Your session expired after 30 minutes of inactivity."}
            return False

        users, _ = self._ensure_user_store()
        user_key = (session_entry.get("user_key") or "").strip().lower()
        user_entry = users.get(user_key)
        if not user_entry:
            del sessions[token]
            self._save_auth_sessions(store)
            self._clear_query_session_token()
            return False

        st.session_state.authenticated = True
        st.session_state.user_role = user_entry.get("role", "user")
        st.session_state.username = user_entry.get("display_name") or user_entry.get("username") or user_key
        st.session_state.user_key = user_key
        st.session_state.user_theme = (user_entry.get("theme") or "light").lower()
        st.session_state.password_reset_required = bool(user_entry.get("must_change_password", False))
        st.session_state.login_attempts = 0
        st.session_state.auth_token = token

        now = time.time()
        st.session_state.session_start_time = float(session_entry.get("created_ts", now) or now)
        st.session_state.last_activity_time = now
        self._touch_persistent_session()
        return True

    def _default_user(
        self,
        username: str,
        password: str,
        role: str,
        first_name: str,
        last_name: str,
        email: str,
        allow_password_reset: bool = True,
        system_account: bool = False,
        theme: str = "light",
    ) -> dict:
        password_hash, salt = self.hash_password(password)
        return {
            "username": username,
            "display_name": username,
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "role": role,
            "password_hash": password_hash,
            "salt": salt,
            "must_change_password": False,
            "temporary_password_active": False,
            "allow_password_reset": allow_password_reset,
            "system_account": system_account,
            "theme": theme,
            "created_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
        }

    def _ensure_user_store(self) -> tuple[dict, dict]:
        store = self._load_user_store()
        users = store.setdefault("users", {})
        changed = False

        if SYSTEM_ADMIN_USERNAME not in users:
            users[SYSTEM_ADMIN_USERNAME] = self._default_user(
                username=SYSTEM_ADMIN_USERNAME,
                password=SYSTEM_ADMIN_PASSWORD,
                role="admin",
                first_name="Admin",
                last_name="User",
                email="admin@example.com",
                allow_password_reset=False,
                system_account=True,
            )
            changed = True
        else:
            admin = users[SYSTEM_ADMIN_USERNAME]
            admin_changed = False
            if not self.verify_password(
                SYSTEM_ADMIN_PASSWORD,
                admin.get("password_hash", ""),
                admin.get("salt", ""),
            ):
                password_hash, salt = self.hash_password(SYSTEM_ADMIN_PASSWORD)
                admin["password_hash"] = password_hash
                admin["salt"] = salt
                admin_changed = True
            if admin.get("role") != "admin":
                admin["role"] = "admin"
                admin_changed = True
            if admin.get("allow_password_reset", True):
                admin["allow_password_reset"] = False
                admin_changed = True
            if not admin.get("system_account", False):
                admin["system_account"] = True
                admin_changed = True
            if admin_changed:
                admin["updated_at"] = _utc_now_iso()
                changed = True

        if "suraj" not in users:
            users["suraj"] = self._default_user(
                username="Suraj",
                password="Suraj@123",
                role="user",
                first_name="Suraj",
                last_name="User",
                email="suraj@example.com",
            )
            changed = True

        for user_key, user_entry in users.items():
            if "theme" not in user_entry:
                user_entry["theme"] = "light"
                user_entry["updated_at"] = _utc_now_iso()
                changed = True
            if user_key != SYSTEM_ADMIN_USERNAME and "allow_password_reset" not in user_entry:
                user_entry["allow_password_reset"] = True
                user_entry["updated_at"] = _utc_now_iso()
                changed = True

        if changed:
            self._save_user_store(store)

        return users, store

    def is_rate_limited(self) -> bool:
        if st.session_state.login_attempts >= 3:
            elapsed = time.time() - st.session_state.last_attempt_time
            if elapsed < 300:
                return True
            st.session_state.login_attempts = 0
        return False

    def is_session_expired(self) -> bool:
        if not st.session_state.authenticated:
            return False
        if not st.session_state.get("last_activity_time"):
            return False
        return (time.time() - st.session_state.last_activity_time) > SESSION_IDLE_TIMEOUT_SECONDS

    def update_activity_time(self):
        st.session_state.last_activity_time = time.time()

    def check_session_timeout(self) -> bool:
        if self.is_session_expired():
            self.logout(rerun=False)
            st.warning("Your session has expired due to inactivity. Please log in again.")
            return False
        self.update_activity_time()
        self._touch_persistent_session()
        return True

    def authenticate_user(self, username: str, password: str) -> tuple[bool, str]:
        if self.is_rate_limited():
            remaining = 300 - (time.time() - st.session_state.last_attempt_time)
            return False, f"Too many failed attempts. Try again in {int(remaining/60)} minutes."

        users, _ = self._ensure_user_store()
        lookup = (username or "").strip().lower()
        user_entry = users.get(lookup)

        if not user_entry:
            st.session_state.login_attempts += 1
            st.session_state.last_attempt_time = time.time()
            return False, "Invalid credentials"

        if not self.verify_password(password or "", user_entry.get("password_hash", ""), user_entry.get("salt", "")):
            st.session_state.login_attempts += 1
            st.session_state.last_attempt_time = time.time()
            return False, "Invalid credentials"

        st.session_state.authenticated = True
        st.session_state.user_role = user_entry.get("role", "user")
        st.session_state.username = user_entry.get("display_name") or user_entry.get("username") or username
        st.session_state.user_key = lookup
        st.session_state.user_theme = (user_entry.get("theme") or "light").lower()
        st.session_state.password_reset_required = bool(user_entry.get("must_change_password", False))
        st.session_state.login_attempts = 0

        now = time.time()
        st.session_state.session_start_time = now
        st.session_state.last_activity_time = now
        self._create_persistent_session(lookup, user_entry)
        return True, st.session_state.user_role

    def register_user(
        self,
        username: str,
        first_name: str,
        last_name: str,
        email: str,
        password: str,
        confirm_password: str,
    ) -> tuple[bool, str]:
        username = (username or "").strip()
        first_name = (first_name or "").strip()
        last_name = (last_name or "").strip()
        email = (email or "").strip().lower()

        if not username or not first_name or not last_name or not email or not password or not confirm_password:
            return False, "Please fill all sign-up fields."
        if not re.match(r"^[A-Za-z0-9_.-]{3,40}$", username):
            return False, "Username must be 3-40 chars and use letters, numbers, _, ., - only."
        if "@" not in email or "." not in email:
            return False, "Please enter a valid email address."
        if len(password) < PASSWORD_MIN_LENGTH:
            return False, f"Password must be at least {PASSWORD_MIN_LENGTH} characters."
        if password != confirm_password:
            return False, "Password and confirm password do not match."

        users, store = self._ensure_user_store()
        username_key = username.lower()

        if username_key in users:
            return False, "Username already exists."
        if any((u.get("email") or "").lower() == email for u in users.values()):
            return False, "Email already registered."

        password_hash, salt = self.hash_password(password)
        users[username_key] = {
            "username": username,
            "display_name": username,
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "role": "user",
            "password_hash": password_hash,
            "salt": salt,
            "must_change_password": False,
            "temporary_password_active": False,
            "allow_password_reset": True,
            "system_account": False,
            "theme": "light",
            "created_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
        }

        if not self._save_user_store(store):
            return False, "Could not save user. Please try again."

        return True, "Sign up successful. Please login with your new credentials."

    def issue_temporary_password(self, identifier: str) -> tuple[bool, str, str]:
        lookup = (identifier or "").strip().lower()
        if not lookup:
            return False, "", "Enter username or email to reset password."

        users, store = self._ensure_user_store()

        target_key = None
        for user_key, user_data in users.items():
            if user_key == lookup or (user_data.get("email") or "").lower() == lookup:
                target_key = user_key
                break

        if not target_key:
            return False, "", "User not found for the provided username/email."
        target_user = users.get(target_key, {})
        if target_key == SYSTEM_ADMIN_USERNAME or target_user.get("role") == "admin":
            return False, "", "Admin password cannot be reset."
        if not target_user.get("allow_password_reset", True):
            return False, "", "Password reset is disabled for this user."

        temp_password = DEFAULT_TEMP_PASSWORD
        password_hash, salt = self.hash_password(temp_password)
        users[target_key]["password_hash"] = password_hash
        users[target_key]["salt"] = salt
        users[target_key]["must_change_password"] = True
        users[target_key]["temporary_password_active"] = True
        users[target_key]["updated_at"] = _utc_now_iso()

        if not self._save_user_store(store):
            return False, "", "Password reset failed while saving."

        return True, temp_password, "Temporary password generated."

    def change_password(self, current_password: str, new_password: str, confirm_password: str) -> tuple[bool, str]:
        if not st.session_state.get("authenticated"):
            return False, "Login required."

        user_key = (st.session_state.get("user_key") or "").strip().lower()
        if not user_key:
            return False, "Invalid user session."

        users, store = self._ensure_user_store()
        user_entry = users.get(user_key)
        if not user_entry:
            return False, "User record not found."

        if not current_password or not new_password or not confirm_password:
            return False, "Please fill current, new, and confirm password."
        if new_password != confirm_password:
            return False, "New password and confirm password do not match."
        if len(new_password) < PASSWORD_MIN_LENGTH:
            return False, f"New password must be at least {PASSWORD_MIN_LENGTH} characters."
        if new_password == current_password:
            return False, "New password must be different from current password."

        if not self.verify_password(current_password, user_entry.get("password_hash", ""), user_entry.get("salt", "")):
            return False, "Current password is incorrect."

        password_hash, salt = self.hash_password(new_password)
        user_entry["password_hash"] = password_hash
        user_entry["salt"] = salt
        user_entry["must_change_password"] = False
        user_entry["temporary_password_active"] = False
        user_entry["updated_at"] = _utc_now_iso()

        if not self._save_user_store(store):
            return False, "Could not save new password."

        st.session_state.password_reset_required = False
        return True, "Password updated successfully."

    def set_user_theme(self, theme: str) -> tuple[bool, str]:
        normalized = (theme or "").strip().lower()
        if normalized not in ("light", "dark"):
            return False, "Invalid theme selection."
        user_key = (st.session_state.get("user_key") or "").strip().lower()
        if not user_key:
            return False, "Invalid user session."

        users, store = self._ensure_user_store()
        user_entry = users.get(user_key)
        if not user_entry:
            return False, "User record not found."

        user_entry["theme"] = normalized
        user_entry["updated_at"] = _utc_now_iso()
        if not self._save_user_store(store):
            return False, "Could not save theme preference."

        st.session_state.user_theme = normalized
        return True, f"{normalized.title()} theme applied."

    def apply_user_theme(self):
        theme = (st.session_state.get("user_theme") or "light").strip().lower()
        if theme != "dark":
            return
        st.markdown(
            """
            <style>
            .stApp, .main {
                background: linear-gradient(135deg, #0b1220, #111827) !important;
                color: #e5e7eb !important;
            }
            .header-card {
                background: transparent !important;
            }
            .header-title, .step-title, h1, h2, h3, h4, h5, h6, label,
            [data-testid="stMarkdownContainer"] p,
            [data-testid="stMarkdownContainer"] li,
            [data-testid="stCaptionContainer"] p {
                color: #e5e7eb !important;
            }
            .header-subtitle {
                color: #cbd5e1 !important;
            }
            .step-card, .auth-card, .info-box, .stDataFrame, .stAlert, [data-testid="stForm"] {
                background: #111827 !important;
                border-color: #334155 !important;
            }
            .step-card.small {
                background: #111827 !important;
                border-color: #334155 !important;
            }
            .step-number {
                background: #1e293b !important;
                color: #e5e7eb !important;
                border: 1px solid #475569 !important;
            }
            .step2-field-label {
                color: #e5e7eb !important;
            }
            .info-box {
                color: #cbd5e1 !important;
                border-color: #334155 !important;
                background: #0f172a !important;
            }
            .stSelectbox > div > div {
                background: #0f172a !important;
                border: 1px solid #334155 !important;
            }
            .stSelectbox > div > div * {
                color: #e5e7eb !important;
                fill: #e5e7eb !important;
            }
            [data-testid="stFileUploaderDropzone"],
            [data-testid="stFileUploaderDropzone"] * {
                background: #0f172a !important;
                color: #e5e7eb !important;
                border-color: #334155 !important;
            }
            .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"], .stNumberInput input {
                background: #0f172a !important;
                color: #e5e7eb !important;
                border-color: #334155 !important;
            }
            .stTextInput input::placeholder, .stTextArea textarea::placeholder {
                color: #94a3b8 !important;
                opacity: 1 !important;
            }
            .stSelectbox div[data-baseweb="select"] *, .stSelectbox svg {
                color: #e5e7eb !important;
                fill: #e5e7eb !important;
            }
            div[data-baseweb="menu"], ul[role="listbox"] {
                background: #0f172a !important;
                color: #e5e7eb !important;
                border: 1px solid #334155 !important;
            }
            div[data-baseweb="menu"] *, ul[role="listbox"] * {
                color: #e5e7eb !important;
            }
            div[data-baseweb="popover"], div[data-baseweb="popover"] * {
                background: #111827 !important;
                color: #e5e7eb !important;
                border-color: #334155 !important;
            }
            .stButton button[kind="primary"], .stButton button[kind="secondary"], .stFormSubmitButton button {
                background: #1e293b !important;
                color: #e5e7eb !important;
                border: 1px solid #475569 !important;
            }
            a {
                color: #7dd3fc !important;
            }
            [data-testid="stCaptionContainer"] p, small {
                color: #cbd5e1 !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

    def authenticate(self, username: str, password: str) -> bool:
        success, role_or_error = self.authenticate_user(username, password)
        if success:
            st.session_state.authenticated = True
            st.session_state.user_role = role_or_error
            return True
        return False

    def get_user_info(self, username: str) -> dict:
        users, _ = self._ensure_user_store()
        lookup = (username or "").strip().lower()
        if lookup in users:
            u = users[lookup]
            return {"username": u.get("display_name", username), "role": u.get("role", "user")}
        return {"username": username, "role": st.session_state.get("user_role", "user")}

    def add_user(self, username: str, password: str, role: str = "user") -> bool:
        ok, _ = self.register_user(username, username, "User", f"{username}@example.com", password, password)
        if ok:
            users, store = self._ensure_user_store()
            key = username.strip().lower()
            if key in users:
                users[key]["role"] = role
                self._save_user_store(store)
        return ok

    def list_users(self) -> list:
        users, _ = self._ensure_user_store()
        out = []
        for u in users.values():
            out.append(
                {
                    "username": u.get("display_name") or u.get("username"),
                    "role": u.get("role", "user"),
                    "email": u.get("email", ""),
                }
            )
        return out

    def delete_user(self, username: str) -> bool:
        users, store = self._ensure_user_store()
        lookup = (username or "").strip().lower()
        if lookup in (SYSTEM_ADMIN_USERNAME,):
            return False
        if lookup not in users:
            return False
        del users[lookup]
        return self._save_user_store(store)

    def _remove_user_settings(self, user_key: str):
        settings_path = os.path.join(os.getcwd(), ".gpt_settings.json")
        if not os.path.exists(settings_path):
            return
        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                root = json.load(f)
            if not isinstance(root, dict):
                return
            users = root.get("users")
            if isinstance(users, dict) and user_key in users:
                del users[user_key]
                root["users"] = users
                with open(settings_path, "w", encoding="utf-8") as f:
                    json.dump(root, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def delete_current_user(self) -> tuple[bool, str]:
        if not st.session_state.get("authenticated"):
            return False, "Login required."
        user_key = (st.session_state.get("user_key") or "").strip().lower()
        if not user_key:
            return False, "Invalid user session."
        if user_key == SYSTEM_ADMIN_USERNAME:
            return False, "System admin user cannot be deleted."

        users, store = self._ensure_user_store()
        if user_key not in users:
            return False, "User not found."

        del users[user_key]
        if not self._save_user_store(store):
            return False, "Could not delete user."
        self._remove_user_settings(user_key)
        self._remove_user_sessions(user_key)
        return True, "User deleted successfully."

    def _render_auth_css(self):
        st.markdown(
            """
            <style>
            .stApp {
                background:
                    linear-gradient(112deg, rgba(7, 16, 33, 0.46), rgba(14, 53, 95, 0.34)),
                    url('https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&w=2200&q=80');
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }
            .main .block-container {
                max-width: 100% !important;
                width: 100% !important;
                padding-top: 0.8rem !important;
                padding-left: 1rem !important;
                padding-right: 1rem !important;
            }
            .auth-hero {
                background: linear-gradient(120deg, rgba(9, 30, 66, 0.96), rgba(12, 89, 160, 0.88));
                border: 1px solid rgba(148, 163, 184, 0.36);
                border-radius: 14px;
                padding: 0.95rem 1rem;
                color: #f8fafc;
                margin-bottom: 0.6rem;
                box-shadow: 0 12px 28px rgba(8, 20, 43, 0.25);
            }
            .auth-hero h2 {
                margin: 0;
                font-size: 1.35rem;
                line-height: 1.2;
                font-weight: 700;
            }
            .auth-hero p {
                margin: 0.35rem 0 0 0;
                font-size: 0.95rem;
                color: rgba(248, 250, 252, 0.92);
            }
            [data-testid="stForm"] {
                border: 1px solid #d9e2ee !important;
                border-radius: 12px !important;
                padding: 0.7rem 0.78rem !important;
                background: rgba(255, 255, 255, 0.985) !important;
                box-shadow: 0 10px 24px rgba(15, 23, 42, 0.14) !important;
            }
            [data-testid="stTextInput"] > label,
            [data-testid="stRadio"] > label,
            [data-testid="stTextInput"] p,
            [data-testid="stRadio"] p {
                color: #0f172a !important;
                font-size: 0.9rem !important;
                font-weight: 600 !important;
            }
            [data-testid="stTextInput"] input {
                height: 42px !important;
                min-height: 42px !important;
                font-size: 1rem !important;
                border-radius: 9px !important;
                border: 1px solid #c9d7e8 !important;
                background: #ffffff !important;
                color: #0f172a !important;
            }
            [data-testid="stRadio"] label,
            [data-testid="stRadio"] p,
            [data-testid="stRadio"] span {
                color: #f8fafc !important;
                font-weight: 600 !important;
                text-shadow: 0 1px 2px rgba(0, 0, 0, 0.45);
            }
            .stFormSubmitButton > button {
                height: 42px !important;
                min-height: 42px !important;
                font-size: 1rem !important;
                border-radius: 9px !important;
                padding: 0.28rem 0.78rem !important;
            }
            .auth-small-note {
                color: #0f172a;
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid #d9e2ee;
                border-radius: 8px;
                font-size: 0.8rem;
                margin-top: 0.35rem;
                padding: 0.28rem 0.5rem;
            }
            @media (max-width: 900px) {
                .main .block-container {
                    padding-left: 0.6rem !important;
                    padding-right: 0.6rem !important;
                }
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

    def _show_auth_feedback(self):
        feedback = st.session_state.get("auth_feedback")
        if not isinstance(feedback, dict):
            return
        level = feedback.get("level", "info")
        text = feedback.get("text", "")
        if level == "success":
            st.success(text)
        elif level == "error":
            st.error(text)
        elif level == "warning":
            st.warning(text)
        else:
            st.info(text)
        st.session_state.auth_feedback = None

    def _render_login_tab(self):
        self._show_auth_feedback()

        if self.is_rate_limited():
            remaining = 300 - (time.time() - st.session_state.last_attempt_time)
            st.error(f"Too many failed attempts. Please wait {int(remaining/60)} minutes before trying again.")
            return

        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            action_col_login, action_col_forgot = st.columns([1, 1], gap="small")
            with action_col_login:
                submit = st.form_submit_button("Login", use_container_width=True)
            with action_col_forgot:
                forgot_click = st.form_submit_button("Forgot Password", use_container_width=True)

            if forgot_click:
                st.session_state.show_forgot_password = True
                st.rerun()

            if submit:
                if not username or not password:
                    st.error("Please enter both username and password.")
                else:
                    success, role_or_error = self.authenticate_user(username, password)
                    if success:
                        st.session_state.show_forgot_password = False
                        if st.session_state.get("password_reset_required"):
                            st.warning("Temporary password detected. Reset password to continue.")
                        else:
                            st.success("Login successful. Redirecting...")
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error(role_or_error)
                        if st.session_state.login_attempts >= 3:
                            st.warning("Account temporarily locked due to multiple failed attempts.")

        if st.session_state.login_attempts > 0 and not self.is_rate_limited():
            remaining_attempts = 3 - st.session_state.login_attempts
            if remaining_attempts > 0:
                st.warning(f"{remaining_attempts} attempts remaining.")

        if st.session_state.get("show_forgot_password", False):
            st.markdown("<div class='auth-small-note'>Reset password using username/email.</div>", unsafe_allow_html=True)
            with st.form("forgot_password_form"):
                identifier = st.text_input("Username or Email", placeholder="Enter username or email")
                reset_col, cancel_col = st.columns([1, 1], gap="small")
                with reset_col:
                    reset_submit = st.form_submit_button("Generate Temporary Password", use_container_width=True)
                with cancel_col:
                    cancel_submit = st.form_submit_button("Cancel", use_container_width=True)
                if cancel_submit:
                    st.session_state.show_forgot_password = False
                    st.rerun()
                if reset_submit:
                    ok, temp_password, msg = self.issue_temporary_password(identifier)
                    if ok:
                        st.success(f"{msg} Temporary password: {temp_password}")
                        st.info("Login with the temporary password. You will be prompted to set a new password.")
                    else:
                        st.error(msg)

    def _render_signup_tab(self):
        st.session_state.show_forgot_password = False
        with st.form("signup_form"):
            username = st.text_input("Username", placeholder="Choose username")
            first_name = st.text_input("First Name", placeholder="Enter first name")
            last_name = st.text_input("Last Name", placeholder="Enter last name")
            email = st.text_input("Email ID", placeholder="Enter email")
            password = st.text_input("Password", type="password", placeholder="Create password")
            confirm_password = st.text_input("Confirm Password", type="password", placeholder="Confirm password")
            submit = st.form_submit_button("Sign Up", use_container_width=True)

            if submit:
                ok, msg = self.register_user(
                    username=username,
                    first_name=first_name,
                    last_name=last_name,
                    email=email,
                    password=password,
                    confirm_password=confirm_password,
                )
                if ok:
                    st.session_state._switch_to_login = True
                    st.session_state.auth_feedback = {"level": "success", "text": msg}
                    st.rerun()
                else:
                    st.error(msg)

    def login_form(self):
        self._render_auth_css()
        if st.session_state.pop("_switch_to_login", False):
            st.session_state.auth_mode = "Login"
        left_spacer, mid_spacer, auth_col = st.columns([1.35, 0.2, 1.45], gap="small")
        with auth_col:
            st.markdown(
                """
                <div class="auth-hero">
                  <h2>Similarity Answer Matcher</h2>
                  <p>Login or sign up to run AI similarity matching with user-isolated settings.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            mode = st.radio(
                "Access",
                options=["Login", "Sign Up"],
                horizontal=True,
                key="auth_mode",
                label_visibility="collapsed",
            )

            if mode == "Login":
                self._render_login_tab()
            else:
                self._render_signup_tab()

    def force_password_change_form(self):
        self._render_auth_css()
        left_spacer, mid_spacer, auth_col = st.columns([1.35, 0.2, 1.45], gap="small")
        with auth_col:
            st.markdown(
                """
                <div class="auth-hero">
                  <h2>Password Reset Required</h2>
                  <p>Temporary password login detected. Set a new password to continue.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with st.form("force_password_change_form"):
                current_password = st.text_input("Current Temporary Password", type="password")
                new_password = st.text_input("New Password", type="password")
                confirm_password = st.text_input("Confirm New Password", type="password")
                submit = st.form_submit_button("Save New Password", use_container_width=True)

                if submit:
                    ok, msg = self.change_password(current_password, new_password, confirm_password)
                    if ok:
                        st.success(msg)
                        time.sleep(0.6)
                        st.rerun()
                    else:
                        st.error(msg)

    def require_auth(self, required_role: str = "user") -> bool:
        if not st.session_state.authenticated:
            return False
        if required_role == "admin" and st.session_state.user_role != "admin":
            return False
        return True

    def logout(self, rerun: bool = True):
        self._clear_persistent_session()
        st.session_state.authenticated = False
        st.session_state.user_role = None
        st.session_state.username = None
        st.session_state.user_key = None
        st.session_state.auth_token = None
        st.session_state.password_reset_required = False
        st.session_state.session_start_time = None
        st.session_state.last_activity_time = None
        st.session_state.auth_mode = "Login"
        st.session_state.show_forgot_password = False
        st.session_state.show_user_menu_change_password = False
        st.session_state.show_user_menu_delete_confirm = False
        st.session_state.show_user_main_menu = False
        st.session_state.show_user_profile_panel = False
        st.session_state.profile_menu_open = False
        st.session_state.user_theme = "light"

        transient_prefixes = (
            "api_key_",
            "provider_model_select_",
            "provider_model_custom_",
        )
        transient_suffixes = (
            "_resolved_model_name",
        )
        transient_exact = {
            "_saved_api_keys",
            "active_provider_model_name",
            "show_advanced_gpt",
        }

        for key in list(st.session_state.keys()):
            if key in transient_exact:
                del st.session_state[key]
                continue
            if key.startswith(transient_prefixes):
                del st.session_state[key]
                continue
            if key.endswith(transient_suffixes):
                del st.session_state[key]

        if rerun:
            st.rerun()



def get_current_user_key() -> str:
    raw = st.session_state.get("user_key") or st.session_state.get("username") or "default"
    return str(raw).strip().lower() or "default"


def check_authentication() -> bool:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "user_role" not in st.session_state:
        st.session_state.user_role = None
    if "login_attempts" not in st.session_state:
        st.session_state.login_attempts = 0
    if "last_attempt_time" not in st.session_state:
        st.session_state.last_attempt_time = 0
    if "session_start_time" not in st.session_state:
        st.session_state.session_start_time = None
    if "last_activity_time" not in st.session_state:
        st.session_state.last_activity_time = None
    if "password_reset_required" not in st.session_state:
        st.session_state.password_reset_required = False
    if "auth_token" not in st.session_state:
        st.session_state.auth_token = None

    auth_enabled = os.getenv("AUTH_ENABLED", "true").lower() in ("1", "true", "yes")
    if not auth_enabled:
        return True

    auth_manager = AuthManager()

    if not st.session_state.authenticated:
        auth_manager._restore_persistent_session()

    if st.session_state.authenticated:
        if not auth_manager.check_session_timeout():
            return False

        auth_manager.apply_user_theme()

        if st.session_state.get("password_reset_required"):
            auth_manager.force_password_change_form()
            return False

    if not st.session_state.authenticated:
        auth_manager.login_form()
        return False

    return True


def show_user_info():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "show_user_menu_change_password" not in st.session_state:
        st.session_state.show_user_menu_change_password = False
    if "show_user_menu_delete_confirm" not in st.session_state:
        st.session_state.show_user_menu_delete_confirm = False
    if "profile_menu_open" not in st.session_state:
        st.session_state.profile_menu_open = False
    if "user_theme" not in st.session_state:
        st.session_state.user_theme = "light"

    if not st.session_state.authenticated:
        return

    username = st.session_state.get("username") or "user"
    auth_mgr = AuthManager()

    feedback = st.session_state.get("auth_feedback")
    if isinstance(feedback, dict):
        level = feedback.get("level", "info")
        text_msg = feedback.get("text", "")
        if level == "success":
            st.success(text_msg)
        elif level == "error":
            st.error(text_msg)
        elif level == "warning":
            st.warning(text_msg)
        else:
            st.info(text_msg)
        st.session_state.auth_feedback = None

    current_theme = (st.session_state.get("user_theme") or "light").strip().lower()
    safe_user = html.escape(str(username))
    is_dark = current_theme == "dark"

    panel_bg = "#0f172a" if is_dark else "#ffffff"
    panel_text = "#e5e7eb" if is_dark else "#0f172a"
    panel_subtext = "#cbd5e1" if is_dark else "#64748b"
    panel_border = "#334155" if is_dark else "#dbe3ee"
    theme_widget_key = "profile_theme_radio_inline"
    if theme_widget_key not in st.session_state:
        st.session_state[theme_widget_key] = "Dark" if current_theme == "dark" else "Light"
    layout_shift_css = ""
    if st.session_state.get("profile_menu_open", False):
        layout_shift_css = """
        .main .block-container {
            padding-right: 360px !important;
            transition: padding-right 0.18s ease;
        }
        @media (max-width: 1200px) {
            .main .block-container {
                padding-right: 1rem !important;
            }
        }
        """

    st.markdown(
        f"""
        <style>
        .st-key-profile_icon_toggle_btn button {{
            position: fixed !important;
            right: 18px !important;
            bottom: 18px !important;
            z-index: 1400 !important;
            min-height: 46px !important;
            height: 46px !important;
            width: 46px !important;
            border-radius: 999px !important;
            border: 1px solid {panel_border} !important;
            background: {panel_bg} !important;
            color: {panel_text} !important;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.26) !important;
            padding: 0 !important;
            font-size: 20px !important;
            line-height: 1 !important;
        }}
        .profile-floating-panel {{
            position: fixed;
            right: 18px;
            bottom: 74px;
            width: 300px;
            z-index: 1390;
            background: {panel_bg};
            color: {panel_text};
            border: 1px solid {panel_border};
            border-radius: 12px;
            box-shadow: 0 12px 24px rgba(15, 23, 42, 0.24);
            min-height: 238px;
            max-height: calc(100vh - 110px);
            overflow-y: auto;
            padding: 14px;
        }}
        .profile-floating-panel .user-row {{
            font-size: 1.1rem;
            font-weight: 700;
            margin-bottom: 12px;
            color: {panel_text};
        }}
        .profile-floating-panel .theme-label {{
            font-size: 0.92rem;
            color: {panel_subtext};
            margin-bottom: 8px;
        }}
        .st-key-profile_theme_radio_inline,
        [class*="st-key-profile_theme_radio_inline"] {{
            position: fixed !important;
            right: 30px !important;
            bottom: 194px !important;
            width: 268px !important;
            z-index: 1395 !important;
            margin: 0 !important;
            padding: 0 !important;
        }}
        .st-key-profile_theme_radio_inline [data-testid="stRadio"] > div,
        [class*="st-key-profile_theme_radio_inline"] [data-testid="stRadio"] > div {{
            display: flex !important;
            flex-direction: row !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
        }}
        .st-key-profile_theme_radio_inline [data-testid="stRadio"] label,
        [class*="st-key-profile_theme_radio_inline"] [data-testid="stRadio"] label {{
            min-height: 26px !important;
            height: 26px !important;
            padding: 0.14rem 0.5rem !important;
            border-radius: 8px !important;
            border: 1px solid {panel_border} !important;
            background: {"#111827" if is_dark else "#f8fafc"} !important;
            transform: none !important;
            box-shadow: none !important;
        }}
        .st-key-profile_theme_radio_inline div[role="radiogroup"] label:has(input[type="radio"]:checked),
        [class*="st-key-profile_theme_radio_inline"] div[role="radiogroup"] label:has(input[type="radio"]:checked) {{
            min-height: 26px !important;
            height: 26px !important;
            padding: 0.14rem 0.5rem !important;
            border-radius: 8px !important;
            transform: none !important;
            box-shadow: none !important;
            font-weight: 600 !important;
        }}
        .st-key-profile_theme_radio_inline [data-testid="stRadio"] span,
        .st-key-profile_theme_radio_inline [data-testid="stRadio"] p,
        [class*="st-key-profile_theme_radio_inline"] [data-testid="stRadio"] span,
        [class*="st-key-profile_theme_radio_inline"] [data-testid="stRadio"] p {{
            color: {panel_text} !important;
            font-size: 0.88rem !important;
        }}
        .st-key-profile_btn_change,
        .st-key-profile_btn_delete,
        .st-key-profile_btn_logout,
        [class*="st-key-profile_btn_change"],
        [class*="st-key-profile_btn_delete"],
        [class*="st-key-profile_btn_logout"] {{
            position: fixed !important;
            right: 30px !important;
            width: 268px !important;
            z-index: 1395 !important;
            margin: 0 !important;
            padding: 0 !important;
        }}
        .st-key-profile_btn_change, [class*="st-key-profile_btn_change"] {{ bottom: 148px !important; }}
        .st-key-profile_btn_delete, [class*="st-key-profile_btn_delete"] {{ bottom: 114px !important; }}
        .st-key-profile_btn_logout, [class*="st-key-profile_btn_logout"] {{ bottom: 80px !important; }}
        .st-key-profile_btn_change button,
        .st-key-profile_btn_delete button,
        .st-key-profile_btn_logout button,
        [class*="st-key-profile_btn_change"] button,
        [class*="st-key-profile_btn_delete"] button,
        [class*="st-key-profile_btn_logout"] button {{
            width: 100% !important;
            min-height: 30px !important;
            height: 30px !important;
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            justify-content: flex-start !important;
            padding: 2px 0 !important;
            margin: 0 !important;
            color: {panel_text} !important;
            text-decoration: none !important;
            font-size: 0.98rem !important;
        }}
        .st-key-profile_btn_change button p,
        .st-key-profile_btn_delete button p,
        .st-key-profile_btn_logout button p,
        [class*="st-key-profile_btn_change"] button p,
        [class*="st-key-profile_btn_delete"] button p,
        [class*="st-key-profile_btn_logout"] button p {{
            color: {panel_text} !important;
            text-decoration: none !important;
            margin: 0 !important;
        }}
        .st-key-profile_btn_change button:hover p,
        .st-key-profile_btn_delete button:hover p,
        .st-key-profile_btn_logout button:hover p {{
            color: #38bdf8 !important;
        }}
        .st-key-dialog_confirm_delete_user button,
        .st-key-dialog_cancel_delete_user button {{
            border-radius: 8px !important;
            min-height: 36px !important;
            font-weight: 600 !important;
            border: 1px solid {panel_border} !important;
            box-shadow: none !important;
        }}
        .st-key-dialog_confirm_delete_user button {{
            background: #dc2626 !important;
            color: #ffffff !important;
            border-color: #dc2626 !important;
        }}
        .st-key-dialog_cancel_delete_user button {{
            background: {"#1e293b" if is_dark else "#ffffff"} !important;
            color: {panel_text} !important;
        }}
        .st-key-profile_icon_toggle_btn {{
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
        }}
        {layout_shift_css}
        </style>
        """,
        unsafe_allow_html=True,
    )

    if st.button("\U0001F464", key="profile_icon_toggle_btn"):
        st.session_state.profile_menu_open = not st.session_state.get("profile_menu_open", False)
        st.rerun()

    if st.session_state.get("profile_menu_open", False):
        st.markdown(
            f"""
            <div class="profile-floating-panel">
              <div class="user-row">Username: {safe_user}</div>
              <div class="theme-label">Theme</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        selected_theme = st.radio(
            "Theme",
            ["Light", "Dark"],
            key=theme_widget_key,
            horizontal=True,
            label_visibility="collapsed",
        )
        selected_theme_norm = selected_theme.lower()
        if selected_theme_norm != current_theme:
            ok, _ = auth_mgr.set_user_theme(selected_theme_norm)
            if not ok:
                st.session_state.auth_feedback = {"level": "error", "text": "Unable to update theme."}
            st.rerun()

        if st.button("Change Password", key="profile_btn_change", type="tertiary"):
            st.session_state.show_user_menu_change_password = True
            st.session_state.show_user_menu_delete_confirm = False
            st.session_state.profile_menu_open = False
            st.rerun()

        if st.button("Delete Profile", key="profile_btn_delete", type="tertiary"):
            st.session_state.show_user_menu_delete_confirm = True
            st.session_state.show_user_menu_change_password = False
            st.session_state.profile_menu_open = False
            st.rerun()

        if st.button("Logout", key="profile_btn_logout", type="tertiary"):
            st.session_state.profile_menu_open = False
            auth_mgr.logout()

    if st.session_state.get("show_user_menu_change_password", False):
        @st.dialog("Change Password")
        def _change_password_dialog():
            with st.form("dialog_change_password_form", clear_on_submit=True):
                current_password = st.text_input("Current Password", type="password")
                new_password = st.text_input("New Password", type="password")
                confirm_password = st.text_input("Confirm New Password", type="password")
                save_col, cancel_col = st.columns([1, 1], gap="small")
                with save_col:
                    save_click = st.form_submit_button("Save", use_container_width=True)
                with cancel_col:
                    cancel_click = st.form_submit_button("Cancel", use_container_width=True)

                if cancel_click:
                    st.session_state.show_user_menu_change_password = False
                    st.session_state.show_user_menu_delete_confirm = False
                    st.session_state.profile_menu_open = False
                    st.rerun()

                if save_click:
                    ok, msg = auth_mgr.change_password(current_password, new_password, confirm_password)
                    if ok:
                        st.session_state.auth_feedback = {"level": "success", "text": msg}
                        st.session_state.show_user_menu_change_password = False
                        st.session_state.show_user_menu_delete_confirm = False
                        st.session_state.profile_menu_open = False
                        st.rerun()
                    st.error(msg)

        _change_password_dialog()

    if st.session_state.get("show_user_menu_delete_confirm", False):
        @st.dialog("Delete Profile")
        def _delete_user_dialog():
            st.warning("Confirm account deletion. This action is permanent.")
            confirm_col, cancel_col = st.columns([1, 1], gap="small")
            with confirm_col:
                if st.button("Confirm Delete", key="dialog_confirm_delete_user", use_container_width=True):
                    ok, msg = auth_mgr.delete_current_user()
                    if ok:
                        st.session_state.auth_feedback = {"level": "success", "text": msg}
                        st.session_state.show_user_menu_delete_confirm = False
                        st.session_state.show_user_menu_change_password = False
                        st.session_state.profile_menu_open = False
                        auth_mgr.logout()
                    else:
                        st.error(msg)
            with cancel_col:
                if st.button("Cancel", key="dialog_cancel_delete_user", use_container_width=True):
                    st.session_state.show_user_menu_delete_confirm = False
                    st.session_state.show_user_menu_change_password = False
                    st.session_state.profile_menu_open = False
                    st.rerun()

        _delete_user_dialog()
