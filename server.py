from datetime import datetime, timedelta, timezone
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import random
import re
import secrets
import sqlite3
import time
from urllib.parse import parse_qs, urlparse
import requests

# ==========================================
# ⚙️ SECURE SERVER & TURSO CONFIGURATION
# ==========================================
DATA_DIR = Path.home() / "Documents"
LOCAL_DB_FILE = DATA_DIR / "yosan_cloud.db"

BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "").strip()
BREVO_SENDER_EMAIL = os.environ.get("BREVO_SENDER_EMAIL", "").strip()

TURSO_DB_URL = os.environ.get("TURSO_DB_URL", "").strip()
TURSO_AUTH_TOKEN = os.environ.get("TURSO_AUTH_TOKEN", "").strip()

PBKDF2_ROUNDS = 600_000
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 600  # 10 Minutes


# ==========================================
# 🗄️ UNIFIED DATABASE ADAPTER (Turso / SQLite)
# ==========================================
class DatabaseSession:
    def __init__(self):
        self.use_turso = bool(TURSO_DB_URL and TURSO_AUTH_TOKEN)
        self.http_url = ""
        if self.use_turso:
            url = TURSO_DB_URL.replace("libsql://", "https://")
            if not url.endswith("/v2/pipeline"):
                url = f"{url.rstrip('/')}/v2/pipeline"
            self.http_url = url
            self.headers = {
                "Authorization": f"Bearer {TURSO_AUTH_TOKEN}",
                "Content-Type": "application/json",
            }
        else:
            self.conn = sqlite3.connect(str(LOCAL_DB_FILE))
            self.conn.row_factory = sqlite3.Row
            self.cursor = self.conn.cursor()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if not self.use_turso and hasattr(self, "conn"):
            if not exc_type:
                self.conn.commit()
            self.conn.close()

    def execute(self, sql: str, params: tuple = ()) -> list[dict]:
        if not self.use_turso:
            self.cursor.execute(sql, params)
            rows = self.cursor.fetchall()
            return [dict(row) for row in rows] if rows else []

        # Execute on Turso via HTTP Pipeline
        args = []
        for p in params:
            if p is None:
                args.append({"type": "null"})
            elif isinstance(p, (int, float)):
                args.append({"type": "integer" if isinstance(p, int) else "float", "value": p})
            else:
                args.append({"type": "text", "value": str(p)})

        req_body = {
            "requests": [
                {"type": "execute", "stmt": {"sql": sql, "args": args}},
                {"type": "close"}
            ]
        }

        try:
            res = requests.post(self.http_url, json=req_body, headers=self.headers, timeout=10)
            if res.status_code == 200:
                data = res.json()
                exec_res = data.get("results", [{}])[0].get("response", {}).get("result", {})
                cols = [c["name"] for c in exec_res.get("cols", [])]
                rows = exec_res.get("rows", [])
                
                result = []
                for row in rows:
                    row_dict = {}
                    for col_name, val_obj in zip(cols, row):
                        row_dict[col_name] = val_obj.get("value", None)
                    result.append(row_dict)
                return result
            else:
                print(f"❌ Turso Query Error: {res.status_code} - {res.text}")
                return []
        except Exception as e:
            print(f"❌ Turso Connection Error: {e}")
            return []

    def commit(self):
        if not self.use_turso and hasattr(self, "conn"):
            self.conn.commit()


def init_cloud_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with DatabaseSession() as db:
        # 1. Users Table
        db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                is_verified INTEGER DEFAULT 0,
                token TEXT UNIQUE,
                token_expiry TEXT,
                created_at TEXT NOT NULL
            )
        """)

        # 2. Hardened OTP Table
        db.execute("""
            CREATE TABLE IF NOT EXISTS otps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                otp_code TEXT NOT NULL,
                purpose TEXT NOT NULL,
                attempts_left INTEGER DEFAULT 3,
                expires_at TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)

        # 3. Security Audit & Rate Limiting Table
        db.execute("""
            CREATE TABLE IF NOT EXISTS login_attempts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                identifier TEXT NOT NULL,
                attempts INTEGER DEFAULT 1,
                last_attempt REAL NOT NULL,
                locked_until REAL DEFAULT 0
            )
        """)

        # 4. Multi-Tenant Budget Cycles
        db.execute("""
            CREATE TABLE IF NOT EXISTS cycles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                month_code TEXT NOT NULL,
                month_name TEXT NOT NULL,
                total_budget REAL NOT NULL,
                is_active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
                UNIQUE(user_id, month_code)
            )
        """)

        # 5. Multi-Tenant Transactions
        db.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                cycle_id INTEGER NOT NULL,
                branch TEXT NOT NULL,
                description TEXT NOT NULL,
                amount REAL NOT NULL,
                type TEXT DEFAULT 'Expense',
                timestamp TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
                FOREIGN KEY (cycle_id) REFERENCES cycles(id) ON DELETE CASCADE
            )
        """)
        db.commit()


init_cloud_db()


# ==========================================
# 🔒 SECURITY & VALIDATION UTILITIES
# ==========================================
def is_valid_email_syntax(email_addr: str) -> bool:
    """Fast, reliable regex syntax check for email formatting."""
    return bool(re.match(r"^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$", email_addr))


def hash_password(password: str, salt: str = None) -> tuple[str, str]:
    if salt is None:
        salt = secrets.token_hex(16)
    key = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        PBKDF2_ROUNDS,
    )
    return key.hex(), salt


def verify_password(password: str, stored_hash: str, salt: str) -> bool:
    key = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        PBKDF2_ROUNDS,
    )
    return secrets.compare_digest(key.hex(), stored_hash)


def check_rate_limit(identifier: str, db: DatabaseSession) -> tuple[bool, int]:
    now = time.time()
    rows = db.execute("SELECT attempts, locked_until FROM login_attempts WHERE identifier = ?", (identifier,))
    if not rows:
        return True, 0

    row = rows[0]
    locked_until = float(row.get("locked_until") or 0)
    if locked_until > now:
        remaining_lockout = int(locked_until - now)
        return False, remaining_lockout

    return True, 0


def record_failed_attempt(identifier: str, db: DatabaseSession):
    now = time.time()
    rows = db.execute("SELECT id, attempts FROM login_attempts WHERE identifier = ?", (identifier,))

    if not rows:
        db.execute(
            "INSERT INTO login_attempts (identifier, attempts, last_attempt, locked_until) VALUES (?, 1, ?, 0)",
            (identifier, now),
        )
    else:
        row = rows[0]
        new_attempts = int(row.get("attempts", 1)) + 1
        locked_until = now + LOCKOUT_DURATION_SECONDS if new_attempts >= MAX_LOGIN_ATTEMPTS else 0
        db.execute(
            "UPDATE login_attempts SET attempts = ?, last_attempt = ?, locked_until = ? WHERE id = ?",
            (new_attempts, now, locked_until, row["id"]),
        )
    db.commit()


def reset_rate_limit(identifier: str, db: DatabaseSession):
    db.execute("DELETE FROM login_attempts WHERE identifier = ?", (identifier,))
    db.commit()


def send_cloud_email_otp(target_email: str, otp: str, purpose: str) -> bool:
    """Dispatches OTP email via Brevo HTTPS API."""
    if not BREVO_API_KEY or not BREVO_SENDER_EMAIL:
        print(f"\n⚠️  [DEV MODE] Brevo not configured. Active OTP code for '{purpose}' is: {otp}\n")
        return True

    html_content = f"""
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 500px; margin: auto; padding: 25px; border-radius: 12px; background: #0f141c; color: #f5f6fa; border: 1px solid #1e293b;">
        <h2 style="color: #00d2ff; margin-top: 0;">Yosan Budget Cloud</h2>
        <p style="color: #a4b0be; font-size: 15px;">Use the verification code below for <strong>{purpose}</strong>:</p>
        <div style="background: #182232; padding: 18px; border-radius: 8px; text-align: center; margin: 25px 0; border: 1px dashed #00d2ff;">
            <span style="font-size: 34px; font-weight: bold; letter-spacing: 6px; color: #2ecc71;">{otp}</span>
        </div>
        <p style="color: #747d8c; font-size: 13px;">• This code will expire in <strong>5 minutes</strong>.<br>• Maximum of 3 attempts permitted.</p>
        <p style="color: #57606f; font-size: 12px; margin-top: 20px;">If you did not initiate this request, you can safely ignore this email.</p>
    </div>
    """

    headers = {
        "api-key": BREVO_API_KEY,
        "Content-Type": "application/json",
        "accept": "application/json",
    }
    payload = {
        "sender": {"name": "Yosan Cloud", "email": BREVO_SENDER_EMAIL},
        "to": [{"email": target_email}],
        "subject": f"[{otp}] Your Yosan Cloud Security Code",
        "htmlContent": html_content,
    }

    try:
        res = requests.post("https://api.brevo.com/v3/smtp/email", json=payload, headers=headers, timeout=10)
        if res.status_code in [200, 201, 202]:
            print(f"📧 [Brevo HTTPS] Security code successfully delivered to {target_email}")
            return True
        else:
            print(f"❌ Brevo API Error: {res.status_code} - {res.text}")
            return False
    except Exception as e:
        print(f"❌ Email Dispatch Error: {e}")
        return False


def authenticate_request(headers: dict, db: DatabaseSession) -> dict:
    auth_header = headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    token = auth_header.split(" ")[1].strip()

    rows = db.execute("""
        SELECT * FROM users 
        WHERE token = ? AND is_verified = 1 AND token_expiry > datetime('now')
    """, (token,))
    return rows[0] if rows else None


# ==========================================
# 🌐 HARDENED API HANDLER
# ==========================================
class YosanAPIHandler(BaseHTTPRequestHandler):
    def _send_json(self, status_code: int, data: dict):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode("utf-8"))

    def do_OPTIONS(self):
        self._send_json(200, {})

    def do_GET(self):
        parsed_url = urlparse(self.path)
        parsed_path = parsed_url.path

        if parsed_path in ["", "/", "/health", "/ping"]:
            self._send_json(200, {
                "status": "online",
                "service": "Yosan Cloud API",
                "database": "Turso Cloud SQLite" if TURSO_DB_URL else "Local SQLite",
                "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
            })
            return

        with DatabaseSession() as db:
            # 1. Username Availability Check
            if parsed_path == "/api/auth/check-username":
                params = parse_qs(parsed_url.query)
                username = params.get("username", [""])[0].strip()
                if not username:
                    self._send_json(400, {"detail": "Username parameter required"})
                    return

                rows = db.execute(
                    "SELECT id, is_verified FROM users WHERE LOWER(username) = LOWER(?)",
                    (username,)
                )
                exists = bool(rows and int(rows[0].get("is_verified", 0)) == 1)
                self._send_json(200, {"username": username, "exists": exists})
                return

            # 2. Email Availability Check (Taken vs Available)
            if parsed_path == "/api/auth/check-email":
                params = parse_qs(parsed_url.query)
                email_addr = params.get("email", [""])[0].strip().lower()
                if not email_addr:
                    self._send_json(400, {"detail": "Email parameter required"})
                    return

                rows = db.execute(
                    "SELECT id, is_verified FROM users WHERE LOWER(email) = LOWER(?)",
                    (email_addr,)
                )
                exists = bool(rows and int(rows[0].get("is_verified", 0)) == 1)
                self._send_json(200, {
                    "email": email_addr,
                    "exists": exists,
                    "status": "taken" if exists else "available"
                })
                return

            if parsed_path == "/api/user/profile":
                user = authenticate_request(dict(self.headers), db)
                if not user:
                    self._send_json(401, {"detail": "Unauthorized session."})
                    return

                self._send_json(200, {
                    "username": user["username"],
                    "email": user["email"],
                    "created_at": user["created_at"],
                    "is_verified": bool(int(user.get("is_verified", 0))),
                })
                return

        self._send_json(404, {"detail": "Endpoint not found"})

    def do_POST(self):
        parsed_path = urlparse(self.path).path
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)
        client_ip = self.client_address[0]

        try:
            payload = json.loads(body.decode("utf-8")) if body else {}
        except Exception:
            self._send_json(400, {"detail": "Malformed JSON payload"})
            return

        with DatabaseSession() as db:
            # 1. REGISTRATION
            if parsed_path == "/api/auth/register":
                username = payload.get("username", "").strip()
                email_addr = payload.get("email", "").strip().lower()
                password = payload.get("password", "")

                if not username or not email_addr or not password:
                    self._send_json(400, {"detail": "All fields are required"})
                    return
                if not is_valid_email_syntax(email_addr):
                    self._send_json(400, {"detail": "Invalid email address format"})
                    return
                if len(password) < 6:
                    self._send_json(400, {"detail": "Password must be at least 6 characters"})
                    return

                # Rate limit OTP creation (1 per 60s)
                otp_rows = db.execute("""
                    SELECT created_at FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND created_at > datetime('now', '-60 seconds')
                """, (email_addr,))
                if otp_rows:
                    self._send_json(429, {"detail": "Please wait 60 seconds before requesting another code."})
                    return

                existing = db.execute(
                    "SELECT id, is_verified FROM users WHERE LOWER(username) = LOWER(?) OR LOWER(email) = LOWER(?)",
                    (username, email_addr),
                )
                if existing:
                    user_row = existing[0]
                    if int(user_row.get("is_verified", 0)) == 0:
                        db.execute("DELETE FROM users WHERE id = ?", (user_row["id"],))
                        db.commit()
                    else:
                        self._send_json(400, {"detail": "Username or Email already registered."})
                        return

                pw_hash, salt = hash_password(password)
                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                db.execute("""
                    INSERT INTO users (username, email, password_hash, salt, is_verified, created_at)
                    VALUES (?, ?, ?, ?, 0, ?)
                """, (username, email_addr, pw_hash, salt, now_str))
                db.commit()

                otp = f"{random.randint(100000, 999999)}"
                exp_str = (datetime.now(timezone.utc) + timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
                db.execute("""
                    INSERT INTO otps (email, otp_code, purpose, attempts_left, expires_at, created_at)
                    VALUES (?, ?, 'REGISTRATION', 3, ?, ?)
                """, (email_addr, otp, exp_str, now_str))
                db.commit()

                send_cloud_email_otp(email_addr, otp, "Account Registration")
                self._send_json(200, {"message": f"Verification OTP sent to {email_addr}"})
                return

            # 2. VERIFY REGISTRATION (With atomic attempts decrement)
            elif parsed_path == "/api/auth/verify-registration":
                email_addr = payload.get("email", "").strip().lower()
                otp_code = payload.get("otp_code", "").strip()

                otp_rows = db.execute("""
                    SELECT * FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND purpose = 'REGISTRATION' AND expires_at > datetime('now')
                    ORDER BY id DESC LIMIT 1
                """, (email_addr,))

                if not otp_rows:
                    self._send_json(400, {"detail": "No active verification code found or code expired."})
                    return

                otp_row = otp_rows[0]
                current_attempts = int(otp_row.get("attempts_left", 3))

                if current_attempts <= 1:
                    db.execute("DELETE FROM otps WHERE id = ?", (otp_row["id"],))
                    db.commit()
                    self._send_json(403, {"detail": "Too many failed attempts. Code revoked. Register again."})
                    return

                if not secrets.compare_digest(str(otp_row.get("otp_code")), str(otp_code)):
                    # Atomic SQL decrement
                    db.execute("UPDATE otps SET attempts_left = attempts_left - 1 WHERE id = ?", (otp_row["id"],))
                    db.commit()
                    remaining = current_attempts - 1
                    self._send_json(400, {"detail": f"Incorrect code. {remaining} attempt(s) remaining."})
                    return

                token = secrets.token_hex(32)
                exp_str = (datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
                db.execute("""
                    UPDATE users SET is_verified = 1, token = ?, token_expiry = ? WHERE LOWER(email) = LOWER(?)
                """, (token, exp_str, email_addr))
                db.execute("DELETE FROM otps WHERE LOWER(email) = LOWER(?)", (email_addr,))
                
                u_rows = db.execute("SELECT username FROM users WHERE LOWER(email) = LOWER(?)", (email_addr,))
                db.commit()

                username = u_rows[0]["username"] if u_rows else ""
                self._send_json(200, {
                    "token": token,
                    "username": username,
                    "message": "Account verified successfully",
                })
                return

            # 3. LOGIN
            elif parsed_path == "/api/auth/login":
                username = payload.get("username", "").strip()
                password = payload.get("password", "")

                ip_ok, ip_wait = check_rate_limit(client_ip, db)
                user_ok, user_wait = check_rate_limit(f"user_{username.lower()}", db)

                if not ip_ok or not user_ok:
                    wait_time = max(ip_wait, user_wait)
                    self._send_json(429, {
                        "detail": f"Account/IP locked due to excessive failed attempts. Try again in {wait_time}s."
                    })
                    return

                u_rows = db.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (username,))
                if not u_rows:
                    record_failed_attempt(client_ip, db)
                    record_failed_attempt(f"user_{username.lower()}", db)
                    self._send_json(400, {"detail": "Invalid username or password"})
                    return

                user = u_rows[0]
                if not verify_password(password, user["password_hash"], user["salt"]):
                    record_failed_attempt(client_ip, db)
                    record_failed_attempt(f"user_{username.lower()}", db)
                    self._send_json(400, {"detail": "Invalid username or password"})
                    return

                if int(user.get("is_verified", 0)) == 0:
                    self._send_json(403, {"detail": "Account unverified. Verify via email OTP."})
                    return

                reset_rate_limit(client_ip, db)
                reset_rate_limit(f"user_{username.lower()}", db)

                token = secrets.token_hex(32)
                exp_str = (datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
                db.execute("UPDATE users SET token = ?, token_expiry = ? WHERE id = ?", (token, exp_str, user["id"]))
                db.commit()

                self._send_json(200, {
                    "token": token,
                    "username": user["username"],
                    "message": "Login successful",
                })
                return

            # 4. LOGOUT
            elif parsed_path == "/api/auth/logout":
                user = authenticate_request(dict(self.headers), db)
                if user:
                    db.execute("UPDATE users SET token = NULL, token_expiry = NULL WHERE id = ?", (user["id"],))
                    db.commit()
                self._send_json(200, {"message": "Session revoked on cloud."})
                return

            # 5. FORGOT PASSWORD
            elif parsed_path == "/api/auth/forgot-password":
                username = payload.get("username", "").strip()
                email_addr = payload.get("email", "").strip().lower()

                u_rows = db.execute(
                    "SELECT * FROM users WHERE LOWER(username) = LOWER(?) AND LOWER(email) = LOWER(?)",
                    (username, email_addr),
                )
                if not u_rows:
                    self._send_json(404, {"detail": "No matching user account found."})
                    return

                otp_rows = db.execute("""
                    SELECT created_at FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND created_at > datetime('now', '-60 seconds')
                """, (email_addr,))
                if otp_rows:
                    self._send_json(429, {"detail": "Please wait 60 seconds before requesting another code."})
                    return

                otp = f"{random.randint(100000, 999999)}"
                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                exp_str = (datetime.now(timezone.utc) + timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
                db.execute("""
                    INSERT INTO otps (email, otp_code, purpose, attempts_left, expires_at, created_at)
                    VALUES (?, ?, 'PASSWORD_RESET', 3, ?, ?)
                """, (email_addr, otp, exp_str, now_str))
                db.commit()

                send_cloud_email_otp(email_addr, otp, "Password Reset")
                self._send_json(200, {"message": f"Password reset OTP sent to {email_addr}"})
                return

            # 6. RESET PASSWORD (With atomic attempts decrement)
            elif parsed_path == "/api/auth/reset-password":
                email_addr = payload.get("email", "").strip().lower()
                otp_code = payload.get("otp_code", "").strip()
                new_password = payload.get("new_password", "")

                if len(new_password) < 6:
                    self._send_json(400, {"detail": "Password must be at least 6 characters."})
                    return

                otp_rows = db.execute("""
                    SELECT * FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND purpose = 'PASSWORD_RESET' AND expires_at > datetime('now')
                    ORDER BY id DESC LIMIT 1
                """, (email_addr,))

                if not otp_rows:
                    self._send_json(400, {"detail": "Invalid or expired OTP."})
                    return

                otp_row = otp_rows[0]
                current_attempts = int(otp_row.get("attempts_left", 3))

                if current_attempts <= 1:
                    db.execute("DELETE FROM otps WHERE id = ?", (otp_row["id"],))
                    db.commit()
                    self._send_json(403, {"detail": "Too many failed attempts. Code revoked. Request a new password reset."})
                    return

                if not secrets.compare_digest(str(otp_row.get("otp_code")), str(otp_code)):
                    # Atomic SQL decrement
                    db.execute("UPDATE otps SET attempts_left = attempts_left - 1 WHERE id = ?", (otp_row["id"],))
                    db.commit()
                    remaining = current_attempts - 1
                    self._send_json(400, {"detail": f"Incorrect code. {remaining} attempt(s) remaining."})
                    return

                new_hash, salt = hash_password(new_password)
                db.execute("""
                    UPDATE users SET password_hash = ?, salt = ?, token = NULL WHERE LOWER(email) = LOWER(?)
                """, (new_hash, salt, email_addr))
                db.execute("DELETE FROM otps WHERE LOWER(email) = LOWER(?)", (email_addr,))
                db.commit()

                self._send_json(200, {"message": "Password updated successfully. Please log in."})
                return

            # 7. FORGOT USERNAME
            elif parsed_path == "/api/auth/forgot-username":
                email_addr = payload.get("email", "").strip().lower()
                u_rows = db.execute("SELECT username FROM users WHERE LOWER(email) = LOWER(?) AND is_verified = 1", (email_addr,))

                if not u_rows:
                    self._send_json(404, {"detail": "No verified account associated with this email address."})
                    return

                send_cloud_email_otp(email_addr, u_rows[0]["username"], "Username Recovery")
                self._send_json(200, {"message": f"Username sent to {email_addr}"})
                return

            # 8. CHANGE PASSWORD
            elif parsed_path == "/api/user/change-password":
                user = authenticate_request(dict(self.headers), db)
                if not user:
                    self._send_json(401, {"detail": "Unauthorized session."})
                    return

                old_pw = payload.get("old_password", "")
                new_pw = payload.get("new_password", "")

                if not verify_password(old_pw, user["password_hash"], user["salt"]):
                    self._send_json(400, {"detail": "Current password does not match."})
                    return

                if len(new_pw) < 6:
                    self._send_json(400, {"detail": "New password must be at least 6 characters."})
                    return

                new_hash, salt = hash_password(new_pw)
                db.execute("UPDATE users SET password_hash = ?, salt = ? WHERE id = ?", (new_hash, salt, user["id"]))
                db.commit()

                self._send_json(200, {"message": "Password changed successfully."})
                return

            # 9. DELETE ACCOUNT
            elif parsed_path == "/api/auth/delete-account":
                user = authenticate_request(dict(self.headers), db)
                if not user:
                    self._send_json(401, {"detail": "Unauthorized session."})
                    return

                password = payload.get("password", "")
                if not verify_password(password, user["password_hash"], user["salt"]):
                    self._send_json(400, {"detail": "Incorrect password. Account deletion aborted."})
                    return

                db.execute("DELETE FROM users WHERE id = ?", (user["id"],))
                db.commit()

                self._send_json(200, {"message": "Account and all associated cloud data deleted successfully."})
                return

        self._send_json(404, {"detail": "Resource not found"})


def run_server():
    port = int(os.environ.get("PORT", 8000))
    host = "0.0.0.0"
    
    server_address = (host, port)
    httpd = HTTPServer(server_address, YosanAPIHandler)
    db_mode = "Turso Cloud SQLite 🌐" if TURSO_DB_URL else "Local Ephemeral SQLite 📂"
    print("\n" + "=" * 65)
    print(f" 🛡️  YOSAN CLOUD SERVER LIVE ON http://{host}:{port}")
    print(f" 🗄️  Storage Engine : {db_mode}")
    print(" 🔒 Security : PBKDF2 (600k) | Brute-Force Lockouts | Active Revocation")
    print("=" * 65 + "\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 Cloud server stopped.")
        httpd.server_close()


if __name__ == "__main__":
    run_server()
