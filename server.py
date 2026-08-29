from datetime import datetime, timedelta, timezone
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import random
import secrets
import sqlite3
import time
from urllib.parse import parse_qs, urlparse
import requests

# ==========================================
# ⚙️ SECURE SERVER CONFIGURATION
# ==========================================
DATA_DIR = Path.home() / "Documents"
DB_FILE = DATA_DIR / "yosan_cloud.db"

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "").strip()
PBKDF2_ROUNDS = 600_000

# Rate Limiting Parameters
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 600  # 10 Minutes


def get_db():
    conn = sqlite3.connect(str(DB_FILE))
    conn.row_factory = sqlite3.Row
    return conn


def init_cloud_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with get_db() as conn:
        cursor = conn.cursor()

        # 1. Users Table
        cursor.execute("""
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
        cursor.execute("""
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
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS login_attempts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                identifier TEXT NOT NULL,
                attempts INTEGER DEFAULT 1,
                last_attempt REAL NOT NULL,
                locked_until REAL DEFAULT 0
            )
        """)

        # 4. Multi-Tenant Budget Cycles
        cursor.execute("""
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
        cursor.execute("""
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
        conn.commit()


init_cloud_db()


# ==========================================
# 🔒 SECURITY & CRYPTO UTILITIES
# ==========================================
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


def check_rate_limit(identifier: str, cursor) -> tuple[bool, int]:
    now = time.time()
    cursor.execute("SELECT attempts, locked_until FROM login_attempts WHERE identifier = ?", (identifier,))
    row = cursor.fetchone()
    if not row:
        return True, 0

    if row["locked_until"] > now:
        remaining_lockout = int(row["locked_until"] - now)
        return False, remaining_lockout

    return True, 0


def record_failed_attempt(identifier: str, cursor, conn):
    now = time.time()
    cursor.execute("SELECT id, attempts FROM login_attempts WHERE identifier = ?", (identifier,))
    row = cursor.fetchone()

    if not row:
        cursor.execute(
            "INSERT INTO login_attempts (identifier, attempts, last_attempt, locked_until) VALUES (?, 1, ?, 0)",
            (identifier, now),
        )
    else:
        new_attempts = row["attempts"] + 1
        locked_until = now + LOCKOUT_DURATION_SECONDS if new_attempts >= MAX_LOGIN_ATTEMPTS else 0
        cursor.execute(
            "UPDATE login_attempts SET attempts = ?, last_attempt = ?, locked_until = ? WHERE id = ?",
            (new_attempts, now, locked_until, row["id"]),
        )
    conn.commit()


def reset_rate_limit(identifier: str, cursor, conn):
    cursor.execute("DELETE FROM login_attempts WHERE identifier = ?", (identifier,))
    conn.commit()


def send_cloud_email_otp(target_email: str, otp: str, purpose: str) -> bool:
    """Dispatches OTP email via Resend HTTPS API (bypasses all ISP & Cloud SMTP port blocks)."""
    if not RESEND_API_KEY:
        print(f"\n⚠️  [DEV MODE] RESEND_API_KEY not configured. Active OTP code for '{purpose}' is: {otp}\n")
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
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "from": "Yosan Cloud <onboarding@resend.dev>",
        "to": [target_email],
        "subject": f"[{otp}] Your Yosan Cloud Security Code",
        "html": html_content,
    }

    try:
        res = requests.post("https://api.resend.com/emails", json=payload, headers=headers, timeout=10)
        if res.status_code in [200, 201]:
            print(f"📧 [Resend HTTPS] Security code successfully delivered to {target_email}")
            return True
        else:
            print(f"❌ Resend API Error: {res.status_code} - {res.text}")
            return False
    except Exception as e:
        print(f"❌ Email Dispatch Error: {e}")
        return False


def authenticate_request(headers: dict, cursor) -> dict:
    auth_header = headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    token = auth_header.split(" ")[1].strip()

    cursor.execute("""
        SELECT * FROM users 
        WHERE token = ? AND is_verified = 1 AND token_expiry > datetime('now')
    """, (token,))
    return cursor.fetchone()


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
                "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
            })
            return

        if parsed_path == "/api/auth/check-username":
            params = parse_qs(parsed_url.query)
            username = params.get("username", [""])[0].strip()
            if not username:
                self._send_json(400, {"detail": "Username parameter required"})
                return

            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id, is_verified FROM users WHERE LOWER(username) = LOWER(?)",
                    (username,)
                )
                user = cursor.fetchone()
                exists = bool(user and user["is_verified"] == 1)
                self._send_json(200, {"username": username, "exists": exists})
                return

        if parsed_path == "/api/user/profile":
            with get_db() as conn:
                cursor = conn.cursor()
                user = authenticate_request(dict(self.headers), cursor)
                if not user:
                    self._send_json(401, {"detail": "Unauthorized session."})
                    return

                self._send_json(200, {
                    "username": user["username"],
                    "email": user["email"],
                    "created_at": user["created_at"],
                    "is_verified": bool(user["is_verified"]),
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

        with get_db() as conn:
            cursor = conn.cursor()

            if parsed_path == "/api/auth/register":
                username = payload.get("username", "").strip()
                email_addr = payload.get("email", "").strip().lower()
                password = payload.get("password", "")

                if not username or not email_addr or not password:
                    self._send_json(400, {"detail": "All fields are required"})
                    return
                if len(password) < 6:
                    self._send_json(400, {"detail": "Password must be at least 6 characters"})
                    return

                cursor.execute("""
                    SELECT created_at FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND created_at > datetime('now', '-60 seconds')
                """, (email_addr,))
                if cursor.fetchone():
                    self._send_json(429, {"detail": "Please wait 60 seconds before requesting another code."})
                    return

                cursor.execute(
                    "SELECT id, is_verified FROM users WHERE LOWER(username) = LOWER(?) OR LOWER(email) = LOWER(?)",
                    (username, email_addr),
                )
                existing = cursor.fetchone()
                if existing:
                    if existing["is_verified"] == 0:
                        cursor.execute("DELETE FROM users WHERE id = ?", (existing["id"],))
                        conn.commit()
                    else:
                        self._send_json(400, {"detail": "Username or Email already registered."})
                        return

                pw_hash, salt = hash_password(password)
                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    INSERT INTO users (username, email, password_hash, salt, is_verified, created_at)
                    VALUES (?, ?, ?, ?, 0, ?)
                """, (username, email_addr, pw_hash, salt, now_str))
                conn.commit()

                otp = f"{random.randint(100000, 999999)}"
                exp_str = (datetime.now(timezone.utc) + timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    INSERT INTO otps (email, otp_code, purpose, attempts_left, expires_at, created_at)
                    VALUES (?, ?, 'REGISTRATION', 3, ?, ?)
                """, (email_addr, otp, exp_str, now_str))
                conn.commit()

                send_cloud_email_otp(email_addr, otp, "Account Registration")
                self._send_json(200, {"message": f"Verification OTP sent to {email_addr}"})
                return

            elif parsed_path == "/api/auth/verify-registration":
                email_addr = payload.get("email", "").strip().lower()
                otp_code = payload.get("otp_code", "").strip()

                cursor.execute("""
                    SELECT * FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND purpose = 'REGISTRATION' AND expires_at > datetime('now')
                    ORDER BY id DESC LIMIT 1
                """, (email_addr,))
                otp_row = cursor.fetchone()

                if not otp_row:
                    self._send_json(400, {"detail": "No active verification code found or code expired."})
                    return

                if otp_row["attempts_left"] <= 0:
                    self._send_json(403, {"detail": "Too many failed attempts. Code revoked. Register again."})
                    return

                if not secrets.compare_digest(otp_row["otp_code"], otp_code):
                    remaining = otp_row["attempts_left"] - 1
                    cursor.execute("UPDATE otps SET attempts_left = ? WHERE id = ?", (remaining, otp_row["id"]))
                    conn.commit()
                    self._send_json(400, {"detail": f"Incorrect code. {remaining} attempt(s) remaining."})
                    return

                token = secrets.token_hex(32)
                exp_str = (datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    UPDATE users SET is_verified = 1, token = ?, token_expiry = ? WHERE LOWER(email) = LOWER(?)
                """, (token, exp_str, email_addr))
                cursor.execute("DELETE FROM otps WHERE LOWER(email) = LOWER(?)", (email_addr,))
                cursor.execute("SELECT username FROM users WHERE LOWER(email) = LOWER(?)", (email_addr,))
                user = cursor.fetchone()
                conn.commit()

                self._send_json(200, {
                    "token": token,
                    "username": user["username"],
                    "message": "Account verified successfully",
                })
                return

            elif parsed_path == "/api/auth/login":
                username = payload.get("username", "").strip()
                password = payload.get("password", "")

                ip_ok, ip_wait = check_rate_limit(client_ip, cursor)
                user_ok, user_wait = check_rate_limit(f"user_{username.lower()}", cursor)

                if not ip_ok or not user_ok:
                    wait_time = max(ip_wait, user_wait)
                    self._send_json(429, {
                        "detail": f"Account/IP locked due to excessive failed attempts. Try again in {wait_time}s."
                    })
                    return

                cursor.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (username,))
                user = cursor.fetchone()

                if not user or not verify_password(password, user["password_hash"], user["salt"]):
                    record_failed_attempt(client_ip, cursor, conn)
                    record_failed_attempt(f"user_{username.lower()}", cursor, conn)
                    self._send_json(400, {"detail": "Invalid username or password"})
                    return

                if user["is_verified"] == 0:
                    self._send_json(403, {"detail": "Account unverified. Verify via email OTP."})
                    return

                reset_rate_limit(client_ip, cursor, conn)
                reset_rate_limit(f"user_{username.lower()}", cursor, conn)

                token = secrets.token_hex(32)
                exp_str = (datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("UPDATE users SET token = ?, token_expiry = ? WHERE id = ?", (token, exp_str, user["id"]))
                conn.commit()

                self._send_json(200, {
                    "token": token,
                    "username": user["username"],
                    "message": "Login successful",
                })
                return

            elif parsed_path == "/api/auth/logout":
                user = authenticate_request(dict(self.headers), cursor)
                if user:
                    cursor.execute("UPDATE users SET token = NULL, token_expiry = NULL WHERE id = ?", (user["id"],))
                    conn.commit()
                self._send_json(200, {"message": "Session revoked on cloud."})
                return

            elif parsed_path == "/api/auth/forgot-password":
                username = payload.get("username", "").strip()
                email_addr = payload.get("email", "").strip().lower()

                cursor.execute(
                    "SELECT * FROM users WHERE LOWER(username) = LOWER(?) AND LOWER(email) = LOWER(?)",
                    (username, email_addr),
                )
                user = cursor.fetchone()

                if not user:
                    self._send_json(404, {"detail": "No matching user account found."})
                    return

                cursor.execute("""
                    SELECT created_at FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND created_at > datetime('now', '-60 seconds')
                """, (email_addr,))
                if cursor.fetchone():
                    self._send_json(429, {"detail": "Please wait 60 seconds before requesting another code."})
                    return

                otp = f"{random.randint(100000, 999999)}"
                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                exp_str = (datetime.now(timezone.utc) + timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    INSERT INTO otps (email, otp_code, purpose, attempts_left, expires_at, created_at)
                    VALUES (?, ?, 'PASSWORD_RESET', 3, ?, ?)
                """, (email_addr, otp, exp_str, now_str))
                conn.commit()

                send_cloud_email_otp(email_addr, otp, "Password Reset")
                self._send_json(200, {"message": f"Password reset OTP sent to {email_addr}"})
                return

            elif parsed_path == "/api/auth/reset-password":
                email_addr = payload.get("email", "").strip().lower()
                otp_code = payload.get("otp_code", "").strip()
                new_password = payload.get("new_password", "")

                if len(new_password) < 6:
                    self._send_json(400, {"detail": "Password must be at least 6 characters."})
                    return

                cursor.execute("""
                    SELECT * FROM otps 
                    WHERE LOWER(email) = LOWER(?) AND purpose = 'PASSWORD_RESET' AND expires_at > datetime('now')
                    ORDER BY id DESC LIMIT 1
                """, (email_addr,))
                otp_row = cursor.fetchone()

                if not otp_row or otp_row["attempts_left"] <= 0:
                    self._send_json(400, {"detail": "Invalid or expired OTP."})
                    return

                if not secrets.compare_digest(otp_row["otp_code"], otp_code):
                    remaining = otp_row["attempts_left"] - 1
                    cursor.execute("UPDATE otps SET attempts_left = ? WHERE id = ?", (remaining, otp_row["id"]))
                    conn.commit()
                    self._send_json(400, {"detail": f"Incorrect code. {remaining} attempt(s) remaining."})
                    return

                new_hash, salt = hash_password(new_password)
                cursor.execute("""
                    UPDATE users SET password_hash = ?, salt = ?, token = NULL WHERE LOWER(email) = LOWER(?)
                """, (new_hash, salt, email_addr))
                cursor.execute("DELETE FROM otps WHERE LOWER(email) = LOWER(?)", (email_addr,))
                conn.commit()

                self._send_json(200, {"message": "Password updated successfully. Please log in."})
                return

            elif parsed_path == "/api/auth/forgot-username":
                email_addr = payload.get("email", "").strip().lower()
                cursor.execute("SELECT username FROM users WHERE LOWER(email) = LOWER(?) AND is_verified = 1", (email_addr,))
                user = cursor.fetchone()

                if not user:
                    self._send_json(404, {"detail": "No verified account associated with this email address."})
                    return

                send_cloud_email_otp(email_addr, user["username"], "Username Recovery")
                self._send_json(200, {"message": f"Username sent to {email_addr}"})
                return

            elif parsed_path == "/api/user/change-password":
                user = authenticate_request(dict(self.headers), cursor)
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
                cursor.execute("UPDATE users SET password_hash = ?, salt = ? WHERE id = ?", (new_hash, salt, user["id"]))
                conn.commit()

                self._send_json(200, {"message": "Password changed successfully."})
                return

            elif parsed_path == "/api/auth/delete-account":
                user = authenticate_request(dict(self.headers), cursor)
                if not user:
                    self._send_json(401, {"detail": "Unauthorized session."})
                    return

                password = payload.get("password", "")
                if not verify_password(password, user["password_hash"], user["salt"]):
                    self._send_json(400, {"detail": "Incorrect password. Account deletion aborted."})
                    return

                cursor.execute("DELETE FROM users WHERE id = ?", (user["id"],))
                conn.commit()

                self._send_json(200, {"message": "Account and all associated cloud data deleted successfully."})
                return

        self._send_json(404, {"detail": "Resource not found"})


def run_server():
    port = int(os.environ.get("PORT", 8000))
    host = "0.0.0.0"
    
    server_address = (host, port)
    httpd = HTTPServer(server_address, YosanAPIHandler)
    print("\n" + "=" * 65)
    print(f" 🛡️  YOSAN CLOUD SERVER LIVE ON http://{host}:{port}")
    print(f" 📂 Database Path : {DB_FILE}")
    print(" 🔒 Security : PBKDF2 (600k) | Brute-Force Lockouts | Active Revocation")
    print("=" * 65 + "\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 Cloud server stopped.")
        httpd.server_close()


if __name__ == "__main__":
    run_server()
