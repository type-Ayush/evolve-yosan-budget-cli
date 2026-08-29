from datetime import datetime
import gc
import getpass
import itertools
import json
import msvcrt
import os
from pathlib import Path
import re
import sys
import threading
import time
import requests

os.system("")

# Cloud Production API URL with local development fallback
API_URL = os.environ.get("YOSAN_API_URL", "https://evolve-yosan-budget-cli.onrender.com/api").rstrip("/")
SESSION_FILE = Path.home() / "Documents" / ".yosan_session.json"


# ==========================================
# 🎨 COLOR PALETTE
# ==========================================
class C:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    CYAN    = "\033[38;2;0;210;255m"
    GREEN   = "\033[38;2;46;204;113m"
    RED     = "\033[38;2;231;76;60m"
    YELLOW  = "\033[38;2;241;196;15m"
    PURPLE  = "\033[38;2;165;94;234m"
    GRAY    = "\033[38;2;127;143;166m"
    WHITE   = "\033[38;2;245;246;250m"


# ==========================================
# ⏳ ANIMATED SPINNER
# ==========================================
class Spinner:
    """Threaded CLI spinner animation for active network dispatches."""
    def __init__(self, message="Sending OTP..."):
        self.message = message
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._spin, daemon=True)

    def _spin(self):
        spinner_cycle = itertools.cycle(["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"])
        while not self._stop_event.is_set():
            sys.stdout.write(f"\r{C.CYAN}{next(spinner_cycle)}{C.RESET} {C.YELLOW}{self.message}{C.RESET}   ")
            sys.stdout.flush()
            time.sleep(0.08)
        sys.stdout.write("\r\033[K")
        sys.stdout.flush()

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._stop_event.set()
        self._thread.join()


# ==========================================
# 🔍 CLOUD VALIDATION UTILITIES
# ==========================================
def check_cloud_username(username: str) -> bool:
    try:
        res = requests.get(
            f"{API_URL}/auth/check-username",
            params={"username": username},
            timeout=4,
        )
        if res.status_code == 200:
            return res.json().get("exists", False)
    except Exception:
        pass
    return False


def check_cloud_email(email_addr: str) -> bool:
    try:
        res = requests.get(
            f"{API_URL}/auth/check-email",
            params={"email": email_addr},
            timeout=4,
        )
        if res.status_code == 200:
            return res.json().get("exists", False)
    except Exception:
        pass
    return False


def is_valid_email_syntax(email: str) -> bool:
    return bool(re.match(r"^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$", email))


# ==========================================
# ⌨️ INTERACTIVE LIVE INPUT HANDLERS
# ==========================================
def live_username_input(prompt: str, check_mode: str = "must_exist") -> str:
    buffer = []
    last_keystroke_time = time.time()
    checked_str = None
    is_valid = False
    badge = ""

    colored_prompt = f"{C.CYAN}{prompt}{C.RESET}"
    sys.stdout.write(f"\r{colored_prompt}\033[K")
    sys.stdout.flush()

    while True:
        current_text = "".join(buffer)

        if (
            current_text
            and current_text != checked_str
            and (time.time() - last_keystroke_time >= 0.8)
        ):
            if current_text.lower() in ["b", "back"]:
                badge = ""
                checked_str = current_text
                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[K")
                sys.stdout.flush()
            else:
                frames = [".  ", ".. ", "..."]
                for frame in frames:
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[s {C.GRAY}{frame}{C.RESET}\033[K\033[u")
                    sys.stdout.flush()
                    time.sleep(0.10)
                    if msvcrt.kbhit():
                        break

                if not msvcrt.kbhit():
                    exists = check_cloud_username(current_text)
                    if check_mode == "must_exist":
                        is_valid = exists
                        badge = f"{C.GREEN}✔ [Verified]{C.RESET}" if exists else f"{C.RED}✖ [Not Found]{C.RESET}"
                    elif check_mode == "must_be_available":
                        is_valid = not exists
                        badge = f"{C.GREEN}✔ [Available]{C.RESET}" if not exists else f"{C.RED}✖ [Taken]{C.RESET}"

                    checked_str = current_text
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[s {badge}\033[K\033[u")
                    sys.stdout.flush()

        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            last_keystroke_time = time.time()

            if ch in ("\r", "\n"):
                if not current_text:
                    continue

                if current_text.lower() in ["b", "back"]:
                    sys.stdout.write("\n")
                    return "BACK"

                if checked_str != current_text:
                    exists = check_cloud_username(current_text)
                    if check_mode == "must_exist":
                        is_valid = exists
                        badge = f"{C.GREEN}✔ [Verified]{C.RESET}" if exists else f"{C.RED}✖ [Not Found]{C.RESET}"
                    elif check_mode == "must_be_available":
                        is_valid = not exists
                        badge = f"{C.GREEN}✔ [Available]{C.RESET}" if not exists else f"{C.RED}✖ [Taken]{C.RESET}"
                    checked_str = current_text

                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET} {badge}\033[K\n")
                sys.stdout.flush()

                if is_valid:
                    return current_text
                else:
                    if check_mode == "must_exist":
                        print(f"  {C.RED}└─ ❌ Username '{current_text}' does not exist.{C.GRAY} (Type 'b' to go back){C.RESET}")
                    else:
                        print(f"  {C.RED}└─ ❌ Username '{current_text}' is already taken.{C.GRAY} (Type 'b' to go back){C.RESET}")

                    buffer = []
                    checked_str = None
                    is_valid = False
                    badge = ""
                    sys.stdout.write(f"\r{colored_prompt}\033[K")
                    sys.stdout.flush()

            elif ch in ("\x08", "\b"):
                if buffer:
                    buffer.pop()
                    checked_str = None
                    is_valid = False
                    badge = ""
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{''.join(buffer)}{C.RESET}\033[K")
                    sys.stdout.flush()

            elif ch == "\x03":
                sys.exit(0)

            elif ch.isprintable():
                buffer.append(ch)
                checked_str = None
                is_valid = False
                badge = ""
                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{''.join(buffer)}{C.RESET}\033[K")
                sys.stdout.flush()

        time.sleep(0.02)


def live_email_input(prompt: str, check_mode: str = "must_be_available") -> str:
    """Provides real-time interactive email syntax and availability validation."""
    buffer = []
    last_keystroke_time = time.time()
    checked_str = None
    is_valid = False
    badge = ""

    colored_prompt = f"{C.CYAN}{prompt}{C.RESET}"
    sys.stdout.write(f"\r{colored_prompt}\033[K")
    sys.stdout.flush()

    while True:
        current_text = "".join(buffer).strip().lower()

        if (
            current_text
            and current_text != checked_str
            and (time.time() - last_keystroke_time >= 0.8)
        ):
            if current_text in ["b", "back"]:
                badge = ""
                checked_str = current_text
                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[K")
                sys.stdout.flush()
            elif not is_valid_email_syntax(current_text):
                is_valid = False
                badge = f"{C.YELLOW}⚠ [Invalid Format]{C.RESET}"
                checked_str = current_text
                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[s {badge}\033[K\033[u")
                sys.stdout.flush()
            else:
                frames = [".  ", ".. ", "..."]
                for frame in frames:
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[s {C.GRAY}{frame}{C.RESET}\033[K\033[u")
                    sys.stdout.flush()
                    time.sleep(0.10)
                    if msvcrt.kbhit():
                        break

                if not msvcrt.kbhit():
                    exists = check_cloud_email(current_text)
                    if check_mode == "must_be_available":
                        is_valid = not exists
                        badge = f"{C.GREEN}✔ [Available]{C.RESET}" if not exists else f"{C.RED}✖ [Already Registered]{C.RESET}"
                    elif check_mode == "must_exist":
                        is_valid = exists
                        badge = f"{C.GREEN}✔ [Found]{C.RESET}" if exists else f"{C.RED}✖ [Not Registered]{C.RESET}"

                    checked_str = current_text
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET}\033[s {badge}\033[K\033[u")
                    sys.stdout.flush()

        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            last_keystroke_time = time.time()

            if ch in ("\r", "\n"):
                if not current_text:
                    continue

                if current_text in ["b", "back"]:
                    sys.stdout.write("\n")
                    return "BACK"

                if not is_valid_email_syntax(current_text):
                    print(f"\n  {C.RED}└─ ❌ Please enter a valid email format (e.g. name@domain.com){C.RESET}")
                    buffer = []
                    checked_str = None
                    is_valid = False
                    badge = ""
                    sys.stdout.write(f"\r{colored_prompt}\033[K")
                    sys.stdout.flush()
                    continue

                if checked_str != current_text:
                    exists = check_cloud_email(current_text)
                    if check_mode == "must_be_available":
                        is_valid = not exists
                        badge = f"{C.GREEN}✔ [Available]{C.RESET}" if not exists else f"{C.RED}✖ [Already Registered]{C.RESET}"
                    elif check_mode == "must_exist":
                        is_valid = exists
                        badge = f"{C.GREEN}✔ [Found]{C.RESET}" if exists else f"{C.RED}✖ [Not Registered]{C.RESET}"
                    checked_str = current_text

                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{current_text}{C.RESET} {badge}\033[K\n")
                sys.stdout.flush()

                if is_valid:
                    return current_text
                else:
                    if check_mode == "must_be_available":
                        print(f"  {C.RED}└─ ❌ Email '{current_text}' is already registered.{C.GRAY} (Type 'b' to go back){C.RESET}")
                    else:
                        print(f"  {C.RED}└─ ❌ No account found for email '{current_text}'.{C.GRAY} (Type 'b' to go back){C.RESET}")

                    buffer = []
                    checked_str = None
                    is_valid = False
                    badge = ""
                    sys.stdout.write(f"\r{colored_prompt}\033[K")
                    sys.stdout.flush()

            elif ch in ("\x08", "\b"):
                if buffer:
                    buffer.pop()
                    checked_str = None
                    is_valid = False
                    badge = ""
                    sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{''.join(buffer)}{C.RESET}\033[K")
                    sys.stdout.flush()

            elif ch == "\x03":
                sys.exit(0)

            elif ch.isprintable():
                buffer.append(ch)
                checked_str = None
                is_valid = False
                badge = ""
                sys.stdout.write(f"\r{colored_prompt}{C.WHITE}{''.join(buffer)}{C.RESET}\033[K")
                sys.stdout.flush()

        time.sleep(0.02)


# ==========================================
# 💾 SESSION MANAGEMENT
# ==========================================
def get_session_data() -> dict:
    if SESSION_FILE.exists():
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def get_token() -> str:
    return get_session_data().get("token", "")


def get_current_user() -> str:
    return get_session_data().get("username", "default")


def save_session(token: str, username: str):
    SESSION_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump({"token": token, "username": username}, f)


def clear_session():
    """Revokes the token on the cloud server and deletes the local session file."""
    token = get_token()
    if token:
        try:
            requests.post(
                f"{API_URL}/auth/logout",
                headers={"Authorization": f"Bearer {token}"},
                timeout=3,
            )
        except Exception:
            pass

    if SESSION_FILE.exists():
        SESSION_FILE.unlink()

    print(f"\n{C.YELLOW}🔒 Logged out of Yosan Cloud. Session cleared.{C.RESET}\n")


# ==========================================
# 🚀 AUTHENTICATION FLOWS
# ==========================================
def register_flow() -> bool:
    print(f"\n{C.CYAN}╔════════════════════════════════════════════════════════╗")
    print(f"║                {C.BOLD}{C.WHITE}YOSAN CLOUD REGISTRATION{C.RESET}{C.CYAN}                ║")
    print(f"║       {C.GRAY}(Type \"b\" or \"back\" at any point to cancel){C.RESET}{C.CYAN}       ║")
    print(f"╚════════════════════════════════════════════════════════╝{C.RESET}")

    username = live_username_input("Enter Desired Username: ", check_mode="must_be_available")
    if username == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    email = live_email_input("Enter Email Address: ", check_mode="must_be_available")
    if email == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    password = getpass.getpass(f"{C.CYAN}Enter Password: {C.RESET}")
    if password in ["b", "back"]:
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    confirm_pw = getpass.getpass(f"{C.CYAN}Confirm Password: {C.RESET}")
    if confirm_pw in ["b", "back"]:
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    if password != confirm_pw:
        print(f"{C.RED}❌ Passwords do not match.{C.RESET}")
        return False

    try:
        with Spinner("Dispatched OTP request... Generating code"):
            res = requests.post(
                f"{API_URL}/auth/register",
                json={"username": username, "email": email, "password": password},
                timeout=15,
            )

        if res.status_code != 200:
            print(f"{C.RED}❌ {res.json().get('detail', 'Registration failed')}{C.RESET}")
            return False

        print(f"\n{C.GREEN}📨 {res.json()['message']}{C.RESET}")
        otp = input(f"{C.YELLOW}Enter 6-digit OTP code sent to your inbox: {C.RESET}").strip()
        if otp in ["b", "back"]:
            print(f"{C.GRAY}↩ Registration aborted.{C.RESET}")
            return False

        with Spinner("Verifying security token..."):
            verify_res = requests.post(
                f"{API_URL}/auth/verify-registration",
                json={"email": email, "otp_code": otp},
                timeout=10,
            )

        if verify_res.status_code == 200:
            data = verify_res.json()
            save_session(data["token"], data["username"])
            print(f"\n{C.GREEN}{C.BOLD}🎉 Welcome to Yosan Cloud, {data['username']}!{C.RESET}\n")
            return True
        else:
            print(f"{C.RED}❌ Verification failed: {verify_res.json().get('detail', 'Invalid OTP')}{C.RESET}")
            return False
    except requests.exceptions.ConnectionError:
        print(f"{C.RED}❌ Could not connect to Yosan Cloud Server.{C.RESET}")
        return False


def login_flow() -> bool:
    print(f"\n{C.CYAN}── {C.BOLD}Log In{C.RESET}{C.CYAN} {C.GRAY}(Type 'b' or 'back' to return){C.CYAN} ──{C.RESET}")
    username = live_username_input("Username: ", check_mode="must_exist")
    if username == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    password = getpass.getpass(f"{C.CYAN}Password: {C.RESET}")
    if password in ["b", "back"]:
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return False

    try:
        with Spinner("Authenticating credentials..."):
            res = requests.post(
                f"{API_URL}/auth/login",
                json={"username": username, "password": password},
                timeout=10,
            )

        if res.status_code == 200:
            data = res.json()
            save_session(data["token"], data["username"])
            print(f"\n{C.GREEN}{C.BOLD}✅ Access Granted. Welcome back, {data['username']}!{C.RESET}\n")
            return True
        elif res.status_code == 429:
            print(f"{C.RED}🔒 {res.json().get('detail')}{C.RESET}")
            return False
        else:
            print(f"{C.RED}❌ Login failed: {res.json().get('detail')}{C.RESET}")
            return False
    except requests.exceptions.ConnectionError:
        print(f"{C.RED}❌ Could not reach Yosan Cloud Server.{C.RESET}")
        return False


def forgot_password_flow():
    print(f"\n{C.PURPLE}═══════════════════════════════════════════════════════")
    print(f"         {C.BOLD}{C.WHITE}YOSAN CLOUD — PASSWORD RECOVERY{C.RESET}{C.PURPLE}")
    print(f"         {C.GRAY}(Type 'b' or 'back' to return){C.RESET}{C.PURPLE}")
    print(f"═══════════════════════════════════════════════════════{C.RESET}")

    username = live_username_input("Enter Username: ", check_mode="must_exist")
    if username == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return

    email = live_email_input("Enter Registered Email: ", check_mode="must_exist")
    if email == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return

    try:
        with Spinner("Sending password recovery OTP..."):
            res = requests.post(
                f"{API_URL}/auth/forgot-password",
                json={"username": username, "email": email},
                timeout=15,
            )

        if res.status_code != 200:
            print(f"{C.RED}❌ {res.json().get('detail')}{C.RESET}")
            return

        print(f"\n{C.GREEN}📨 {res.json()['message']}{C.RESET}")
        otp = input(f"{C.YELLOW}Enter 6-digit OTP from email: {C.RESET}").strip()
        if otp in ["b", "back"]:
            return

        new_pw = getpass.getpass(f"{C.CYAN}Enter New Password: {C.RESET}")
        if new_pw in ["b", "back"]:
            return

        confirm_pw = getpass.getpass(f"{C.CYAN}Confirm New Password: {C.RESET}")
        if new_pw != confirm_pw:
            print(f"{C.RED}❌ Passwords do not match.{C.RESET}")
            return

        with Spinner("Updating cloud security credentials..."):
            reset_res = requests.post(
                f"{API_URL}/auth/reset-password",
                json={"email": email, "otp_code": otp, "new_password": new_pw},
                timeout=10,
            )

        if reset_res.status_code == 200:
            print(f"\n{C.GREEN}🎉 {reset_res.json()['message']}{C.RESET}\n")
        else:
            print(f"{C.RED}❌ {reset_res.json().get('detail')}{C.RESET}")
    except requests.exceptions.ConnectionError:
        print(f"{C.RED}❌ Cloud server unreachable.{C.RESET}")


def forgot_username_flow():
    print(f"\n{C.PURPLE}═══════════════════════════════════════════════════════")
    print(f"         {C.BOLD}{C.WHITE}YOSAN CLOUD — USERNAME RECOVERY{C.RESET}{C.PURPLE}")
    print(f"         {C.GRAY}(Type 'b' or 'back' to return){C.RESET}{C.PURPLE}")
    print(f"═══════════════════════════════════════════════════════{C.RESET}")

    email = live_email_input("Enter Registered Email Address: ", check_mode="must_exist")
    if email == "BACK":
        print(f"{C.GRAY}↩ Returning to menu...{C.RESET}")
        return

    try:
        with Spinner("Recovering account username..."):
            res = requests.post(
                f"{API_URL}/auth/forgot-username",
                json={"email": email},
                timeout=15,
            )

        if res.status_code == 200:
            print(f"\n{C.GREEN}📨 {res.json().get('message')}{C.RESET}\n")
        else:
            print(f"\n{C.RED}❌ {res.json().get('detail')}{C.RESET}\n")
    except requests.exceptions.ConnectionError:
        print(f"{C.RED}❌ Cloud server unreachable.{C.RESET}")


def delete_account_flow():
    token = get_token()
    uname = get_current_user()
    if not token:
        print(f"\n{C.RED}❌ You must be logged in to delete an account.{C.RESET}\n")
        return

    print(f"\n{C.RED}!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print(f" ⚠️  WARNING: PERMANENT ACCOUNT DELETION FOR [{uname.upper()}]")
    print(f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!{C.RESET}")
    print(f"{C.YELLOW}This will permanently delete your credentials, cloud records, and local ledger.{C.RESET}")
    confirm = input(f"Type '{C.RED}DELETE{C.RESET}' to confirm (or 'b' to go back): ").strip()

    if confirm.lower() in ["b", "back"] or confirm != "DELETE":
        print(f"{C.GRAY}❌ Deletion canceled.{C.RESET}\n")
        return

    password = getpass.getpass("Enter your current password to confirm: ")
    if password.lower() in ["b", "back"]:
        print(f"{C.GRAY}❌ Deletion canceled.{C.RESET}\n")
        return

    try:
        headers = {"Authorization": f"Bearer {token}"}
        with Spinner("Wiping account and cloud records..."):
            res = requests.post(
                f"{API_URL}/auth/delete-account",
                json={"password": password},
                headers=headers,
                timeout=10,
            )

        if res.status_code == 200:
            clear_session()
            gc.collect()
            time.sleep(0.3)

            data_dir = Path.home() / "Documents"
            safe_name = "".join(c for c in uname if c.isalnum() or c in ("-", "_")).lower()

            for filename in [f"yosan_{safe_name}.db", f"budget_book_{safe_name}.xlsx"]:
                file_target = data_dir / filename
                if file_target.exists():
                    try:
                        file_target.unlink(missing_ok=True)
                    except Exception:
                        try:
                            with open(file_target, "w") as f:
                                f.truncate(0)
                        except Exception:
                            pass

            print(f"{C.GREEN}✅ Account [{uname}] and local ledger successfully deleted.{C.RESET}\n")
            sys.exit(0)
        else:
            print(f"{C.RED}❌ {res.json().get('detail')}{C.RESET}\n")
    except requests.exceptions.ConnectionError:
        print(f"{C.RED}❌ Cloud server unreachable.{C.RESET}")


def show_profile_view():
    token = get_token()
    if not token:
        print(f"\n{C.RED}❌ You must be logged in to view your profile.{C.RESET}\n")
        return

    headers = {"Authorization": f"Bearer {token}"}
    try:
        with Spinner("Fetching user profile..."):
            res = requests.get(f"{API_URL}/user/profile", headers=headers, timeout=10)

        if res.status_code != 200:
            print(f"\n{C.RED}❌ Failed to fetch profile: {res.json().get('detail', 'Session error')}{C.RESET}")
            return
        user_info = res.json()
    except requests.exceptions.ConnectionError:
        print(f"\n{C.RED}❌ Could not connect to Yosan Cloud Server.{C.RESET}")
        return

    safe_name = "".join(c for c in user_info["username"] if c.isalnum() or c in ("-", "_")).lower()
    local_db = Path.home() / "Documents" / f"yosan_{safe_name}.db"
    local_xl = Path.home() / "Documents" / f"budget_book_{safe_name}.xlsx"

    while True:
        print(f"\n{C.CYAN}╔══════════════════════════════════════════════════════════╗")
        print(f"║                {C.BOLD}{C.WHITE}YOSAN USER ACCOUNT PROFILE{C.RESET}{C.CYAN}                ║")
        print(f"╚══════════════════════════════════════════════════════════╝{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Username{C.RESET}          : {C.GREEN}{user_info['username']}{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Registered Email{C.RESET}  : {C.WHITE}{user_info['email']}{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Account Created{C.RESET}   : {C.GRAY}{user_info['created_at']}{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Cloud Status{C.RESET}      : {C.GREEN if user_info['is_verified'] else C.RED}{'Verified ✅' if user_info['is_verified'] else 'Unverified ❌'}{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Local DB Ledger{C.RESET}   : {C.YELLOW}{local_db.name}{C.RESET}")
        print(f" {C.CYAN}•{C.RESET} {C.BOLD}Local Excel Book{C.RESET}  : {C.YELLOW}{local_xl.name}{C.RESET}")
        print(f"{C.GRAY}────────────────────────────────────────────────────────────{C.RESET}")
        print(f"  {C.CYAN}[1]{C.RESET} Change Password")
        print(f"  {C.CYAN}[2]{C.RESET} Log Out")
        print(f"  {C.CYAN}[3]{C.RESET} {C.RED}Delete Account{C.RESET}")
        print(f"  {C.CYAN}[4]{C.RESET} Back to CLI {C.GRAY}(or 'b'){C.RESET}\n")

        choice = input(f"{C.CYAN}Select an option (1-4) [Default 4]: {C.RESET}").strip()

        if choice == "1":
            old_pw = getpass.getpass("Enter current password: ")
            if old_pw.lower() in ["b", "back"]:
                continue
            new_pw = getpass.getpass("Enter new password: ")
            if new_pw.lower() in ["b", "back"]:
                continue
            confirm_pw = getpass.getpass("Confirm new password: ")

            if new_pw != confirm_pw:
                print(f"{C.RED}❌ Passwords do not match.{C.RESET}")
                continue

            try:
                with Spinner("Updating cloud password..."):
                    chg_res = requests.post(
                        f"{API_URL}/user/change-password",
                        headers=headers,
                        json={"old_password": old_pw, "new_password": new_pw},
                        timeout=10,
                    )
                if chg_res.status_code == 200:
                    print(f"\n{C.GREEN}🎉 {chg_res.json()['message']}{C.RESET}\n")
                else:
                    print(f"\n{C.RED}❌ {chg_res.json().get('detail')}{C.RESET}\n")
            except requests.exceptions.ConnectionError:
                print(f"{C.RED}❌ Cloud server unreachable.{C.RESET}")
        elif choice == "2":
            clear_session()
            sys.exit(0)
        elif choice == "3":
            delete_account_flow()
            return
        elif choice in ["4", "", "b", "back", "q", "exit"]:
            break


def require_login() -> bool:
    if get_token():
        return True

    while True:
        print(f"\n{C.CYAN}╔════════════════════════════════════════════════╗")
        print(f"║      {C.BOLD}{C.WHITE}YOSAN CLOUD ACCESS — AUTHENTICATION{C.RESET}{C.CYAN}       ║")
        print(f"╚════════════════════════════════════════════════╝{C.RESET}")
        print(f"  {C.CYAN}[1]{C.RESET} Log In")
        print(f"  {C.CYAN}[2]{C.RESET} Sign Up {C.GRAY}(New Account + Email OTP){C.RESET}")
        print(f"  {C.CYAN}[3]{C.RESET} Forgot Password {C.GRAY}(Email OTP Recovery){C.RESET}")
        print(f"  {C.CYAN}[4]{C.RESET} Forgot Username {C.GRAY}(Send to Email){C.RESET}")
        print(f"  {C.CYAN}[5]{C.RESET} Exit\n")

        choice = input(f"{C.CYAN}Select option (1-5) [Default 1]: {C.RESET}").strip()

        if choice in ["", "1"]:
            if login_flow():
                return True
        elif choice == "2":
            if register_flow():
                return True
        elif choice == "3":
            forgot_password_flow()
        elif choice == "4":
            forgot_username_flow()
        elif choice in ["5", "q", "exit"]:
            sys.exit(0)
