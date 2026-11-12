from __future__ import annotations
import os, sqlite3
from pathlib import Path
from contextlib import contextmanager

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = Path(os.environ.get("APP_DB_PATH", str(BASE_DIR / "users_subscriptions.db")))

@contextmanager
def db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def init_db():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS accounts (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          email TEXT NOT NULL UNIQUE COLLATE NOCASE,
          full_name TEXT NOT NULL DEFAULT '',
          password_hash TEXT,
          email_verified INTEGER NOT NULL DEFAULT 0,
          is_blocked INTEGER NOT NULL DEFAULT 0,
          free_uses_left INTEGER NOT NULL DEFAULT 2,
          subscription_expires_at TEXT,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          last_login_at TEXT
        );
        CREATE TABLE IF NOT EXISTS otp_codes (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          email TEXT NOT NULL COLLATE NOCASE,
          code_hash TEXT NOT NULL,
          purpose TEXT NOT NULL,
          expires_at TEXT NOT NULL,
          attempts INTEGER NOT NULL DEFAULT 0,
          used INTEGER NOT NULL DEFAULT 0,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_otp_email_created ON otp_codes(email, created_at);
        CREATE TABLE IF NOT EXISTS login_attempts (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          identity TEXT NOT NULL,
          attempted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          success INTEGER NOT NULL DEFAULT 0
        );
        CREATE INDEX IF NOT EXISTS idx_login_identity_time ON login_attempts(identity, attempted_at);
        CREATE TABLE IF NOT EXISTS access_codes (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          code_hash TEXT NOT NULL UNIQUE,
          label TEXT NOT NULL DEFAULT '',
          max_redemptions INTEGER NOT NULL DEFAULT 1,
          redeemed_count INTEGER NOT NULL DEFAULT 0,
          free_uses INTEGER NOT NULL DEFAULT 2,
          valid_until TEXT,
          is_active INTEGER NOT NULL DEFAULT 1,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS access_code_redemptions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          code_id INTEGER NOT NULL REFERENCES access_codes(id),
          user_id INTEGER NOT NULL REFERENCES accounts(id),
          redeemed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          UNIQUE(code_id, user_id)
        );
        CREATE TABLE IF NOT EXISTS usage_sessions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          user_id INTEGER NOT NULL REFERENCES accounts(id),
          session_token TEXT NOT NULL UNIQUE,
          started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          charged INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS admin_audit (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          admin_identity TEXT NOT NULL,
          action TEXT NOT NULL,
          target TEXT,
          details TEXT,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """)

def get_user_by_email(email: str):
    with db() as c:
        return c.execute("SELECT * FROM accounts WHERE email=?", (email.strip().lower(),)).fetchone()

def create_user(email: str, full_name: str, password_hash: str | None = None, verified: bool = False):
    with db() as c:
        c.execute("INSERT INTO accounts(email,full_name,password_hash,email_verified) VALUES(?,?,?,?)",
                  (email.strip().lower(), full_name.strip(), password_hash, int(verified)))
        return c.execute("SELECT * FROM accounts WHERE email=?", (email.strip().lower(),)).fetchone()

def log_attempt(identity: str, success: bool):
    with db() as c:
        c.execute("INSERT INTO login_attempts(identity,success) VALUES(?,?)", (identity.lower(), int(success)))
        c.execute("DELETE FROM login_attempts WHERE attempted_at < datetime('now','-2 days')")

def is_rate_limited(identity: str, window_minutes: int = 15, max_attempts: int = 8) -> bool:
    with db() as c:
        row = c.execute("SELECT COUNT(*) n FROM login_attempts WHERE identity=? AND success=0 AND attempted_at >= datetime('now',?)",
                         (identity.lower(), f'-{window_minutes} minutes')).fetchone()
        return row['n'] >= max_attempts
