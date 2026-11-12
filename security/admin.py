from __future__ import annotations
import secrets
from datetime import date, timedelta
from .database import db
from .passwords import hash_access_code

def create_access_code(label: str, max_redemptions: int, free_uses: int, valid_days: int | None, custom_code: str | None = None):
    raw = (custom_code.strip().upper() if custom_code and custom_code.strip()
           else secrets.token_urlsafe(9).replace('-','').replace('_','').upper())
    if len(raw) < 16:
        raise ValueError('يجب أن يكون الرمز المخصص 16 خانة على الأقل.')
    expiry = (date.today()+timedelta(days=valid_days)).isoformat() if valid_days else None
    with db() as c:
        try:
            c.execute('INSERT INTO access_codes(code_hash,label,max_redemptions,free_uses,valid_until) VALUES(?,?,?,?,?)',
                      (hash_access_code(raw),label,max(1,max_redemptions),max(1,free_uses),expiry))
        except Exception as exc:
            if 'UNIQUE' in str(exc).upper():
                raise ValueError('هذا الرمز مستخدم مسبقاً. اختر رمزاً مختلفاً.') from exc
            raise
    return raw, expiry

def grant_free_uses(user_id: int, uses: int):
    with db() as c: c.execute('UPDATE accounts SET free_uses_left=free_uses_left+? WHERE id=?',(max(1,uses),user_id))

def grant_subscription(user_id: int, days: int):
    with db() as c:
        row = c.execute('SELECT subscription_expires_at FROM accounts WHERE id=?',(user_id,)).fetchone()
        if not row: return False
        base = date.today()
        if row['subscription_expires_at']:
            base = max(base,date.fromisoformat(row['subscription_expires_at']))
        c.execute('UPDATE accounts SET subscription_expires_at=? WHERE id=?',((base+timedelta(days=days)).isoformat(),user_id))
        return True

def set_blocked(user_id: int, blocked: bool):
    with db() as c: c.execute('UPDATE accounts SET is_blocked=? WHERE id=?',(int(blocked),user_id))

def disable_code(code_id: int):
    with db() as c: c.execute('UPDATE access_codes SET is_active=0 WHERE id=?',(code_id,))
