from __future__ import annotations
from datetime import date
import secrets
from .database import db
from .passwords import verify_access_code

FREE_USES = 2

def ensure_user(email: str, full_name: str = '', password_hash: str | None = None, verified: bool = False):
    with db() as c:
        row = c.execute('SELECT * FROM accounts WHERE email=?',(email.lower(),)).fetchone()
        if row:
            if full_name and not row['full_name']:
                c.execute('UPDATE accounts SET full_name=? WHERE id=?',(full_name,row['id']))
            if verified and not row['email_verified']:
                c.execute('UPDATE accounts SET email_verified=1 WHERE id=?',(row['id'],))
            return c.execute('SELECT * FROM accounts WHERE id=?',(row['id'],)).fetchone()
        c.execute('INSERT INTO accounts(email,full_name,password_hash,email_verified,free_uses_left) VALUES(?,?,?,?,?)',
                  (email.lower(),full_name,password_hash,int(verified),FREE_USES))
        return c.execute('SELECT * FROM accounts WHERE email=?',(email.lower(),)).fetchone()

def access_status(user_id: int):
    with db() as c:
        u = c.execute('SELECT * FROM accounts WHERE id=?',(user_id,)).fetchone()
        if not u: return False,'الحساب غير موجود',0
        if u['is_blocked']: return False,'الحساب موقوف من الإدارة.',0
        expiry = u['subscription_expires_at']
        if expiry and date.fromisoformat(expiry) >= date.today(): return True,f'الاشتراك ساري حتى {expiry}',u['free_uses_left']
        if u['free_uses_left'] > 0: return True,f'المتبقي {u["free_uses_left"]} استخدام مجاني',u['free_uses_left']
        return False,'انتهت الاستخدامات المجانية. يرجى تفعيل اشتراك.',0

def start_usage(user_id: int) -> tuple[bool,str]:
    """Charge once when user explicitly starts a new audit session."""
    token = secrets.token_urlsafe(24)
    with db() as c:
        u = c.execute('SELECT * FROM accounts WHERE id=?',(user_id,)).fetchone()
        if not u or u['is_blocked']: return False,'الحساب غير متاح.'
        expiry = u['subscription_expires_at']
        paid = bool(expiry and date.fromisoformat(expiry) >= date.today())
        if not paid:
            cur = c.execute('UPDATE accounts SET free_uses_left=free_uses_left-1 WHERE id=? AND free_uses_left>0 AND is_blocked=0',(user_id,))
            if cur.rowcount != 1: return False,'لا توجد استخدامات مجانية متبقية.'
        c.execute('INSERT INTO usage_sessions(user_id,session_token,charged) VALUES(?,?,?)',(user_id,token,int(not paid)))
    return True,token

def redeem_access_code(user_id: int, raw_code: str):
    with db() as c:
        candidates = c.execute('SELECT * FROM access_codes WHERE is_active=1 AND redeemed_count < max_redemptions').fetchall()
        code = next((item for item in candidates if verify_access_code(raw_code,item['code_hash'])),None)
        if not code: return False,'رمز غير صحيح أو ملغى.'
        if code['valid_until'] and code['valid_until'] < date.today().isoformat(): return False,'انتهت صلاحية الرمز.'
        if code['redeemed_count'] >= code['max_redemptions']: return False,'اكتمل عدد المستفيدين من هذا الرمز.'
        try:
            c.execute('INSERT INTO access_code_redemptions(code_id,user_id) VALUES(?,?)',(code['id'],user_id))
        except Exception: return False,'سبق استخدام هذا الرمز لهذا الحساب.'
        c.execute('UPDATE access_codes SET redeemed_count=redeemed_count+1 WHERE id=?',(code['id'],))
        c.execute('UPDATE accounts SET free_uses_left=free_uses_left+? WHERE id=?',(code['free_uses'],user_id))
    return True,f"تمت إضافة {code['free_uses']} استخدامات مجانية إلى حسابك."
