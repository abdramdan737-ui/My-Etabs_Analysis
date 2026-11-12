from __future__ import annotations
import hashlib, hmac, secrets, smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta, timezone
import streamlit as st
from .database import db

UTC = timezone.utc

def _now(): return datetime.now(UTC)
def _secret():
    return str(st.secrets.get('OTP_PEPPER', st.secrets.get('APP_SECRET', '')))

def _digest(email: str, code: str) -> str:
    pepper = _secret()
    if not pepper: raise RuntimeError('لم يتم إعداد OTP_PEPPER في Streamlit Secrets.')
    return hmac.new(pepper.encode(), f'{email.lower()}:{code}'.encode(), hashlib.sha256).hexdigest()

def send_otp(email: str, purpose: str = 'login'):
    cfg = st.secrets.get('email', {})
    required = ('smtp_host','smtp_port','smtp_user','smtp_password','from_email')
    if any(not cfg.get(k) for k in required):
        raise RuntimeError('إعدادات البريد غير مكتملة في Streamlit Secrets.')
    code = f'{secrets.randbelow(1_000_000):06d}'
    expiry = (_now() + timedelta(minutes=5)).isoformat()
    with db() as c:
        recent = c.execute("SELECT COUNT(*) n FROM otp_codes WHERE email=? AND created_at >= datetime('now','-10 minutes')", (email.lower(),)).fetchone()['n']
        if recent >= 3: raise RuntimeError('تجاوزت الحد المسموح لإرسال الرموز. حاول بعد 10 دقائق.')
        c.execute("INSERT INTO otp_codes(email,code_hash,purpose,expires_at) VALUES(?,?,?,?)", (email.lower(),_digest(email,code),purpose,expiry))
    msg = EmailMessage()
    msg['Subject'] = 'رمز التحقق - منصة إنشاءات'
    msg['From'] = cfg['from_email']; msg['To'] = email
    msg.set_content(f'رمز التحقق الخاص بك هو: {code}\nصلاحية الرمز خمس دقائق. لا تشاركه مع أي شخص.')
    with smtplib.SMTP(str(cfg['smtp_host']), int(cfg['smtp_port']), timeout=20) as server:
        server.starttls(); server.login(str(cfg['smtp_user']), str(cfg['smtp_password']))
        server.send_message(msg)

def verify_otp(email: str, code: str, purpose: str = 'login') -> bool:
    with db() as c:
        row = c.execute("SELECT * FROM otp_codes WHERE email=? AND purpose=? AND used=0 ORDER BY id DESC LIMIT 1", (email.lower(),purpose)).fetchone()
        if not row or row['attempts'] >= 5: return False
        valid = datetime.fromisoformat(row['expires_at']) > _now() and hmac.compare_digest(row['code_hash'], _digest(email,code.strip()))
        if valid: c.execute('UPDATE otp_codes SET used=1 WHERE id=?',(row['id'],))
        else: c.execute('UPDATE otp_codes SET attempts=attempts+1 WHERE id=?',(row['id'],))
        return valid
