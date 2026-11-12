from __future__ import annotations
import re, streamlit as st
from .database import init_db, get_user_by_email, log_attempt, is_rate_limited, db
from .passwords import hash_password, verify_password
from .email_otp import send_otp, verify_otp
from .usage import ensure_user

EMAIL_RE = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')

def _set_user(row, provider='email'):
    with db() as c: c.execute("UPDATE accounts SET last_login_at=CURRENT_TIMESTAMP WHERE id=?",(row['id'],))
    st.session_state.auth_user_id = int(row['id'])
    st.session_state.auth_email = row['email']
    st.session_state.auth_name = row['full_name'] or row['email']
    st.session_state.auth_provider = provider
    st.session_state.authenticated = True

def process_google_login():
    if not st.user.is_logged_in: return False
    email = str(st.user.get('email','')).lower().strip()
    if not email or not EMAIL_RE.match(email): return False
    if not bool(st.user.get('email_verified', False)):
        st.error('تعذر التحقق من أن بريد Google موثّق.')
        st.logout()
        return False
    row = ensure_user(email, str(st.user.get('name','')), verified=True)
    if row['is_blocked']:
        st.error('هذا الحساب موقوف. تواصل مع إدارة المنصة.')
        st.logout(); return False
    _set_user(row,'google')
    return True

def render_email_auth():
    st.subheader('الدخول بالبريد الإلكتروني')
    tab_login, tab_register = st.tabs(['تسجيل الدخول','إنشاء حساب'])
    with tab_register:
        with st.form('register_form'):
            name = st.text_input('الاسم الكامل')
            email = st.text_input('البريد الإلكتروني').strip().lower()
            password = st.text_input('كلمة المرور',type='password')
            confirm = st.text_input('تأكيد كلمة المرور',type='password')
            send = st.form_submit_button('إنشاء الحساب وإرسال رمز التحقق')
        if send:
            if not name.strip() or not EMAIL_RE.match(email): st.error('أدخل الاسم وبريداً إلكترونياً صحيحاً.')
            elif len(password)<10 or not re.search(r'[A-Za-z]',password) or not re.search(r'\d',password): st.error('اجعل كلمة المرور 10 خانات على الأقل وتحتوي حروفاً وأرقاماً.')
            elif password != confirm: st.error('كلمتا المرور غير متطابقتين.')
            elif get_user_by_email(email): st.error('هذا البريد مسجل مسبقاً. استخدم تبويب تسجيل الدخول.')
            else:
                st.session_state.pending_registration = {'name':name.strip(),'email':email,'password_hash':hash_password(password)}
                try:
                    send_otp(email,'register'); st.success('أُرسل رمز التحقق. أدخله في الحقل أدناه.')
                except Exception as e: st.error(f'تعذر إرسال الرمز: {e}')
        pending = st.session_state.get('pending_registration')
        if pending:
            with st.form('verify_register_form'):
                code = st.text_input('رمز التحقق المرسل إلى بريدك',max_chars=6)
                verify = st.form_submit_button('تأكيد البريد وإنشاء الحساب')
            if verify:
                if verify_otp(pending['email'],code,'register'):
                    row = ensure_user(pending['email'],pending['name'],pending['password_hash'],True)
                    _set_user(row); del st.session_state.pending_registration; st.rerun()
                else: st.error('الرمز غير صحيح أو منتهي الصلاحية.')
    with tab_login:
        with st.form('email_login_form'):
            email = st.text_input('البريد الإلكتروني',key='login_email').strip().lower()
            password = st.text_input('كلمة المرور',type='password',key='login_password')
            submit = st.form_submit_button('التحقق وإرسال OTP')
        if submit:
            if is_rate_limited(email): st.error('محاولات كثيرة. انتظر 15 دقيقة قبل المحاولة مجدداً.')
            else:
                row = get_user_by_email(email)
                if not row or not verify_password(password,row['password_hash']):
                    log_attempt(email,False); st.error('بيانات الدخول غير صحيحة.')
                elif row['is_blocked']: st.error('الحساب موقوف من الإدارة.')
                else:
                    try:
                        send_otp(email,'login'); st.session_state.pending_login_email=email
                        st.success('أُرسل رمز التحقق إلى بريدك. أدخله أدناه.')
                    except Exception as e: st.error(f'تعذر إرسال رمز التحقق: {e}')
        pending_email = st.session_state.get('pending_login_email')
        if pending_email:
            with st.form('verify_login_form'):
                code = st.text_input('رمز OTP',max_chars=6)
                verify = st.form_submit_button('تأكيد الدخول')
            if verify:
                if verify_otp(pending_email,code,'login'):
                    row = get_user_by_email(pending_email); log_attempt(pending_email,True); _set_user(row)
                    del st.session_state.pending_login_email; st.rerun()
                else: log_attempt(pending_email,False); st.error('رمز غير صحيح أو منتهي الصلاحية.')

def logout():
    provider = st.session_state.get('auth_provider')
    for key in ('authenticated','auth_user_id','auth_email','auth_name','auth_provider','active_usage_session'):
        st.session_state.pop(key,None)
    if provider == 'google' and st.user.is_logged_in:
        st.logout()
