from __future__ import annotations
import streamlit as st
from .database import init_db, db
from .auth import process_google_login, render_email_auth, logout
from .usage import access_status, start_usage, redeem_access_code
from .admin import create_access_code, grant_free_uses, grant_subscription, set_blocked, disable_code


def _admin_ok():
    expected = str(st.secrets.get('ADMIN_EMAIL','')).lower().strip()
    return bool(expected and st.session_state.get('auth_provider')=='google' and st.session_state.get('auth_email','').lower()==expected and st.session_state.get('authenticated'))

def render_auth_gate() -> bool:
    init_db()
    st.session_state.setdefault('authenticated',False)
    if not st.session_state.authenticated:
        try: process_google_login()
        except Exception as e: st.warning(f'تعذر إكمال تسجيل Google: {e}')
    if not st.session_state.authenticated:
        st.title('منصة إنشاءات للاستشارات الهندسية')
        st.info('سجّل الدخول باستخدام Google أو البريد الإلكتروني. يلزم إعداد Google OAuth والبريد قبل النشر الفعلي.')
        if st.button('الدخول باستخدام Google',use_container_width=True):
            try: st.login()
            except Exception as e: st.error(f'تعذر بدء تسجيل Google: {e}')
        render_email_auth()
        st.stop()
    user_id = int(st.session_state.auth_user_id)
    if _admin_ok(): render_admin_panel()
    else:
        allowed,msg,remaining = access_status(user_id)
        with db() as c:
            account = c.execute('SELECT is_blocked FROM accounts WHERE id=?',(user_id,)).fetchone()
        if not account or account['is_blocked']:
            st.error('الحساب موقوف أو غير متاح. تواصل مع إدارة المنصة.')
            st.stop()
        st.sidebar.markdown(f"**{st.session_state.auth_name}**  \n{st.session_state.auth_email}")
        if st.sidebar.button('تسجيل الخروج',key='auth_logout'):
            logout(); st.rerun()
        if not allowed and not st.session_state.get('active_usage_session'):
            st.warning(msg)
            st.markdown('### خطط الاشتراك بالدولار الأمريكي')
            st.markdown('- شهري: **15 USD**  |  ربع سنوي: **40 USD**  |  سنوي: **100 USD**')
            st.caption('تُفعّل الاشتراكات يدوياً من إدارة المنصة بعد تأكيد الدفع. الأسعار قابلة للتعديل من الإدارة عند اعتماد سياسة التسعير النهائية.')
            st.markdown('[التواصل مع الإدارة عبر واتساب](https://wa.me/9647740884754)')
            code = st.text_input('لديك رمز دعوة مجاني؟')
            if st.button('تفعيل الرمز'):
                ok,info=redeem_access_code(user_id,code); (st.success if ok else st.error)(info); st.rerun()
            st.stop()
        if not st.session_state.get('active_usage_session'):
            st.sidebar.info(msg)
            if st.button('بدء جلسة تدقيق هندسي',type='primary'):
                ok,token=start_usage(user_id)
                if ok: st.session_state.active_usage_session=token; st.rerun()
                else: st.error(token); st.stop()
            st.caption('يُحتسب استخدام مجاني واحد عند بدء جلسة التدقيق، وليس عند تحديث الصفحة. الاشتراك الساري لا يستهلك الرصيد المجاني.')
            st.stop()
        if st.sidebar.button('إنهاء جلسة التدقيق'):
            st.session_state.pop('active_usage_session',None); st.rerun()
    return True


def render_admin_panel():
    st.sidebar.success('دخول المشرف')
    if st.sidebar.button('تسجيل الخروج',key='admin_logout_new'):
        logout(); st.rerun()
    with st.sidebar.expander('إدارة المستخدمين',expanded=False):
        with db() as c: users=c.execute('SELECT id,email,full_name,free_uses_left,subscription_expires_at,is_blocked FROM accounts ORDER BY id DESC').fetchall()
        if users:
            st.dataframe([dict(x) for x in users],use_container_width=True,hide_index=True)
            options={f"{x['email']} (ID {x['id']})":x['id'] for x in users}
            selected=st.selectbox('اختر مستخدماً',list(options),key='admin_user_select')
            uid=options[selected]
            col1,col2=st.columns(2)
            with col1:
                days=st.number_input('أيام الاشتراك',min_value=1,max_value=3650,value=30)
                if st.button('منح اشتراك',key='grant_sub'): grant_subscription(uid,int(days)); st.success('تم منح الاشتراك.'); st.rerun()
                uses=st.number_input('عدد الاستخدامات المجانية',min_value=1,max_value=1000,value=2)
                if st.button('إضافة رصيد مجاني',key='grant_uses'): grant_free_uses(uid,int(uses)); st.success('تمت إضافة الرصيد.'); st.rerun()
            with col2:
                if st.button('إيقاف الحساب',key='block_user'): set_blocked(uid,True); st.rerun()
                if st.button('رفع الإيقاف',key='unblock_user'): set_blocked(uid,False); st.rerun()
        else: st.info('لا يوجد مستخدمون بعد.')
    with st.sidebar.expander('إنشاء رمز مجاني استثنائي'):
        label=st.text_input('وصف الرمز (مثلاً: عميل مشروع X)')
        custom=st.text_input('رمز تحدده بنفسك (اتركه فارغاً للتوليد التلقائي)',type='password',help='16 خانة على الأقل. سيُحفظ بصيغته المشتقة فقط ولن يظهر مجدداً بعد مغادرة الجلسة.')
        count=st.number_input('عدد المستفيدين',min_value=1,max_value=1000,value=1)
        uses=st.number_input('استخدامات لكل مستفيد',min_value=1,max_value=1000,value=2)
        days=st.number_input('مدة صلاحية الرمز بالأيام (0 = دون انتهاء)',min_value=0,max_value=3650,value=30)
        if st.button('إنشاء رمز الدعوة',key='create_access_code'):
            try:
                raw,expiry=create_access_code(label,int(count),int(uses),int(days) or None,custom or None)
                st.session_state.last_created_access_code=raw
                st.success('تم إنشاء الرمز. انسخه الآن؛ لن نخزن النص الصريح للرمز.')
            except ValueError as e:
                st.error(str(e))
        if st.session_state.get('last_created_access_code'):
            st.code(st.session_state.last_created_access_code)
    with st.sidebar.expander('الرموز المنشأة'):
        with db() as c: codes=c.execute('SELECT id,label,max_redemptions,redeemed_count,free_uses,valid_until,is_active FROM access_codes ORDER BY id DESC').fetchall()
        if codes:
            st.dataframe([dict(x) for x in codes],use_container_width=True,hide_index=True)
            active=[x['id'] for x in codes if x['is_active']]
            if active:
                cid=st.selectbox('رمز لإلغائه',active,key='disable_code_select')
                if st.button('إلغاء الرمز',key='disable_code_btn'): disable_code(cid); st.rerun()
