from collections.abc import Mapping

import streamlit as st
from streamlit.errors import StreamlitSecretNotFoundError


def _is_configured(value: object) -> bool:
    if not isinstance(value, str) or not value.strip():
        return False
    normalized = value.upper()
    return not any(
        marker in normalized
        for marker in ("YOUR_", "YOUR-", "REPLACE_WITH", "PLACEHOLDER")
    )


def auth_configuration_issue(
    auth_settings: Mapping[str, object] | None,
) -> str | None:
    if not isinstance(auth_settings, Mapping):
        return "لم يُعثر على إعدادات المصادقة [auth]."

    for key in ("redirect_uri", "cookie_secret"):
        if not _is_configured(auth_settings.get(key)):
            return f"إعداد المصادقة {key} غير موجود أو ما زال قيمة نموذجية."

    provider = auth_settings.get("auth0")
    if not isinstance(provider, Mapping):
        return "إعداد مزوّد Auth0 [auth.auth0] غير موجود."

    for key in ("client_id", "client_secret", "server_metadata_url"):
        if not _is_configured(provider.get(key)):
            return f"إعداد Auth0 {key} غير موجود أو ما زال قيمة نموذجية."

    return None


def require_authentication() -> None:
    try:
        auth_settings = st.secrets.get("auth")
    except StreamlitSecretNotFoundError:
        auth_settings = None

    issue = auth_configuration_issue(auth_settings)
    if issue:
        st.error("التطبيق غير متاح قبل إعداد تسجيل الدخول.")
        st.info(
            f"{issue} راجع قسم إعداد Auth0 في README، ثم أضف القيم "
            "إلى أسرار بيئة التشغيل. لا تضع كلمات المرور أو مفاتيح Auth0 "
            "في المستودع أو ترسلها في المحادثة."
        )
        st.stop()

    if not st.user.is_logged_in:
        st.title("تسجيل الدخول أو إنشاء حساب")
        st.write(
            "يمكن إنشاء حساب ذاتي عبر صفحة Auth0. يلزم تفعيل التسجيل "
            "والتحقق من البريد الإلكتروني في إعدادات Auth0."
        )
        st.warning(
            "ستُرسل ملفات ETABS إلى خادم الاستضافة لمعالجتها. لا ترفع "
            "ملفات مشاريع سرية قبل مراجعة سياسة الخصوصية والاحتفاظ "
            "وتأمين بيئة الاستضافة."
        )
        st.button(
            "المتابعة إلى التسجيل أو تسجيل الدخول",
            on_click=st.login,
            args=("auth0",),
            type="primary",
        )
        st.stop()

    if not st.user.get("email_verified", False):
        st.error(
            "تحقق من بريدك الإلكتروني لدى Auth0 قبل استخدام التطبيق. "
            "إذا أكدت البريد للتو، سجّل الخروج ثم أعد تسجيل الدخول."
        )
        st.button(
            "تسجيل الخروج",
            on_click=st.logout,
            key="unverified_auth_logout",
        )
        st.stop()

    email = st.user.get("email")
    with st.sidebar:
        st.caption(f"مسجل الدخول: {email or 'مستخدم Auth0'}")
        st.button(
            "تسجيل الخروج",
            on_click=st.logout,
            key="auth_logout",
        )
