import io
import os
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
# ==========================================
# تسجيل مجاني مبسط: الاسم ورقم الهاتف فقط
# لا توجد كلمات مرور أو OTP أو Google OAuth أو اشتراكات أو حدود استخدام
# ==========================================
import sqlite3
import re
st.set_page_config(page_title="منصة تدقيق نتائج الايتابس وتوليد النوتة الحسابية", layout="wide")

USERS_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "platform_users.db")

def normalize_phone(raw):
    """توحيد الأرقام العربية والهندية وإزالة الفواصل مع إبقاء الرقم كنص."""
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    value = str(raw or "").translate(trans).strip()
    value = re.sub(r"[\s()\-]", "", value)
    return value

def init_users_database():
    with sqlite3.connect(USERS_DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS platform_users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                phone TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def render_free_access_gate():
    init_users_database()

    if st.session_state.get("platform_logged_in", False):
        user = st.session_state.get("platform_user", {})
        with st.sidebar:
            st.success(f"مرحباً {user.get('full_name', '')}")
            st.caption(f"رقم الهاتف: {user.get('phone', '')}")
            if st.button("تسجيل الخروج", key="platform_logout", use_container_width=True):
                st.session_state.pop("platform_logged_in", None)
                st.session_state.pop("platform_user", None)
                st.rerun()
        return True

    st.title("منصة تدقيق نتائج الايتابس وتوليد النوتة الحسابية")
    st.subheader("الدخول المجاني إلى المنصة")
    st.info("التسجيل متاح بالاسم ورقم الهاتف فقط. جميع أدوات المنصة متاحة مجاناً حالياً دون اشتراك .")

    tab_register, tab_login = st.tabs(["تسجيل مستخدم جديد", "دخول مستخدم مسجل"] )

    with tab_register:
        name = st.text_input("الاسم الكامل", key="platform_register_name")
        phone = st.text_input("رقم الهاتف", key="platform_register_phone", placeholder="07XXXXXXXXX")
        if st.button("إنشاء حساب مجاني", type="primary", key="platform_register_button", use_container_width=True):
            clean_name = name.strip()
            clean_phone = normalize_phone(phone)
            if not clean_name or not clean_phone:
                st.error("يرجى إدخال الاسم ورقم الهاتف.")
            else:
                try:
                    with sqlite3.connect(USERS_DB_PATH) as conn:
                        conn.execute(
                            "INSERT INTO platform_users (full_name, phone) VALUES (?, ?)",
                            (clean_name, clean_phone)
                        )
                        conn.commit()
                    st.session_state["platform_user"] = {"full_name": clean_name, "phone": clean_phone}
                    st.session_state["platform_logged_in"] = True
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("رقم الهاتف مسجل مسبقاً. استخدم تبويب الدخول.")
                except sqlite3.Error as exc:
                    st.error(f"تعذر إنشاء الحساب بسبب خطأ في قاعدة البيانات: {exc}")

    with tab_login:
        login_name = st.text_input("الاسم الكامل", key="platform_login_name")
        login_phone = st.text_input("رقم الهاتف", key="platform_login_phone")
        if st.button("دخول إلى المنصة", key="platform_login_button", use_container_width=True):
            clean_name = login_name.strip()
            clean_phone = normalize_phone(login_phone)
            if not clean_name or not clean_phone:
                st.error("يرجى إدخال الاسم ورقم الهاتف.")
            else:
                with sqlite3.connect(USERS_DB_PATH) as conn:
                    row = conn.execute(
                        "SELECT full_name, phone FROM platform_users WHERE full_name = ? AND phone = ?",
                        (clean_name, clean_phone)
                    ).fetchone()
                if row:
                    st.session_state["platform_user"] = {"full_name": row[0], "phone": row[1]}
                    st.session_state["platform_logged_in"] = True
                    st.rerun()
                else:
                    st.error("بيانات الدخول غير موجودة. تحقق من الاسم والرقم أو أنشئ حساباً جديداً.")

    st.caption("تنبيه: هذا تسجيل مبسط دون التحقق من ملكية رقم الهاتف؛ لا تستخدمه لحماية بيانات حساسة.")
    return False

if not render_free_access_gate():
    st.stop()

st.success(f"مرحباً بك {st.session_state['platform_user']['full_name']} في منصة إنشاءات للاستشارات الهندسية.")

# ==========================================
# يبدأ هنا محتوى المنصة وأدوات التدقيق دون فرض أي حدود استخدام
# ==========================================

# إدارة الملفات المرفوعة بأمان في الذاكرة المؤقتة (Session State)
# ==========================================
st.sidebar.header("📁 إدارة ملفات البيانات (ETABS Results)")

uploaded_file = st.sidebar.file_uploader(
    "1️⃣ قم برفع ملف البيانات الأولية (results.xlsx)", 
    type=["xlsx", "xls"],
    help="يتم قراءة الملف مؤقتاً في الذاكرة الجلسوية للحفاظ على السرية وعدم حفظه على القرص."
)

# تهيئة الذاكرة المؤقتة لبيانات الجلسة إن لمג تكن موجودة
if 'session_sheets_data' not in st.session_state:
    st.session_state['session_sheets_data'] = None
if 'session_sheet_names' not in st.session_state:
    st.session_state['session_sheet_names'] = []

# التعامل مع الملف المرفوع في الذاكرة المؤقتة حصرياً دون حفظه على القرص
if uploaded_file is not None:
    try:
        # قراءة ملف الإكسيل مباشرة من الذاكرة (BytesIO) دون كتابته على الهارد
        bytes_data = uploaded_file.getvalue()
        excel_buffer = io.BytesIO(bytes_data)
        xls_obj = pd.ExcelFile(excel_buffer)
        
        # تخزين أسماء الأوراق والبيانات في الذاكرة المؤقتة للجلسة فقط
        st.session_state['session_sheet_names'] = xls_obj.sheet_names
        
        # قراءة الأوراق الاحتياطية لتكون متاحة للجداول
        sheets_dict = {}
        for s in xls_obj.sheet_names:
            sheets_dict[s] = pd.read_excel(excel_buffer, sheet_name=s)
        st.session_state['session_sheets_data'] = sheets_dict
        
        st.sidebar.success("✅ تمت معالجة وقراءة الملف مؤقتاً في الذاكرة بنجاح!")
    except Exception as e:
        st.sidebar.error(f"❌ حدث خطأ أثناء قراءة الملف: {e}")

# التحقق من توفر البيانات (سواء المرفوعة مؤقتاً في الذاكرة أو الملف الافتراضي المحلي إن وجد)
active_sheets = st.session_state['session_sheet_names']
active_data = st.session_state['session_sheets_data']

# إذا لم يتم رفع أي ملف ولا يوجد ملف افتراضي، يتم التنبيه وإيقاف التنفيذ
local_default_path = os.path.join(os.getcwd(), "result_tables.xlsx")

if not active_sheets and not os.path.exists(local_default_path):
    st.error("🚨 **تنبيه هام جداً:** لا يمكن المتابعة أو الحصول على أي نتائج هندسية لعدم توفر بيانات الإدخال!")
    
    with st.expander("📌 تعليمات هامة لطريقة الرفع وأسماء الأوراق (Sheets) الضرورية", expanded=True):
        st.markdown("""
        لضمان عمل المنصة واستخراج النوتة الحسابية وتدقيق النتائج بدقة، يرجى الالتزام بالآتي:
        
        1. **طريقة الرفع:** قم بسحب وإفلات ملف الإكسيل الخام بكامل أوراقه أو النقر على زر **(Browse files)** في الشريط الجانبي لتصفح جهازك ورفع الملف المسمى باسم **`results.xlsx`**.
        2. **أسماء الأوراق الضرورية في ملف الإكسيل:**
           يجب أن يحتوي ملف الإكسيل على الأوراق التالية مطابقة لتنفيذ المعالجات والتدقيق لضمان التدقيق الأولي:
           * ورقة **`Story Stiffness`**: تدقيق صلابة الطوابق الجانبية.
           * ورقة **`Story Max Over Avg Disp`**: تدقيق نسبة الانجرافات القصوى إلى المتوسطة لكل طابق.
           * ورقة **`Modal Participating Mass Ratios`**: تدقيق نسبة مشاركة الكتلة لكل نمط.
           * ورقة **`Story Drifts`**: تدقيق انجرافات الطوابق النسبية.
           * ورقة **`Centers Of Mass And Rigidity`**: تدقيق مراكز الكتلة والصلابة لكل طابق.
           * ورقة **`Base Reactions`**: تدقيق قوى القص القاعدي ومقارنة الاستاتيكي والديناميكي.
        """)
    st.stop()

# تحديد مصدر قراءة الجداول (الذاكرة المؤقتة للجلسة لها الأولوية المطلقة لكي لا تتأثر الملفات الافتراضية)
def get_dynamic_sheet_df(base_name):
    if active_data and active_sheets:
        target_sheet = get_sheet_by_name(active_sheets, base_name)
        if target_sheet and target_sheet in active_data:
            return active_data[target_sheet]
    # العودة للملف الافتراضي المحلي إن لم يُرفع ملف جديد
    if os.path.exists(local_default_path):
        try:
            return pd.read_excel(local_default_path, sheet_name=get_sheet_by_name(pd.ExcelFile(local_default_path).sheet_names, base_name))
        except:
            return pd.DataFrame()
    return pd.DataFrame()
# ==========================================
import os
import pandas as pd
import streamlit as st

# ==========================================
# نظام إدارة ورفع الملف المعتمد (results.xlsx)
# ==========================================
st.sidebar.header("📁 إدارة الملف المعتمد للمعالجة")

current_dir = os.getcwd()
target_file_path = os.path.join(current_dir, "results.xlsx")

# إذا قام المستخدم برفع ملف جديد، يتم استبدال الملف الحالي المعتمد في مسار العمل
if uploaded_file is not None:
    try:
        with open(target_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ تم استبدال وتحديث الملف المعتمد (results.xlsx) بنجاح!")
    except Exception as e:
        st.sidebar.error(f"❌ حدث خطأ أثناء حفظ الملف: {e}")

# التحقق الإلزامي من وجود الملف المعتمد للمعالجة
if not os.path.exists(target_file_path):
    st.error("🚨 **تنبيه هام جداً:** لا يمكن المتابعة أو الحصول على أي نتائج لعدم توفر ملف البيانات المعتمد (`results.xlsx`)!")
    
    with st.expander("📌 تعليمات هامة لطريقة الرفع وأسماء الأوراق  الضرورية", expanded=True):
        st.markdown(
            "لضمان عمل التطبيق واستخراج النوتة الحسابية وتدقيق النتائج بدقة، يرجى الالتزام بالآتي:\n\n"
            "1. **طريقة الرفع:** قم بسحب وإفلات ملف الإكسيل الخام بكامل أوراقه أو النقر على زر **(Browse files)** في الشريط الجانبي لتصفح جهازك ورفع الملف وتسميته باسم **`results.xlsx`**.\n"
            "2. **أسماء الأوراق الضرورية في ملف الإكسيل:**\n"
            "يجب أن يحتوي ملف الإكسيل على الأوراق التالية على الاقل لضمان التدقيق الأولي\n"
            "* ورقة **`Story Stiffness`**: تدقيق صلابة الطوابق الجانبية.\n"
            "* ورقة **`Story Max Over Avg Disp`**: تدقيق نسبة الانجرافات القصوى إلى المتوسطة لكل طابق.\n"
            "* ورقة **`Modal Participating Mass Ratios`**: تدقيق نسبة مشاركة الكتلة لكل نمط.\n"
            "* ورقة **`Story Drifts`**: تدقيق انجرافات الطوابق النسبية.\n"
            "* ورقة **`Centers Of Mass And Rigidity`**: تدقيق مراكز الكتلة والصلابة لكل طابق.\n"
            "* ورقة **`Base Reactions`**: تدقيق قوى القص القاعدي ومقارنة الاستاتيكي والديناميكي."
        )
    st.stop()

# قراءة الأوراق من الملف المعتمد (results.xlsx) فور توفره
try:
    xls_file = pd.ExcelFile(target_file_path)
    sheet_names = xls_file.sheet_names
except Exception as e:
    st.error(f"⚠️ تعذّر قراءة محتويات ملف الإكسيل المعتمد: {e}")
    sheet_names = []
# ==========================================
# إعدادات واجهة التطبيق
# ==========================================
def _load_css():
    """يقرأ ملف style.css من نفس مجلد هذا السكربت ويُدرجه في الصفحة.
    محمي بـ try/except عمداً: إن غاب الملف أو تعذّرت قراءته لأي سبب (مسار
    خاطئ، نسيان رفعه)، يستمر التطبيق بالعمل بالمظهر الافتراضي بدل التعطّل
    الكامل، مع تنبيه واضح في الصفحة نفسه بدل فشل صامت."""
    try:
        css_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "style.css")
        with open(css_path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except Exception as exc:
        st.warning(
            f"⚠️ تعذّر تحميل ملف التنسيق style.css ({exc}). سيعمل التطبيق "
            "بالمظهر الافتراضي. تأكد من وجود style.css في نفس مجلد هذا الملف."
        )


_load_css()

st.title("🏗️ تدقيق التحليل الإنشائي لبرنامج الايتابس وتوليد النوتة الحسابية الهندسية")
st.write(
    "منصة هندسية متكاملة لتدقيق مخرجات الايتابس وفق مدونة الزلازل العراقية والكود الامريكي "
    "مع إدخال بيانات المشروع وتصدير تقرير هندسي احترافي ورسوم تفاعلية."
)

files_dir = os.getcwd()


# ==========================================
# دوال مساعدة عامة
# ==========================================
def get_excel_file_path():
    current_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else files_dir
    for ext in [".xlsx", ".xls", ".XLSX", ".XLS"]:
        full_path = os.path.join(current_dir, "result_tables" + ext)
        if os.path.exists(full_path):
            return full_path
    if os.path.exists(current_dir):
        for f in os.listdir(current_dir):
            if f.lower().endswith(('.xlsx', '.xls')) and ('result' in f.lower() or 'table' in f.lower()):
                return os.path.join(current_dir, f)
    return None


def _clean_sheet_token(name: str) -> str:
    return name.lower().replace("_", "").replace(" ", "").replace("-", "")


def get_sheet_by_name(sheet_names, base_name):
    clean_base = _clean_sheet_token(base_name)
    for s in sheet_names:
        clean_sheet = _clean_sheet_token(s)
        if clean_base in clean_sheet or clean_sheet in clean_base:
            return s
    return sheet_names[0]


def strip_and_dedupe_columns(df):
    """يزيل الفراغات الزائدة من أسماء الأعمدة، ثم يضمن عدم تكرار أي اسم بعد
    الإزالة (بإضافة لاحقة رقمية لأي تكرار جديد ناتج عن الإزالة نفسها).
    هذا تحديداً ما كان ينقص عند قراءة ورقة "P-Delta " ذات العمودين 'P Kn'
    و'P Kn ' (بفراغ زائد) — فقد كانا مختلفين تماماً لـ pandas قبل الإزالة،
    فأصبحا متطابقين حرفياً بعدها، فرفض pyarrow/streamlit عرض الجدول بخطأ
    "Duplicate column names found" كما ظهر فعلياً عند التشغيل."""
    stripped = [str(c).strip() for c in df.columns]
    seen = {}
    unique = []
    for name in stripped:
        if name not in seen:
            seen[name] = 0
            unique.append(name)
        else:
            seen[name] += 1
            unique.append(f"{name}.dup{seen[name]}")
    df = df.copy()
    df.columns = unique
    return df


def clean_dataframe(df):
    unwanted_patterns = ['^unnamed', '^index', 'تسلسل', 'case type', 'step type', 'step number']
    pattern = '|'.join(unwanted_patterns)
    df = df.loc[:, ~df.columns.str.contains(pattern, case=False, na=False)]
    return df


def filter_lateral_cases(df):
    if df is None or df.empty:
        return df
    case_col = next((c for c in df.columns if "case" in c.lower() or "output" in c.lower() or "load" in c.lower()), None)
    if case_col:
        mask = df[case_col].astype(str).str.upper().str.contains("EQ|RS|WIND|SPEC", na=False)
        filtered_df = df[mask]
        return filtered_df if not filtered_df.empty else df
    return df


def color_results(val):
    val_str = str(val).upper()
    if any(k in val_str for k in ["SAFE", "NO P-DELTA", "محقق", "صحيح", "مُحَقَّق"]):
        return "background-color: #1e4620; color: #8be89b; font-weight: bold; font-size: 14px;"
    elif any(k in val_str for k in ["UNSAFE", "SOFT", "❌", "يتطلب تعديل"]):
        return "background-color: #5c1f24; color: #ff9aa2; font-weight: bold; font-size: 14px;"
    return "font-size: 14px; color: #e8ecf1;"


def color_mass_ratios(val):
    try:
        num = float(val)
        if num >= 90.0 or (0 <= num <= 1 and num * 100 >= 90.0):
            return "background-color: #1e4620; color: #8be89b; font-weight: bold;"
    except Exception:
        pass
    return "color: #e8ecf1;"


def classify_status(status: str) -> str:
    """يصنّف نص الحالة إلى: pass / fail / warn / pending.
    مهم: يجب فحص "غير محقق" قبل "محقق" لأن الكلمة الثانية جزء حرفي من الأولى
    (هذا كان سبب خطأ احتساب لوحة المؤشرات في النسخة السابقة)."""
    s = str(status)
    s_up = s.upper()
    if "غير محقق" in s or "UNSAFE" in s_up or "تجاوز" in s:
        return "fail"
    if "خطأ" in s:
        return "error"
    if "يتطلب" in s:
        return "warn"
    if "غير محسوب" in s or "⏳" in s:
        return "pending"
    if "محقق" in s or "SAFE" in s_up or "NO P-DELTA" in s_up:
        return "pass"
    return "pending"


def status_suffix(status: str) -> str:
    kind = classify_status(status)
    return {"pass": "✅", "fail": "❌", "warn": "⚠️", "error": "⚠️", "pending": "➖"}.get(kind, "⏳")


# ==========================================
# معادلات ومنهجية كل قسم (تُعرض تحت اسم القسم مباشرة)
# ==========================================
def render_section_equations(section_key, context):
    with st.expander("📐 المعادلات والمنهجية المستخدمة في هذا الفحص", expanded=True):
        if section_key == "centers":
            st.latex(r"e_x = |X_r - X_m|,\qquad e_y = |Y_r - Y_m|")
            st.latex(r"r_x = 100\times\frac{e_x}{D_x}\ (\%),\qquad r_y = 100\times\frac{e_y}{D_y}\ (\%)")
            st.markdown(
                "- `X_r, Y_r`: إحداثيات مركز الصلابة (Center of Rigidity) لكل طابق/حجاب حاجز.\n"
                "- `X_m, Y_m`: إحداثيات مركز الكتلة (Center of Mass).\n"
                "- `D_x, D_y`: البعدان الكليان للمسقط الأفقي للمبنى (ملم) كما أُدخلا في الشريط الجانبي.\n"
                "- **تصحيح هام بعد مراجعة ASCE 7-10 مباشرة:** الكود **لا يضع** أي حد لنسبة "
                "اللامركزية بين مركز الكتلة ومركز الصلابة كنسبة من بُعد المبنى (لا 15% ولا أي "
                "رقم آخر). لذلك أُزيل حكم النجاح/الفشل من هذا القسم، وتُعرض `e_x, e_y` "
                "والنسب هنا **للعرض الإرشادي فقط**.\n"
                "- المعيار الكودي الفعلي لتقييم الالتواء الناتج عن هذه اللامركزية هو **نسبة "
                "إزاحات الحجاب الحاجز** (δmax/δavg مقابل 1.2/1.4) — انظر القسم 5 "
                "'الإزاحة والالتواء العرضي' حيث يُطبَّق هذا المعيار فعلياً وفق جدول 12.3-1."
            )
        elif section_key == "base_reactions":
            st.latex(r"V_{s,\mathrm{req}} = V_{\mathrm{static}} \times \mathrm{target\%}")
            st.latex(r"\text{Check: } V_{\mathrm{dynamic}} \ge V_{s,\mathrm{req}}")
            st.latex(r"SF_{\mathrm{recommended}} = \frac{V_{\mathrm{static}} \times \mathrm{target\%}}{V_{\mathrm{dynamic}}}\times\frac{I_e}{R}")
            st.markdown(
                "- `V_static`: أقصى قص قاعدي من حالة الحمل الاستاتيكي EQX/EQY.\n"
                "- `V_dynamic`: أقصى قص قاعدي من حالة الطيف الديناميكي RSX/RSY.\n"
                f"- `target%`: النسبة المستهدفة المختارة في الشريط الجانبي (حالياً {context.get('asce_percentage', 0)*100:.0f}%).\n"
                "- `Ie`, `R`: عامل الأهمية ومعامل تعديل الاستجابة المختاران في الشريط الجانبي.\n"
                "- معامل التحجيم المعروض توصية حسابية أولية فقط وليس اعتماداً نهائياً."
            )
        elif section_key == "story_drifts":
            st.latex(r"\Delta = \dfrac{C_d\,\Delta_{xe}}{I_e}\qquad\text{(المعادلة 12.8-15)}")
            st.latex(r"\Delta_a = \dfrac{\Delta_{a,\mathrm{base}}}{\rho}\qquad\text{(الجدول 12.12-1)}")
            st.latex(r"\text{Check: } \Delta \le \Delta_a")
            st.markdown(
                "- `Δxe`: نسبة الانجراف المرنة (elastic) كما وردت في جدول ETABS، بلا وحدة.\n"
                "- `Cd`, `Ie`: معامل تكبير الانحراف وعامل الأهمية، المحددان من فئة المخاطر "
                "والنظام الإنشائي المختارين في الشريط الجانبي.\n"
                "- `Δa,base`: حد الانجراف الأساسي حسب فئة المخاطر (صف 'جميع المنشآت الأخرى' "
                "من جدول 12.12-1): **0.020** لفئة I/II، **0.015** لفئة III، **0.010** لفئة IV "
                "(الفئة IV نصف الفئة I/II تماماً كما ينص الكود).\n"
                "- `ρ`: عامل التكرار (Redundancy Factor) حسب §12.3.4، المُدخل في الشريط الجانبي.\n"
                "- **تصحيح هام:** النسخة السابقة من هذا التطبيق استخدمت صيغة مرتبطة بالفترة "
                "الأساسية T ومعامل R (`0.02 أو 0.025 ÷ 0.7R`)، وهي صيغة **غير موجودة في "
                "ASCE 7-10** (أقرب لصيغة قديمة من UBC-1997). تم استبدالها بالصيغة الصحيحة أعلاه "
                "بعد التحقق المباشر من جدولي 12.12-1 والمعادلة 12.8-15."
            )
        elif section_key == "modal_mass":
            st.latex(r"M_{\mathrm{cum}} = \max_i\big(\mathrm{SumUX}_i\ \text{أو}\ \mathrm{SumUY}_i\big)")
            st.latex(r"\text{Check: } M_{\mathrm{cum}} \ge 90\%")
            st.markdown(
                "- `SumUX`, `SumUY`: مشاركة الكتلة التراكمية لكل حالة نمطية في جدول ETABS.\n"
                "- يُؤخذ أكبر مقدار تراكمي لكل اتجاه عبر جميع الحالات النمطية.\n"
                "- حد 90% معيار شائع لكفاية عدد الأنماط المستخدمة في التحليل، وليس نسخة "
                "موثقة من نص الكود العراقي في هذا التطبيق."
            )
        elif section_key == "displacements":
            st.latex(r"\mathrm{ratio} = \left|\dfrac{\delta_{\max}}{\delta_{\mathrm{avg}}}\right|")
            st.latex(r"A_x = \min\left[\left(\dfrac{\delta_{\max}}{1.2\,\delta_{\mathrm{avg}}}\right)^{2},\ 3.0\right]\qquad\text{(المعادلة 12.8-14)}")
            st.markdown(
                "**هذا هو المعيار الكودي الفعلي لعدم الانتظام الالتوائي وفق جدول 12.3-1 في "
                "ASCE 7-10** (وليس حداً داخلياً تقريبياً كما في النسخة السابقة من هذا القسم):\n"
                "- `ratio ≤ 1.2`: لا يوجد عدم انتظام التوائي (Torsional Irregularity).\n"
                "- `1.2 < ratio ≤ 1.4`: عدم انتظام نوع 1a — يجب تطبيق معامل التكبير `Ax`.\n"
                "- `ratio > 1.4`: عدم انتظام حاد نوع 1b (Extreme Torsional Irregularity) — "
                "قيود تصميمية إضافية تُفرض بموجب جدول 12.6-1 (قد تمنع استخدام طريقة القوة "
                "الجانبية المكافئة في فئات تصميم زلزالي معينة).\n"
                "- `Ax` معروض هنا كقيمة **مرجعية/إرشادية** فقط؛ تطبيقه الفعلي على قوى التصميم "
                "يتطلب التأكد من أن حالات ETABS تتضمن اللامركزية العرضية المطلوبة، وهو أمر لا "
                "يتحقق منه هذا التطبيق تلقائياً."
            )
        elif section_key == "story_stiffness":
            st.latex(r"\text{Type 1a (لين): } K_i < 0.70\,K_{\text{above}}\ \ \text{أو}\ \ K_i < 0.80\,\overline{K}_{3\,\text{above}}")
            st.latex(r"\text{Type 1b (لين حاد): } K_i < 0.60\,K_{\text{above}}\ \ \text{أو}\ \ K_i < 0.70\,\overline{K}_{3\,\text{above}}")
            st.markdown(
                "- `Ki`: الصلابة الجانبية للطابق قيد الفحص (Stiff X أو Stiff Y من ETABS).\n"
                "- `K_above`: صلابة الطابق الذي يعلوه مباشرة.\n"
                "- `K̄3above`: متوسط صلابة أعلى 3 طوابق فوقه (أو أقل إن لم تتوفر 3).\n"
                "- هذا جدول **12.3-2** في ASCE 7-10 (Soft Story / Stiffness Irregularity)، "
                "مطبَّق هنا مباشرة على بيانات ETABS. **تبسيط مقصود:** الكود يسمح بتجاوز هذا "
                "الفحص كلياً إن تحقق شرط بديل أبسط مبني على الانجراف فقط "
                "(إذا كانت نسبة انجراف أي طابق ≤ 1.3 ضعف نسبة انجراف الطابق الذي يعلوه)، "
                "وهذا الشرط البديل **غير مُطبَّق هنا**؛ لذلك قد يُصنَّف التطبيق طابقاً كـ"
                "'لين' بينما يُعفيه الكود فعلياً عبر الشرط البديل — النتيجة هنا تحفظية وليست "
                "نهائية. الطابقان الأعلى في المبنى لا يُفحصان (لا يوجد عدد كافٍ من الطوابق "
                "فوقهما للمقارنة)."
            )
        elif section_key == "pdelta":
            st.latex(r"\theta = \dfrac{P\,\Delta}{V\,h_s\,C_d}")
            st.latex(r"\theta_{\max} = \min\left(0.25,\ \dfrac{0.5}{\beta\,C_d}\right)")
            st.markdown(
                "- `θ ≤ 0.1`: يمكن إهمال تأثير P-Delta.\n"
                "- `0.1 < θ ≤ θmax`: يجب إدراج تأثير P-Delta في التحليل.\n"
                "- `θ > θmax`: تجاوز حد الاستقرار؛ يلزم إجراء هندسي إضافي.\n"
                "- يعرض هذا القسم جدول P-Delta كما ورد من ETABS. إذا تضمّن الجدول عموداً "
                "رقمياً لنسبة الاستقرار (θ) يعتمده التطبيق تلقائياً لتحديد الحالة أعلاه؛ "
                "وإلا تبقى الحالة **غير محسوبة** حتى يتوفر هذا العمود صراحة."
            )
        elif section_key == "wind_drift":
            st.latex(r"\text{انجراف نسبي (طابق): } \dfrac{\Delta_i \times 0.7}{h_i} \le \dfrac{1}{360}")
            st.latex(r"\text{انجراف كلي (مبنى): } \dfrac{\delta_{\text{من القاعدة}}}{H_{\text{تراكمي من القاعدة}}} \le \dfrac{1}{360}")
            st.markdown(
                "**هذا فحص خدمي (Serviceability) لانجراف الرياح، وليس جزءاً من ASCE 7-10 "
                "Chapter 12 الزلزالي** — أكّد المستخدم أن هذا هو القصد من جدولَي "
                "'Drift Ratio X/Y' (الانجراف النسبي بين كل طابق وما يليه) و"
                "'Displacement X/Y' (الانجراف الكلي للمبنى من القاعدة حتى كل طابق).\n"
                "- الحد المستخدم **1/360** في كلتا الحالتين، كما ورد في ملف المشروع نفسه.\n"
                "- عامل **0.7** في الانجراف النسبي محفوظ كما ورد في ملف المشروع (غالباً "
                "لتحويل القيمة من مستوى التصميم إلى المستوى الخدمي)؛ لم أتحقق من توثيقه "
                "كوداً، بل نقلته كما هو من جدول المشروع.\n"
                "- **الارتفاع التراكمي لفحص الانجراف الكلي يُعاد حسابه داخلياً** من ارتفاع "
                "كل طابق (H)، وليس من عمود 'Cumulative' الأصلي في الملف — لأنه وُجد غير "
                "متّسق رقمياً عند المراجعة المباشرة.\n"
                "- **تنبيه بيانات:** إذا كانت كل قيم الانجراف/الإزاحة الفعلية صفراً في "
                "الجدول، فهذا يعني أن حالة الرياح غير مرتبطة بعد بهذا الجدول في الملف "
                "المرفوع، فتبقى الحالة 'غير محسوب' ولا تُعرض كنجاح زائف."
            )


# ==========================================
# دوال الرسوم البيانية
# ==========================================
def plot_center_of_mass_vs_rigidity(centers_df):
    story_col = 'Story' if 'Story' in centers_df.columns else centers_df.columns[0]
    fig = go.Figure()
    if 'X_CM' in centers_df.columns and 'X_CR' in centers_df.columns:
        fig.add_trace(go.Scatter(x=centers_df['X_CM'], y=centers_df[story_col], mode='lines+markers', name='مركز الكتلة X (X_CM)', line=dict(color='#1F4E79', width=3)))
        fig.add_trace(go.Scatter(x=centers_df['X_CR'], y=centers_df[story_col], mode='lines+markers', name='مركز الصلابة X (X_CR)', line=dict(color='#D9534F', width=3, dash='dash')))
    if 'Y_CM' in centers_df.columns and 'Y_CR' in centers_df.columns:
        fig.add_trace(go.Scatter(x=centers_df['Y_CM'], y=centers_df[story_col], mode='lines+markers', name='مركز الكتلة Y (Y_CM)', line=dict(color='#2CA02C', width=3)))
        fig.add_trace(go.Scatter(x=centers_df['Y_CR'], y=centers_df[story_col], mode='lines+markers', name='مركز الصلابة Y (Y_CR)', line=dict(color='#F0AD4E', width=3, dash='dash')))
    fig.update_layout(title="مقارنة مراكز الكتلة (CM) والصلابة (CR) عبر الطوابق", xaxis_title="إحداثيات المركز (مم)", yaxis_title="الطابق", template="plotly_white", height=500)
    return fig


def plot_base_shear_comparison(v_st_x, v_st_y, v_dyn_x, v_dyn_y, target_pct):
    fig = go.Figure()
    directions = ['الاتجاه الأفقي (X)', 'الاتجاه العمودي (Y)']
    static_vals = [v_st_x, v_st_y]
    dynamic_vals = [v_dyn_x, v_dyn_y]
    required_vals = [s * target_pct for s in static_vals]

    fig.add_trace(go.Bar(name='القص الاستاتيكي (Static)', x=directions, y=static_vals, marker_color='#1F4E79'))
    fig.add_trace(go.Bar(name='القص الديناميكي (Dynamic)', x=directions, y=dynamic_vals, marker_color='#D9534F'))
    fig.add_trace(go.Bar(name=f'الحد الأدنى المطلوب ({target_pct*100:.0f}%)', x=directions, y=required_vals, marker_color='#F0AD4E'))
    fig.update_layout(title="مقارنة قوى القص القاعدي الساكن والديناميكي", yaxis_title="قوة القص (kN)", barmode='group', template="plotly_white", height=450)
    return fig


def plot_story_drifts(drifts_df, allowable_drift=0.020):
    x_col = next((c for c in drifts_df.columns if "drift" in c.lower() or "ratio" in c.lower()), drifts_df.columns[min(1, len(drifts_df.columns) - 1)])
    y_col = next((c for c in drifts_df.columns if "story" in c.lower()), drifts_df.columns[0])

    fig = px.line(drifts_df, x=x_col, y=y_col, markers=True, title="تحليل انجرافات الطوابق النسبية (Story Drifts)")
    fig.add_vline(x=allowable_drift, line_dash="dash", line_color="red", annotation_text=f"حد مرجعي ({allowable_drift})")
    fig.update_layout(xaxis_title="نسبة الانجراف (Drift Ratio)", yaxis_title="الطابق", template="plotly_white", height=500)
    return fig


def plot_story_stiffness(stiff_df):
    x_col = next((c for c in stiff_df.columns if "kx" in c.lower() or "stiff" in c.lower()), stiff_df.columns[min(1, len(stiff_df.columns) - 1)])
    y_col = next((c for c in stiff_df.columns if "story" in c.lower()), stiff_df.columns[0])

    fig = px.line(stiff_df, x=x_col, y=y_col, markers=True, title="تغير صلابة الطوابق (Story Stiffness Analysis)")
    fig.update_layout(xaxis_title="الصلابة الجانبية", yaxis_title="الطابق", template="plotly_white", height=450)
    return fig


# ==========================================
# دوال التصدير
# ==========================================
def export_audit_results_to_excel(project_info, statuses, centers_df, base_reactions_df, pdelta_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_summary = pd.DataFrame({"بيان المشروع": list(project_info.keys()), "القيمة": list(project_info.values())})
        df_summary.to_excel(writer, sheet_name="ملخص المشروع", index=False)
        df_statuses = pd.DataFrame([{"بند التدقيق": k, "حالة التحقق": v} for k, v in statuses.items()])
        df_statuses.to_excel(writer, sheet_name="حالة الفحوصات", index=False)
        if isinstance(centers_df, pd.DataFrame) and not centers_df.empty:
            centers_df.to_excel(writer, sheet_name="مراكز الكتلة والصلابة", index=False)
        if isinstance(base_reactions_df, pd.DataFrame) and not base_reactions_df.empty:
            base_reactions_df.to_excel(writer, sheet_name="القص القاعدي", index=False)
        if isinstance(pdelta_df, pd.DataFrame) and not pdelta_df.empty:
            pdelta_df.to_excel(writer, sheet_name="فحص P-Delta", index=False)
    output.seek(0)
    return output


def generate_word_report(project_name, project_location, consultant_name, contact_info, building_usage,
                          num_stories, total_height, fc_vert, fc_horiz, rebar, zone_z, soil_factor_s,
                          importance_factor, r_val, cd_val, status_dict, df_comp, fig_list=None,
                          methodology_notes=None):
    doc = Document()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"نوتة حسابية وتقارير التدقيق الإنشائي - {project_name}")
    title_run.font.name = 'Tahoma'
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 121)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"الموقع: {project_location} | المصمم: {consultant_name} | هاتف: {contact_info}")
    sub_run.font.name = 'Tahoma'
    sub_run.font.size = Pt(11)

    doc.add_paragraph("---")
    doc.add_heading("أولاً: بيانات المشروع والخواص الهندسية والمواد", level=2)
    p_info = doc.add_paragraph()
    p_info.add_run(
        f"• طبيعة الاستخدام: {building_usage}\n"
        f"• عدد الطوابق: {num_stories} طوابق (الارتفاع الكلي: {total_height} متر)\n"
        f"• خرسانة العناصر الشاقولية: {fc_vert} | الأفقية: {fc_horiz}\n"
        f"• حديد التسليح: {rebar}\n"
    )

    doc.add_heading("ثانياً: المعاملات الزلزالية والأحمال التصميمية", level=2)
    p_seismic = doc.add_paragraph()
    p_seismic.add_run(
        f"• معامل المنطقة الزلزالية (Z): {zone_z} | معامل التربة (S): {soil_factor_s}\n"
        f"• معامل الأهمية (Ie): {importance_factor} | معامل R: {r_val} | عامل Cd: {cd_val}\n"
    )

    doc.add_heading("ثالثاً: ملخص حالة التدقيق لأقسام النموذج (ETABS)", level=2)
    table_status = doc.add_table(rows=1, cols=2)
    table_status.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_status.rows[0].cells[0].text = "بند التدقيق (Check Item)"
    table_status.rows[0].cells[1].text = "حالة التحقق (Status)"
    for item, status in status_dict.items():
        row_cells = table_status.add_row().cells
        row_cells[0].text = str(item)
        row_cells[1].text = str(status)
    doc.add_paragraph()

    doc.add_heading("رابعاً: مقارنة وزيادة القص القاعدي (Base Reactions & Scale Factors)", level=2)
    if df_comp is not None and not df_comp.empty:
        table_comp = doc.add_table(rows=1, cols=len(df_comp.columns))
        table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, col_name in enumerate(df_comp.columns):
            table_comp.rows[0].cells[i].text = str(col_name)
        for _, row in df_comp.iterrows():
            row_cells = table_comp.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
    else:
        doc.add_paragraph("جدول مقارنة القص القاعدي غير متوفر.")
    doc.add_paragraph()

    if fig_list:
        doc.add_heading("خامساً: التمثيل البصري والرسوم البيانية التحليلية", level=2)
        for caption, fig in fig_list:
            try:
                img_bytes = fig.to_image(format="png", width=900, height=500, scale=2)
                image_stream = io.BytesIO(img_bytes)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(image_stream, width=Inches(6.0))
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_cap = p_cap.add_run(f"شكل توضيحي: {caption}")
                run_cap.font.name = "Tahoma"
                run_cap.font.size = Pt(10)
                run_cap.font.italic = True
                doc.add_paragraph()
            except Exception:
                pass

    if methodology_notes:
        doc.add_heading("سادساً: ملاحظات على منهجية الحساب والمعادلات المستخدمة", level=2)
        for note in methodology_notes:
            doc.add_paragraph(note, style="List Bullet")

    doc.add_heading("سابعاً: الخلاصة والتوصيات الهندسية", level=2)
    p_concl = doc.add_paragraph()
    p_concl.add_run(
        "هذا التقرير يلخص الفحوص الحسابية الأولية المتاحة في الملف المرفوع وفق المعادلات "
        "الموضحة أعلاه. الحالات المصنّفة \"غير محسوب\" لم تُقيَّم لعدم توفر بيانات أو حد قبول "
        "معتمد في هذا التطبيق، ولا تمثل اجتيازاً للفحص. لا يُغني هذا التقرير عن مراجعة "
        "واعتماد مهندس إنشائي مؤهل."
    )

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# ==========================================
# 🎛️ الشريط الجانبي: بيانات المشروع والإعدادات (قبل أي معالجة)
# ==========================================
master_excel_path = get_excel_file_path()

# حاوية فارغة في أعلى الشريط الجانبي؛ تُملأ لاحقاً بقائمة التنقل بعد معالجة
# الملف وحساب حالة كل قسم، لكنها تبقى ظاهرة في أعلى الشريط الجانبي بصرياً.
nav_placeholder = st.sidebar.empty()
st.sidebar.markdown("---")

st.sidebar.header("📋 بيانات المشروع الأساسية")
project_name = st.sidebar.text_input("اسم المشروع", value="مبنى تجاري")
project_location = st.sidebar.text_input("موقع المشروع", value="العراق - الموصل")
consultant_name = st.sidebar.text_input("المصمم الانشائي", value="عبدالغفور رمضان عيدان")
contact_info = st.sidebar.text_input("معلومات الاتصال", value="+9647740884754")

building_usage = st.sidebar.selectbox("طبيعة الاستخدام (Use Type)", [
    "مبنى إداري", "مبنى سكني", "مبنى تجاري", "منشأ تعليمي / خدمي", "منشأ صناعي / مستودع"
])

ground_floor_height = st.sidebar.number_input("ارتفاع الطابق الأرضي (م)", value=4.5, step=0.5, min_value=0.0)
typical_floor_height = st.sidebar.number_input("ارتفاع الطابق المتكرر (م)", value=3.5, step=0.5, min_value=0.0)
num_typical_floors = st.sidebar.number_input("عدد الطوابق المتكررة", value=3, step=1, min_value=0)
total_height = ground_floor_height + num_typical_floors * typical_floor_height
num_stories = num_typical_floors + 1
st.sidebar.caption(f"الارتفاع الإجمالي المحسوب: {total_height:.2f} م | عدد الطوابق الكلي: {num_stories}")

st.sidebar.markdown("---")
st.sidebar.header("🧱 خواص المواد الإنشائية")
fc_vertical = st.sidebar.selectbox("مقاومة خرسانة العناصر الشاقولية (أعمدة/جدران)", ["f'c = 30 MPa", "f'c = 35 MPa", "f'c = 40 MPa", "f'c = 45 MPa"])
fc_horizontal = st.sidebar.selectbox("مقاومة خرسانة العناصر الأفقية (بلاطات/جسور)", ["f'c = 25 MPa", "f'c = 28 MPa", "f'c = 30 MPa", "f'c = 35 MPa"])
rebar_grade = st.sidebar.selectbox("حديد التسليح الرئيسي", [
    "Grade 60 (fy = 420 MPa)", "Grade 40 (fy = 280 MPa)", "Grade 50 (fy = 350 MPa)", "Grade 75 (fy = 520 MPa)"
])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ إعدادات الزلازل والتربة (مدونة العراق)")
zone_options = {
    "المنطقة الزلزالية الثانية (الموصل / نينوى) - Z = 0.15": 0.15,
    "المنطقة الزلزالية الأولى (بغداد / وسط العراق) - Z = 0.10": 0.10,
    "المنطقة الزلزالية الثالثة (الشمال الشرقي) - Z = 0.20": 0.20,
}
selected_zone = st.sidebar.selectbox("المنطقة الزلزالية (Z)", list(zone_options.keys()))
zone_z = zone_options[selected_zone]

soil_options = {
    "النوع (II): تربة متوسطة / كثيفة (Stiff Soil - S_C)": 1.2,
    "النوع (I): صخور صلبة (Rock - S_A/S_B)": 1.0,
    "النوع (III): تربة طينية سميكة (Soft Soil - S_D)": 1.5,
}
selected_soil = st.sidebar.selectbox("نوع التربة (Soil Profile)", list(soil_options.keys()))
soil_factor_s = soil_options[selected_soil]

risk_category_options = {
    "فئة المخاطر I أو II (معظم المنشآت الاعتيادية)": {"ie": 1.0, "drift_limit": 0.020},
    "فئة المخاطر III (تجمعات كبيرة / خطورة متوسطة)": {"ie": 1.25, "drift_limit": 0.015},
    "فئة المخاطر IV (منشآت حرجة / سلامة عامة)": {"ie": 1.5, "drift_limit": 0.010},
}
selected_risk_category = st.sidebar.selectbox(
    "فئة المخاطر (Risk Category) — ASCE 7-10 جدولا 1.5-2 و12.12-1",
    list(risk_category_options.keys()),
)
importance_factor = risk_category_options[selected_risk_category]["ie"]
base_drift_limit = risk_category_options[selected_risk_category]["drift_limit"]

redundancy_options = {
    "ρ = 1.0 (تحقق شروط التخفيض في §12.3.4)": 1.0,
    "ρ = 1.3 (الافتراضي عند عدم تحقق الشروط)": 1.3,
}
selected_redundancy = st.sidebar.selectbox(
    "عامل التكرار ρ (Redundancy Factor) — ASCE 7-10 §12.3.4",
    list(redundancy_options.keys()),
)
redundancy_rho = redundancy_options[selected_redundancy]

# SDC منفصلة عمداً عن فئة المخاطر: تشتق من SDS/SD1 وليس من فئة المخاطر وحدها،
# ووجوب تكبير Ax (§12.8.4.3) مرتبط بـ SDC تحديداً لا بفئة المخاطر.
seismic_design_category = st.sidebar.selectbox(
    "فئة التصميم الزلزالي SDC — ASCE 7-10 §11.6 (لتطبيق Ax فقط، §12.8.4.3)",
    ["A", "B", "C", "D", "E", "F"],
    index=3,
)

seismic_systems = {
    "جدران قص خرسانية مسلحة [R=5.0, Cd=4.5]": (5.0, 4.5),
    "إطارات خرسانية مقاومة للعزوم خاصة [R=8.0, Cd=5.5]": (8.0, 5.5),
    "نظام مزدوج (Dual System) [R=7.0, Cd=5.5]": (7.0, 5.5),
}
selected_system = st.sidebar.selectbox("النظام الإنشائي المقاوم للزلازل", list(seismic_systems.keys()))
response_mod_r, deflection_cd = seismic_systems[selected_system]

asce_percentage = st.sidebar.slider("النسبة المطلوبة للقص الديناميكي من الاستاتيكي (%)", 80.0, 100.0, 100.0, 5.0) / 100.0

st.sidebar.markdown("---")
st.sidebar.header("📐 أبعاد المبنى الإجمالية")
dx_val = st.sidebar.number_input("البعد الأفقي (Dx - ملم)", value=16000.0, step=100.0)
dy_val = st.sidebar.number_input("البعد العمودي (Dy - ملم)", value=25000.0, step=100.0)

# ==========================================
# معالجة القراءة وإجراء الفحوص (مرة واحدة، قبل بناء قائمة التنقل)
# ==========================================
status_dict = {
    "Centers Of Mass": "⏳ بانتظار التحميل",
    "Base Reactions": "⏳ بانتظار التحميل",
    "Story Drifts": "⏳ بانتظار التحميل",
    "Modal Mass Ratios": "⏳ بانتظار التحميل",
    "Displacements": "⏳ بانتظار التحميل",
    "Story Stiffness": "⏳ بانتظار التحميل",
    "P-Delta Check": "⏳ بانتظار التحميل",
    "Wind Drift": "⏳ بانتظار التحميل",
}

res_df = pd.DataFrame()
df_comp = pd.DataFrame()
df_pdelta_filtered = pd.DataFrame()
df_drifts_display = pd.DataFrame()
df_modal = pd.DataFrame()
df_disp = pd.DataFrame()
df_stiff = pd.DataFrame()
v_static_x = v_static_y = v_dyn_x = v_dyn_y = 0.0
centers_error = base_error = drift_note = modal_note = disp_error = disp_note = stiff_error = stiff_note = pdelta_note = None
wind_error = wind_note = None
df_wind_story = pd.DataFrame()
df_wind_total = pd.DataFrame()
figures_for_report = []
methodology_notes = []
top_level_error = None

if master_excel_path is not None:
    try:
        xls_file = pd.ExcelFile(master_excel_path)
        sheet_names = xls_file.sheet_names

        # 1. مراكز الكتلة والصلابة
        try:
            t1 = get_sheet_by_name(sheet_names, "Check_centers_of_mass_and_rigidity")
            df_rigidity = pd.read_excel(master_excel_path, sheet_name=t1, index_col=None)
            df_rigidity = clean_dataframe(df_rigidity)
            df_rigidity = strip_and_dedupe_columns(df_rigidity)
            story_col = next((c for c in df_rigidity.columns if "story" in c.lower()), df_rigidity.columns[0])
            diaph_col = next((c for c in df_rigidity.columns if "diaph" in c.lower()), df_rigidity.columns[1] if len(df_rigidity.columns) > 1 else df_rigidity.columns[0])
            xr_col = next((c for c in df_rigidity.columns if "x" in c.lower() and ("rig" in c.lower() or "1" in c.lower())), df_rigidity.columns[2] if len(df_rigidity.columns) > 2 else df_rigidity.columns[0])
            yr_col = next((c for c in df_rigidity.columns if "y" in c.lower() and ("rig" in c.lower() or "2" in c.lower())), df_rigidity.columns[3] if len(df_rigidity.columns) > 3 else df_rigidity.columns[0])
            xm_col = next((c for c in df_rigidity.columns if "x" in c.lower() and "mas" in c.lower()), df_rigidity.columns[4] if len(df_rigidity.columns) > 4 else df_rigidity.columns[0])
            ym_col = next((c for c in df_rigidity.columns if "y" in c.lower() and "mas" in c.lower()), df_rigidity.columns[5] if len(df_rigidity.columns) > 5 else df_rigidity.columns[0])

            res_df["Story"] = df_rigidity[story_col].astype(str)
            res_df["Diaphragm"] = df_rigidity[diaph_col].astype(str)
            res_df["Xr"] = pd.to_numeric(df_rigidity[xr_col], errors="coerce").fillna(0).astype(float)
            res_df["Yr"] = pd.to_numeric(df_rigidity[yr_col], errors="coerce").fillna(0).astype(float)
            res_df["Xm"] = pd.to_numeric(df_rigidity[xm_col], errors="coerce").fillna(0).astype(float)
            res_df["Ym"] = pd.to_numeric(df_rigidity[ym_col], errors="coerce").fillna(0).astype(float)
            res_df["Dx"] = float(dx_val)
            res_df["Dy"] = float(dy_val)
            res_df["ex"] = np.abs(res_df["Xr"] - res_df["Xm"])
            res_df["ey"] = np.abs(res_df["Yr"] - res_df["Ym"])
            res_df["ex/Dx %"] = (res_df["ex"] / res_df["Dx"]) * 100
            res_df["ey/Dy %"] = (res_df["ey"] / res_df["Dy"]) * 100
            res_df[["Xr", "Yr", "Xm", "Ym", "ex/Dx %", "ey/Dy %"]] = res_df[["Xr", "Yr", "Xm", "Ym", "ex/Dx %", "ey/Dy %"]].round(2)
            # تأكّد بالمراجعة المباشرة لدليل ASCE 7-10 أن هذا الكود لا يضع حداً لنسبة
            # اللامركزية إلى بُعد المبنى؛ لذلك لا يصدر هذا القسم حكم نجاح/فشل، ويبقى
            # معلوماتياً فقط. تقييم الالتواء الفعلي منقول إلى قسم 5 (δmax/δavg).
            status_dict["Centers Of Mass"] = "معلوماتي — لا يصدر حكم مطابقة"
            fig_cm_cr = plot_center_of_mass_vs_rigidity(res_df)
            figures_for_report.append(("مقارنة مراكز الكتلة والصلابة عبر الطوابق", fig_cm_cr))
            methodology_notes.append(
                "مراكز الكتلة والصلابة: ex=|Xr-Xm|, ey=|Yr-Ym|؛ النسبة = 100×e/D. عرض "
                "معلوماتي فقط؛ لا يضع ASCE 7-10 حداً لهذه النسبة (تم التحقق مباشرة من الكود)."
            )
        except Exception as exc:
            status_dict["Centers Of Mass"] = "خطأ ⚠️"
            centers_error = str(exc)

        # 2. القص القاعدي
        try:
            t2 = get_sheet_by_name(sheet_names, "base_reactions")
            df_react = pd.read_excel(master_excel_path, sheet_name=t2, index_col=None)
            df_react = clean_dataframe(df_react)
            df_react = strip_and_dedupe_columns(df_react)
            r_case_col = next((c for c in df_react.columns if "case" in c.lower()), None)
            fx_col = next((c for c in df_react.columns if "fx" in c.lower() or "f_x" in c.lower()), None)
            fy_col = next((c for c in df_react.columns if "fy" in c.lower() or "f_y" in c.lower()), None)
            if r_case_col and fx_col:
                df_react["Clean_Case"] = df_react[r_case_col].astype(str).str.upper().str.strip()
                eqx_row = df_react[df_react["Clean_Case"] == "EQX"]
                eqy_row = df_react[df_react["Clean_Case"] == "EQY"] if fy_col and not df_react[df_react["Clean_Case"] == "EQY"].empty else eqx_row
                rsx_row = df_react[df_react["Clean_Case"].str.contains("RSX", na=False)]
                rsy_row = df_react[df_react["Clean_Case"].str.contains("RSY", na=False)] if fy_col and not df_react[df_react["Clean_Case"].str.contains("RSY", na=False)].empty else rsx_row
                v_static_x = abs(float(pd.to_numeric(eqx_row[fx_col], errors="coerce").values[0])) if not eqx_row.empty else 0.0
                v_static_y = abs(float(pd.to_numeric(eqy_row[fy_col if fy_col else fx_col], errors="coerce").values[0])) if not eqy_row.empty else 0.0
                v_dyn_x = abs(float(pd.to_numeric(rsx_row[fx_col], errors="coerce").abs().max())) if not rsx_row.empty else 0.0
                v_dyn_y = abs(float(pd.to_numeric(rsy_row[fy_col if fy_col else fx_col], errors="coerce").abs().max())) if not rsy_row.empty else 0.0
                status_x_val = v_dyn_x >= (v_static_x * asce_percentage) if v_static_x > 0 else None
                status_y_val = v_dyn_y >= (v_static_y * asce_percentage) if v_static_y > 0 else None
                if status_x_val is None and status_y_val is None:
                    status_dict["Base Reactions"] = "غير محسوب ⚠️"
                else:
                    overall_ok = all(v for v in (status_x_val, status_y_val) if v is not None)
                    status_dict["Base Reactions"] = "محقق ✅" if overall_ok else "يتطلب تعديل ⚠️"
                comparison_data = {
                    "Direction": ["الاتجاه الأفقي (X)", "الاتجاه العمودي (Y)"],
                    "Static Base Shear (V_st)": [round(v_static_x, 2), round(v_static_y, 2)],
                    "Dynamic Base Shear (V_dyn)": [round(v_dyn_x, 2), round(v_dyn_y, 2)],
                    "Target Ratio (%)": [asce_percentage * 100, asce_percentage * 100],
                    "Check Status": [
                        "محقق ✅" if status_x_val else ("يتطلب تعديل ⚠️" if status_x_val is False else "غير محسوب ⚠️"),
                        "محقق ✅" if status_y_val else ("يتطلب تعديل ⚠️" if status_y_val is False else "غير محسوب ⚠️"),
                    ],
                }
                df_comp = pd.DataFrame(comparison_data)
                fig_bs = plot_base_shear_comparison(v_static_x, v_static_y, v_dyn_x, v_dyn_y, asce_percentage)
                figures_for_report.append(("مقارنة القص القاعدي الساكن والديناميكي", fig_bs))
                methodology_notes.append(
                    f"القص القاعدي: يُشترط V_dynamic ≥ V_static×{asce_percentage*100:.0f}% لكل اتجاه."
                )
            else:
                status_dict["Base Reactions"] = "غير محسوب ⚠️"
        except Exception as exc:
            status_dict["Base Reactions"] = "خطأ ⚠️"
            base_error = str(exc)

        # 3. انجراف الطوابق
        try:
            t3 = get_sheet_by_name(sheet_names, "story_drifts")
            df_drifts_raw = pd.read_excel(master_excel_path, sheet_name=t3, index_col=None)
            df_drifts_raw = clean_dataframe(df_drifts_raw)
            df_drifts_raw = strip_and_dedupe_columns(df_drifts_raw)
            df_drifts_display = filter_lateral_cases(df_drifts_raw)

            drift_col = next((c for c in df_drifts_display.columns if "drift" in c.lower() or "ratio" in c.lower()), None)

            if drift_col is not None:
                work = df_drifts_display.copy()
                work["_elastic_abs"] = pd.to_numeric(work[drift_col], errors="coerce").abs()
                # Eq. 12.8-15: Δ = Cd·Δxe/Ie — الانجراف المكبّر (التصميمي) من الانجراف المرن.
                work["Design Drift (Δ=Cd·Δxe/Ie)"] = work["_elastic_abs"] * deflection_cd / importance_factor
                # Table 12.12-1: Δa = Δa,base/ρ — لا علاقة للفترة T أو R بهذا الحد إطلاقاً.
                drift_limit = base_drift_limit / redundancy_rho
                max_design_drift = work["Design Drift (Δ=Cd·Δxe/Ie)"].max(skipna=True)
                df_drifts_display = work

                if pd.notna(max_design_drift):
                    overall = bool(max_design_drift <= drift_limit)
                    status_dict["Story Drifts"] = "محقق ✅" if overall else "غير محقق ❌"
                    drift_note = (
                        f"أقصى انجراف تصميم Δ = {max_design_drift:.5f} مقابل الحد "
                        f"Δa/ρ = {drift_limit:.5f} (جدول 12.12-1، ρ={redundancy_rho:g})."
                    )
                    methodology_notes.append(f"انجراف الطوابق: {drift_note}")
                else:
                    status_dict["Story Drifts"] = "غير محسوب ⚠️"
                    drift_note = "تعذر استخراج قيم انجراف رقمية صالحة من العمود المكتشف."
                try:
                    fig_drift = plot_story_drifts(df_drifts_display, allowable_drift=drift_limit)
                    figures_for_report.append(("انجرافات الطوابق النسبية", fig_drift))
                except Exception:
                    pass
            else:
                status_dict["Story Drifts"] = "غير محسوب ⚠️"
                drift_note = "تعذر تحديد عمود الانجراف في الجدول."
        except Exception as exc:
            status_dict["Story Drifts"] = "خطأ ⚠️"
            drift_note = f"خطأ أثناء قراءة/معالجة جدول انجراف الطوابق: {exc}"

        # 4. مشاركة الكتلة النمطية
        try:
            t4 = get_sheet_by_name(sheet_names, "modal_participating_mass_ratios")
            df_modal = pd.read_excel(master_excel_path, sheet_name=t4, index_col=None)
            df_modal = clean_dataframe(df_modal)
            df_modal = strip_and_dedupe_columns(df_modal)
            # §12.9.1 يشترط 90% تحديداً لمشاركة الكتلة الأفقية في كل من X وY
            # (SumUX, SumUY). أعمدة أخرى تحوي كلمة "Sum" مثل SumUZ (رأسي، صفر
            # عادةً) وSumRX/SumRY/SumRZ (دورانية) لا علاقة لها بهذا الشرط، وكان
            # تضمينها سابقاً يُفشل الفحص زوراً رغم أن SumUX/SumUY كانا محققين فعلاً.
            sum_cols = [c for c in df_modal.columns if c.strip().lower() in ("sumux", "sumuy")]
            if sum_cols:
                max_vals = {}
                for c in sum_cols:
                    numeric = pd.to_numeric(df_modal[c], errors="coerce")
                    if numeric.max(skipna=True) is not None and pd.notna(numeric.max(skipna=True)) and numeric.max(skipna=True) <= 1.0:
                        numeric = numeric * 100
                    max_vals[c] = numeric.max(skipna=True)
                valid_vals = {c: v for c, v in max_vals.items() if pd.notna(v)}
                if valid_vals:
                    modal_pass = all(v >= 90.0 for v in valid_vals.values())
                    status_dict["Modal Mass Ratios"] = "محقق ✅" if modal_pass else "غير محقق ❌"
                    modal_note = " | ".join(f"{c}: {v:.1f}%" for c, v in valid_vals.items())
                    methodology_notes.append(f"مشاركة الكتلة النمطية: {modal_note} (الحد 90%).")
                else:
                    status_dict["Modal Mass Ratios"] = "غير محسوب ⚠️"
                    modal_note = "لم تتوفر قيم رقمية صالحة في أعمدة SumUX/SumUY."
            else:
                status_dict["Modal Mass Ratios"] = "غير محسوب ⚠️"
                modal_note = "تعذر إيجاد أعمدة SumUX/SumUY في الجدول."
        except Exception as exc:
            status_dict["Modal Mass Ratios"] = "خطأ ⚠️"
            modal_note = f"خطأ أثناء قراءة جدول مشاركة الكتلة: {exc}"

        # 5. الالتواء العرضي الفعلي وفق جدول 12.3-1 (من نسبة الانجراف) ومعادلة Ax
        # (من نسبة الإزاحة) — الدليل ينص حرفياً أن كمية Ax "based on story
        # displacements, not interstory drifts"، لذا المصدران مختلفان عمداً.
        def _find_ratio_col(df):
            return next(
                (c for c in df.columns
                 if "ratio" in c.lower()
                 or "max/avg" in c.lower().replace(" ", "")
                 or ("max" in c.lower() and "avg" in c.lower())),
                None,
            )

        try:
            t5d = get_sheet_by_name(sheet_names, "story_max_over_avg_drifts")
            df_drift_ratio = pd.read_excel(master_excel_path, sheet_name=t5d, index_col=None)
            df_drift_ratio = clean_dataframe(df_drift_ratio)
            df_drift_ratio = strip_and_dedupe_columns(df_drift_ratio)
            df_drift_ratio = filter_lateral_cases(df_drift_ratio)
            ratio_col_drift = _find_ratio_col(df_drift_ratio)

            t5v = get_sheet_by_name(sheet_names, "story_max_over_avg_displacements")
            df_disp = pd.read_excel(master_excel_path, sheet_name=t5v, index_col=None)
            df_disp = clean_dataframe(df_disp)
            df_disp = strip_and_dedupe_columns(df_disp)
            df_disp_lateral = filter_lateral_cases(df_disp)
            ratio_col_disp = _find_ratio_col(df_disp_lateral)

            max_drift_ratio = (
                pd.to_numeric(df_drift_ratio[ratio_col_drift], errors="coerce").max(skipna=True)
                if ratio_col_drift is not None else None
            )
            max_disp_ratio = (
                pd.to_numeric(df_disp_lateral[ratio_col_disp], errors="coerce").max(skipna=True)
                if ratio_col_disp is not None else None
            )

            if max_drift_ratio is not None and pd.notna(max_drift_ratio):
                if max_drift_ratio <= 1.2:
                    status_dict["Displacements"] = "محقق ✅ (لا يوجد التواء عرضي)"
                elif max_drift_ratio <= 1.4:
                    status_dict["Displacements"] = "يتطلب تطبيق Ax (التواء Type 1a) ⚠️"
                else:
                    status_dict["Displacements"] = "غير محقق ❌ (التواء حاد Type 1b)"
                ax_text = ""
                applicability_text = ""
                if max_drift_ratio > 1.2:
                    ax_value = None
                    if max_disp_ratio is not None and pd.notna(max_disp_ratio):
                        ax_value = min((max_disp_ratio / 1.2) ** 2, 3.0)
                        ax_text = f" | Ax المرجعي (من نسبة الإزاحة {max_disp_ratio:.3f}) = {ax_value:.3f}"

                    # اكتشاف طريقة إدخال الالتواء العرضي فعلياً من بيانات القص
                    # القاعدي: حالات ساكنة (LinStatic) متعددة لكل اتجاه تعني أن
                    # الالتواء العرضي أُدخل كحمل ساكن منفصل (مثل EQx1/EQx2)، وهذا
                    # تحديداً ما يستوجب تكبير Ax وفق §12.8.4.3 عند SDC C وما فوق.
                    try:
                        t_br = get_sheet_by_name(sheet_names, "base_reactions")
                        df_br = pd.read_excel(master_excel_path, sheet_name=t_br, index_col=None)
                        df_br = strip_and_dedupe_columns(df_br)
                        case_col_br = next((c for c in df_br.columns if "case" in c.lower() and "type" not in c.lower()), None)
                        type_col_br = next((c for c in df_br.columns if "case type" in c.lower() or c.lower() == "case type"), None)
                        static_x_variants = static_y_variants = dynamic_without_torsion = False
                        if case_col_br is not None and type_col_br is not None:
                            is_static = df_br[type_col_br].astype(str).str.upper().str.contains("STATIC", na=False)
                            is_dynamic = df_br[type_col_br].astype(str).str.upper().str.contains("RESPSPEC|MODAL", na=False)
                            cu = df_br[case_col_br].astype(str).str.upper()
                            x_static_names = cu[is_static & cu.str.contains("X", na=False) & ~cu.str.contains("Y", na=False)].unique()
                            y_static_names = cu[is_static & cu.str.contains("Y", na=False) & ~cu.str.contains("X", na=False)].unique()
                            static_x_variants = len(x_static_names) > 1
                            static_y_variants = len(y_static_names) > 1
                            dynamic_without_torsion = bool(is_dynamic.any())
                        static_method_detected = static_x_variants or static_y_variants

                        if seismic_design_category in ("A", "B"):
                            applicability_text = (
                                f" Ax غير مطلوب أصلاً (SDC={seismic_design_category} — §12.8.4.3 "
                                "يشترط C فما فوق)."
                            )
                        elif static_method_detected:
                            extra = (
                                " (مع ملاحظة: وُجد أيضاً تحليل ديناميكي RS بلا حالات التواء "
                                "مرافقة، وهو لا يُدرج الالتواء العرضي أصلاً وفق §12.9.5؛ "
                                "التكبير يخص فقط الحالات الساكنة ذات الالتواء مثل EQx1/EQx2.)"
                                if dynamic_without_torsion else ""
                            )
                            applicability_text = (
                                f" Ax **مطلوب تطبيقه فعلياً** (SDC={seismic_design_category}، "
                                "والالتواء العرضي مُدخَل كحمل ساكن منفصل في النموذج — تأكدت من "
                                "وجود حالات EQx1/EQx2 أو ما يماثلها). لا يتحقق هذا التطبيق من أن "
                                f"النموذج استخدم القيمة الصحيحة لـ Ax فعلياً؛ يلزم تدقيق ذلك يدوياً "
                                f"في تعريف حالات التحميل.{extra}"
                            )
                        else:
                            applicability_text = (
                                f" تعذّر تأكيد طريقة إدخال الالتواء العرضي (SDC={seismic_design_category})؛ "
                                "إن كانت ديناميكية بإزاحة فعلية لمركز الكتلة (§12.9.5) فلا تكبير مطلوب، "
                                "وإلا فيلزم تطبيق Ax يدوياً."
                            )
                    except Exception:
                        applicability_text = ""

                disp_note = (
                    f"أقصى نسبة انجراف الحافة/المتوسط = {max_drift_ratio:.3f} (حالات زلزالية فقط، "
                    f"العمود '{ratio_col_drift}')؛ الحدود: 1.2 (Type 1a) و1.4 (Type 1b){ax_text}."
                    f"{applicability_text}"
                )
                methodology_notes.append(f"الالتواء العرضي (جدول 12.3-1، §12.8.4.3): {disp_note}")
            else:
                status_dict["Displacements"] = "غير محسوب ⚠️"
                disp_note = (
                    "تعذر تحديد عمود نسبة الانجراف (δmax/δavg) في جدول "
                    "'Story Max Over Avg Drifts'، أو لم تتبقَّ حالات زلزالية بعد الفلترة."
                )
            df_disp = df_disp_lateral
        except Exception as exc:
            status_dict["Displacements"] = "خطأ ⚠️"
            disp_error = str(exc)

        # 6. صلابة الطوابق — فحص الطابق اللين الفعلي وفق جدول ASCE 7-10 رقم 12.3-2
        try:
            t6 = get_sheet_by_name(sheet_names, "story_stiffness")
            df_stiff = pd.read_excel(master_excel_path, sheet_name=t6, index_col=None)
            df_stiff = clean_dataframe(df_stiff)
            df_stiff = strip_and_dedupe_columns(df_stiff)

            story_col_s = next((c for c in df_stiff.columns if c.lower() == "story" or "story" in c.lower()), None)
            case_col_s = next((c for c in df_stiff.columns if "case" in c.lower() or "output" in c.lower()), None)
            stiff_x_col = next((c for c in df_stiff.columns if "stiff" in c.lower() and "x" in c.lower()), None)
            stiff_y_col = next((c for c in df_stiff.columns if "stiff" in c.lower() and "y" in c.lower()), None)

            soft_story_findings = []
            worst_classification = "none"  # none < soft < extreme

            def _per_story_stiffness(direction_letter, stiff_col):
                if story_col_s is None or case_col_s is None or stiff_col is None:
                    return None
                case_upper = df_stiff[case_col_s].astype(str).str.upper()
                dir_mask = case_upper.str.contains(direction_letter, na=False) & \
                    ~case_upper.str.contains("X" if direction_letter == "Y" else "Y", na=False)
                lateral_mask = case_upper.str.contains("EQ|RS|WIND|SPEC", na=False)
                sub = df_stiff[dir_mask & lateral_mask].copy()
                if sub.empty:
                    return None
                sub["_k"] = pd.to_numeric(sub[stiff_col], errors="coerce").abs()
                # يُحتفظ بأول قيمة لكل طابق (ترتيب الورقة الأصلي من الأعلى إلى الأسفل
                # كما تصدّره ETABS افتراضياً)، لتفادي اضطراب الترتيب الأبجدي للتجميع.
                per_story = sub.dropna(subset=["_k"]).drop_duplicates(subset=[story_col_s], keep="first")
                return list(zip(per_story[story_col_s], per_story["_k"]))

            for direction_letter, stiff_col, dir_label in (("X", stiff_x_col, "X"), ("Y", stiff_y_col, "Y")):
                series = _per_story_stiffness(direction_letter, stiff_col)
                if not series:
                    continue
                # الترتيب في الورقة من الأعلى إلى الأسفل؛ الطابقان الأعلى فقط
                # (i=0,1) لا يُقيَّمان، ويُفحص كل ما تبقى نزولاً حتى الطابق الأرضي.
                for i in range(2, len(series)):
                    story_i, k_i = series[i]  # الطابق قيد الفحص
                    _, k_above = series[i - 1]  # الطابق الذي يعلوه مباشرة
                    above3 = [series[j][1] for j in range(max(0, i - 3), i)]
                    avg_above3 = sum(above3) / len(above3)
                    if k_above <= 0 or avg_above3 <= 0:
                        continue
                    r1 = k_i / k_above
                    r2 = k_i / avg_above3
                    if r1 < 0.60 or r2 < 0.70:
                        worst_classification = "extreme"
                        soft_story_findings.append(f"{story_i} ({dir_label}): Ki/Kabove={r1:.2f}، Ki/avg3={r2:.2f} — Type 1b حاد")
                    elif r1 < 0.70 or r2 < 0.80:
                        if worst_classification != "extreme":
                            worst_classification = "soft"
                        soft_story_findings.append(f"{story_i} ({dir_label}): Ki/Kabove={r1:.2f}، Ki/avg3={r2:.2f} — Type 1a")

            if worst_classification == "none" and (stiff_x_col is None and stiff_y_col is None):
                status_dict["Story Stiffness"] = "غير محسوب ⚠️"
                stiff_note = "تعذر تحديد أعمدة الصلابة (Stiff X/Stiff Y) في هذا الجدول."
            elif worst_classification == "extreme":
                status_dict["Story Stiffness"] = "غير محقق ❌ (طابق لين حاد Type 1b)"
                stiff_note = "؛ ".join(soft_story_findings)
                methodology_notes.append(f"الطابق اللين (جدول 12.3-2): {stiff_note}")
            elif worst_classification == "soft":
                status_dict["Story Stiffness"] = "يتطلب مراجعة (طابق لين Type 1a) ⚠️"
                stiff_note = "؛ ".join(soft_story_findings)
                methodology_notes.append(f"الطابق اللين (جدول 12.3-2): {stiff_note}")
            else:
                status_dict["Story Stiffness"] = "محقق ✅ (لا يوجد طابق لين)"
                stiff_note = "لا توجد نسبة صلابة بين طابق وما يعلوه أقل من الحدود الكودية (70%/80%)."

            try:
                fig_stiff = plot_story_stiffness(df_stiff)
                figures_for_report.append(("تغير صلابة الطوابق", fig_stiff))
            except Exception:
                pass
        except Exception as exc:
            status_dict["Story Stiffness"] = "خطأ ⚠️"
            stiff_error = str(exc)

        # 7. فحص P-Delta
        try:
            t7 = get_sheet_by_name(sheet_names, "P_DELTA_CHECK")
            df_pdelta_raw = pd.read_excel(master_excel_path, sheet_name=t7, index_col=None)
            df_pdelta_raw = clean_dataframe(df_pdelta_raw)
            df_pdelta_raw = strip_and_dedupe_columns(df_pdelta_raw)
            df_pdelta_filtered = filter_lateral_cases(df_pdelta_raw)

            theta_col = next(
                (c for c in df_pdelta_filtered.columns if any(k in c.lower() for k in ["theta", "θ", "stability", "p-delta", "pdelta"])),
                None,
            )
            if theta_col is not None:
                theta_numeric = pd.to_numeric(df_pdelta_filtered[theta_col], errors="coerce")
                if theta_numeric.notna().any():
                    max_theta = float(theta_numeric.max())
                    theta_max_limit = min(0.25, 0.5 / deflection_cd)
                    if max_theta <= 0.1:
                        status_dict["P-Delta Check"] = "محقق ✅ (يمكن الإهمال)"
                    elif max_theta <= theta_max_limit:
                        status_dict["P-Delta Check"] = "يتطلب إدراج P-Delta ⚠️"
                    else:
                        status_dict["P-Delta Check"] = "تجاوز الحد ❌"
                    pdelta_note = f"أقصى θ = {max_theta:.4f} (من العمود '{theta_col}') مقابل θmax = {theta_max_limit:.4f}."
                    methodology_notes.append(f"P-Delta: {pdelta_note}")
                else:
                    status_dict["P-Delta Check"] = "غير محسوب ⚠️"
                    pdelta_note = f"العمود المكتشف '{theta_col}' لا يحتوي قيماً رقمية صالحة."
            else:
                status_dict["P-Delta Check"] = "غير محسوب ⚠️"
                pdelta_note = "لا يحتوي الجدول على عمود يحدد نسبة الاستقرار θ صراحة."
        except Exception as exc:
            status_dict["P-Delta Check"] = "خطأ ⚠️"
            pdelta_note = f"خطأ أثناء قراءة جدول P-Delta: {exc}"

        # 8. انجراف الرياح — فحص خدمي منفصل تماماً عن ASCE 7-10 الزلزالي (Chapter
        # 12)؛ أُكِّد من المستخدم أن "Drift Ratio X/Y" يمثل الانجراف النسبي بين
        # الطابق وما يليه (interstory)، و"Displacement X/Y" يمثل الانجراف الكلي
        # للمبنى (من القاعدة حتى كل طابق). كلاهما بحد شائع H/360.
        try:
            def _two_row_header(sheet_name):
                raw = pd.read_excel(master_excel_path, sheet_name=sheet_name, header=None)
                # تحويل آمن للنص: بعض خلايا NaN لا تتحول فعلياً عبر astype(str)
                # في بيئات pandas معينة وتبقى float، فنتعامل معها صراحة هنا.
                row1 = ["" if pd.isna(v) else str(v) for v in raw.iloc[1]]
                data = raw.iloc[2:].reset_index(drop=True)
                return row1, data

            def _find_col(row1, needle):
                for idx, val in enumerate(row1):
                    if needle.strip() in val.strip():
                        return idx
                return None

            wind_findings = []
            worst_wind = "none"  # none < soft-fail-like escalation: none -> fail
            any_nonzero_data = False

            # --- 8أ: الانجراف النسبي (interstory) من Drift Ratio X/Y ---
            for sheet_key, dir_label in (("drift_ratio_x", "X"), ("drift_ratio_y", "Y")):
                try:
                    t_w = get_sheet_by_name(sheet_names, sheet_key)
                    row1, data = _two_row_header(t_w)
                    ratio_idx = _find_col(row1, "* 0.7")
                    allow_idx = _find_col(row1, "∆a")
                    if ratio_idx is None or allow_idx is None:
                        continue
                    ratio_vals = pd.to_numeric(data.iloc[:, ratio_idx], errors="coerce").abs()
                    allow_vals = pd.to_numeric(data.iloc[:, allow_idx], errors="coerce")
                    story_vals = data.iloc[:, 0]
                    df_wind_story = pd.concat(
                        [df_wind_story, pd.DataFrame({
                            "الاتجاه": dir_label, "الطابق": story_vals,
                            "النسبة (مع عامل 0.7)": ratio_vals, "المسموح": allow_vals,
                        })],
                        ignore_index=True,
                    )
                    if ratio_vals.notna().any() and ratio_vals.max(skipna=True) > 0:
                        any_nonzero_data = True
                    max_r = ratio_vals.max(skipna=True)
                    a_r = allow_vals.max(skipna=True)
                    if pd.notna(max_r) and pd.notna(a_r) and a_r > 0:
                        if max_r > a_r:
                            worst_wind = "fail"
                            wind_findings.append(f"انجراف نسبي {dir_label}: {max_r:.5f} > المسموح {a_r:.5f}")
                except Exception:
                    continue

            # --- 8ب: الانجراف الكلي (من القاعدة) من Displacement X/Y، بإعادة
            # حساب الارتفاع التراكمي بنفسي من أعمدة H لكل طابق، متجاوزاً عمود
            # "Cumulative" الموجود في الملف (وجدته غير متسق/معطوباً سابقاً). ---
            for sheet_key, dir_label in (("displacement_x", "X"), ("displacement_y", "Y")):
                try:
                    t_w = get_sheet_by_name(sheet_names, sheet_key)
                    row1, data = _two_row_header(t_w)
                    h_idx = 1  # العمود الثاني: H(m) الخاص بكل طابق (وليس عمود Cumulative المعطوب)
                    disp_idx = _find_col(row1, "δ")
                    if disp_idx is None:
                        continue
                    heights = pd.to_numeric(data.iloc[:, h_idx], errors="coerce").tolist()
                    disp_vals = pd.to_numeric(data.iloc[:, disp_idx], errors="coerce")
                    story_vals = data.iloc[:, 0]
                    # إعادة بناء الارتفاع التراكمي من القاعدة صعوداً (الورقة
                    # مرتبة من الأعلى إلى الأسفل، فنجمع من آخر صف نحو الأول).
                    cum = [None] * len(heights)
                    running = 0.0
                    for i in range(len(heights) - 1, -1, -1):
                        h = heights[i]
                        if h is None or pd.isna(h):
                            cum[i] = None
                            continue
                        running += h
                        cum[i] = running
                    allow_recalc = [
                        (c / 360.0) if (c is not None) else None for c in cum
                    ]
                    df_wind_total = pd.concat(
                        [df_wind_total, pd.DataFrame({
                            "الاتجاه": dir_label, "الطابق": story_vals,
                            "الإزاحة (مم)": disp_vals,
                            "الارتفاع التراكمي المُعاد حسابه (مم)": cum,
                            "المسموح المُعاد حسابه = تراكمي/360 (مم)": allow_recalc,
                        })],
                        ignore_index=True,
                    )
                    if disp_vals.notna().any() and disp_vals.max(skipna=True) > 0:
                        any_nonzero_data = True
                    for i in range(len(heights)):
                        d = disp_vals.iloc[i] if i < len(disp_vals) else None
                        a = allow_recalc[i]
                        if d is not None and pd.notna(d) and a is not None and d > a:
                            worst_wind = "fail"
                            wind_findings.append(
                                f"انجراف كلي {dir_label} ({story_vals.iloc[i]}): {d:.2f} ملم > "
                                f"المسموح المُعاد حسابه {a:.2f} ملم"
                            )
                except Exception:
                    continue

            if not any_nonzero_data:
                status_dict["Wind Drift"] = "غير محسوب ⚠️"
                wind_note = (
                    "جميع قيم الانجراف/الإزاحة الفعلية في جداول Drift Ratio X/Y وDisplacement "
                    "X/Y تساوي صفراً في كل الصفوف — يبدو أن حالة تحميل الرياح غير مرتبطة فعلياً "
                    "بهذه الجداول في الملف الحالي. هذا ليس حكماً بالنجاح؛ ارفع ملفاً يحوي نتائج "
                    "فعلية لحالة الرياح لإجراء الفحص بثقة."
                )
            elif worst_wind == "fail":
                status_dict["Wind Drift"] = "غير محقق ❌"
                wind_note = "؛ ".join(wind_findings)
                methodology_notes.append(f"انجراف الرياح (فحص خدمي، H/360): {wind_note}")
            else:
                status_dict["Wind Drift"] = "محقق ✅"
                wind_note = "جميع نسب/قيم انجراف الرياح المرصودة ضمن الحد المسموح H/360."
                methodology_notes.append(f"انجراف الرياح (فحص خدمي، H/360): {wind_note}")
        except Exception as exc:
            status_dict["Wind Drift"] = "خطأ ⚠️"
            wind_error = f"خطأ أثناء قراءة جداول انجراف الرياح: {exc}"

    except Exception as err:
        top_level_error = str(err)

# ==========================================
# 🧭 تعبئة قائمة التنقل (المعلّقة في أعلى الشريط الجانبي) بحالة كل قسم
# ==========================================
nav_items = [
    ("📊 لوحة المؤشرات العامة (KPI Dashboard)", None),
    ("1. مراكز الكتلة والصلابة", "Centers Of Mass"),
    ("2. القص القاعدي ومعامل التحجيم", "Base Reactions"),
    ("3. انجراف الطوابق", "Story Drifts"),
    ("4. مشاركة الكتلة النمطية", "Modal Mass Ratios"),
    ("5. أقصى/متوسط الإزاحة", "Displacements"),
    ("6. صلابة الطوابق", "Story Stiffness"),
    ("7. فحص P-Delta", "P-Delta Check"),
    ("8. انجراف الرياح (خدمي)", "Wind Drift"),
]
nav_labels = [item[0] for item in nav_items]
nav_keys = [item[1] for item in nav_items]


def _format_nav(i):
    label = nav_labels[i]
    key = nav_keys[i]
    if key is None:
        return label
    return f"{label}   {status_suffix(status_dict[key])}"


with nav_placeholder.container():
    st.markdown("##### 🧭 أقسام التدقيق")
    selected_index = st.radio(
        "اختر القسم للعرض:",
        options=list(range(len(nav_items))),
        format_func=_format_nav,
        label_visibility="collapsed",
    )
section_navigation = nav_labels[selected_index]

# ==========================================
# 🖥️ عرض القسم المطلوب
# ==========================================
if master_excel_path is None:
    st.warning("⚠️ ملف النتائج (`result_tables.xlsx`) غير موجود في مجلد المشروع. يرجى التأكد من وجوده في نفس مسار السكريبت.")
elif top_level_error:
    st.error(f"حدث خطأ أثناء قراءة ملف النتائج: {top_level_error}")
else:
    if section_navigation == "📊 لوحة المؤشرات العامة (KPI Dashboard)":
        total_checks = len(status_dict)
        passed_count = sum(1 for s in status_dict.values() if classify_status(s) == "pass")
        failed_count = sum(1 for s in status_dict.values() if classify_status(s) == "fail")
        warn_count = sum(1 for s in status_dict.values() if classify_status(s) in ("warn", "error"))
        pending_count = sum(1 for s in status_dict.values() if classify_status(s) == "pending")

        st.markdown("### 📊 لوحة مؤشرات الأداء الإنشائي (KPI Dashboard)")
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        with kpi1:
            st.metric("إجمالي الفحوصات", f"{total_checks} فحص")
        with kpi2:
            st.metric("الفحوصات المحققة (Pass)", f"{passed_count}")
        with kpi3:
            st.metric("الفحوصات غير المحققة", f"{failed_count}", delta=f"-{failed_count} مخالفات" if failed_count > 0 else "مستقر", delta_color="inverse" if failed_count > 0 else "off")
        with kpi4:
            st.metric("تحتاج مراجعة/خطأ", f"{warn_count}")
        with kpi5:
            st.metric("غير محسوبة", f"{pending_count}")
        st.markdown("---")
        st.info("👈 استخدم القائمة الجانبية للانتقال واستعراض تفاصيل أي قسم هندسي، وستلاحظ رمز الحالة بجانب اسم كل قسم.")

    elif section_navigation == "1. مراكز الكتلة والصلابة":
        st.header("1. Centers Of Mass And Rigidity")
        render_section_equations("centers", {})
        if centers_error:
            st.error(f"تعذر حساب هذا القسم: {centers_error}")
        elif not res_df.empty:
            st.caption("عرض معلوماتي فقط — ASCE 7-10 لا يضع حداً لهذه النسبة (راجع التصحيح أعلاه).")
            display_df = res_df[["Story", "Diaphragm", "Xr", "Yr", "Xm", "Ym", "ex/Dx %", "ey/Dy %"]]
            st.dataframe(display_df, use_container_width=True)
            st.plotly_chart(plot_center_of_mass_vs_rigidity(res_df), use_container_width=True)
        else:
            st.warning("لا توجد بيانات متاحة لعرض جدول مراكز الكتلة والصلابة.")

    elif section_navigation == "2. القص القاعدي ومعامل التحجيم":
        st.header("2. Base Reactions & Dynamic Scale Factor Adjustment")
        render_section_equations("base_reactions", {"asce_percentage": asce_percentage})
        if base_error:
            st.error(f"تعذر حساب هذا القسم: {base_error}")
        elif not df_comp.empty:
            st.dataframe(df_comp, use_container_width=True)
            st.plotly_chart(plot_base_shear_comparison(v_static_x, v_static_y, v_dyn_x, v_dyn_y, asce_percentage), use_container_width=True)
            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.markdown("#### الاتجاه X")
                if v_dyn_x > 0 and v_static_x > 0:
                    req_scale_x = ((v_static_x * asce_percentage) / v_dyn_x) * (importance_factor / response_mod_r)
                    st.metric(label="Scale Factor المقترح لـ RSX", value=f"{req_scale_x:.4f}")
            with col_res2:
                st.markdown("#### الاتجاه Y")
                if v_dyn_y > 0 and v_static_y > 0:
                    req_scale_y = ((v_static_y * asce_percentage) / v_dyn_y) * (importance_factor / response_mod_r)
                    st.metric(label="Scale Factor المقترح لـ RSY", value=f"{req_scale_y:.4f}")
        else:
            st.warning("جدول تفاعلات القاعدة غير متوفر أو تعذر قراءته.")

    elif section_navigation == "3. انجراف الطوابق":
        st.header("3. Story Drifts")
        render_section_equations("story_drifts", {})
        if drift_note:
            _kind = classify_status(status_dict["Story Drifts"])
            (st.success if _kind == "pass" else st.error if _kind == "fail" else st.warning)(drift_note)
        if not df_drifts_display.empty:
            if "Result" in df_drifts_display.columns or "CHECK" in df_drifts_display.columns:
                st.caption(
                    "⚠️ عمود 'Result'/'CHECK' (إن وُجد في الجدول أدناه) محفوظ كما هو من "
                    "ملف الإكسل المصدر، وقد يعتمد على معادلة مختلفة عن الحكم أعلاه الذي "
                    "يحسبه هذا التطبيق بنفسه وفق ASCE 7-10 (جدول 12.12-1). الحكم الموثوق "
                    "هو ما يظهر أعلاه، وليس هذا العمود."
                )
            st.dataframe(df_drifts_display, use_container_width=True)
            try:
                st.plotly_chart(plot_story_drifts(df_drifts_display), use_container_width=True)
            except Exception as e:
                st.error(f"تعذر رسم مخطط الانجراف: {e}")
        else:
            st.warning("لا توجد بيانات متاحة لعرض جدول انجراف الطوابق.")

    elif section_navigation == "4. مشاركة الكتلة النمطية":
        st.header("4. Modal Participating Mass Ratios")
        render_section_equations("modal_mass", {})
        if modal_note:
            _kind = classify_status(status_dict["Modal Mass Ratios"])
            (st.success if _kind == "pass" else st.error if _kind == "fail" else st.warning)(modal_note)
        if not df_modal.empty:
            st.dataframe(df_modal.style.map(color_mass_ratios, subset=[c for c in df_modal.columns if c.strip().lower() in ("sumux", "sumuy")]), use_container_width=True)
        else:
            st.warning("لا توجد بيانات متاحة لعرض جدول مشاركة الكتلة النمطية.")

    elif section_navigation == "5. أقصى/متوسط الإزاحة":
        st.header("5. Story Max Over Avg Displacements — الالتواء العرضي")
        render_section_equations("displacements", {})
        if disp_error:
            st.error(f"تعذر قراءة هذا القسم: {disp_error}")
        else:
            if disp_note:
                _kind = classify_status(status_dict["Displacements"])
                (st.success if _kind == "pass" else st.error if _kind == "fail" else st.warning)(disp_note)
            if not df_disp.empty:
                st.dataframe(df_disp, use_container_width=True)
            else:
                st.warning("لا توجد بيانات متاحة لعرض جدول الإزاحات.")

    elif section_navigation == "6. صلابة الطوابق":
        st.header("6. Story Stiffness — الطابق اللين")
        render_section_equations("story_stiffness", {})
        if stiff_error:
            st.error(f"تعذر قراءة هذا القسم: {stiff_error}")
        else:
            if stiff_note:
                _kind = classify_status(status_dict["Story Stiffness"])
                (st.success if _kind == "pass" else st.error if _kind == "fail" else st.warning)(stiff_note)
            if not df_stiff.empty:
                st.dataframe(df_stiff, use_container_width=True)
                try:
                    st.plotly_chart(plot_story_stiffness(df_stiff), use_container_width=True)
                except Exception as e:
                    st.error(f"تعذر رسم مخطط الصلابة: {e}")
            else:
                st.warning("لا توجد بيانات متاحة لعرض جدول الصلابة.")

    elif section_navigation == "7. فحص P-Delta":
        st.header("7. P-Delta Effect Check")
        render_section_equations("pdelta", {})
        if pdelta_note:
            (st.success if classify_status(status_dict["P-Delta Check"]) == "pass" else st.warning)(pdelta_note)
        if not df_pdelta_filtered.empty:
            st.dataframe(df_pdelta_filtered.head(15), use_container_width=True)
        else:
            st.warning("لا توجد بيانات متاحة لعرض جدول P-Delta.")

    elif section_navigation == "8. انجراف الرياح (خدمي)":
        st.header("8. Wind Drift — فحص خدمي (ليس زلزالياً)")
        render_section_equations("wind_drift", {})
        if wind_error:
            st.error(f"تعذر قراءة هذا القسم: {wind_error}")
        else:
            if wind_note:
                _kind = classify_status(status_dict["Wind Drift"])
                (st.success if _kind == "pass" else st.error if _kind == "fail" else st.warning)(wind_note)
            if not df_wind_story.empty:
                st.markdown("##### الانجراف النسبي بين الطوابق (Interstory)")
                st.dataframe(df_wind_story, use_container_width=True)
            if not df_wind_total.empty:
                st.markdown("##### الانجراف الكلي من القاعدة (Cumulative، مُعاد حسابه)")
                st.dataframe(df_wind_total, use_container_width=True)
            if df_wind_story.empty and df_wind_total.empty:
                st.warning("لا توجد بيانات متاحة لعرض جداول انجراف الرياح.")

    # ==========================================
    # أزرار التصدير في الشريط الجانبي
    # ==========================================
    st.sidebar.markdown("---")
    st.sidebar.header("📄 تصدير المخرجات النهائية")
    project_info_dict = {
        "اسم المشروع": project_name, "الموقع": project_location, "المصمم": consultant_name,
        "الاستخدام": building_usage, "عدد الطوابق": num_stories, "الارتفاع الإجمالي": total_height,
        "النظام الزلزالي": selected_system, "معامل R": response_mod_r, "معامل Cd": deflection_cd,
        "فئة المخاطر (Risk Category)": selected_risk_category, "عامل الأهمية Ie": importance_factor,
        "عامل التكرار ρ": redundancy_rho, "حد الانجراف الأساسي Δa,base": base_drift_limit,
    }

    excel_stream = export_audit_results_to_excel(project_info_dict, status_dict, res_df, df_comp, df_pdelta_filtered)
    st.sidebar.download_button(
        label="📥 تحميل جداول التحليل (Excel)",
        data=excel_stream,
        file_name=f"{project_name}_ETABS_Audit_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    docx_stream = generate_word_report(
        project_name, project_location, consultant_name, contact_info,
        building_usage, num_stories, total_height,
        fc_vertical, fc_horizontal, rebar_grade,
        zone_z, soil_factor_s, importance_factor,
        response_mod_r, deflection_cd,
        status_dict, df_comp, figures_for_report, methodology_notes,
    )
    st.sidebar.download_button(
        label="📥 تحميل النوتة الحسابية والتقرير (Word)",
        data=docx_stream,
        file_name=f"{project_name}_Calculation_Note_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
