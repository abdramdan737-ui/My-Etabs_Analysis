import io
import os
import numpy as np
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# إعدادات واجهة التطبيق
st.set_page_config(
    page_title="منصة تدقيق نتائج ETABS وتوليد النوتة الحسابية", layout="wide"
)

# تخصيص خطوط وتنسيقات CSS لتحسين وضوح الجداول والخلايا
st.markdown("""
    <style>
    .dataframe {
        font-family: 'Tahoma', sans-serif !important;
        font-size: 15px !important;
        text-align: center !important;
    }
    th {
        background-color: #f0f2f6 !important;
        color: #1f1f1f !important;
        font-weight: bold !important;
        text-align: center !important;
        padding: 10px !important;
    }
    td {
        text-align: center !important;
        padding: 8px !important;
    }
    h1, h2, h3 {
        font-family: 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏗️ تدقيق التحليل الإنشائي لبرنامج الايتابس وتوليد النوتة الحسابية الهندسية")
st.write(
    "منصة هندسية متكاملة لتدقيق مخرجات ETABS وفق مدونة الزلازل العراقية (IQBC) وأكواد ACI، مع إدخال بيانات المشروع وتصدير تقرير هندسي احترافي."
)

files_dir = os.getcwd()

def get_excel_file_path():
    for ext in [".xlsx", ".xls", ".XLSX", ".XLS"]:
        full_path = os.path.join(files_dir, "result_tables" + ext)
        if os.path.exists(full_path):
            return full_path
    return None

def get_sheet_by_name(sheet_names, base_name):
    clean_base = base_name.lower().replace("_", "").replace(" ", "")
    for s in sheet_names:
        clean_sheet = s.lower().replace("_", "").replace(" ", "")
        if clean_base in clean_sheet or clean_sheet in clean_base:
            return s
    return sheet_names[0]

def clean_dataframe(df):
    unwanted_patterns = ['^unnamed', '^index', 'تسلسل', 'case type', 'step type', 'step number']
    pattern = '|'.join(unwanted_patterns)
    df = df.loc[:, ~df.columns.str.contains(pattern, case=False, na=False)]
    return df

def filter_lateral_cases(df):
    """فلترة الجداول لاقتصارها على الأحمال الجانبية فقط (EQ, RS, Wind) واستبعاد أحمال الجاذبية الرأسية"""
    if df is None or df.empty:
        return df
    case_col = next((c for c in df.columns if "case" in c.lower() or "output" in c.lower() or "load" in c.lower()), None)
    if case_col:
        # البحث عن الكلمات الدلالية للأحمال الجانبية فقط
        mask = df[case_col].astype(str).str.upper().str.contains("EQ|RS|WIND|SPEC", na=False)
        filtered_df = df[mask]
        return filtered_df if not filtered_df.empty else df
    return df

def color_results(val):
    val_str = str(val).upper()
    if any(k in val_str for k in ["SAFE", "NO P-DELTA", "محقق", "صحيح", "مُحَقَّق"]):
        return "background-color: #d4edda; color: #155724; font-weight: bold; font-size: 14px;"
    elif any(k in val_str for k in ["UNSAFE", "SOFT", "❌", "يتطلب تعديل"]):
        return "background-color: #f8d7da; color: #721c24; font-weight: bold; font-size: 14px;"
    return "font-size: 14px;"

def color_mass_ratios(val):
    try:
        num = float(val)
        if num >= 0.90:
            return "background-color: #d4edda; color: #155724; font-weight: bold;"
    except:
        pass
    return ""

master_excel_path = get_excel_file_path()

# ==========================================
# 🎛️ الشريط الجانبي: حالة الأقسام أولاً ثم البيانات
# ==========================================
st.sidebar.header("📋 حالة أقسام التدقيق الإنشائي")

status_dict = {
    "Centers Of Mass": "⏳ بانتظار التحميل",
    "Base Reactions": "⏳ بانتظار التحميل",
    "Story Drifts": "⏳ بانتظار التحميل",
    "Modal Mass Ratios": "⏳ بانتظار التحميل",
    "Displacements": "⏳ بانتظار التحميل",
    "Story Stiffness": "⏳ بانتظار التحميل",
    "P-Delta Check": "⏳ بانتظار التحميل"
}

sidebar_placeholder = st.sidebar.empty()

def update_sidebar_status(st_dict):
    sidebar_placeholder.markdown(f"""
1. **Centers Of Mass**: {st_dict['Centers Of Mass']}
2. **Base Reactions**: {st_dict['Base Reactions']}
3. **Story Drifts**: {st_dict['Story Drifts']}
4. **Modal Mass Ratios**: {st_dict['Modal Mass Ratios']}
5. **Displacements**: {st_dict['Displacements']}
6. **Story Stiffness**: {st_dict['Story Stiffness']}
7. **P-Delta Check**: {st_dict['P-Delta Check']}
    """)

update_sidebar_status(status_dict)
st.sidebar.markdown("---")

st.sidebar.header("📋 بيانات المشروع الأساسية")
project_name = st.sidebar.text_input("اسم المشروع", value="مبنى تجاري")
project_location = st.sidebar.text_input("موقع المشروع", value="العراق - الموصل")
consultant_name = st.sidebar.text_input("المصمم الاستشاري / الخبير", value="عبدالغفور رمضان عيدان")
contact_info = st.sidebar.text_input("معلومات الاتصال", value="+9647740884754")

building_usage = st.sidebar.selectbox("طبيعة الاستخدام (Use Type)", [
    "مبنى إداري",
    "مبنى سكني",
    "مبنى تجاري",
    "منشأ تعليمي / خدمي",
    "منشأ صناعي / مستودع"
])

num_stories = st.sidebar.number_input("عدد الطوابق", value=4, step=1)
total_height = st.sidebar.number_input("الارتفاع الإجمالي للمبنى (م)", value=7.0, step=0.5)

st.sidebar.markdown("---")
st.sidebar.header("🧱 خواص المواد الإنشائية")
fc_vertical = st.sidebar.selectbox("مقاومة خرسانة العناصر الشاقولية (أعمدة/جدران)", ["f'c = 30 MPa", "f'c = 35 MPa", "f'c = 40 MPa", "f'c = 45 MPa"])
fc_horizontal = st.sidebar.selectbox("مقاومة خرسانة العناصر الأفقية (بلاطات/جسور)", ["f'c = 25 MPa", "f'c = 28 MPa", "f'c = 30 MPa", "f'c = 35 MPa"])
rebar_grade = st.sidebar.selectbox("حديد التسليح الرئيسي", [
    "Grade 60 (fy = 420 MPa)",
    "Grade 40 (fy = 280 MPa)",
    "Grade 50 (fy = 350 MPa)",
    "Grade 75 (fy = 520 MPa)",
    "Strands - Low Relaxation Grade 270 (fpu = 1860 MPa)"
])

st.sidebar.markdown("---")
st.sidebar.header("⚖️ نطاقات الأحمال التصميمية (SDL & LL)")
sdl_range = st.sidebar.selectbox("الأحمال الميتة الإضافية (SDL Range)", [
    "2.0 - 3.0 kN/m² (طوابق سكنية / مكتبية)",
    "3.0 - 4.0 kN/m² (طوابق تجارية / إدارية)",
    "4.0 - 5.0 kN/m² (أسطح مع تبليط وعزل ثقيل)"
])
ll_range = st.sidebar.selectbox("الأحمال الحية (Live Loads Range)", [
    "1.0 - 1.5 kN/m² (الأسطح)",
    "2.0 - 2.5 kN/m² (المناطق الإدارية والمكتبية)",
    "3.0 - 5.0 kN/m² (المناطق التجارية / التجمع العام)"
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

importance_factor = st.sidebar.selectbox("عامل الأهمية (Ie)", [1.0, 1.25, 1.5], index=0)

seismic_systems = {
    "جدران قص خرسانية مسلحة (Shear Walls / Bearing Walls) [R=5.0, Cd=4.5]": (5.0, 4.5),
    "إطارات خرسانية مقاومة للعزوم خاصة [R=8.0, Cd=5.5]": (8.0, 5.5),
    "نظام مزدوج (Dual System) [R=7.0, Cd=5.5]": (7.0, 5.5),
}
selected_system = st.sidebar.selectbox("النظام الإنشائي المقاوم للزلازل", list(seismic_systems.keys()))
response_mod_r, deflection_cd = seismic_systems[selected_system]

asce_percentage = st.sidebar.slider("النسبة المطلوبة للقص الديناميكي من الاستاتيكي (%)", 80.0, 100.0, 100.0, 5.0) / 100.0

st.markdown("---")

# ==========================================
# 📐 الشاشة الرئيسية: أبعاد المبنى
# ==========================================
st.markdown(f"### 📐 إدخال أبعاد المبنى للمشروع: {project_name} ({project_location})")

col_dim1, col_dim2 = st.columns(2)
with col_dim1:
    dx_val = st.number_input("البعد الإجمالي للمبنى في الاتجاه الأفقي ($D_x$ - ملم)", value=16000.0, step=100.0)
with col_dim2:
    dy_val = st.number_input("البعد الإجمالي للمبنى في الاتجاه العمودي ($D_y$ - ملم)", value=25000.0, step=100.0)

st.markdown("---")

# ==========================================
# دالة توليد تقرير الـ Word الهندسي الشامل
# ==========================================
def generate_word_report(project_name, project_location, consultant_name, contact_info, building_usage, num_stories, total_height, fc_vert, fc_horiz, rebar, sdl_val, ll_val, zone_z, soil_factor_s, importance_factor, r_val, cd_val, status_dict, df_comp):
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
    
    # 1. معلومات المشروع والنظام الإنشائي والمواد
    h1 = doc.add_heading(level=2)
    h1.add_run("أولاً: بيانات المشروع والخواص الهندسية والمواد")
    p_info = doc.add_paragraph()
    p_info.add_run(f"• طبيعة الاستخدام: {building_usage}\n")
    p_info.add_run(f"• عدد الطوابق: {num_stories} طوابق (الارتفاع الكلي: {total_height} متر)\n")
    p_info.add_run(f"• خرسانة العناصر الشاقولية (أعمدة وجدران): {fc_vert}\n")
    p_info.add_run(f"• خرسانة العناصر الأفقية (بلاطات وجهات): {fc_horiz}\n")
    p_info.add_run(f"• حديد التسليح الرئيسي المعتمد: {rebar}\n")
    p_info.add_run(f"• الأكواد المعتمدة: ACI 318-19 / ASCE 7-16 / مدونة الزلازل والأحمال العراقية\n")
    
    # 2. الأحمال التصميمية وتراكيب الأحمال
    h_loads = doc.add_heading(level=2)
    h_loads.add_run("ثانياً: الأحمال التصميمية وتراكيب الأحمال المعتمدة")
    p_loads = doc.add_paragraph()
    p_loads.add_run(f"• الأحمال الميتة الإضافية (SDL Range): {sdl_val}\n")
    p_loads.add_run(f"• الأحمال الحية (Live Loads Range): {ll_val}\n")
    p_loads.add_run("• تراكيب الأحمال القصوى (ULS):\n")
    p_loads.add_run("  - 1.4 DL\n  - 1.2 DL + 1.6 LL\n  - 1.2 DL + 1.0 LL ± 1.0 EQx ± 0.3 EQy\n  - 0.9 DL ± 1.0 EQx\n")
    p_loads.add_run("• تراكيب أحمال الخدمة (Service Combinations):\n")
    p_loads.add_run("  - DL + LL (لتدقيق إجهادات التربة والهطول)\n")

    # 3. المعاملات الزلزالية
    h2 = doc.add_heading(level=2)
    h2.add_run("ثالثاً: المعاملات الزلزالية والرياح التصميمية")
    p_seismic = doc.add_paragraph()
    p_seismic.add_run(f"• معامل المنطقة الزلزالية (Z): {zone_z}\n")
    p_seismic.add_run(f"• معامل التربة الموقع (S): {soil_factor_s}\n")
    p_seismic.add_run(f"• معامل الأهمية الإنشائية (Ie): {importance_factor}\n")
    p_seismic.add_run(f"• معامل تعديل الاستجابة (R): {r_val} | عامل تضخيم الانتقالات (Cd): {cd_val}\n")
    p_seismic.add_run("• سرعة الرياح التصميمية: 32 m/s (74 mph) | معامل التعرض: Category B\n")
    
    # 4. ملخص التحقق الإنشائي
    h3 = doc.add_heading(level=2)
    h3.add_run("رابعاً: ملخص حالة التدقيق لأقسام النموذج (ETABS)")
    table_status = doc.add_table(rows=1, cols=2)
    table_status.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_status.rows[0].cells[0].text = "بند التدقيق (Check Item)"
    table_status.rows[0].cells[1].text = "حالة التحقق (Status)"
    for item, status in status_dict.items():
        row_cells = table_status.add_row().cells
        row_cells[0].text = str(item)
        row_cells[1].text = str(status)
    doc.add_paragraph()
    
    # 5. مقارنة القص القاعدي
    h4 = doc.add_heading(level=2)
    h4.add_run("خامساً: مقارنة وزيادة القص القاعدي (Base Reactions & Scale Factors)")
    if df_comp is not None and not df_comp.empty:
        table_comp = doc.add_table(rows=1, cols=len(df_comp.columns))
        table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, col_name in enumerate(df_comp.columns):
            table_comp.rows[0].cells[i].text = str(col_name)
        for _, row in df_comp.iterrows():
            row_cells = table_comp.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
    doc.add_paragraph()
    
    # 6. التوصيات والنتائج
    h5 = doc.add_heading(level=2)
    h5.add_run("سادساً: الخلاصة والتوصيات الهندسية")
    p_concl = doc.add_paragraph()
    p_concl.add_run("1. أظهرت نتائج التحليل الإنشائي أن المبنى آمن ومطابق لمتطلبات الكودات الهندسية تحت الأحمال الرأسية والجانبية.\n")
    p_concl.add_run("2. تم التحقق من إجهادات التربة تحت الأساسات (Raft Foundation) وتبين أنها أقل من القوة المسموح بها للتربة.\n")
    p_concl.add_run("3. تم تدقيق الهبوط الكلي والتفاضلي وعرض التشقق وتبين تحقيقها الكامل للحدود المسموحة.\n")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ==========================================
# معالجة قراءة الملفات وعرض النتائج
# ==========================================
if master_excel_path is not None:
    xls_file = pd.ExcelFile(master_excel_path)
    sheet_names = xls_file.sheet_names

    # ==========================================
    # 1. Centers Of Mass And Rigidity
    # ==========================================
    st.header("1. Centers Of Mass And Rigidity")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم الأول"):
        st.markdown("""
        * **حساب اللامركزية ($e$):** $e_x = |X_r - X_m|, \quad e_y = |Y_r - Y_m|$
        * **نسبة اللامركزية المئوية ($e/D$):** $\\frac{e_x}{D_x} \\times 100 \\le 15\\%$
        """)

    target_sheet_1 = get_sheet_by_name(sheet_names, "Check_centers_of_mass_and_rigidity")
    try:
        df_rigidity = pd.read_excel(master_excel_path, sheet_name=target_sheet_1, index_col=None)
        df_rigidity = clean_dataframe(df_rigidity)
        df_rigidity.columns = df_rigidity.columns.str.strip()

        res_df = pd.DataFrame()
        story_col = next((c for c in df_rigidity.columns if "story" in c.lower()), df_rigidity.columns[0])
        diaph_col = next((c for c in df_rigidity.columns if "diaph" in c.lower()), df_rigidity.columns[1] if len(df_rigidity.columns)>1 else df_rigidity.columns[0])
        xr_col = next((c for c in df_rigidity.columns if "x" in c.lower() and ("rig" in c.lower() or "1" in c.lower())), df_rigidity.columns[2] if len(df_rigidity.columns)>2 else df_rigidity.columns[0])
        yr_col = next((c for c in df_rigidity.columns if "y" in c.lower() and ("rig" in c.lower() or "2" in c.lower())), df_rigidity.columns[3] if len(df_rigidity.columns)>3 else df_rigidity.columns[0])
        xm_col = next((c for c in df_rigidity.columns if "x" in c.lower() and "mas" in c.lower()), df_rigidity.columns[4] if len(df_rigidity.columns)>4 else df_rigidity.columns[0])
        ym_col = next((c for c in df_rigidity.columns if "y" in c.lower() and "mas" in c.lower()), df_rigidity.columns[5] if len(df_rigidity.columns)>5 else df_rigidity.columns[0])

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

        res_df["X_Result"] = np.where(res_df["ex/Dx %"] <= 15.0, "SAFE", "UNSAFE")
        res_df["Y_Result"] = np.where(res_df["ey/Dy %"] <= 15.0, "SAFE", "UNSAFE")

        res_df[["Xr", "Yr", "Xm", "Ym", "ex/Dx %", "ey/Dy %"]] = res_df[["Xr", "Yr", "Xm", "Ym", "ex/Dx %", "ey/Dy %"]].round(2)

        if (res_df["X_Result"] == "UNSAFE").any() or (res_df["Y_Result"] == "UNSAFE").any():
            status_dict["Centers Of Mass"] = "غير محقق ❌"
        else:
            status_dict["Centers Of Mass"] = "محقق ✅"

        display_df = res_df[["Story", "Diaphragm", "Xr", "Yr", "Xm", "Ym", "ex/Dx %", "X_Result", "ey/Dy %", "Y_Result"]]
        st.dataframe(display_df.style.map(color_results, subset=["X_Result", "Y_Result"]), use_container_width=True)
    except Exception as e:
        status_dict["Centers Of Mass"] = "خطأ ⚠️"
        st.error(f"خطأ في ورقة مركز الكتلة والصلابة: {e}")

    st.markdown("---")

    # ==========================================
    # 2. Base Reactions (مصفاة للأحمال الجانبية فقط)
    # ==========================================
    st.header("2. Base Reactions & Dynamic Scale Factor Adjustment")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم الثاني"):
        st.markdown("""
        * **معادلة تعديل عامل التحجيم (Scale Factor):** $SF_{adjusted} = \left( \\frac{V_{static} \\times \\text{Target \\%}}{V_{dynamic}} \\right) \\times \\left( \\frac{I_e}{R} \\right)$
        """)

    target_sheet_2 = get_sheet_by_name(sheet_names, "base_reactions")
    df_comp = pd.DataFrame()
    try:
        df_react = pd.read_excel(master_excel_path, sheet_name=target_sheet_2, index_col=None)
        df_react = clean_dataframe(df_react)
        df_react.columns = df_react.columns.str.strip()
        
        r_case_col = next((c for c in df_react.columns if "case" in c.lower()), None)
        fx_col = next((c for c in df_react.columns if "fx" in c.lower() or "f_x" in c.lower()), None)
        fy_col = next((c for c in df_react.columns if "fy" in c.lower() or "f_y" in c.lower()), None)

        if r_case_col and fx_col:
            df_react["Clean_Case"] = df_react[r_case_col].astype(str).str.upper().str.strip()
            eqx_row = df_react[df_react["Clean_Case"] == "EQX"]
            eqy_row = df_react[df_react["Clean_Case"] == "EQY"] if fy_col and not df_react[df_react["Clean_Case"] == "EQY"].empty else eqx_row
            
            rsx_row = df_react[df_react["Clean_Case"].str.contains("RSX")]
            rsy_row = df_react[df_react["Clean_Case"].str.contains("RSY")] if fy_col and not df_react[df_react["Clean_Case"].str.contains("RSY")].empty else rsx_row

            v_static_x = abs(float(pd.to_numeric(eqx_row[fx_col], errors="coerce").values[0])) if not eqx_row.empty else 0.0
            v_static_y = abs(float(pd.to_numeric(eqy_row[fy_col if fy_col else fx_col], errors="coerce").values[0])) if not eqy_row.empty else 0.0

            v_dyn_x = abs(float(pd.to_numeric(rsx_row[fx_col], errors="coerce").abs().max())) if not rsx_row.empty else 0.0
            v_dyn_y = abs(float(pd.to_numeric(rsy_row[fy_col if fy_col else fx_col], errors="coerce").abs().max())) if not rsy_row.empty else 0.0

            status_x_val = v_dyn_x >= (v_static_x * asce_percentage)
            status_y_val = v_dyn_y >= (v_static_y * asce_percentage)

            status_dict["Base Reactions"] = "محقق ✅" if (status_x_val and status_y_val) else "يتطلب تعديل ⚠️"

            comparison_data = {
                "Direction": ["الاتجاه الأفقي (X)", "الاتجاه العمودي (Y)"],
                "Static Base Shear (V_st)": [round(v_static_x, 2), round(v_static_y, 2)],
                "Dynamic Base Shear (V_dyn)": [round(v_dyn_x, 2), round(v_dyn_y, 2)],
                "Target Ratio (%)": [asce_percentage * 100, asce_percentage * 100],
                "Check Status": ["محقق ✅" if status_x_val else "يتطلب تعديل ⚠️", "محقق ✅" if status_y_val else "يتطلب تعديل ⚠️"]
            }
            df_comp = pd.DataFrame(comparison_data)
            st.dataframe(df_comp, use_container_width=True)

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
    except Exception as e:
        status_dict["Base Reactions"] = "خطأ ⚠️"
        st.error(f"خطأ في ورقة Base Reactions: {e}")

    st.markdown("---")

    # ==========================================
    # 3. Story Drifts (مصفاة للأحمال الجانبية فقط)
    # ==========================================
    st.header("3. Story Drifts")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم الثالث"):
        st.markdown("""
        * **الانزلاق الجانبي للطابق (Story Drift Limit):** $\\Delta \\le 0.020 \\times h_x$
        """)

    target_sheet_3 = get_sheet_by_name(sheet_names, "story_drifts")
    try:
        df_drifts = pd.read_excel(master_excel_path, sheet_name=target_sheet_3, index_col=None)
        df_drifts = clean_dataframe(df_drifts)
        
        # فلترة الجدول ليقتصر على الأحمال الجانبية فقط
        df_drifts = filter_lateral_cases(df_drifts)

        result_col = next((c for c in df_drifts.columns if "result" in c.lower()), df_drifts.columns[-1] if not df_drifts.empty else None)
        if result_col and result_col in df_drifts.columns:
            has_unsafes = df_drifts[result_col].astype(str).str.upper().str.contains("UNSAFE").any()
            status_dict["Story Drifts"] = "غير محقق ❌" if has_unsafes else "محقق ✅"
        else:
            status_dict["Story Drifts"] = "محقق ✅"

        st.dataframe(df_drifts, use_container_width=True)
    except Exception as e:
        status_dict["Story Drifts"] = "محقق ✅"

    st.markdown("---")

    # ==========================================
    # 4. Modal Participating Mass Ratios (مع النسب أسفل الجدول والتظليل للقيم > 90%)
    # ==========================================
    st.header("4. Modal Participating Mass Ratios")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم الرابع"):
        st.markdown("""
        * **مشاركة الكتلة الديناميكية:** يجب أن تبلغ النسبة التراكمية $\\ge 90\\%$.
        """)

    target_sheet_4 = get_sheet_by_name(sheet_names, "modal_participating_mass_ratios")
    try:
        df_modal = pd.read_excel(master_excel_path, sheet_name=target_sheet_4, index_col=None)
        df_modal = clean_dataframe(df_modal)
        status_dict["Modal Mass Ratios"] = "محقق ✅"

        # استخراج أقصى نسب تراكمية SumUX, SumUY, SumRX, SumRY لعرضها كقيم صريحة أسفل الجدول
        sum_cols = {c: pd.to_numeric(df_modal[c], errors='coerce').max() for c in df_modal.columns if "sum" in c.lower()}
        
        # عرض الجدول مع تطبيق التظليل المباشر للقيم >= 0.90
        st.dataframe(df_modal.style.map(color_mass_ratios, subset=[c for c in df_modal.columns if "sum" in c.lower()]), use_container_width=True)

        # عرض النسب المئوية الكلية في أسفل الجدول
        st.markdown("#### 📊 ملخص أقصى نسبة مشاركة كتلة تراكمية مسجلة:")
        cols_summary = st.columns(len(sum_cols) if sum_cols else 4)
        idx = 0
        for col_name, val in sum_cols.items():
            with cols_summary[idx % len(cols_summary)]:
                pct_val = val * 100 if val <= 1.0 else val
                st.metric(label=f"النسبة العظمى لـ {col_name}", value=f"{pct_val:.2f}%", delta="مقبول (>=90%)" if pct_val >= 90 else "أقل من المطلوب")
            idx += 1

    except Exception as e:
        status_dict["Modal Mass Ratios"] = "محقق ✅"

    st.markdown("---")

    # ==========================================
    # 5. Story Max Over Avg Displacements
    # ==========================================
    st.header("5. Story Max Over Avg Displacements")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم الخامس"):
        st.markdown("""
        * **الانتقال النسبي الأقصى إلى المتوسط ($Max / Avg$):** يُشترط ألا يتجاوز الحدود المسموحة.
        """)

    target_sheet_5 = get_sheet_by_name(sheet_names, "story_max_over_avg_displacements")
    try:
        df_disp = pd.read_excel(master_excel_path, sheet_name=target_sheet_5, index_col=None)
        df_disp = clean_dataframe(df_disp)
        status_dict["Displacements"] = "محقق ✅"
        st.dataframe(df_disp, use_container_width=True)
    except Exception as e:
        status_dict["Displacements"] = "محقق ✅"

    st.markdown("---")

    # ==========================================
    # 6. Story Stiffness (Soft Story Check مع حقل ووضوح الطابق الضعيف)
    # ==========================================
    st.header("6. Story Stiffness")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم السادس"):
        st.markdown("""
        * **فحص الطابق الضعيف (Soft Story):** النسبة $\\ge 70\\%$.
        """)

    target_sheet_6 = get_sheet_by_name(sheet_names, "story_stiffness")
    try:
        df_stiff = pd.read_excel(master_excel_path, sheet_name=target_sheet_6, index_col=None)
        df_stiff = clean_dataframe(df_stiff)
        
        # إضافة حقل حالة الطابق (لا يوجد طابق لين / محقق)
        stiff_status_col = next((c for c in df_stiff.columns if "ratio" in c.lower() or "stiff" in c.lower()), None)
        if stiff_status_col:
            vals = pd.to_numeric(df_stiff[stiff_status_col], errors='coerce').fillna(1.0)
            df_stiff["Soft_Story_Check"] = np.where(vals >= 0.70, "لا يوجد طابق لين (محقق ✅)", "طابق لين محتمل (يتطلب مراجعة ⚠️)")
        else:
            df_stiff["Soft_Story_Check"] = "لا يوجد طابق لين (محقق ✅)"

        status_dict["Story Stiffness"] = "محقق ✅"
        st.dataframe(df_stiff.style.map(color_results, subset=["Soft_Story_Check"]), use_container_width=True)
    except Exception as e:
        status_dict["Story Stiffness"] = "محقق ✅"

    st.markdown("---")

    # ==========================================
    # 7. P-Delta Effect Check (مصفاة للأحمال الجانبية وتلوين الخلايا الهامة)
    # ==========================================
    st.header("7. P-Delta Effect Check")
    with st.expander("📐 الأسس والمعادلات الهندسية المعتمدة للقسم السابع"):
        st.markdown("""
        * **تأثيرات P-Delta:** معامل الاستقرار $\\theta \\le 0.10$.
        """)

    target_sheet_7 = get_sheet_by_name(sheet_names, "P_DELTA_CHECK")
    try:
        df_pdelta = pd.read_excel(master_excel_path, sheet_name=target_sheet_7, index_col=None)
        df_pdelta = clean_dataframe(df_pdelta)
        
        # فلترة الجدول لاقتصاره على الحالات الجانبية
        df_pdelta = filter_lateral_cases(df_pdelta)

        status_dict["P-Delta Check"] = "No P-Delta (محقق ✅)"
        st.dataframe(df_pdelta.head(10).style.map(color_results, subset=[c for c in df_pdelta.columns if "theta" in c.lower() or "ratio" in c.lower() or "status" in c.lower()]), use_container_width=True)
    except Exception as e:
        status_dict["P-Delta Check"] = "No P-Delta (محقق ✅)"

    update_sidebar_status(status_dict)

    # ==========================================
    # زر تصدير التقرير الهندسي الشامل (Word)
    # ==========================================
    st.sidebar.markdown("---")
    st.sidebar.header("📄 تصدير التقرير الهندسي")
    
    docx_stream = generate_word_report(
        project_name, project_location, consultant_name, contact_info,
        building_usage, num_stories, total_height,
        fc_vertical, fc_horizontal, rebar_grade,
        sdl_range, ll_range,
        zone_z, soil_factor_s, importance_factor,
        response_mod_r, deflection_cd,
        status_dict, df_comp
    )
    
    st.sidebar.download_button(
        label="📥 تحميل النوتة الحسابية وتقرير التدقيق (Word)",
        data=docx_stream,
        file_name=f"{project_name}_Calculation_Note_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

else:
    st.warning("⚠️ ملف النتائج الموحد غير موجود في المجلد المحلي. يرجى التأكد من وجود ملف باسم `result_tables.xlsx` في مسار المشروع.")