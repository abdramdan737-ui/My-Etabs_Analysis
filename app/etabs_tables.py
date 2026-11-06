import io

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, RGBColor

from app.etabs_processing import (
    OPTIONAL_PDELTA_SHEETS,
    OPTIONAL_SHEET_NAMES,
    REQUIRED_SHEET_NAMES,
    WorkbookValidationError,
    calculate_base_reactions,
    calculate_eccentricity,
    calculate_modal_mass,
    calculate_iraqi_2019_story_drifts,
    calculate_iraqi_2019_torsion_eccentricities,
    calculate_iraqi_2019_torsional_irregularity,
    calculate_story_drifts,
    get_cached_workbook,
    pdelta_assessment,
    validate_upload_filename,
)


st.set_page_config(
    page_title="منصة تدقيق نتائج الايتابس وتوليد النوتة الحسابية", layout="wide"
)

st.markdown(
    """
    <style>
    .dataframe { font-family: Tahoma, sans-serif !important; font-size: 15px !important; }
    th { background-color: #f0f2f6 !important; color: #1f1f1f !important; font-weight: bold !important; text-align: center !important; }
    td { text-align: center !important; }
    h1 { font-family: Tahoma, sans-serif !important; color: #1F4E79 !important; }
    h2 { font-family: Tahoma, sans-serif !important; color: #0056B3 !important; }
    h3 { font-family: Tahoma, sans-serif !important; color: #2CA02C !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    "<h2 style='text-align: center; font-size: 40px;'>🏗️ تدقيق التحليل الإنشائي لنتائج برنامج الايتابس وتوليد النوتة الحسابية</h2>",
    unsafe_allow_html=True,
)
st.write(
    "ارفع مصنف ETABS الخام لقراءة الجداول والتحقق من الحسابات المتاحة وتصدير ملخص النتائج."
)
st.subheader("🛠️ المنصة من إعداد وتصميم المهندس الاستشاري عبدالغفور رمضان اللهيبي - 2026")

STATUS_LABELS = {
    "WAITING": "⏳ بانتظار الرفع",
    "PASS": "محقق ✅",
    "FAIL": "غير محقق ❌",
    "INCOMPLETE": "غير مكتمل ⚠️",
    "NOT_EVALUATED": "غير محسوب ⚠️",
    "ERROR": "خطأ في المدخلات ⚠️",
    "INFORMATIONAL": "معلوماتي — لا يصدر حكم مطابقة",
}
SECTION_NAMES = {
    "Centers Of Mass": "مراكز الكتلة والصلابة",
    "Base Reactions": "القص القاعدي ومعامل التحجيم",
    "Story Drifts": "انجراف الطوابق",
    "Modal Mass Ratios": "مشاركة الكتلة النمطية",
    "Displacements": "أقصى/متوسط الإزاحة",
    "Story Stiffness": "صلابة الطوابق",
    "P-Delta Check": "فحص P-Delta",
    "Iraqi 2019 Drift": "انجراف الطوابق — الكود العراقي 2019",
    "Iraqi Torsion": "الالتواء واللامركزية — الكود العراقي 2019",
}


def status_label(status: str) -> str:
    return STATUS_LABELS.get(status, status)


def generate_word_report(
    project: dict,
    statuses: dict[str, str],
    pdelta: dict,
    findings: list[str],
) -> io.BytesIO:
    document = Document()
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("ملخص تدقيق نتائج ETABS\nETABS Results Audit Summary")
    title_run.font.name = "Tahoma"
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 121)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(
        f"المشروع: {project['name']} | الموقع: {project['location']}\n"
        f"المصمم: {project['consultant']} | الاتصال: {project['contact']}"
    )

    document.add_heading("بيانات وإعدادات المشروع", level=2)
    for label, value in (
        ("الاستخدام", project["usage"]),
        ("عدد الطوابق", project["stories"]),
        ("الارتفاع الإجمالي (ملم)", project["height"]),
        ("خرسانة العناصر الشاقولية", project["fc_vertical"]),
        ("خرسانة العناصر الأفقية", project["fc_horizontal"]),
        ("حديد التسليح", project["rebar"]),
        ("الأحمال الميتة الإضافية", project["sdl"]),
        ("الأحمال الحية", project["ll"]),
        ("النظام الزلزالي", project["system"]),
        ("المنطقة الزلزالية Z", project["zone"]),
        ("معامل التربة S", project["soil"]),
        ("R", project["r"]),
        ("Cd", project["cd"]),
        ("Ie", project["ie"]),
        ("فئة المخاطر وفق الكود العراقي 2019", project["risk_category"] or "غير محددة"),
        ("فئة التصميم الزلزالي SDC", project["seismic_design_category"] or "غير محددة"),
        (
            "انطباق حد فئة المنشآت الخرسانية الأخرى",
            "مؤكد من المستخدم" if project["drift_applicability_confirmed"] else "غير مؤكد",
        ),
        (
            "نسبة اللامركزية العرضية المختارة",
            (
                f"{project['accidental_eccentricity_percent']:.2f}%"
                if project["accidental_eccentricity_percent"] is not None
                else "غير محددة"
            ),
        ),
        (
            "تأكيد حالات الالتواء العرضي",
            "مؤكد من المستخدم" if project["accidental_torsion_confirmed"] else "غير مؤكد",
        ),
        (
            "تأكيد انطباق معامل Ax",
            "مؤكد من المستخدم" if project["ax_applicability_confirmed"] else "غير مؤكد",
        ),
        ("Tx (sec)", project["tx"]),
        ("Ty (sec)", project["ty"]),
        ("Dx (ملم)", project["dx"]),
        ("Dy (ملم)", project["dy"]),
        ("نسبة القص المستهدفة", f"{project['target']:.1f}%"),
    ):
        document.add_paragraph(f"{label}: {value}")

    document.add_heading("حالة الفحوص", level=2)
    table = document.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0].cells[0].text = "الفحص"
    table.rows[0].cells[1].text = "الحالة"
    for name, status in statuses.items():
        cells = table.add_row().cells
        cells[0].text = SECTION_NAMES[name]
        cells[1].text = status_label(status)

    document.add_heading("النتائج الرقمية المحسوبة", level=2)
    if findings:
        for finding in findings:
            document.add_paragraph(finding, style="List Bullet")
    else:
        document.add_paragraph("لم تتوفر نتائج رقمية مكتملة لإدراجها.")

    document.add_heading("ملاحظات على منهجية الحساب", level=2)
    document.add_paragraph(
        "حد الانجراف المستخدم هو قاعدة القالب: 0.02/(0.7×R) عندما تكون الفترة "
        "أكبر من 0.7 ثانية، وإلا 0.025/(0.7×R). هذه ترجمة لقاعدة القالب وليست "
        "شهادة مستقلة بمطابقة الكود."
    )
    document.add_paragraph(
        "فحص الصلابة ونسبة أقصى/متوسط الإزاحة معروضان كبيانات خام دون حكم؛ "
        "لم يتوفر في القالب تعريف مكتمل لحدود التحقق."
    )
    document.add_paragraph(
        "الأصل في هذه الفحوص هو الكود الزلزالي العراقي 2019. فحص الانجراف يفترض "
        "أن قيم Story Drifts خطية/مرنة، ويحسب انجراف التصميم = الانجراف المرن × "
        "Cd / Ie وفق معادلة الانجراف في الفصل 3، §3-9-7، المعادلة (3-9-14). "
        "حدود الصف المختار للمنشآت الخرسانية الأخرى: 0.020 لفئتي المخاطر I/II، "
        "و0.015 للفئة III، و0.010 للفئة IV (جدول حدود الانجراف، الفصل 3، "
        "الصفحة المطبوعة 39/3). لا يُحسب الفحص قبل تحديد الفئة وتأكيد انطباق "
        "هذا النوع من المنشآت."
    )
    document.add_paragraph(
        "حد 90% لمقارنة SumUX وSumUY هو المعيار الذي قدمه المستخدم، ويطبق "
        "لكل حالة نمطية على حدة؛ لم يُتحقق رقمياً من نص الحد في نسخة الكود "
        "المحلية، لذلك لا يُنسب هذا الرقم إلى الكود العراقي. اللامركزية الجوهرية "
        "تحسب من إحداثيات CM/CR. نسبة اللامركزية العرضية "
        "لا تُفترض: مستخرج النص المحلي لصفحة الالتواء من نسخة الكود المقدمة "
        "يعرض قيمة ملتبسة؛ لذلك يلزم اختيار النسبة صراحة لكل تحليل. لا تثبت "
        "الإزاحة النظرية وحدها إدراج اللامركزية العرضية في حالات التحليل."
    )
    document.add_paragraph(
        "تُعرض عتبتا 1.2 و1.4 كنسب فحص مرشحة من معايير ASCE 7-16 التي قدمها "
        "المستخدم، ولا تُنسبان إلى الكود العراقي ما لم يثبت ذلك من نصه. معادلة "
        "Ax وحده الأقصى 3.0 مأخوذان من الكود العراقي، الفصل 3، §3-9-5، "
        "المعادلة (3-9-12). لا يُحسب Ax إلا بعد تأكيد تضمين اللامركزية العرضية "
        "في حالات ETABS، واختيار SDC وتأكيد انطباق شروطه؛ ولا يتحقق التطبيق من "
        "تنفيذ الحالات أو المعامل فعلياً."
    )
    document.add_paragraph(str(pdelta["message"]))
    for item in pdelta["required_inputs"]:
        document.add_paragraph(item, style="List Bullet")
    document.add_paragraph(
        "هذا التقرير يلخص الفحوص المحسوبة من الملف المرفوع ولا يستبدل مراجعة "
        "المهندس الإنشائي أو اعتماد المعايير التصميمية للمشروع."
    )

    stream = io.BytesIO()
    document.save(stream)
    stream.seek(0)
    return stream


status_dict = {name: "WAITING" for name in SECTION_NAMES}
status_placeholder = st.sidebar.empty()


def render_sidebar_status(statuses: dict[str, str]) -> None:
    with status_placeholder.container():
        st.header("📋 حالة أقسام التدقيق")
        for name, status in statuses.items():
            st.write(f"**{name}**: {status_label(status)}")


render_sidebar_status(status_dict)
st.sidebar.markdown("---")

st.sidebar.header("📋 بيانات المشروع الأساسية")
project_name = st.sidebar.text_input("اسم المشروع", value="مبنى تجاري")
project_location = st.sidebar.text_input("موقع المشروع", value="الموصل")
consultant_name = st.sidebar.text_input(
    "المصمم الاستشاري", value="عبدالغفور رمضان اللهيبي"
)
contact_info = st.sidebar.text_input("معلومات الاتصال", value="")
building_usage = st.sidebar.selectbox(
    "طبيعة الاستخدام",
    ["مبنى صناعي", "مبنى إداري", "مبنى سكني", "مبنى تجاري", "منشأ تعليمي / خدمي"],
)
num_floors = st.sidebar.number_input("عدد الطوابق المتكررة", value=2, step=1)
ground_height = st.sidebar.number_input(
    "ارتفاع الطابق الأرضي (ملم)", value=4500, step=500
)
typical_height = st.sidebar.number_input(
    "ارتفاع الطابق المتكرر (ملم)", value=3500, step=500
)
total_height = ground_height + num_floors * typical_height
num_stories = num_floors + 1
st.sidebar.write(f"الارتفاع الإجمالي للمبنى: {total_height:.0f} ملم")

st.sidebar.markdown("---")
st.sidebar.header("🧱 خواص المواد الإنشائية")
fc_vertical = st.sidebar.selectbox(
    "مقاومة خرسانة العناصر الشاقولية",
    ["f'c = 30 MPa", "f'c = 35 MPa", "f'c = 40 MPa", "f'c = 45 MPa"],
)
fc_horizontal = st.sidebar.selectbox(
    "مقاومة خرسانة العناصر الأفقية",
    ["f'c = 25 MPa", "f'c = 28 MPa", "f'c = 30 MPa", "f'c = 35 MPa"],
)
rebar_grade = st.sidebar.selectbox(
    "حديد التسليح الرئيسي",
    [
        "Grade 60 (fy = 420 MPa)",
        "Grade 40 (fy = 280 MPa)",
        "Grade 50 (fy = 350 MPa)",
        "Grade 75 (fy = 520 MPa)",
        "Strands - Low Relaxation Grade 270 (fpu = 1860 MPa)",
    ],
)

st.sidebar.markdown("---")
st.sidebar.header("⚖️ نطاقات الأحمال التصميمية")
sdl_range = st.sidebar.selectbox(
    "الأحمال الميتة الإضافية (SDL)",
    [
        "2.0 - 4.0 kN/m² (مخازن / أحمال تشطيبية وجدران)",
        "3.0 - 4.0 kN/m² (أحمال السلالم)",
        "2.0 - 3.0 kN/m² (طوابق سكنية / مكتبية)",
        "4.0 - 5.0 kN/m² (أسطح وعزل ثقيل)",
    ],
)
ll_range = st.sidebar.selectbox(
    "الأحمال الحية (LL)",
    [
        "5.0 - 10.0 kN/m² (مخازن واستخدامات ثقيلة)",
        "3.0 - 3.5 kN/m² (البالكونات)",
        "1.0 - 1.5 kN/m² (الأسطح)",
        "3.0 kN/m² (المطابخ والحمامات)",
        "3.0 kN/m² (السلالم السكنية)",
        "4.8 kN/m² (السلالم الإدارية والفندقية)",
        "2.0 kN/m² (الغرف السكنية)",
    ],
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ إعدادات الزلازل والتربة")
zone_options = {
    "المنطقة الزلزالية الثانية (الموصل / نينوى) - Z = 0.15": 0.15,
    "المنطقة الزلزالية الأولى (بغداد / وسط العراق) - Z = 0.10": 0.10,
    "المنطقة الزلزالية الثالثة (الشمال الشرقي) - Z = 0.20": 0.20,
}
selected_zone = st.sidebar.selectbox("المنطقة الزلزالية (Z)", list(zone_options))
zone_z = zone_options[selected_zone]
soil_options = {
    "النوع (II): تربة متوسطة / كثيفة (Stiff Soil - S_C/Class D)": 1.2,
    "النوع (I): صخور صلبة (Rock - S_A/S_B)": 1.0,
    "النوع (III): تربة طينية سميكة (Soft Soil - S_D)": 1.5,
}
selected_soil = st.sidebar.selectbox("نوع التربة (Site Class)", list(soil_options))
soil_factor_s = soil_options[selected_soil]
importance_factor = st.sidebar.selectbox("عامل الأهمية (Ie)", [1.0, 1.25, 1.5])
risk_category_option = st.sidebar.selectbox(
    "فئة المخاطر وفق الكود العراقي 2019",
    ["— اختر الفئة لكل تحليل —", "I", "II", "III", "IV"],
)
risk_category = None if risk_category_option.startswith("—") else risk_category_option
sdc_option = st.sidebar.selectbox(
    "فئة التصميم الزلزالي SDC",
    ["— اختر الفئة لكل تحليل —", "A", "B", "C", "D", "E", "F"],
)
seismic_design_category = None if sdc_option.startswith("—") else sdc_option
drift_applicability_confirmed = st.sidebar.checkbox(
    "أؤكد انطباق حد «المنشآت الخرسانية الأخرى» للانجراف"
)
accidental_percent_option = st.sidebar.selectbox(
    "نسبة اللامركزية العرضية حسب المرجع المختار",
    [
        "— اختر النسبة لكل تحليل —",
        "2% (من استخراج نص الكود العراقي؛ يلزم التحقق من الصفحة)",
        "5% (معيار ASCE 7-16 المرشح)",
        "نسبة أخرى",
    ],
)
if accidental_percent_option.startswith("2%"):
    accidental_eccentricity_percent = 2.0
elif accidental_percent_option.startswith("5%"):
    accidental_eccentricity_percent = 5.0
elif accidental_percent_option == "نسبة أخرى":
    custom_accidental_percent = st.sidebar.number_input(
        "نسبة اللامركزية العرضية المخصصة (%) — الصفر يعني غير محددة",
        min_value=0.0,
        max_value=100.0,
        value=0.0,
        step=0.5,
    )
    accidental_eccentricity_percent = (
        custom_accidental_percent if custom_accidental_percent > 0 else None
    )
else:
    accidental_eccentricity_percent = None
accidental_torsion_confirmed = st.sidebar.checkbox(
    "أؤكد أن الحالات المختارة في ETABS تشمل النسبة المحددة وبالإشارتين",
    disabled=accidental_eccentricity_percent is None,
)
ax_applicability_confirmed = st.sidebar.checkbox(
    "أؤكد انطباق شروط Ax على فئة SDC والنظام المختارين",
    disabled=not accidental_torsion_confirmed or seismic_design_category is None,
)
st.sidebar.caption(
    "تُختار فئة المخاطر وSDC والانطباق لكل تحليل. يظهر استخراج نص صفحة الالتواء "
    "رقم 2% لكن تعذر التحقق البصري من الصفحة؛ لذلك لا تُختار أي نسبة تلقائياً. "
    "حساب Ax يحتاج تأكيد إدراج اللامركزية في الحالات وانطباق شروطه."
)
seismic_systems = {
    "جدران قص / جدران حاملة [R=5.0, Cd=4.5]": (5.0, 4.5),
    "إطارات خرسانية خاصة [R=8.0, Cd=5.5]": (8.0, 5.5),
    "نظام مزدوج [R=7.0, Cd=5.5]": (7.0, 5.5),
}
selected_system = st.sidebar.selectbox(
    "النظام الإنشائي المقاوم للزلازل", list(seismic_systems)
)
response_mod_r, deflection_cd = seismic_systems[selected_system]
tx_val = st.sidebar.number_input(
    "الفترة الأساسية Tx (ثانية)", min_value=0.0, value=0.0, step=0.01, format="%.3f"
)
ty_val = st.sidebar.number_input(
    "الفترة الأساسية Ty (ثانية)", min_value=0.0, value=0.0, step=0.01, format="%.3f"
)
st.sidebar.caption("أدخل الفترتين من التصميم؛ الصفر يعني أن فحص الانجراف غير محسوب.")
asce_percentage = st.sidebar.slider(
    "النسبة المطلوبة للقص الديناميكي من الاستاتيكي (%)",
    80.0,
    100.0,
    85.0,
    5.0,
)

uploaded_file = st.file_uploader(
    "ارفع ملف نتائج ETABS الخام باسم results.xlsx", type=["xlsx"]
)
with st.expander("أسماء الأوراق المطلوبة لإجراء الحسابات"):
    st.markdown(
        "يجب أن يحتوي الملف على الأوراق التالية (تُقبل كتابة الاسم مع اختلاف المسافات أو الشرطة السفلية):"
    )
    for _key, aliases in REQUIRED_SHEET_NAMES:
        st.markdown(f"- `{ ' أو '.join(aliases) }`")
    st.markdown(
        "**جداول إضافية اختيارية:** "
        + "، ".join(f"`{label}`" for _key, label, _aliases in OPTIONAL_SHEET_NAMES)
        + ". غيابها لا يمنع الفحوص الأساسية، لكنه يقلل البيانات المتاحة للمراجعة."
    )
    st.markdown(
        "**مصادر P-Delta الاختيارية:** "
        + "، ".join(f"`{name}`" for name in OPTIONAL_PDELTA_SHEETS)
        + ". حتى مع وجودها يبقى فحص P-Delta غير محسوب إلى أن تُؤكد معادلة ومدخلات الحالة."
    )

st.markdown("---")
st.markdown("### 📐 أبعاد المبنى")
col_dx, col_dy = st.columns(2)
with col_dx:
    dx_val = st.number_input(
        "البعد الإجمالي للمبنى Dx (ملم)", min_value=0.0, value=0.0, step=100.0
    )
with col_dy:
    dy_val = st.number_input(
        "البعد الإجمالي للمبنى Dy (ملم)", min_value=0.0, value=0.0, step=100.0
    )
st.caption("أدخل أبعاد المخطط الفعلية؛ الصفر يعني أن حسابات اللامركزية الهندسية غير محسوبة.")

if uploaded_file is None:
    st.session_state.pop("etabs_workbook", None)
    st.session_state.pop("workbook_key", None)
    st.session_state.pop("upload_blocked", None)
    st.info("ارفع مصنف ETABS لبدء التحليل. لا توجد نتائج ثابتة مضمّنة في التطبيق.")
else:
    try:
        validate_upload_filename(uploaded_file.name)
    except WorkbookValidationError as exc:
        st.session_state.pop("etabs_workbook", None)
        st.session_state.pop("workbook_key", None)
        st.error(str(exc))
        st.info("اسم الملف الصحيح المطلوب: `results.xlsx`.")
        st.session_state["upload_blocked"] = True
    else:
        st.session_state.pop("upload_blocked", None)
        workbook_bytes = uploaded_file.getvalue()
        try:
            workbook_key, parsed_workbook = get_cached_workbook(
                st.session_state, workbook_bytes
            )
            st.session_state["etabs_workbook"] = parsed_workbook
            st.session_state["workbook_key"] = workbook_key
        except WorkbookValidationError as exc:
            st.session_state.pop("etabs_workbook", None)
            st.session_state["workbook_key"] = None
            st.session_state["upload_blocked"] = True
            st.error(str(exc))
            for name in status_dict:
                status_dict[name] = "ERROR"
            status_dict["P-Delta Check"] = "NOT_EVALUATED"
            render_sidebar_status(status_dict)

    st.caption(
        "تُحفظ الجداول المقروءة مؤقتاً حسب محتوى الملف (آخر أربع نسخ). "
        "اختر نسخة سابقة من القائمة لاستعادة قيمها بعد تعديل الملف."
    )

    workbook_versions = st.session_state.get("_etabs_workbook_versions", {})
    if workbook_versions and not st.session_state.get("upload_blocked"):
        if st.session_state.get("last_uploaded_workbook_key") != workbook_key:
            st.session_state["last_uploaded_workbook_key"] = workbook_key
            st.session_state["selected_cached_workbook_key"] = workbook_key

        version_keys = list(reversed(workbook_versions))
        if st.session_state.get("selected_cached_workbook_key") not in version_keys:
            st.session_state["selected_cached_workbook_key"] = version_keys[0]
        st.selectbox(
            "نسخة الملف المستخدمة (يمكن الرجوع إلى نسخة سابقة)",
            options=version_keys,
            format_func=lambda key: f"نسخة محفوظة: {key[:8]}",
            key="selected_cached_workbook_key",
        )
        selected_key = st.session_state["selected_cached_workbook_key"]
        st.session_state["workbook_key"] = selected_key
        st.session_state["etabs_workbook"] = workbook_versions[selected_key]

    workbook = st.session_state.get("etabs_workbook")
    if st.session_state.get("upload_blocked"):
        workbook = None
    if workbook is not None:
        tables = workbook.tables
        report_findings = []

        st.header("1. مراكز الكتلة والصلابة")
        st.caption("eₓ = |XCR − XCM|، eᵧ = |YCR − YCM|؛ النسبة تقارن بأبعاد Dx وDy المدخلة.")
        if dx_val <= 0 or dy_val <= 0:
            status_dict["Centers Of Mass"] = "NOT_EVALUATED"
            st.info("أدخل Dx وDy الفعليين لحساب النسب الهندسية.")
            st.dataframe(tables["centers"], width="stretch")
        else:
            try:
                centers, status_dict["Centers Of Mass"] = calculate_eccentricity(
                    tables["centers"], dx_val, dy_val
                )
                st.dataframe(centers.round(3), width="stretch")
                max_x = centers["ex_dx_percent"].dropna().max()
                max_y = centers["ey_dy_percent"].dropna().max()
                if pd.notna(max_x) and pd.notna(max_y):
                    report_findings.append(
                        f"أقصى لا مركزية نسبية: X = {max_x:.2f}%، Y = {max_y:.2f}%."
                    )
            except WorkbookValidationError as exc:
                status_dict["Centers Of Mass"] = "ERROR"
                st.error(str(exc))
        st.markdown("---")

        st.header("2. القص القاعدي ومعامل التحجيم")
        st.caption(
            "يقارن أكبر قيمة مطلقة لـ RSX/RSY مع EQX/EQY مضروبة بنسبة القص المستهدفة."
        )
        try:
            base_reactions, status_dict["Base Reactions"] = calculate_base_reactions(
                tables["base_reactions"],
                asce_percentage / 100,
                importance_factor,
                response_mod_r,
            )
            st.dataframe(base_reactions.round(4), width="stretch")
            for row in base_reactions.itertuples(index=False):
                scale = (
                    f"{row.recommended_scale_factor:.4f}"
                    if pd.notna(row.recommended_scale_factor)
                    else "غير متاح"
                )
                report_findings.append(
                    f"الاتجاه {row.direction}: القص الاستاتيكي {row.static_shear:.2f}، "
                    f"القص الديناميكي {row.dynamic_shear:.2f}، معامل التحجيم المقترح "
                    f"{scale}."
                )
        except WorkbookValidationError as exc:
            status_dict["Base Reactions"] = "ERROR"
            st.error(str(exc))
        st.markdown("---")

        st.header("3. انجراف الطوابق")
        st.caption(
            "حد القالب: 0.02/(0.7×R) إذا كانت الفترة أكبر من 0.7 ثانية، وإلا "
            "0.025/(0.7×R). هذا حد القالب ولا يمثل شهادة مطابقة كودية."
        )
        try:
            drifts, status_dict["Story Drifts"] = calculate_story_drifts(
                tables["story_drifts"], response_mod_r, tx_val, ty_val
            )
            st.dataframe(drifts.round(6), width="stretch")
            for direction, direction_rows in drifts.groupby("direction"):
                max_drift = direction_rows["drift_ratio"].dropna().max()
                limit = direction_rows["limit_ratio"].dropna().max()
                if pd.notna(max_drift) and pd.notna(limit):
                    report_findings.append(
                        f"الانجراف الأقصى في {direction}: {max_drift:.6f}؛ "
                        f"حد القالب: {limit:.6f}."
                    )
        except WorkbookValidationError as exc:
            status_dict["Story Drifts"] = "NOT_EVALUATED"
            st.warning(f"تعذر إصدار حكم على الانجراف: {exc}")
            st.dataframe(tables["story_drifts"], width="stretch")
        st.markdown("---")

        st.header("فحص الانجراف وفق الكود الزلزالي العراقي 2019")
        st.caption(
            "يفترض هذا الفحص أن قيم ETABS في Story Drifts انجرافات مرنة، "
            "ويحسب انجراف التصميم = القيمة المرنة × Cd / Ie وفق الفصل 3، "
            "§3-9-7، المعادلة (3-9-14). يستخدم صف «المنشآت الخرسانية الأخرى» "
            "من جدول حدود الانجراف، الصفحة المطبوعة 39/3: I/II = 0.020، "
            "III = 0.015، IV = 0.010. يلزم اختيار فئة المخاطر وتأكيد انطباق الصف."
        )
        iraqi_drifts, status_dict["Iraqi 2019 Drift"] = calculate_iraqi_2019_story_drifts(
            tables["story_drifts"],
            risk_category,
            drift_applicability_confirmed,
            deflection_cd,
            importance_factor,
        )
        if iraqi_drifts.empty:
            st.warning("لم تتوفر صفوف جانبية صالحة في Story Drifts لإجراء الفحص.")
        else:
            st.dataframe(iraqi_drifts.round(6), width="stretch")
            for direction, direction_rows in iraqi_drifts.groupby("direction"):
                max_drift = direction_rows["design_drift_ratio"].dropna().max()
                limit = direction_rows["limit_ratio"].dropna().max()
                if pd.notna(max_drift) and pd.notna(limit):
                    report_findings.append(
                        f"الكود العراقي 2019، الفئة {risk_category}، الاتجاه {direction}: "
                        f"أقصى انجراف تصميم {max_drift:.6f} مقابل حد {limit:.6f} "
                        f"(Cd/Ie = {deflection_cd / importance_factor:.3f})."
                    )
        if status_dict["Iraqi 2019 Drift"] == "NOT_EVALUATED":
            st.info(
                "فحص الكود العراقي غير محسوب حتى اختيار فئة المخاطر وتأكيد انطباق "
                "حد «المنشآت الخرسانية الأخرى»."
            )
        st.markdown("---")

        st.header("4. مشاركة الكتلة النمطية — معيار 90% المقدم")
        st.caption(
            "تُفحص أكبر مشاركة كتلة تراكمية SumUX/SumUY لكل حالة في كل اتجاه "
            "مقابل 90% لكل حالة نمطية. هذا الحد مقدم من المستخدم، وليس نسخاً "
            "موثقاً من نص الكود العراقي."
        )
        modal, status_dict["Modal Mass Ratios"] = calculate_modal_mass(
            tables["modal_mass"]
        )
        st.dataframe(modal.round(3), width="stretch")
        for row in modal.itertuples(index=False):
            if pd.notna(row.maximum_percent):
                report_findings.append(
                    f"الحالة النمطية {row.case}، الاتجاه {row.direction}: "
                    f"أقصى مشاركة كتلة تراكمية {row.maximum_percent:.2f}%."
                )
        st.markdown("---")

        st.header("اللامركزية والالتواء — الكود العراقي 2019")
        st.caption(
            "لامركزية X تستخدم فرق YCM−YCR والبعد Dy، ولامركزية Y تستخدم "
            "XCM−XCR والبعد Dx. لا تُفترض نسبة اللامركزية العرضية: مستخرج نص "
            "الصفحة المطبوعة 16/3 يظهر 2% لكن تعذر التحقق البصري منه. اختر النسبة "
            "المعتمدة قبل حساب الإزاحات النظرية."
        )
        try:
            eccentricities, eccentricity_status = (
                calculate_iraqi_2019_torsion_eccentricities(
                    tables["centers"],
                    dx_val,
                    dy_val,
                    accidental_eccentricity_percent,
                )
            )
            if eccentricities.empty:
                st.warning("لا توجد إحداثيات CM/CR كافية لحساب اللامركزية.")
            else:
                st.dataframe(eccentricities.round(3), width="stretch")
                available_eccentricities = eccentricities[
                    eccentricities["check"] == "INFORMATIONAL"
                ]
                for direction, direction_rows in available_eccentricities.groupby(
                    "direction"
                ):
                    max_inherent = direction_rows["inherent_absolute"].max()
                    max_accidental = direction_rows["accidental_magnitude"].max()
                    accidental_summary = (
                        f"مقدار ±{accidental_eccentricity_percent:.2f}% النظري "
                        f"{max_accidental:.2f} ملم."
                        if pd.notna(max_accidental)
                        else "نسبة اللامركزية العرضية لم تُحدد."
                    )
                    report_findings.append(
                        f"الاتجاه {direction}: أكبر لامركزية جوهرية "
                        f"{max_inherent:.2f} ملم؛ {accidental_summary}"
                    )
        except WorkbookValidationError as exc:
            eccentricity_status = "NOT_EVALUATED"
            st.warning(f"تعذر حساب اللامركزية وفق الكود العراقي: {exc}")
        if "diaphragm_drifts" in tables:
            torsional, status_dict["Iraqi Torsion"] = (
                calculate_iraqi_2019_torsional_irregularity(
                    tables["diaphragm_drifts"],
                    accidental_torsion_confirmed,
                    accidental_eccentricity_percent,
                    seismic_design_category,
                    ax_applicability_confirmed,
                )
            )
            if torsional.empty:
                st.warning(
                    "لم توجد صفوف زلزالية صالحة في Diaphragm Max Over Avg Drifts."
                )
            else:
                st.dataframe(torsional.round(4), width="stretch")
                valid_torsion = torsional[torsional["ratio"].notna()]
                if not valid_torsion.empty:
                    max_ratio = valid_torsion["ratio"].max()
                    has_ax = valid_torsion["ax"].notna().any()
                    report_findings.append(
                        f"أقصى نسبة أقصى/متوسط انجراف للحجاب الحاجز: "
                        f"{max_ratio:.3f}. عتبتا 1.2/1.4 من معايير ASCE 7-16 "
                        f"المرشحة. Ax "
                        f"{'حُسب بعد تأكيد الشروط' if has_ax else 'غير محسوب؛ شروط الانطباق أو بياناته غير مؤكدة'}."
                    )
                if not accidental_torsion_confirmed:
                    st.info(
                        "تعرض النسب كتشخيص أولي فقط. لا يصدر تصنيف لالتواء كودي "
                        "ولا يحسب Ax قبل تأكيد أن الحالات تشمل اللامركزية العرضية "
                        "بالنسبة المختارة وبالإشارتين."
                    )
                elif not ax_applicability_confirmed:
                    st.info(
                        "تم تصنيف النسبة كإشارة معلوماتية، لكن Ax غير محسوب حتى "
                        "اختيار SDC وتأكيد انطباق شروطه."
                    )
                elif status_dict["Iraqi Torsion"] == "INFORMATIONAL":
                    st.caption(
                        "وفق الكود العراقي 2019، الفصل 3، §3-9-5، المعادلة "
                        "(3-9-12): Ax = min((δmax/(1.2δavg))², 3.0). "
                        "لا يتحقق التطبيق من تنفيذ Ax فعلياً في حالات ETABS."
                    )
        else:
            status_dict["Iraqi Torsion"] = "NOT_EVALUATED"
            issue = workbook.optional_table_issues.get("diaphragm_drifts")
            if issue:
                st.warning(
                    "تعذر استخدام جدول Diaphragm Max Over Avg Drifts: " + issue
                )
            else:
                st.info(
                    "جدول Diaphragm Max Over Avg Drifts غير موجود؛ لا يمكن تقييم "
                    "نسبة الالتواء العرضي أو Ax."
                )
        st.caption(
            f"حالة حساب اللامركزية الهندسية: {status_label(eccentricity_status)}. "
            "هذه قيم هندسية/استشارية وليست حكماً مستقلاً بالمطابقة."
        )
        st.markdown("---")

        st.header("5. أقصى/متوسط الإزاحة")
        st.caption("تُعرض قيم ETABS الخام فقط؛ حد قبول هذه النسبة غير محدد في القالب.")
        st.dataframe(tables["displacements"], width="stretch")
        status_dict["Displacements"] = "NOT_EVALUATED"
        st.markdown("---")

        st.header("6. صلابة الطوابق")
        st.caption("تُعرض قيم الجساءة الخام؛ لا يُستنتج فحص الطابق اللين من قيم الجساءة وحدها.")
        st.dataframe(tables["story_stiffness"], width="stretch")
        status_dict["Story Stiffness"] = "NOT_EVALUATED"
        st.markdown("---")

        st.header("7. فحص P-Delta")
        pdelta = pdelta_assessment(workbook.pdelta_source_issue)
        status_dict["P-Delta Check"] = pdelta["status"]
        st.warning(pdelta["message"])
        st.write("المدخلات/التعريفات المطلوبة قبل إصدار نتيجة:")
        for item in pdelta["required_inputs"]:
            st.markdown(f"- {item}")
        st.markdown("---")
        st.header("8. جداول ETABS الإضافية")
        st.caption(
            "تعرض هذه الجداول كما وردت في الملف. وجودها يتيح مراجعة بيانات إضافية، "
            "ولا ينشئ فحص نجاح/فشل تلقائياً."
        )
        for key, label, _aliases in OPTIONAL_SHEET_NAMES:
            if key in tables:
                with st.expander(f"{label} — متوفر"):
                    st.dataframe(tables[key], width="stretch")
            else:
                st.info(f"الجدول الاختياري `{label}` غير موجود في الملف المرفوع.")
                report_findings.append(f"الجدول الاختياري {label}: غير موجود.")

        render_sidebar_status(status_dict)

        st.sidebar.markdown("---")
        st.sidebar.header("📄 تصدير التقرير")
        project = {
            "name": project_name,
            "location": project_location,
            "consultant": consultant_name,
            "contact": contact_info,
            "usage": building_usage,
            "stories": num_stories,
            "height": total_height,
            "fc_vertical": fc_vertical,
            "fc_horizontal": fc_horizontal,
            "rebar": rebar_grade,
            "sdl": sdl_range,
            "ll": ll_range,
            "system": selected_system,
            "zone": zone_z,
            "soil": soil_factor_s,
            "r": response_mod_r,
            "cd": deflection_cd,
            "ie": importance_factor,
            "risk_category": risk_category,
            "seismic_design_category": seismic_design_category,
            "drift_applicability_confirmed": drift_applicability_confirmed,
            "accidental_eccentricity_percent": accidental_eccentricity_percent,
            "accidental_torsion_confirmed": accidental_torsion_confirmed,
            "ax_applicability_confirmed": ax_applicability_confirmed,
            "tx": tx_val,
            "ty": ty_val,
            "dx": dx_val,
            "dy": dy_val,
            "target": asce_percentage,
        }
        report = generate_word_report(project, status_dict, pdelta, report_findings)
        st.sidebar.download_button(
            "📥 تحميل ملخص تدقيق ETABS (Word)",
            data=report,
            file_name=f"{project_name}_ETABS_Audit_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
