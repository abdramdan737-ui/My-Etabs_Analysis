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
}
SECTION_NAMES = {
    "Centers Of Mass": "مراكز الكتلة والصلابة",
    "Base Reactions": "القص القاعدي ومعامل التحجيم",
    "Story Drifts": "انجراف الطوابق",
    "Modal Mass Ratios": "مشاركة الكتلة النمطية",
    "Displacements": "أقصى/متوسط الإزاحة",
    "Story Stiffness": "صلابة الطوابق",
    "P-Delta Check": "فحص P-Delta",
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
        "البعد الإجمالي للمبنى Dx (ملم)", min_value=0.0, value=16000.0, step=100.0
    )
with col_dy:
    dy_val = st.number_input(
        "البعد الإجمالي للمبنى Dy (ملم)", min_value=0.0, value=25000.0, step=100.0
    )

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

        st.header("4. مشاركة الكتلة النمطية")
        st.caption("تُفحص أكبر مشاركة كتلة تراكمية في كل اتجاه مقابل 90%.")
        modal, status_dict["Modal Mass Ratios"] = calculate_modal_mass(
            tables["modal_mass"]
        )
        st.dataframe(modal.round(3), width="stretch")
        for row in modal.itertuples(index=False):
            if pd.notna(row.maximum_percent):
                report_findings.append(
                    f"أقصى مشاركة كتلة تراكمية في {row.direction}: "
                    f"{row.maximum_percent:.2f}%."
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
