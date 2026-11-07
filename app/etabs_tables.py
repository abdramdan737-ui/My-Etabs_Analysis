import io

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, RGBColor

from app.authentication import require_authentication
from app.etabs_processing import (
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
    calculate_pdelta_screening,
    get_cached_optional_table,
    get_cached_workbook,
    pdelta_combination_options,
    validate_upload_filename,
)


st.set_page_config(
    page_title="منصة تدقيق نتائج الايتابس وتوليد النوتة الحسابية", layout="wide"
)

require_authentication()

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

with st.expander("📘 دليل المستخدم — خطوات التسجيل ورفع البيانات", expanded=True):
    st.markdown(
        """
        **اتبع الخطوات بالترتيب:**

        1. **إنشاء الحساب:** اختر المتابعة إلى Auth0، وأنشئ حساباً ببريدك ثم أكّد رسالة التحقق. بعد ذلك ارجع إلى التطبيق وسجّل الدخول. الحساب غير المؤكد لا يستطيع فتح أدوات التحليل.
        2. **احمِ بيانات المشروع:** لا ترفع إلا ملفات مسموحاً لك بمشاركتها مع خدمة الاستضافة. تُرسل الملفات إلى خادم التطبيق للتحليل؛ لا ترفع بيانات سرية قبل مراجعة سياسة الخصوصية والاحتفاظ.
        3. **جهّز الملف الأساسي:** صدّر جداول ETABS إلى مصنف `.xlsx` وسمّه بالضبط `results.xlsx`. يجب أن يتضمن الأوراق الأساسية المعروضة أدناه. يتوقع التطبيق عنوان الجدول في الصف الأول، وأسماء الأعمدة في الصف الثاني، والوحدات في الصف الثالث، ثم البيانات. لا تغيّر أسماء الأعمدة أو تحذف بياناتها المطلوبة. الأوراق الإضافية داخل هذا الملف تُهمل.
        4. **ارفع الجداول الاختيارية منفصلة:** من الشريط الجانبي، ارفع ملف `.xlsx` مستقلاً لكل جدول، في المجموعة المرتبطة بفحصه. يجب أن يحتوي الملف على ورقة الجدول المطلوب؛ الأوراق الأخرى فيه لا تُستخدم. لا يلزم رفع جدول اختياري إذا لم تكن تريد الفحص الذي يعتمد عليه.
        5. **أدخل معلومات التحليل:** أدخل الأبعاد والفترات وفئات المخاطر وSDC والنظام الإنشائي والنسبة العرضية وفق مستندات المشروع. لا تترك قيمة مطلوبة على افتراض أنها استنتجت من الملف. لفحص P-Delta، اختر التركيبات وأكّد صلاحية معاملاتها و`Cd`/`Ie` وأن الانجراف مرن غير مكبّر مسبقاً.
        6. **راجع النتائج والحالة:** اقرأ التحذيرات وحالة كل فحص. **«غير محسوب» أو «غير مكتمل» لا تعني النجاح**؛ أكمل البيانات أو ارفع الجدول المطلوب ثم أعد المراجعة.
        7. **نزّل التقرير:** استخدم زر تنزيل تقرير Word في أسفل الشريط الجانبي، واحتفظ به وفق سياسة بيانات مشروعك. سجّل الخروج بعد الانتهاء، خصوصاً على جهاز مشترك.

        **مهم:** النتائج مساعدة للمراجعة الهندسية وليست اعتماداً أو شهادة سلامة، ولا تغني عن مراجعة مهندس إنشائي مؤهل. إذا تعذر التحليل بسبب اسم ورقة أو عمود، راجع رسالة الخطأ وقارن تصدير ETABS بمتطلبات الجداول أدناه.
        """
    )

STATUS_LABELS = {
    "WAITING": "⏳ بانتظار الرفع",
    "PASS": "محقق ✅",
    "FAIL": "غير محقق ❌",
    "INCOMPLETE": "غير مكتمل ⚠️",
    "NOT_EVALUATED": "غير محسوب ⚠️",
    "ERROR": "خطأ في المدخلات ⚠️",
    "INFORMATIONAL": "معلوماتي — لا يصدر حكم مطابقة",
    "PDELTA_NEGLIGIBLE": "يمكن إهمال P-Delta حسب الفحص",
    "PDELTA_REQUIRED": "يجب إدراج P-Delta في التحليل",
    "PDELTA_LIMIT_EXCEEDED": "تجاوز حد الاستقرار؛ يلزم إجراء هندسي",
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
    document.add_paragraph(
        "فحص P-Delta وفق الكود الزلزالي العراقي 2019، الفصل 3، §3-9/7/1: "
        "معادلة (3-9/12) تستخدم θ = P × Δ / (V × hs × Cd)، ومعادلة "
        "(3-9/16) تحدد θmax = min(0.25, 0.5/(β × Cd)). بما أن Drift في "
        "Story Drifts نسبة انجراف مرن بلا وحدة، ومعادلة انجراف التصميم في "
        "الكود هي Cd × الانجراف المرن / Ie، تختزل معادلة الاستقرار إلى "
        "θ = P × |Drift| / (|V| × Ie)، مع اعتبار P الموجب ضغطاً. استُخدم "
        "β = 1.0 تحفظياً كما يسمح الكود. إذا تجاوز θ قيمة θmax فهذه نتيجة "
        "غير مقبولة بهذا الفحص وتتطلب تحليلاً/مراجعة إنشائية إضافية وفق الكود؛ "
        "وإلا فـ θ ≤ 0.1 يسمح بإهمال الأثر، و0.1 < θ ≤ θmax "
        "يتطلب إدراج P-Delta."
    )
    document.add_paragraph(
        "استخدم الحساب صف Location = Bottom من Story Forces، وCase Type = "
        "Combination فقط، وطابق Output Case حرفيًا مع Story Drifts. استُبعدت "
        "مخرجات Max/Min وMax_envelope والصفوف ذات Step Number لأنها ليست "
        "حالة متزامنة صالحة للمقارنة، "
        "واستخدم أكبر قيمة مطلقة للانجراف بين نقاط الطابق. رُتبت الطوابق بحسب "
        "منسوب Z تنازليًا. لا يحوي الملف تعريفات معاملات التركيبات، لذا يعتمد "
        "الحساب على اختيار المستخدم وتأكيده، ولا يتحقق من صلاحية عواملها."
    )
    document.add_paragraph(
        "تُعرض فقط التركيبات التي لها محصلة قص قاعدي جانبية في Base Reactions. "
        "يُعد اتجاه X/Y مؤثراً إذا تجاوزت قيمة FX/FY المطلقة 1e-6 × |FZ|؛ "
        "هذه سماحية عددية لتمييز ضوضاء التقريب عن الحمل الجانبي وليست حداً "
        "للكود العراقي. ويظل الاتجاه الذي لا يحققها غير محسوب، لا ناجحاً."
    )
    document.add_paragraph(
        "التركيبات المختارة: "
        + (
            ", ".join(pdelta["selected_combinations"])
            if pdelta["selected_combinations"]
            else "لم تُحدد"
        )
    )
    document.add_paragraph(
        f"تأكيد معاملات التركيبات: {'نعم' if pdelta['combinations_confirmed'] else 'لا'}; "
        f"تأكيد Cd وIe: {'نعم' if pdelta['coefficients_confirmed'] else 'لا'}; "
        f"تأكيد أن Drift مرن وغير مكبّر مسبقاً: "
        f"{'نعم' if pdelta['drift_confirmed'] else 'لا'}."
    )
    governing = pdelta.get("governing_results")
    if isinstance(governing, pd.DataFrame) and not governing.empty:
        document.add_heading("نتائج P-Delta الحاكمة لكل طابق واتجاه", level=3)
        pdelta_table = document.add_table(rows=1, cols=11)
        pdelta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headings = (
            "الطابق",
            "Z (mm)",
            "الاتجاه",
            "التركيبة الحاكمة",
            "|قص القاعدة|",
            "P",
            "|V|",
            "|Drift|",
            "θ",
            "θmax",
            "التصنيف / السبب",
        )
        for cell, heading in zip(pdelta_table.rows[0].cells, headings):
            cell.text = heading
        classifications = {
            "NEGLIGIBLE": "يمكن الإهمال",
            "INCLUDE_PDELTA": "يجب إدراج P-Delta",
            "LIMIT_EXCEEDED": "تجاوز θmax",
            "NOT_EVALUATED": "غير محسوب",
        }
        for row in governing.itertuples(index=False):
            cells = pdelta_table.add_row().cells
            values = (
                row.story,
                f"{row.story_z:.0f}" if pd.notna(row.story_z) else "غير متاح",
                row.direction,
                row.output_case,
                f"{row.base_shear:.5f}"
                if pd.notna(row.base_shear)
                else "غير مؤثر/غير محسوب",
                f"{row.p:.5f}" if pd.notna(row.p) else "غير محسوب",
                f"{row.story_shear:.5f}"
                if pd.notna(row.story_shear)
                else "غير محسوب",
                f"{row.elastic_drift_ratio:.7f}"
                if pd.notna(row.elastic_drift_ratio)
                else "غير محسوب",
                f"{row.theta:.5f}" if pd.notna(row.theta) else "غير محسوب",
                f"{row.theta_max:.5f}" if pd.notna(row.theta_max) else "غير متاح",
                classifications.get(row.classification, row.classification)
                + (f": {row.reason}" if row.reason else ""),
            )
            for cell, value in zip(cells, values):
                cell.text = str(value)
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

optional_upload_groups = (
    ("الالتواء واللامركزية — الكود العراقي 2019", ("diaphragm_drifts",)),
    ("فحص P-Delta", ("story_forces",)),
    (
        "جداول إضافية للمراجعة",
        ("diaphragm_displacements", "story_mass", "joint_displacements"),
    ),
)
optional_table_labels = {
    key: label for key, label, _aliases in OPTIONAL_SHEET_NAMES
}
optional_tables = {}
optional_table_issues = {}
st.sidebar.markdown("---")
st.sidebar.header("📎 رفع الجداول الاختيارية")
st.sidebar.caption(
    "ارفع ملف Excel منفصلاً لكل جدول، بجانب الفحص المرتبط به. "
    "تُقرأ الورقة المطلوبة فقط من كل ملف."
)
for group_label, table_keys in optional_upload_groups:
    with st.sidebar.expander(group_label):
        for table_key in table_keys:
            table_label = optional_table_labels[table_key]
            accepted_sheets = next(
                aliases
                for key, _label, aliases in OPTIONAL_SHEET_NAMES
                if key == table_key
            )
            st.caption(
                f"الورقة المطلوبة: `{'` أو `'.join(accepted_sheets)}`"
            )
            optional_file = st.file_uploader(
                f"ملف {table_label} (اختياري)",
                type=["xlsx"],
                key=f"optional_upload_{table_key}",
            )
            if optional_file is None:
                st.caption("لم يُرفع هذا الجدول؛ سيبقى الفحص المعتمد عليه غير محسوب.")
                continue
            try:
                _file_hash, optional_table = get_cached_optional_table(
                    st.session_state, optional_file.getvalue(), table_key
                )
                optional_tables[table_key] = optional_table
            except WorkbookValidationError as exc:
                optional_table_issues[table_key] = str(exc)
                st.error(str(exc))

uploaded_file = st.file_uploader(
    "ارفع ملف نتائج ETABS الخام باسم results.xlsx", type=["xlsx"]
)
with st.expander("أسماء الأوراق المطلوبة لإجراء الحسابات"):
    st.markdown(
        "يجب أن يحتوي الملف الأساسي على الأوراق التالية (تُقبل كتابة الاسم مع اختلاف المسافات أو الشرطة السفلية). "
        "أي أوراق أخرى داخله تُهمل ولا تُقرأ؛ ارفع الجداول الاختيارية منفصلة من الشريط الجانبي."
    )
    for _key, aliases in REQUIRED_SHEET_NAMES:
        st.markdown(f"- `{ ' أو '.join(aliases) }`")
    st.markdown(
        "**الجداول الاختيارية:** تُرفع كل ورقة في ملف مستقل من المجموعة المرتبطة بها في الشريط الجانبي. "
        "غياب أي ملف لا يمنع التحليل الأساسي، لكنه يترك الفحص المعتمد على ذلك الجدول غير محسوب."
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
        tables = {**workbook.tables, **optional_tables}
        report_findings = []

        st.header("1. مراكز الكتلة والصلابة")
        with st.expander("معادلات اللامركزية وحدّ المقارنة", expanded=False):
            st.latex(
                r"e_x = |X_{CR}-X_{CM}|,\quad "
                r"e_y = |Y_{CR}-Y_{CM}|"
            )
            st.latex(
                r"r_x = 100\,\frac{e_x}{D_x},\quad "
                r"r_y = 100\,\frac{e_y}{D_y}"
            )
            st.markdown(
                "`CM` مركز الكتلة، و`CR` مركز الصلابة، و`Dx` و`Dy` البعدان "
                "المدخلان بالملم. يجب أن تكون إحداثيات CM/CR والأبعاد بوحدة "
                "طول واحدة. يقارن التطبيق كل نسبة بحده الداخلي 15% "
                "(نجاح عند `r ≤ 15%`)؛ هذا حد التطبيق ولا يُنسب هنا إلى مادة "
                "كودية محددة. لا تُحسب النسبة إذا غابت الإحداثيات أو كان أحد "
                "البعدين غير موجب."
            )
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
        with st.expander("معادلات القص القاعدي ومعامل التحجيم", expanded=False):
            st.latex(
                r"V_{s,\mathrm{req}} = V_{EQ,\max}\,q,\quad "
                r"\mathrm{Check}: V_{RS,\max} \ge V_{s,\mathrm{req}}"
            )
            st.latex(
                r"S_{\mathrm{recommended}} = "
                r"\frac{V_{s,\mathrm{req}}}{V_{RS,\max}}\,\frac{I_e}{R}"
            )
            st.markdown(
                "`V_EQ,max` هو أكبر مقدار مطلق لـ `FX` في حالات `EQX` أو "
                "`FY` في `EQY`، و`V_RS,max` أكبر مقدار مطلق للقوة المقابلة "
                "في حالات `RSX`/`RSY`. `q` النسبة التي يحددها المستخدم، و`Ie` "
                "عامل الأهمية، و`R` معامل تعديل الاستجابة؛ القوى بوحدة ETABS "
                "نفسها، والنسب والمعاملات بلا وحدة. النسبة و`Ie/R` مدخلان "
                "للحساب وليسا بحد ذاتهما إثباتاً لمتطلب كودي؛ معامل التحجيم "
                "المعروض توصية حسابية فقط، ولا يُحسب إذا كان القص الديناميكي "
                "صفراً."
            )
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
        with st.expander("معادلة حد القالب وطريقة المقارنة", expanded=False):
            st.latex(
                r"d = |\mathrm{Drift}|,\quad "
                r"d_{\mathrm{lim}} = "
                r"\begin{cases}"
                r"\dfrac{0.02}{0.7R}, & T>0.7\,\mathrm{s}\\"
                r"\dfrac{0.025}{0.7R}, & T\le 0.7\,\mathrm{s}"
                r"\end{cases}"
            )
            st.markdown(
                "`T` الفترة المدخلة لكل اتجاه بالثواني، و`R` معامل تعديل "
                "الاستجابة، و`Drift` قيمة ETABS؛ يقارن التطبيق مقدار الانجراف "
                "بلا وحدة مع `d_lim` (نجاح إذا `d ≤ d_lim`). هذه قاعدة القالب "
                "المطبقة كما هي وليست حداً كودياً موثقاً أو شهادة مطابقة. "
                "يبقى الفحص غير محسوب إذا لم تتوفر قيمة انجراف رقمية أو فترة "
                "موجبة للاتجاه."
            )
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
        with st.expander("معادلة انجراف التصميم وحدّه حسب فئة المخاطر", expanded=False):
            st.latex(
                r"d_{\mathrm{design}} = |\mathrm{Drift}_{ETABS}|\,"
                r"\frac{C_d}{I_e},\quad "
                r"\mathrm{Check}: d_{\mathrm{design}}\le d_{\mathrm{lim}}"
            )
            st.markdown(
                "**حد «المنشآت الخرسانية الأخرى»:** فئة المخاطر I أو II: "
                "`0.020`؛ III: `0.015`؛ IV: `0.010`."
            )
            st.markdown(
                "`Drift_ETABS` نسبة مرنة بلا وحدة، و`Cd` معامل تكبير "
                "الانحراف، و`Ie` معامل الأهمية. المرجع المذكور في التطبيق: "
                "الفصل 3، §3-9-7، المعادلة (3-9-14)، وجدول حدود الانجراف "
                "(الصفحة المطبوعة 39/3). فئة المخاطر يختارها المستخدم، ويجب "
                "تأكيد انطباق صف المنشآت الخرسانية الأخرى؛ من دون الفئة أو "
                "التأكيد لا يُحسب حد/حكم، وكذلك لا يُحسب عند غياب انجراف رقمي."
            )
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
        with st.expander("معادلة المشاركة التراكمية وحد المقارنة", expanded=False):
            st.latex(
                r"m_x=\max_i(SumUX_i),\quad m_y=\max_i(SumUY_i)"
            )
            st.latex(
                r"M_{\mathrm{cum},d} = "
                r"\begin{cases}"
                r"100\,m_d, & 0\le m_d\le1\\"
                r"m_d, & 1<m_d\le100"
                r"\end{cases},\quad d\in\{x,y\}"
            )
            st.latex(
                r"\mathrm{Check\ per\ case,\ direction}: "
                r"M_{\mathrm{cum}}\ge 90\%"
            )
            st.markdown(
                "`SumUX` و`SumUY` مشاركتا الكتلة التراكميتان من جدول ETABS، "
                "ويؤخذ أكبر مقدار صالح لكل اتجاه ولكل حالة على حدة. إذا كانت "
                "القيمة ضمن `0..1` يحولها التطبيق إلى نسبة مئوية بضربها في "
                "100؛ وإذا كانت أكبر من 1 وحتى 100 فتُعامل كنسبة مئوية كما "
                "وردت. حد 90% مقدم في التطبيق وليس نسبة موثقة من نص الكود "
                "العراقي. القيم خارج النطاقين أو غير الرقمية تبقى غير محسوبة."
            )
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
        with st.expander("معادلات اللامركزية وفحص الالتواء", expanded=False):
            st.markdown("**اللامركزية الهندسية واللامركزية العرضية:**")
            st.latex(
                r"e_{X,\mathrm{signed}} = Y_{CM}-Y_{CR},\quad "
                r"e_{Y,\mathrm{signed}} = X_{CM}-X_{CR}"
            )
            st.latex(
                r"e_{X,\mathrm{acc}} = D_y\,p/100,\quad "
                r"e_{Y,\mathrm{acc}} = D_x\,p/100,\quad "
                r"e_{\mathrm{total}} = e_{\mathrm{signed}}\pm e_{\mathrm{acc}}"
            )
            st.markdown(
                "`p` النسبة العرضية التي يحددها المستخدم، و`Dx`/`Dy` "
                "الأبعاد بالملم؛ تُستخدم قيمة البعد المتعامد مع اتجاه التحليل. "
                "تعرض الإزاحات المحسوبة للاتجاهين ولا يصدر هذا الجزء نجاح/فشل "
                "للامركزية. لا تحسب الإزاحة العرضية قبل اختيار نسبة صالحة، "
                "ولا تحسب اللامركزية الهندسية عند نقص إحداثيات CM/CR."
            )
            st.markdown("**نسبة انجراف الحجاب الحاجز والتصنيف المرشح:**")
            st.latex(
                r"\rho = \left|\frac{\delta_{\max}}{\delta_{\mathrm{avg}}}\right|"
            )
            st.markdown(
                "إذا تعذر اشتقاق النسبة من قيمتي أقصى/متوسط الانجراف (لنقص "
                "قيمة رقمية أو لأن المتوسط صفر)، يستخدم التطبيق نسبة ETABS "
                "المقدمة مباشرة؛ وإذا لم تتوفر فتبقى النسبة غير محسوبة. "
                "القيمتان لهما وحدة واحدة فتكون النسبة بلا وحدة. حسب "
                "المقارنة المرشحة من ASCE 7-16: "
                "`ρ > 1.4` نوع 1b؛ `1.2 < ρ ≤ 1.4` نوع 1a؛ و`ρ ≤ 1.2` "
                "لا يتجاوز العتبتين. هذه العتبات مرشحة ولا تُنسب إلى الكود "
                "العراقي دون تحقق نصي."
            )
            st.markdown("**معامل التكبير العرضي `Ax`:**")
            st.latex(
                r"A_x = \min\left[\left(\frac{\rho}{1.2}\right)^2,\,3.0\right]"
            )
            st.markdown(
                "المرجع المذكور لـ`Ax`: الكود العراقي 2019، الفصل 3، "
                "§3-9-5، المعادلة (3-9-12). لا يُحسب إلا عند تجاوز `ρ` "
                "عتبة 1.2 وتأكيد إدراج اللامركزية العرضية في حالتي الإشارة "
                "واختيار SDC وتأكيد انطباق شروط `Ax`. التطبيق لا يتحقق من "
                "بناء حالات ETABS أو تنفيذ `Ax` فعلياً؛ القيم والتصنيف "
                "معلوماتيان وليسا شهادة مطابقة."
            )
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
            issue = optional_table_issues.get("diaphragm_drifts")
            if issue:
                st.warning(
                    "تعذر استخدام جدول Diaphragm Max Over Avg Drifts: " + issue
                )
            else:
                st.info(
                    "لم يُرفع ملف Diaphragm Max Over Avg Drifts؛ لا يمكن تقييم "
                    "نسبة الالتواء العرضي أو Ax."
                )
        st.caption(
            f"حالة حساب اللامركزية الهندسية: {status_label(eccentricity_status)}. "
            "هذه قيم هندسية/استشارية وليست حكماً مستقلاً بالمطابقة."
        )
        st.markdown("---")

        st.header("5. أقصى/متوسط الإزاحة")
        with st.expander("نطاق العرض وما لا يقيّمه التطبيق", expanded=False):
            st.markdown(
                "يعرض هذا القسم قيم ETABS الخام دون اشتقاق نسبة أو تطبيق "
                "معادلة قبول. لا يوجد في التطبيق حد قبول أو حكم نجاح/فشل "
                "لهذه القيم؛ يلزم تفسيرها وفق متطلبات التحليل والتصميم "
                "المعتمدة للمشروع."
            )
        st.caption("تُعرض قيم ETABS الخام فقط؛ حد قبول هذه النسبة غير محدد في القالب.")
        st.dataframe(tables["displacements"], width="stretch")
        status_dict["Displacements"] = "NOT_EVALUATED"
        st.markdown("---")

        st.header("6. صلابة الطوابق")
        with st.expander("نطاق العرض وما لا يقيّمه التطبيق", expanded=False):
            st.markdown(
                "يعرض هذا القسم قيم الجساءة الخام من ETABS فقط. لا يحسب "
                "التطبيق معادلة قبول أو حدّاً للطابق اللين/الضعيف، ولا يصدر "
                "حكم نجاح/فشل من هذه القيم وحدها."
            )
        st.caption("تُعرض قيم الجساءة الخام؛ لا يُستنتج فحص الطابق اللين من قيم الجساءة وحدها.")
        st.dataframe(tables["story_stiffness"], width="stretch")
        status_dict["Story Stiffness"] = "NOT_EVALUATED"
        st.markdown("---")

        st.header("7. فحص P-Delta")
        with st.expander("معادلة معامل الاستقرار وحدود القرار", expanded=False):
            st.latex(
                r"\mathrm{Drift}_{\mathrm{design}} = "
                r"\frac{C_d}{I_e}\,|\mathrm{Drift}_{\mathrm{elastic}}|"
            )
            st.latex(
                r"\theta = \frac{P\,|\mathrm{Drift}|}{|V|\,I_e},\quad "
                r"\theta_{\max} = "
                r"\min\left(0.25,\,\frac{0.5}{\beta C_d}\right)"
            )
            st.latex(
                r"\theta\le0.1,\quad "
                r"0.1<\theta\le\theta_{\max},\quad "
                r"\theta>\theta_{\max}"
            )
            st.markdown(
                "تفسير الحدود، بالترتيب: يمكن إهمال P-Delta إذا "
                "`θ ≤ 0.1`؛ يجب إدراجه إذا `0.1 < θ ≤ θmax`؛ وتجاوز "
                "`θmax` يستلزم استكمال التحليل/المراجعة وفق الكود."
            )
            st.markdown(
                "`P` الحمل المحوري الموجب ضغطاً، و`V` قص الطابق المطلق عند "
                "`Bottom`، و`Drift` أكبر انجراف مرن مطلق مطابق للطابق/الاتجاه "
                "(نسبة بلا وحدة)، و`Ie` معامل الأهمية، و`Cd` معامل تكبير "
                "الانحراف، و`β=1.0` في هذا التطبيق. يستخدم `P` و`V` وحدة "
                "القوة نفسها فتكون `θ` بلا وحدة. المرجع: الكود الزلزالي "
                "العراقي 2019، الفصل 3، §3-9/7/1، المعادلتان (3-9/12) "
                "و(3-9/16). صيغة `θ` المختزلة أعلاه تستخدم نسبة الانجراف "
                "مباشرة؛ لا يعيد التطبيق ضربها في ارتفاع الطابق."
            )
            st.markdown(
                "هذا فحص فرز أولي وليس بديلاً عن تحليل ETABS من المرتبة "
                "الثانية أو حكماً عاماً بالسلامة. لا يُحسب حتى اختيار "
                "تركيبة Combination مشتركة صالحة وتأكيد معاملات الأحمال "
                "وانطباق `Cd`/`Ie` وأن `Drift` مرن غير مكبّر مسبقاً؛ كما "
                "يتطلب بيانات متطابقة، قصاً غير صفري، و`P > 0`. النقص أو "
                "عدم تحقق هذه الشروط يبقي الحالة غير محسوبة، لا ناجحة."
            )
        pdelta_options = (
            pdelta_combination_options(
                tables.get("story_forces", pd.DataFrame()),
                tables.get("story_drifts", pd.DataFrame()),
                tables.get("base_reactions", pd.DataFrame()),
            )
            if "story_forces" in tables
            else []
        )
        pdelta_source_issue = optional_table_issues.get("story_forces")
        if pdelta_source_issue is None and "story_forces" not in tables:
            pdelta_source_issue = "لم يُرفع ملف جدول Story Forces."
        selected_pdelta_combinations = st.multiselect(
            "تركيبات التصميم المعتمدة لفحص P-Delta (يمكن اختيار عدة تركيبات)",
            options=pdelta_options,
            key="pdelta_selected_combinations",
            help=(
                "تُعرض فقط حالات Combination المشتركة التي لا تحمل مخرجات "
                "Max/Min أو Envelope أو Step Number. لا يستطيع التطبيق استنتاج معاملات "
                "التركيبة من اسمها. ويُشترط وجود قص قاعدي جانبي أعلى من 1e-6 "
                "من |FZ|؛ هذه سماحية عددية وليست حداً للكود."
            ),
        )
        st.caption(
            "تُستخدم قيم Story Forces عند Bottom والانجراف المطلق الأقصى لكل "
            "طابق/اتجاه. تُرتب النتائج بحسب منسوب Z من الأعلى إلى الأسفل؛ "
            "ولا يُستخدم ارتفاع الطابق لأن Drift المصدر نسبة بلا وحدة. تُقبل "
            "قيمة P الموجبة ضغطاً؛ القيم السالبة أو الصفرية لا تُحسب. تُراجع "
            "FX/FY في Base Reactions لاستبعاد التركيبات الجاذبية والاتجاهات "
            "دون حمل جانبي مؤثر؛ السماحية العددية هي 1e-6 × |FZ| وليست حد كود."
        )
        pdelta_combinations_confirmed = st.checkbox(
            "أؤكد أن معاملات التركيبات المختارة تحقق متطلبات الكود العراقي"
        )
        pdelta_coefficients_confirmed = st.checkbox(
            f"أؤكد انطباق Cd = {deflection_cd:g} وIe = {importance_factor:g} على هذه التركيبات"
        )
        pdelta_drift_confirmed = st.checkbox(
            "أؤكد أن قيم Story Drifts.Drift انجراف مرن غير مضروب مسبقاً في Cd/Ie"
        )
        st.markdown(
            "**المعادلة:** θ = P × |Drift| / (|V| × Ie)، بعد اختزال معادلة "
            "الانجراف التصميمي في الكود العراقي. "
            "**الحد:** θmax = min(0.25, 0.5/(β × Cd)) مع β = 1.0 تحفظيًا. "
            "إذا تجاوز θ قيمة θmax فهذا تجاوز لحد الاستقرار؛ وإلا إذا "
            "θ ≤ 0.1 يمكن إهمال الأثر، وإذا 0.1 < θ ≤ θmax وجب إدراجه."
        )
        pdelta = calculate_pdelta_screening(
            tables.get("story_forces", pd.DataFrame()),
            tables["story_drifts"],
            tables.get("base_reactions", pd.DataFrame()),
            selected_pdelta_combinations,
            deflection_cd,
            importance_factor,
            combinations_confirmed=pdelta_combinations_confirmed,
            coefficients_confirmed=pdelta_coefficients_confirmed,
            drift_confirmed=pdelta_drift_confirmed,
            source_issue=pdelta_source_issue,
        )
        status_dict["P-Delta Check"] = pdelta["status"]
        if pdelta["status"] in {
            "PDELTA_LIMIT_EXCEEDED",
            "PDELTA_REQUIRED",
            "INCOMPLETE",
            "NOT_EVALUATED",
        }:
            st.warning(pdelta["message"])
        else:
            st.success(pdelta["message"])
        governing = pdelta["governing_results"]
        if not governing.empty:
            display = governing.rename(
                columns={
                    "story": "الطابق",
                    "story_z": "منسوب Z (ملم)",
                    "direction": "الاتجاه",
                    "base_shear": "|قص القاعدة|",
                    "output_case": "التركيبة الحاكمة",
                    "p": "P عند Bottom",
                    "story_shear": "|V| عند Bottom",
                    "elastic_drift_ratio": "أقصى |Drift| مرن",
                    "theta": "θ",
                    "theta_max": "θmax",
                    "classification": "قرار الكود",
                    "reason": "سبب عدم الحساب",
                }
            )
            display["قرار الكود"] = display["قرار الكود"].map(
                {
                    "NEGLIGIBLE": "يمكن الإهمال (θ ≤ 0.1)",
                    "INCLUDE_PDELTA": "يجب إدراج P-Delta",
                    "LIMIT_EXCEEDED": "تجاوز θmax",
                    "NOT_EVALUATED": "غير محسوب",
                }
            )
            st.dataframe(
                display[
                    [
                        "الطابق",
                        "منسوب Z (ملم)",
                        "الاتجاه",
                        "التركيبة الحاكمة",
                        "|قص القاعدة|",
                        "P عند Bottom",
                        "|V| عند Bottom",
                        "أقصى |Drift| مرن",
                        "θ",
                        "θmax",
                        "قرار الكود",
                        "سبب عدم الحساب",
                    ]
                ].round(6),
                width="stretch",
            )
            for row in governing.itertuples(index=False):
                if pd.notna(row.theta):
                    report_findings.append(
                        f"P-Delta — {row.story}، الاتجاه {row.direction}، "
                        f"التركيبة الحاكمة {row.output_case}: θ = {row.theta:.5f} "
                        f"مقابل θmax = {row.theta_max:.5f} "
                        f"({row.classification})."
                    )
            with st.expander("تفاصيل جميع التركيبات المختارة"):
                st.dataframe(pdelta["results"], width="stretch")
        if pdelta["required_inputs"]:
            st.write("متطلبات/تنبيهات الفحص:")
            for item in pdelta["required_inputs"]:
                st.markdown(f"- {item}")
        st.markdown("---")
        st.header("8. جداول ETABS الإضافية")
        with st.expander("نطاق عرض الجداول الإضافية", expanded=False):
            st.markdown(
                "تُعرض البيانات الاختيارية كما وردت في الملف، دون معادلات "
                "مشتقة أو حدود قبول أو تصنيف تلقائي. غياب جدول لا يُحوّل "
                "أي قيمة إلى صفر ولا يعني اجتياز فحص."
            )
        st.caption(
            "تعرض هذه الجداول من ملفاتها المنفصلة كما وردت. وجودها يتيح مراجعة بيانات إضافية، "
            "ولا ينشئ فحص نجاح/فشل تلقائياً."
        )
        for key, label, _aliases in OPTIONAL_SHEET_NAMES:
            if key in tables:
                with st.expander(f"{label} — متوفر"):
                    st.dataframe(tables[key], width="stretch")
            else:
                issue = optional_table_issues.get(key)
                if issue:
                    st.warning(f"تعذر استخدام الجدول الاختياري `{label}`: {issue}")
                    report_findings.append(f"الجدول الاختياري {label}: تعذر قراءته.")
                else:
                    st.info(f"لم يُرفع ملف الجدول الاختياري `{label}`.")
                    report_findings.append(f"الجدول الاختياري {label}: لم يُرفع.")

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
