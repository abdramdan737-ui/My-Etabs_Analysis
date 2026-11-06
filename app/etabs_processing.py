import io
import hashlib
import re
from dataclasses import dataclass
from typing import BinaryIO, MutableMapping

import numpy as np
import pandas as pd


class WorkbookValidationError(ValueError):
    pass


@dataclass
class EtabsWorkbook:
    tables: dict[str, pd.DataFrame]
    pdelta_source_issue: str | None = None
    missing_optional_sheets: tuple[str, ...] = ()


REQUIRED_SHEETS = {
    "centers": ("Centers Of Mass And Rigidity",),
    "base_reactions": ("Base Reactions",),
    "story_drifts": ("Story Drifts",),
    "modal_mass": ("Modal Participating Mass Ratios",),
    "story_stiffness": ("Story Stiffness",),
    "displacements": ("Story Max Over Avg Disp", "Story Max Over Avg Displacements"),
}

REQUIRED_SHEET_NAMES = tuple(
    (key, aliases) for key, aliases in REQUIRED_SHEETS.items()
)

OPTIONAL_TABLES = {
    "story_forces": {
        "label": "Story Forces",
        "aliases": ("Story Forces",),
        "pdelta_source": True,
    },
    "diaphragm_displacements": {
        "label": "Diaphragm CM Displacements",
        "aliases": (
            "Diaphragm CM Displacements",
            "Diaphragm Center Of Mass Displacements",
        ),
        "pdelta_source": True,
    },
    "story_mass": {
        "label": "Story Masses",
        "aliases": ("Story Masses", "Story Mass", "Story Mass Summary"),
        "pdelta_source": False,
    },
    "joint_displacements": {
        "label": "Joint Displacements",
        "aliases": ("Joint Displacements",),
        "pdelta_source": False,
    },
}

SHEET_ALIASES = {
    **{key: names for key, names in REQUIRED_SHEETS.items()},
    **{
        key: table["aliases"] for key, table in OPTIONAL_TABLES.items()
    },
}

OPTIONAL_SHEET_NAMES = tuple(
    (key, table["label"], table["aliases"])
    for key, table in OPTIONAL_TABLES.items()
)

OPTIONAL_PDELTA_SHEETS = tuple(
    table["label"]
    for table in OPTIONAL_TABLES.values()
    if table["pdelta_source"]
)

MAX_CACHED_WORKBOOK_VERSIONS = 4
CACHE_SESSION_KEY = "_etabs_workbook_versions"

COLUMNS = {
    "centers": {
        "story": ("Story",),
        "diaphragm": ("Diaphragm",),
        "xcm": ("XCM", "X CM"),
        "ycm": ("YCM", "Y CM"),
        "xcr": ("XCR", "X CR"),
        "ycr": ("YCR", "Y CR"),
    },
    "base_reactions": {
        "output_case": ("Output Case", "Case"),
        "fx": ("FX", "F_X"),
        "fy": ("FY", "F_Y"),
    },
    "story_drifts": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "direction": ("Direction", "Dirction"),
        "drift": ("Drift",),
    },
    "modal_mass": {
        "period": ("Period",),
        "ux": ("UX",),
        "uy": ("UY",),
        "sumux": ("SumUX", "Sum UX"),
        "sumuy": ("SumUY", "Sum UY"),
    },
    "story_stiffness": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "stiff_x": ("Stiff X", "StiffX"),
        "stiff_y": ("Stiff Y", "StiffY"),
    },
    "displacements": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "direction": ("Direction", "Dirction"),
        "maximum": ("Maximum", "Max"),
        "average": ("Average", "Avg"),
        "ratio": ("Ratio",),
    },
    "story_forces": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "location": ("Location",),
        "p": ("P",),
        "vx": ("VX", "V_X"),
        "vy": ("VY", "V_Y"),
    },
    "diaphragm_displacements": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "ux": ("UX",),
        "uy": ("UY",),
        "z": ("Z",),
    },
}


def _normalized_name(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def _find_sheet(sheet_names: list[str], aliases: tuple[str, ...]) -> str | None:
    normalized = {_normalized_name(name): name for name in sheet_names}
    for alias in aliases:
        found = normalized.get(_normalized_name(alias))
        if found is not None:
            return found
    return None


def missing_required_sheets(sheet_names: list[str]) -> list[tuple[str, tuple[str, ...]]]:
    return [
        (key, aliases)
        for key, aliases in REQUIRED_SHEET_NAMES
        if _find_sheet(sheet_names, aliases) is None
    ]


def missing_optional_sheets(
    sheet_names: list[str],
) -> list[tuple[str, str, tuple[str, ...]]]:
    return [
        (key, label, aliases)
        for key, label, aliases in OPTIONAL_SHEET_NAMES
        if _find_sheet(sheet_names, aliases) is None
    ]


def validate_upload_filename(filename: str) -> None:
    if filename.strip().casefold() != "results.xlsx":
        raise WorkbookValidationError(
            "يجب أن يكون اسم الملف المرفوع results.xlsx بالضبط."
        )


def get_cached_workbook(
    session_state: MutableMapping[str, object], content: bytes
) -> tuple[str, EtabsWorkbook]:
    content_hash = hashlib.sha256(content).hexdigest()
    versions = session_state.setdefault(CACHE_SESSION_KEY, {})
    if not isinstance(versions, dict):
        versions = {}
        session_state[CACHE_SESSION_KEY] = versions

    cached = versions.get(content_hash)
    if isinstance(cached, EtabsWorkbook):
        return content_hash, cached

    workbook = load_etabs_workbook(content)
    versions[content_hash] = workbook
    while len(versions) > MAX_CACHED_WORKBOOK_VERSIONS:
        oldest_hash = next(iter(versions))
        del versions[oldest_hash]
    session_state[CACHE_SESSION_KEY] = versions
    return content_hash, workbook


def _find_column(columns: list[object], aliases: tuple[str, ...]) -> object | None:
    for alias in aliases:
        exact = [column for column in columns if str(column).strip().casefold() == alias.casefold()]
        if len(exact) == 1:
            return exact[0]
        if len(exact) > 1:
            raise WorkbookValidationError(f"عمود مكرر أو ملتبس: {alias}")

    normalized_aliases = {_normalized_name(alias) for alias in aliases}
    matches = [
        column for column in columns if _normalized_name(column) in normalized_aliases
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise WorkbookValidationError(
            "أعمدة متشابهة لا يمكن تمييزها: " + ", ".join(map(str, matches))
        )
    return None


def _canonicalize_table(
    frame: pd.DataFrame, table_name: str, required_columns: dict[str, tuple[str, ...]]
) -> pd.DataFrame:
    frame = frame.dropna(how="all").copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    renames = {}
    missing = []
    for canonical, aliases in required_columns.items():
        column = _find_column(list(frame.columns), aliases)
        if column is None:
            missing.append(f"{canonical} ({' / '.join(aliases)})")
        else:
            renames[column] = canonical
    if missing:
        raise WorkbookValidationError(
            f"الأعمدة المطلوبة غير موجودة في ورقة {table_name}: {', '.join(missing)}"
        )
    return frame.rename(columns=renames).reset_index(drop=True)


def _read_bytes(source: bytes | bytearray | BinaryIO) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "getvalue"):
        return source.getvalue()
    if hasattr(source, "read"):
        position = source.tell() if hasattr(source, "tell") else None
        data = source.read()
        if position is not None and hasattr(source, "seek"):
            source.seek(position)
        return data
    raise WorkbookValidationError("تعذر قراءة الملف المرفوع بصيغة Excel.")


def load_etabs_workbook(source: bytes | bytearray | BinaryIO) -> EtabsWorkbook:
    content = _read_bytes(source)
    try:
        excel = pd.ExcelFile(io.BytesIO(content))
    except Exception as exc:
        raise WorkbookValidationError(
            "تعذر فتح الملف. يرجى رفع مصنف ETABS بصيغة .xlsx صالح."
        ) from exc

    sheet_names = excel.sheet_names
    selected_sheets = {
        key: _find_sheet(sheet_names, aliases)
        for key, aliases in SHEET_ALIASES.items()
    }
    missing = missing_required_sheets(sheet_names)
    if missing:
        missing_labels = [
            " / ".join(aliases) for _key, aliases in missing
        ]
        raise WorkbookValidationError(
            "الأوراق المطلوبة غير موجودة: " + ", ".join(missing_labels)
        )

    tables = {}
    missing_optional = missing_optional_sheets(sheet_names)
    optional_issues = [
        f"الورقة الاختيارية {label} غير موجودة"
        for key, label, _aliases in missing_optional
        if OPTIONAL_TABLES[key]["pdelta_source"]
    ]
    for key, sheet_name in selected_sheets.items():
        if sheet_name is None:
            continue
        try:
            raw_frame = pd.read_excel(
                excel, sheet_name=sheet_name, skiprows=[0, 2], header=0
            )
            tables[key] = _canonicalize_table(
                raw_frame, sheet_name, COLUMNS.get(key, {})
            )
        except WorkbookValidationError as exc:
            if key in REQUIRED_SHEETS:
                raise
            if OPTIONAL_TABLES[key]["pdelta_source"]:
                optional_issues.append(f"{sheet_name}: {exc}")
        except Exception as exc:
            if key in REQUIRED_SHEETS:
                raise WorkbookValidationError(
                    f"تعذر قراءة ورقة {sheet_name}: {exc}"
                ) from exc
            if OPTIONAL_TABLES[key]["pdelta_source"]:
                optional_issues.append(f"تعذر قراءة ورقة {sheet_name}: {exc}")

    return EtabsWorkbook(
        tables=tables,
        pdelta_source_issue="؛ ".join(optional_issues) if optional_issues else None,
        missing_optional_sheets=tuple(
            label for _key, label, _aliases in missing_optional
        ),
    )


def _positive_number(value: float, name: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise WorkbookValidationError(f"قيمة {name} يجب أن تكون رقمية.") from exc
    if not np.isfinite(number) or number <= 0:
        raise WorkbookValidationError(f"قيمة {name} يجب أن تكون أكبر من صفر.")
    return number


def calculate_eccentricity(
    table: pd.DataFrame, dx: float, dy: float, limit_percent: float = 15.0
) -> tuple[pd.DataFrame, str]:
    dx = _positive_number(dx, "Dx")
    dy = _positive_number(dy, "Dy")
    result = table[["story", "diaphragm"]].copy()
    for column in ("xcm", "ycm", "xcr", "ycr"):
        result[column] = pd.to_numeric(table[column], errors="coerce")

    result["ex"] = (result["xcr"] - result["xcm"]).abs()
    result["ey"] = (result["ycr"] - result["ycm"]).abs()
    result["ex_dx_percent"] = result["ex"] / dx * 100
    result["ey_dy_percent"] = result["ey"] / dy * 100
    valid = result[["ex_dx_percent", "ey_dy_percent"]].notna().all(axis=1)
    result["x_result"] = np.where(
        valid,
        np.where(result["ex_dx_percent"] <= limit_percent, "PASS", "FAIL"),
        "NOT_EVALUATED",
    )
    result["y_result"] = np.where(
        valid,
        np.where(result["ey_dy_percent"] <= limit_percent, "PASS", "FAIL"),
        "NOT_EVALUATED",
    )
    if not valid.all():
        status = "INCOMPLETE"
    elif (result[["x_result", "y_result"]] == "FAIL").any().any():
        status = "FAIL"
    else:
        status = "PASS"
    return result, status


def _maximum_absolute(
    rows: pd.DataFrame, column: str, case_description: str
) -> float:
    values = pd.to_numeric(rows[column], errors="coerce").dropna().abs()
    if values.empty:
        raise WorkbookValidationError(
            f"لم توجد قيمة رقمية صالحة لـ {case_description} في Base Reactions."
        )
    return float(values.max())


def calculate_base_reactions(
    table: pd.DataFrame,
    target_ratio: float,
    importance_factor: float,
    response_modification_factor: float,
) -> tuple[pd.DataFrame, str]:
    target_ratio = _positive_number(target_ratio, "نسبة القص المستهدفة")
    importance_factor = _positive_number(importance_factor, "Ie")
    response_modification_factor = _positive_number(
        response_modification_factor, "R"
    )
    if target_ratio > 1:
        raise WorkbookValidationError("نسبة القص المستهدفة يجب ألا تتجاوز 100%.")

    cases = table["output_case"].astype(str).str.upper().str.replace(
        r"[^A-Z0-9]", "", regex=True
    )
    rows = []
    for direction, static_case, dynamic_pattern, force_column in (
        ("X", "EQX", "RSX", "fx"),
        ("Y", "EQY", "RSY", "fy"),
    ):
        static_rows = table[cases == static_case]
        dynamic_rows = table[cases.str.contains(dynamic_pattern, regex=False)]
        if static_rows.empty:
            raise WorkbookValidationError(
                f"لم توجد حالة استاتيكية {static_case} في Base Reactions."
            )
        if dynamic_rows.empty:
            raise WorkbookValidationError(
                f"لم توجد حالات طيفية {dynamic_pattern} في Base Reactions."
            )

        static_shear = _maximum_absolute(
            static_rows, force_column, static_case
        )
        dynamic_shear = _maximum_absolute(
            dynamic_rows, force_column, dynamic_pattern
        )
        required_shear = static_shear * target_ratio
        rows.append(
            {
                "direction": direction,
                "static_shear": static_shear,
                "dynamic_shear": dynamic_shear,
                "target_ratio_percent": target_ratio * 100,
                "required_shear": required_shear,
                "check": "PASS" if dynamic_shear >= required_shear else "FAIL",
                "recommended_scale_factor": (
                    (required_shear / dynamic_shear)
                    * (importance_factor / response_modification_factor)
                    if dynamic_shear > 0
                    else np.nan
                ),
            }
        )

    result = pd.DataFrame(rows)
    status = "PASS" if (result["check"] == "PASS").all() else "FAIL"
    return result, status


def calculate_story_drifts(
    table: pd.DataFrame,
    response_modification_factor: float,
    tx: float,
    ty: float,
) -> tuple[pd.DataFrame, str]:
    response_modification_factor = _positive_number(
        response_modification_factor, "R"
    )
    periods = {
        "X": _positive_number(tx, "Tx"),
        "Y": _positive_number(ty, "Ty"),
    }
    cases = table["output_case"].astype(str)
    lateral = table[cases.str.contains(r"EQ|RS|WIND|SPEC", case=False, na=False)]
    if lateral.empty:
        raise WorkbookValidationError(
            "لم توجد حالات جانبية (EQ/RS/WIND/SPEC) في Story Drifts."
        )

    result = lateral[["story", "output_case", "direction"]].copy()
    result["direction"] = result["direction"].astype(str).str.strip().str.upper()
    result["drift_ratio"] = pd.to_numeric(lateral["drift"], errors="coerce").abs()
    result["period_seconds"] = result["direction"].map(periods)
    result["limit_ratio"] = np.nan
    for direction, period in periods.items():
        mask = result["direction"] == direction
        limit = (0.02 if period > 0.7 else 0.025) / (
            0.7 * response_modification_factor
        )
        result.loc[mask, "limit_ratio"] = limit

    valid = result["drift_ratio"].notna() & result["limit_ratio"].notna()
    result["check"] = np.where(
        valid,
        np.where(result["drift_ratio"] <= result["limit_ratio"], "PASS", "FAIL"),
        "NOT_EVALUATED",
    )
    if not valid.all():
        status = "INCOMPLETE"
    elif (result["check"] == "FAIL").any():
        status = "FAIL"
    else:
        status = "PASS"
    return result.reset_index(drop=True), status


def calculate_modal_mass(table: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    result = []
    for direction, column in (("X", "sumux"), ("Y", "sumuy")):
        values = pd.to_numeric(table[column], errors="coerce").dropna()
        if values.empty:
            result.append(
                {"direction": direction, "maximum_percent": np.nan, "check": "NOT_EVALUATED"}
            )
            continue
        maximum = float(values.max())
        maximum_percent = maximum if maximum > 1 else maximum * 100
        result.append(
            {
                "direction": direction,
                "maximum_percent": maximum_percent,
                "check": "PASS" if maximum_percent >= 90 else "FAIL",
            }
        )
    result_frame = pd.DataFrame(result)
    if (result_frame["check"] == "NOT_EVALUATED").any():
        status = "INCOMPLETE"
    elif (result_frame["check"] == "FAIL").any():
        status = "FAIL"
    else:
        status = "PASS"
    return result_frame, status


def pdelta_assessment(source_issue: str | None = None) -> dict[str, object]:
    message = (
        "غير محسوب: يحتوي قالب Excel على مراجع خارجية مكسورة، كما تشير معادلات "
        "الحالة إلى خلايا معاملات فارغة؛ لذلك لا يمكن اعتماد نتيجته."
    )
    if source_issue:
        message += f" مصدر بيانات إضافي غير متاح أو غير صالح: {source_issue}."
    return {
        "status": "NOT_EVALUATED",
        "message": message,
        "required_inputs": [
            "حالة/تركيبة الجاذبية الحاكمة لتحديد P لكل طابق.",
            "حالة الزلزال المطابقة لتحديد V والانجراف النسبي بين الطوابق.",
            "ارتفاع كل طابق وربط متسق بين قيم القوة والإزاحة واتجاه التحليل.",
            "معامل الاستقرار وحده الحاكم المعتمد، بما فيه معامل β عند انطباقه.",
        ],
    }
