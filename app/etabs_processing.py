import io
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, MutableMapping

import numpy as np
import pandas as pd


class WorkbookValidationError(ValueError):
    pass


@dataclass
class EtabsWorkbook:
    tables: dict[str, pd.DataFrame]
    optional_issues: dict[str, str] = field(default_factory=dict)


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
PDELTA_BASE_SHEAR_RELATIVE_TOLERANCE = 1e-6

OPTIONAL_TABLES = {
    "story_forces": {
        "label": "Story Forces",
        "aliases": ("Story Forces",),
    },
    "diaphragm_displacements": {
        "label": "Diaphragm CM Displacements",
        "aliases": (
            "Diaphragm CM Displacements",
            "Diaphragm Center Of Mass Displacements",
        ),
    },
    "story_mass": {
        "label": "Story Masses",
        "aliases": ("Story Masses", "Story Mass", "Story Mass Summary"),
    },
    "joint_displacements": {
        "label": "Joint Displacements",
        "aliases": ("Joint Displacements",),
    },
    "diaphragm_drifts": {
        "label": "Diaphragm Max Over Avg Drifts",
        "aliases": ("Diaphragm Max Over Avg Drifts",),
    },
}

OPTIONAL_SHEET_NAMES = tuple(
    (key, table["label"], table["aliases"])
    for key, table in OPTIONAL_TABLES.items()
)

MAX_CACHED_WORKBOOK_VERSIONS = 4
CACHE_SESSION_KEY = "_etabs_workbook_versions"
OPTIONAL_TABLE_CACHE_SESSION_KEY = "_etabs_optional_table_versions"
DEFAULT_RESULTS_PATH = Path(r"E:\etabs_tables\003\results.xlsx")

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
        "case": ("Case",),
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
    "diaphragm_drifts": {
        "story": ("Story",),
        "output_case": ("Output Case", "Case"),
        "item": ("Item",),
        "max_drift": ("Max Drift",),
        "avg_drift": ("Avg Drift",),
        "ratio": ("Ratio",),
    },
}

OPTIONAL_COLUMNS = {
    "base_reactions": {
        "case_type": ("Case Type",),
        "step_type": ("Step Type",),
        "step_number": ("Step Number",),
        "fz": ("FZ", "F_Z"),
    },
    "story_drifts": {
        "case_type": ("Case Type",),
        "step_type": ("Step Type",),
        "step_number": ("Step Number",),
        "z": ("Z",),
    },
    "story_forces": {
        "case_type": ("Case Type",),
        "step_type": ("Step Type",),
        "step_number": ("Step Number",),
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


def read_local_workbook_bytes(path: str | Path) -> bytes:
    workbook_path = Path(path).expanduser()
    if workbook_path.name.strip().casefold() != "results.xlsx":
        raise WorkbookValidationError(
            "يجب أن يكون اسم المصنف المحلي results.xlsx بالضبط."
        )
    if not workbook_path.is_file():
        raise WorkbookValidationError(
            f"لم يُعثر على ملف النتائج المحلي: {workbook_path}. "
            "ضع تصدير ETABS الخام بهذا الاسم في المسار المحدد."
        )
    try:
        return workbook_path.read_bytes()
    except OSError as exc:
        raise WorkbookValidationError(
            f"تعذر قراءة ملف النتائج المحلي {workbook_path}: {exc}"
        ) from exc


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


def get_cached_optional_table(
    session_state: MutableMapping[str, object],
    content: bytes,
    table_key: str,
) -> tuple[str, pd.DataFrame]:
    content_hash = hashlib.sha256(content).hexdigest()
    versions = session_state.setdefault(OPTIONAL_TABLE_CACHE_SESSION_KEY, {})
    if not isinstance(versions, dict):
        versions = {}
        session_state[OPTIONAL_TABLE_CACHE_SESSION_KEY] = versions

    cache_key = (table_key, content_hash)
    cached = versions.get(cache_key)
    if isinstance(cached, pd.DataFrame):
        return content_hash, cached

    table = load_optional_etabs_table(content, table_key)
    versions[cache_key] = table
    max_versions = MAX_CACHED_WORKBOOK_VERSIONS * len(OPTIONAL_TABLES)
    while len(versions) > max_versions:
        oldest_key = next(iter(versions))
        del versions[oldest_key]
    session_state[OPTIONAL_TABLE_CACHE_SESSION_KEY] = versions
    return content_hash, table


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
    frame: pd.DataFrame,
    table_name: str,
    required_columns: dict[str, tuple[str, ...]],
    optional_columns: dict[str, tuple[str, ...]] | None = None,
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
    for canonical, aliases in (optional_columns or {}).items():
        column = _find_column(list(frame.columns), aliases)
        if column is not None:
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
        for key, aliases in REQUIRED_SHEETS.items()
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
    for key, sheet_name in selected_sheets.items():
        try:
            raw_frame = pd.read_excel(
                excel, sheet_name=sheet_name, skiprows=[0, 2], header=0
            )
            tables[key] = _canonicalize_table(
                raw_frame,
                sheet_name,
                COLUMNS.get(key, {}),
                OPTIONAL_COLUMNS.get(key, {}),
            )
        except WorkbookValidationError:
            raise
        except Exception as exc:
            raise WorkbookValidationError(
                f"تعذر قراءة ورقة {sheet_name}: {exc}"
            ) from exc

    optional_issues = {}
    for key, optional_table in OPTIONAL_TABLES.items():
        sheet_name = _find_sheet(sheet_names, optional_table["aliases"])
        if sheet_name is None:
            continue
        try:
            raw_frame = pd.read_excel(
                excel, sheet_name=sheet_name, skiprows=[0, 2], header=0
            )
            tables[key] = _canonicalize_table(
                raw_frame,
                sheet_name,
                COLUMNS.get(key, {}),
                OPTIONAL_COLUMNS.get(key, {}),
            )
        except WorkbookValidationError as exc:
            optional_issues[key] = str(exc)
        except Exception as exc:
            optional_issues[key] = f"تعذر قراءة ورقة {sheet_name}: {exc}"

    return EtabsWorkbook(tables=tables, optional_issues=optional_issues)


def load_optional_etabs_table(
    source: bytes | bytearray | BinaryIO,
    table_key: str,
) -> pd.DataFrame:
    optional_table = OPTIONAL_TABLES.get(table_key)
    if optional_table is None:
        raise WorkbookValidationError(f"جدول ETABS اختياري غير معروف: {table_key}")

    content = _read_bytes(source)
    try:
        excel = pd.ExcelFile(io.BytesIO(content))
    except Exception as exc:
        raise WorkbookValidationError(
            "تعذر فتح الملف الإضافي. يرجى رفع مصنف Excel بصيغة .xlsx صالح."
        ) from exc

    sheet_name = _find_sheet(excel.sheet_names, optional_table["aliases"])
    if sheet_name is None:
        expected_names = " / ".join(optional_table["aliases"])
        raise WorkbookValidationError(
            f"الورقة المطلوبة للجدول الاختياري غير موجودة: {expected_names}."
        )

    try:
        raw_frame = pd.read_excel(
            excel, sheet_name=sheet_name, skiprows=[0, 2], header=0
        )
        return _canonicalize_table(
            raw_frame,
            sheet_name,
            COLUMNS.get(table_key, {}),
            OPTIONAL_COLUMNS.get(table_key, {}),
        )
    except WorkbookValidationError:
        raise
    except Exception as exc:
        raise WorkbookValidationError(
            f"تعذر قراءة ورقة {sheet_name}: {exc}"
        ) from exc


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
    if "case" in table:
        cases = table["case"].fillna("").astype(str).str.strip()
    else:
        cases = pd.Series("Unspecified", index=table.index)
    for case_name, case_rows in table.groupby(cases, sort=False, dropna=False):
        for direction, column in (("X", "sumux"), ("Y", "sumuy")):
            values = pd.to_numeric(case_rows[column], errors="coerce").dropna()
            maximum_percent = np.nan
            check = "NOT_EVALUATED"
            if not values.empty:
                maximum = float(values.max())
                if 0 <= maximum <= 1:
                    maximum_percent = maximum * 100
                    check = "PASS" if maximum_percent >= 90 else "FAIL"
                elif 1 < maximum <= 100:
                    maximum_percent = maximum
                    check = "PASS" if maximum_percent >= 90 else "FAIL"
            result.append(
                {
                    "case": case_name or "Unspecified",
                    "direction": direction,
                    "maximum_percent": maximum_percent,
                    "check": check,
                }
            )
    result_frame = pd.DataFrame(result)
    if result_frame.empty or (result_frame["check"] == "NOT_EVALUATED").any():
        status = "INCOMPLETE" if not result_frame.empty else "NOT_EVALUATED"
    elif (result_frame["check"] == "FAIL").any():
        status = "FAIL"
    else:
        status = "PASS"
    return result_frame, status


IRAQI_2019_DRIFT_LIMITS = {
    "I": 0.020,
    "II": 0.020,
    "III": 0.015,
    "IV": 0.010,
}


def calculate_iraqi_2019_story_drifts(
    table: pd.DataFrame,
    risk_category: str | None,
    applicability_confirmed: bool,
    cd: float,
    ie: float,
) -> tuple[pd.DataFrame, str]:
    cd = _positive_number(cd, "Cd")
    ie = _positive_number(ie, "Ie")
    category = str(risk_category or "").strip().upper()
    limit = IRAQI_2019_DRIFT_LIMITS.get(category)
    cases = table["output_case"].astype(str)
    lateral = table[cases.str.contains(r"EQ|RS|SPEC", case=False, na=False)].copy()
    result = lateral[["story", "output_case", "direction"]].copy()
    result["elastic_drift_ratio"] = pd.to_numeric(
        lateral["drift"], errors="coerce"
    ).abs()
    result["cd_ie_factor"] = cd / ie
    result["design_drift_ratio"] = np.nan
    result["limit_ratio"] = limit if applicability_confirmed else np.nan
    result["check"] = "NOT_EVALUATED"

    if limit is None:
        result["reason"] = "Risk Category I–IV is not selected."
    elif not applicability_confirmed:
        result["reason"] = (
            "Applicability of the 'other concrete structures' limit is unconfirmed."
        )
    else:
        result["design_drift_ratio"] = (
            result["elastic_drift_ratio"] * result["cd_ie_factor"]
        )
        valid = (
            result["elastic_drift_ratio"].notna()
            & np.isfinite(result["design_drift_ratio"])
        )
        result["check"] = np.where(
            valid,
            np.where(result["design_drift_ratio"] <= limit, "PASS", "FAIL"),
            "NOT_EVALUATED",
        )
        result["reason"] = np.where(valid, "", "The ETABS drift value is not numeric.")

    if result.empty:
        status = "NOT_EVALUATED"
    elif (result["check"] == "FAIL").any():
        status = "FAIL"
    elif (result["check"] == "NOT_EVALUATED").any():
        status = "INCOMPLETE" if limit is not None and applicability_confirmed else "NOT_EVALUATED"
    else:
        status = "PASS"
    return result.reset_index(drop=True), status


def calculate_iraqi_2019_torsion_eccentricities(
    table: pd.DataFrame,
    dx: float,
    dy: float,
    accidental_percent: float | None,
) -> tuple[pd.DataFrame, str]:
    dx = _positive_number(dx, "Dx")
    dy = _positive_number(dy, "Dy")
    if accidental_percent is not None:
        accidental_percent = float(accidental_percent)
        if not np.isfinite(accidental_percent) or not 0 < accidental_percent <= 100:
            raise WorkbookValidationError(
                "نسبة اللامركزية العرضية يجب أن تكون أكبر من صفر ولا تتجاوز 100%."
            )
    rows = []
    incomplete = False
    for row in table.itertuples(index=False):
        values = {
            "xcm": pd.to_numeric(pd.Series([row.xcm]), errors="coerce").iloc[0],
            "ycm": pd.to_numeric(pd.Series([row.ycm]), errors="coerce").iloc[0],
            "xcr": pd.to_numeric(pd.Series([row.xcr]), errors="coerce").iloc[0],
            "ycr": pd.to_numeric(pd.Series([row.ycr]), errors="coerce").iloc[0],
        }
        if all(pd.isna(value) for value in values.values()):
            continue
        for direction, coordinate, dimension in (
            ("X", "y", dy),
            ("Y", "x", dx),
        ):
            cm = values[f"{coordinate}cm"]
            cr = values[f"{coordinate}cr"]
            if pd.isna(cm) or pd.isna(cr) or not np.isfinite([cm, cr]).all():
                incomplete = True
                rows.append(
                    {
                        "story": row.story,
                        "diaphragm": row.diaphragm,
                        "direction": direction,
                        "inherent_signed": np.nan,
                        "inherent_absolute": np.nan,
                        "accidental_magnitude": (
                            dimension * accidental_percent / 100
                            if accidental_percent is not None
                            else np.nan
                        ),
                        "total_plus": np.nan,
                        "total_minus": np.nan,
                        "check": "NOT_EVALUATED",
                    }
                )
                continue
            inherent = float(cm - cr)
            accidental = (
                dimension * accidental_percent / 100
                if accidental_percent is not None
                else np.nan
            )
            rows.append(
                {
                    "story": row.story,
                    "diaphragm": row.diaphragm,
                    "direction": direction,
                    "inherent_signed": inherent,
                    "inherent_absolute": abs(inherent),
                    "accidental_magnitude": accidental,
                    "total_plus": (
                        inherent + accidental if pd.notna(accidental) else np.nan
                    ),
                    "total_minus": (
                        inherent - accidental if pd.notna(accidental) else np.nan
                    ),
                    "check": "INFORMATIONAL",
                }
            )

    result = pd.DataFrame(rows)
    if result.empty:
        return result, "NOT_EVALUATED"
    if incomplete or (result["check"] == "NOT_EVALUATED").any():
        status = "INCOMPLETE"
    else:
        status = "INFORMATIONAL"
    return result, status


def calculate_iraqi_2019_torsional_irregularity(
    table: pd.DataFrame,
    accidental_torsion_confirmed: bool,
    accidental_percent: float | None,
    seismic_design_category: str | None,
    ax_applicability_confirmed: bool,
) -> tuple[pd.DataFrame, str]:
    cases = table["output_case"].astype(str)
    seismic = table[cases.str.contains(r"EQ|RS|SPEC", case=False, na=False)].copy()
    if accidental_percent is not None:
        accidental_percent = float(accidental_percent)
        if not np.isfinite(accidental_percent) or not 0 < accidental_percent <= 100:
            raise WorkbookValidationError(
                "نسبة اللامركزية العرضية يجب أن تكون أكبر من صفر ولا تتجاوز 100%."
            )
    cases_confirmed = accidental_torsion_confirmed and accidental_percent is not None
    sdc = str(seismic_design_category or "").strip().upper()
    sdc_selected = sdc in {"A", "B", "C", "D", "E", "F"}
    rows = []
    for row in seismic.itertuples(index=False):
        direction_match = re.search(r"\b([XY])\s*$", str(row.item).upper())
        if direction_match is None:
            continue
        direction = direction_match.group(1)
        maximum = pd.to_numeric(pd.Series([row.max_drift]), errors="coerce").iloc[0]
        average = pd.to_numeric(pd.Series([row.avg_drift]), errors="coerce").iloc[0]
        supplied_ratio = pd.to_numeric(
            pd.Series([row.ratio]), errors="coerce"
        ).iloc[0]
        if pd.notna(maximum) and pd.notna(average) and average != 0:
            ratio = abs(float(maximum / average))
            ratio_source = "Max Drift / Avg Drift"
        elif pd.notna(supplied_ratio):
            ratio = abs(float(supplied_ratio))
            ratio_source = "ETABS Ratio"
        else:
            ratio = np.nan
            ratio_source = "Unavailable"

        valid = pd.notna(ratio) and np.isfinite(ratio) and ratio >= 0
        if not valid:
            candidate = "NOT_EVALUATED"
            ax = np.nan
            ax_status = "NOT_EVALUATED"
        elif ratio > 1.4:
            candidate = "Type 1b"
            ax_status = "NOT_EVALUATED"
            ax = np.nan
            if cases_confirmed and sdc_selected and ax_applicability_confirmed:
                ax = min((ratio / 1.2) ** 2, 3.0)
                ax_status = "COMPUTED"
        elif ratio > 1.2:
            candidate = "Type 1a"
            ax_status = "NOT_EVALUATED"
            ax = np.nan
            if cases_confirmed and sdc_selected and ax_applicability_confirmed:
                ax = min((ratio / 1.2) ** 2, 3.0)
                ax_status = "COMPUTED"
        else:
            candidate = "No torsional irregularity threshold exceeded"
            ax = np.nan
            ax_status = "NOT_REQUIRED"
        rows.append(
            {
                "story": row.story,
                "output_case": row.output_case,
                "direction": direction,
                "item": row.item,
                "ratio": ratio,
                "ratio_source": ratio_source,
                "candidate_classification": candidate,
                "ax": ax,
                "ax_status": ax_status,
                "check": (
                    "INFORMATIONAL"
                    if valid and cases_confirmed
                    else "NOT_EVALUATED"
                ),
            }
        )

    result = pd.DataFrame(rows)
    if result.empty:
        return result, "NOT_EVALUATED"
    if not cases_confirmed or result["ratio"].isna().any():
        status = "NOT_EVALUATED" if not cases_confirmed else "INCOMPLETE"
    else:
        status = "INFORMATIONAL"
    return result, status


PDELTA_RESULT_COLUMNS = (
    "story",
    "story_z",
    "direction",
    "base_shear",
    "output_case",
    "p",
    "story_shear",
    "elastic_drift_ratio",
    "cd",
    "ie",
    "beta",
    "theta",
    "theta_max",
    "classification",
    "reason",
)


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _usable_pdelta_rows(table: pd.DataFrame) -> pd.DataFrame:
    required = {"output_case", "case_type", "step_type", "step_number"}
    if table is None or not required.issubset(table.columns):
        return pd.DataFrame()
    result = table.copy()
    case_types = result["case_type"].map(_clean_text).str.casefold()
    step_types = result["step_type"].map(_clean_text)
    step_numbers = result["step_number"].map(_clean_text)
    output_cases = result["output_case"].map(_clean_text)
    return result[
        (case_types == "combination")
        & (step_types == "")
        & (step_numbers == "")
        & ~output_cases.str.contains("envelope", case=False, regex=False)
        & (output_cases != "")
    ].copy()


def _pdelta_base_shears(base_reactions: pd.DataFrame) -> dict[str, dict[str, float]]:
    required = {"output_case", "case_type", "step_type", "step_number", "fx", "fy", "fz"}
    if base_reactions is None or not required.issubset(base_reactions.columns):
        return {}
    rows = _usable_pdelta_rows(base_reactions)
    if rows.empty:
        return {}
    results: dict[str, dict[str, float]] = {}
    for case, group in rows.groupby(rows["output_case"].map(_clean_text), sort=False):
        if not case or len(group) != 1:
            continue
        values = group.loc[:, ("fx", "fy", "fz")].apply(
            pd.to_numeric, errors="coerce"
        )
        vector = values.to_numpy(dtype=float)
        if not np.isfinite(vector).all():
            continue
        fx, fy, fz = (float(value) for value in vector[0])
        numerical_tolerance = abs(fz) * PDELTA_BASE_SHEAR_RELATIVE_TOLERANCE
        active = {}
        if abs(fx) > numerical_tolerance:
            active["X"] = abs(fx)
        if abs(fy) > numerical_tolerance:
            active["Y"] = abs(fy)
        if active:
            results[case] = active
    return results


def pdelta_combination_options(
    story_forces: pd.DataFrame,
    story_drifts: pd.DataFrame,
    base_reactions: pd.DataFrame,
) -> list[str]:
    force_rows = _usable_pdelta_rows(story_forces)
    drift_rows = _usable_pdelta_rows(story_drifts)
    active_base_reactions = _pdelta_base_shears(base_reactions)
    if force_rows.empty or drift_rows.empty or not active_base_reactions:
        return []

    drift_cases = set(drift_rows["output_case"].map(_clean_text))
    return [
        case
        for case in dict.fromkeys(force_rows["output_case"].map(_clean_text))
        if case in drift_cases and case in active_base_reactions
    ]


def _numeric_values(frame: pd.DataFrame, column: str) -> pd.Series:
    return pd.to_numeric(frame[column], errors="coerce")


def _single_force_record(
    rows: pd.DataFrame, value_columns: tuple[str, ...]
) -> tuple[dict[str, float] | None, str | None]:
    if rows.empty:
        return None, "لا يوجد صف قوة مطابق عند Location = Bottom."
    missing_columns = [column for column in value_columns if column not in rows]
    if missing_columns:
        return None, "أعمدة القوة المطلوبة غير متاحة: " + ", ".join(missing_columns)

    numeric = rows.loc[:, value_columns].apply(pd.to_numeric, errors="coerce")
    values = numeric.to_numpy(dtype=float)
    if not np.isfinite(values).all():
        return None, "قيمة P أو القص مفقودة أو غير رقمية في مصدر القوى."
    if not np.allclose(values, values[0], rtol=1e-8, atol=1e-6):
        return None, "صفوف القوة المكررة عند Bottom متعارضة؛ لم تُدمج."
    return {
        column: float(numeric.iloc[0][column]) for column in value_columns
    }, None


def _story_elevations(drift_rows: pd.DataFrame) -> dict[str, float]:
    if "z" not in drift_rows:
        return {}
    elevations: dict[str, float] = {}
    for story, rows in drift_rows.groupby(
        drift_rows["story"].map(_clean_text), sort=False
    ):
        values = _numeric_values(rows, "z").dropna()
        if not values.empty and np.isfinite(values.to_numpy(dtype=float)).all():
            elevations[story] = float(values.max())
    return elevations


def calculate_pdelta_screening(
    story_forces: pd.DataFrame,
    story_drifts: pd.DataFrame,
    base_reactions: pd.DataFrame,
    selected_combinations: list[str] | tuple[str, ...],
    cd: float,
    ie: float,
    *,
    combinations_confirmed: bool,
    coefficients_confirmed: bool,
    drift_confirmed: bool,
    source_issue: str | None = None,
    beta: float = 1.0,
) -> dict[str, object]:
    required_inputs = [
        "اختر تركيبة تصميم واحدة أو أكثر من الحالات المشتركة بين Story Forces وStory Drifts.",
        "أكد أن معاملات الأحمال في التركيبات المختارة تحقق متطلبات الكود؛ لا يمكن التحقق منها من اسم التركيبة فقط.",
        "أكد أن Cd وIe المحددين للنظام الإنشائي المختار ينطبقان على التركيبات المختارة.",
        "أكد أن Story Drifts.Drift يمثل نسبة انجراف مرن غير مضروبة مسبقاً في Cd/Ie.",
        "تحقق التطبيق من وجود محصلة قص قاعدي جانبية تتجاوز ضوضاء التقريب قبل قبول التركيبة.",
    ]
    base = {
        "status": "NOT_EVALUATED",
        "message": "غير محسوب: لم تُحدد بعد جميع مدخلات وفحوص P-Delta المطلوبة.",
        "required_inputs": required_inputs,
        "selected_combinations": list(selected_combinations or ()),
        "combinations_confirmed": combinations_confirmed,
        "coefficients_confirmed": coefficients_confirmed,
        "drift_confirmed": drift_confirmed,
        "cd": None,
        "ie": None,
        "beta": None,
        "theta_max": None,
        "results": pd.DataFrame(columns=PDELTA_RESULT_COLUMNS),
        "governing_results": pd.DataFrame(columns=PDELTA_RESULT_COLUMNS),
    }

    if source_issue:
        base["message"] = f"غير محسوب: مصدر P-Delta غير متاح أو غير صالح: {source_issue}."
        return base
    options = pdelta_combination_options(
        story_forces, story_drifts, base_reactions
    )
    if not selected_combinations:
        missing_case_metadata = any(
            not {
                "output_case",
                "case_type",
                "step_type",
                "step_number",
            }.issubset(table.columns)
            for table in (story_forces, story_drifts)
        )
        missing_case_metadata = missing_case_metadata or not {
            "output_case",
            "case_type",
            "step_type",
            "step_number",
            "fx",
            "fy",
            "fz",
        }.issubset(base_reactions.columns)
        if missing_case_metadata:
            base["message"] = (
                "غير محسوب: يلزم وجود أعمدة Output Case وCase Type وStep Type "
                "وStep Number في الجداول، وFX/FY/FZ في Base Reactions للتحقق "
                "من وجود حمل جانبي فعلي."
            )
        elif not options:
            base["message"] = (
                "غير محسوب: لا توجد تركيبات Combination مشتركة ذات محصلة قص "
                "قاعدي جانبية فعلية بين الجداول الثلاثة؛ تُستبعد تركيبات الجاذبية "
                "والقيم التي لا تتجاوز 1e-6 من محصلة FZ كضوضاء تقريب."
            )
        return base
    if not combinations_confirmed or not coefficients_confirmed or not drift_confirmed:
        missing_confirmations = []
        if not combinations_confirmed:
            missing_confirmations.append("تأكيد معاملات التركيبات المختارة")
        if not coefficients_confirmed:
            missing_confirmations.append("تأكيد انطباق Cd وIe")
        if not drift_confirmed:
            missing_confirmations.append("تأكيد أن Drift مرن وغير مكبّر مسبقاً")
        base["message"] = (
            "غير محسوب: يلزم " + " و".join(missing_confirmations) + "."
        )
        return base

    try:
        cd_value = _positive_number(cd, "Cd")
        ie_value = _positive_number(ie, "Ie")
        beta_value = _positive_number(beta, "β")
    except WorkbookValidationError as exc:
        base["message"] = f"غير محسوب: {exc}"
        return base

    invalid_cases = [
        case for case in selected_combinations if case not in options
    ]
    if invalid_cases:
        base["message"] = (
            "غير محسوب: التركيبات التالية لا تحتوي بيانات Combination متطابقة "
            "غير مغلّفة مع قص قاعدي جانبي مؤثر في الجداول المطلوبة: "
            + ", ".join(map(str, invalid_cases))
            + "."
        )
        return base

    required_force_columns = {
        "story",
        "output_case",
        "case_type",
        "step_type",
        "step_number",
        "location",
        "p",
        "vx",
        "vy",
    }
    required_drift_columns = {
        "story",
        "output_case",
        "case_type",
        "step_type",
        "step_number",
        "direction",
        "drift",
    }
    required_base_columns = {
        "output_case",
        "case_type",
        "step_type",
        "step_number",
        "fx",
        "fy",
        "fz",
    }
    if not required_force_columns.issubset(story_forces.columns):
        base["message"] = (
            "غير محسوب: أعمدة Story Forces اللازمة للحساب غير مكتملة."
        )
        return base
    if not required_drift_columns.issubset(story_drifts.columns):
        base["message"] = (
            "غير محسوب: أعمدة Story Drifts اللازمة للحساب غير مكتملة."
        )
        return base
    if not required_base_columns.issubset(base_reactions.columns):
        base["message"] = (
            "غير محسوب: أعمدة Base Reactions اللازمة للتحقق من القص الجانبي غير مكتملة."
        )
        return base

    force_rows = _usable_pdelta_rows(story_forces)
    drift_rows = _usable_pdelta_rows(story_drifts)
    base_shears = _pdelta_base_shears(base_reactions)
    force_rows["_story_key"] = force_rows["story"].map(_clean_text)
    force_rows["_case_key"] = force_rows["output_case"].map(_clean_text)
    force_rows["_location_key"] = force_rows["location"].map(_clean_text).str.casefold()
    drift_rows["_story_key"] = drift_rows["story"].map(_clean_text)
    drift_rows["_case_key"] = drift_rows["output_case"].map(_clean_text)
    drift_rows["_direction_key"] = drift_rows["direction"].map(_clean_text).str.upper()
    selected_drift_rows = drift_rows[
        drift_rows["_case_key"].isin(selected_combinations)
    ]
    elevations = _story_elevations(selected_drift_rows)
    theta_max = min(0.25, 0.5 / (beta_value * cd_value))
    selected_base_shears = {
        case: base_shears.get(case, {}) for case in selected_combinations
    }
    active_directions = {
        direction
        for case_shears in selected_base_shears.values()
        for direction in case_shears
    }

    stories_by_case: dict[str, list[str]] = {}
    for case in selected_combinations:
        stories = set(
            force_rows.loc[force_rows["_case_key"] == case, "_story_key"]
        ) | set(drift_rows.loc[drift_rows["_case_key"] == case, "_story_key"])
        stories_by_case[case] = sorted(
            stories,
            key=lambda story: (
                story not in elevations,
                -elevations.get(story, 0.0),
                story.casefold(),
            ),
        )

    rows = []
    for case in selected_combinations:
        for story in stories_by_case[case]:
            matching_forces = force_rows[
                (force_rows["_case_key"] == case)
                & (force_rows["_story_key"] == story)
                & (force_rows["_location_key"] == "bottom")
            ]
            force_values, force_issue = _single_force_record(
                matching_forces, ("p", "vx", "vy")
            )
            matching_drifts = drift_rows[
                (drift_rows["_case_key"] == case)
                & (drift_rows["_story_key"] == story)
                & drift_rows["_direction_key"].isin(("X", "Y"))
            ]

            for direction, shear_column in (("X", "vx"), ("Y", "vy")):
                base_shear = base_shears.get(case, {}).get(direction, np.nan)
                if pd.isna(base_shear):
                    continue
                direction_drifts = matching_drifts[
                    matching_drifts["_direction_key"] == direction
                ]
                reasons = []
                theta = np.nan
                p_value = np.nan
                shear_value = np.nan
                drift_value = np.nan

                if force_issue:
                    reasons.append(force_issue)
                elif force_values is not None:
                    p_value = force_values["p"]
                    shear_value = abs(force_values[shear_column])
                    if p_value <= 0:
                        reasons.append("P غير موجب؛ لا يُفترض أن قيمته صفر ضغط.")
                    if shear_value <= 0:
                        reasons.append("قص الطابق صفري؛ لا يمكن حساب θ.")

                if direction_drifts.empty:
                    reasons.append(f"لا يوجد انجراف مطابق للاتجاه {direction}.")
                elif direction_drifts["drift"].isna().any():
                    reasons.append("قيمة انجراف مفقودة أو غير رقمية.")
                else:
                    drift_values = _numeric_values(direction_drifts, "drift")
                    if not np.isfinite(drift_values.to_numpy(dtype=float)).all():
                        reasons.append("قيمة انجراف غير منتهية.")
                    else:
                        drift_value = float(drift_values.abs().max())

                classification = "NOT_EVALUATED"
                if not reasons and p_value > 0 and shear_value > 0:
                    theta = p_value * drift_value / (shear_value * ie_value)
                    if not np.isfinite(theta):
                        theta = np.nan
                        reasons.append("تعذر حساب θ كقيمة منتهية.")
                    elif theta > theta_max:
                        classification = "LIMIT_EXCEEDED"
                    elif theta <= 0.1:
                        classification = "NEGLIGIBLE"
                    else:
                        classification = "INCLUDE_PDELTA"
                rows.append(
                    {
                        "story": story,
                        "story_z": elevations.get(story, np.nan),
                        "direction": direction,
                        "base_shear": base_shear,
                        "output_case": case,
                        "p": p_value,
                        "story_shear": shear_value,
                        "elastic_drift_ratio": drift_value,
                        "cd": cd_value,
                        "ie": ie_value,
                        "beta": beta_value,
                        "theta": theta,
                        "theta_max": theta_max,
                        "classification": classification,
                        "reason": "؛ ".join(reasons),
                    }
                )

    missing_directions = {"X", "Y"} - active_directions
    all_stories = dict.fromkeys(
        story for case_stories in stories_by_case.values() for story in case_stories
    )
    for direction in sorted(missing_directions):
        for story in all_stories:
            rows.append(
                {
                    "story": story,
                    "story_z": elevations.get(story, np.nan),
                    "direction": direction,
                    "base_shear": np.nan,
                    "output_case": "",
                    "p": np.nan,
                    "story_shear": np.nan,
                    "elastic_drift_ratio": np.nan,
                    "cd": cd_value,
                    "ie": ie_value,
                    "beta": beta_value,
                    "theta": np.nan,
                    "theta_max": theta_max,
                    "classification": "NOT_EVALUATED",
                    "reason": (
                        f"لا تتضمن التركيبات المختارة محصلة قص قاعدي مؤثرة "
                        f"في الاتجاه {direction}؛ هذا الاتجاه غير محسوب."
                    ),
                }
            )

    results = pd.DataFrame(rows, columns=PDELTA_RESULT_COLUMNS)
    if results.empty:
        base["message"] = "غير محسوب: لم توجد طوابق مشتركة في التركيبات المختارة."
        return base

    results["_valid"] = results["theta"].notna()
    governing_rows = []
    for (story, direction), group in results.groupby(
        ["story", "direction"], sort=False
    ):
        valid = group[group["_valid"]]
        if valid.empty:
            governing_rows.append(group.iloc[0].drop(labels="_valid").to_dict())
            continue
        governing_rows.append(
            valid.loc[valid["theta"].idxmax()].drop(labels="_valid").to_dict()
        )
    governing = pd.DataFrame(governing_rows, columns=PDELTA_RESULT_COLUMNS)
    governing["_sort_z"] = pd.to_numeric(governing["story_z"], errors="coerce")
    governing = governing.sort_values(
        ["_sort_z", "story", "direction"],
        ascending=[False, True, True],
        na_position="last",
    ).drop(columns="_sort_z").reset_index(drop=True)
    results["_sort_z"] = pd.to_numeric(results["story_z"], errors="coerce")
    results = results.sort_values(
        ["_sort_z", "story", "direction", "output_case"],
        ascending=[False, True, True, True],
        na_position="last",
    ).drop(columns=["_sort_z", "_valid"]).reset_index(drop=True)

    evaluated = results[results["theta"].notna()]
    if evaluated.empty:
        status = "INCOMPLETE"
        message = "غير محسوب: لم تتوفر بيانات متطابقة وصالحة لأي طابق واتجاه."
    elif (evaluated["classification"] == "LIMIT_EXCEEDED").any():
        status = "PDELTA_LIMIT_EXCEEDED"
        message = (
            "تجاوز θ الحد θmax في نتيجة واحدة أو أكثر؛ لا تُعد هذه النتيجة "
            "مقبولة بهذا الفحص، ويلزم إجراء التحليل/المراجعة الإنشائية الإضافية "
            "وفق متطلبات الكود."
        )
        if results["theta"].isna().any():
            message += " توجد أيضاً حالات أو اتجاهات غير محسوبة؛ لا يشملها هذا الحكم."
    elif (evaluated["classification"] == "INCLUDE_PDELTA").any():
        status = "PDELTA_REQUIRED"
        message = (
            "تجاوز θ حد الإهمال 0.1 في نتيجة واحدة أو أكثر؛ يجب إدراج تأثير "
            "P-Delta في التحليل وفق الكود العراقي."
        )
        if results["theta"].isna().any():
            message += " توجد أيضاً حالات أو اتجاهات غير محسوبة؛ لا يشملها هذا الحكم."
    elif results["theta"].isna().any():
        status = "INCOMPLETE"
        message = (
            "النتائج المحسوبة لا تتجاوز حد الإهمال 0.1، لكن بعض الطوابق أو "
            "الاتجاهات غير مكتملة؛ لا يصدر حكم شامل."
        )
    else:
        status = "PDELTA_NEGLIGIBLE"
        message = (
            "جميع قيم θ المحسوبة لا تتجاوز 0.1؛ يمكن إهمال تأثير P-Delta "
            "وفق فحص الكود العراقي للتركيبات المختارة."
        )
    unresolved_inputs = (
        ["توجد صفوف غير محسوبة؛ راجع عمود السبب لكل طابق واتجاه."]
        if results["theta"].isna().any()
        else []
    )

    return {
        "status": status,
        "message": message,
        "required_inputs": unresolved_inputs,
        "selected_combinations": list(selected_combinations),
        "combinations_confirmed": combinations_confirmed,
        "coefficients_confirmed": coefficients_confirmed,
        "drift_confirmed": drift_confirmed,
        "cd": cd_value,
        "ie": ie_value,
        "beta": beta_value,
        "theta_max": theta_max,
        "results": results,
        "governing_results": governing,
    }


def pdelta_assessment(source_issue: str | None = None) -> dict[str, object]:
    message = (
        "غير محسوب: اختر تركيبات التصميم المؤهلة من بيانات ETABS وأكد معاملات "
        "التركيبات وCd وIe قبل إصدار نتيجة P-Delta."
    )
    if source_issue:
        message += f" مصدر بيانات إضافي غير متاح أو غير صالح: {source_issue}."
    return {
        "status": "NOT_EVALUATED",
        "message": message,
        "required_inputs": [
            "اختيار تركيبة أو أكثر من التركيبات المشتركة غير المغلّفة بين Story Forces وStory Drifts وBase Reactions.",
            "تأكيد أن معاملات الأحمال في التركيبات المختارة تحقق متطلبات الكود.",
            "تأكيد انطباق Cd وIe المحددين على التحليل.",
            "تأكيد أن Story Drifts.Drift انجراف مرن غير مكبّر مسبقاً في Cd/Ie.",
            "توافر محصلات القوى الجانبية FX/FY وFZ في Base Reactions.",
        ],
        "selected_combinations": [],
        "combinations_confirmed": False,
        "coefficients_confirmed": False,
        "drift_confirmed": False,
        "cd": None,
        "ie": None,
        "beta": None,
        "theta_max": None,
        "results": pd.DataFrame(columns=PDELTA_RESULT_COLUMNS),
        "governing_results": pd.DataFrame(columns=PDELTA_RESULT_COLUMNS),
    }
