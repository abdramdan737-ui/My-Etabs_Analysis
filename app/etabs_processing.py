import io
import hashlib
import re
from dataclasses import dataclass, field
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
    optional_table_issues: dict[str, str] = field(default_factory=dict)


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
    "diaphragm_drifts": {
        "label": "Diaphragm Max Over Avg Drifts",
        "aliases": ("Diaphragm Max Over Avg Drifts",),
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
    optional_issues = {}
    missing_optional = missing_optional_sheets(sheet_names)
    pdelta_issues = [
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
            optional_issues[key] = str(exc)
            if OPTIONAL_TABLES[key]["pdelta_source"]:
                pdelta_issues.append(f"{sheet_name}: {exc}")
        except Exception as exc:
            if key in REQUIRED_SHEETS:
                raise WorkbookValidationError(
                    f"تعذر قراءة ورقة {sheet_name}: {exc}"
                ) from exc
            optional_issues[key] = str(exc)
            if OPTIONAL_TABLES[key]["pdelta_source"]:
                pdelta_issues.append(f"تعذر قراءة ورقة {sheet_name}: {exc}")

    return EtabsWorkbook(
        tables=tables,
        pdelta_source_issue="؛ ".join(pdelta_issues) if pdelta_issues else None,
        missing_optional_sheets=tuple(
            label for _key, label, _aliases in missing_optional
        ),
        optional_table_issues=optional_issues,
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
