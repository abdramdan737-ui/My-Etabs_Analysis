import io
import unittest

import pandas as pd
from openpyxl import Workbook

from app.etabs_processing import (
    WorkbookValidationError,
    calculate_base_reactions,
    calculate_eccentricity,
    calculate_modal_mass,
    calculate_iraqi_2019_story_drifts,
    calculate_iraqi_2019_torsion_eccentricities,
    calculate_iraqi_2019_torsional_irregularity,
    calculate_story_drifts,
    get_cached_workbook,
    load_etabs_workbook,
    MAX_CACHED_WORKBOOK_VERSIONS,
    missing_optional_sheets,
    missing_required_sheets,
    pdelta_assessment,
    validate_upload_filename,
)


def build_workbook(
    include_stiffness=True,
    include_pdelta_sheets=True,
    include_additional_sheets=True,
    omit_base_fx=False,
    centers_xcm=1000,
):
    workbook = Workbook()
    workbook.remove(workbook.active)
    sheets = {
        "Centers Of Mass And Rigidity": (
            ["Story", "Diaphragm", "Mass X", "Mass Y", "XCM", "YCM", "Cum Mass X", "Cum Mass Y", "XCCM", "YCCM", "XCR", "YCR"],
            ["", "", "kN-s2/mm", "kN-s2/mm", "mm", "mm", "kN-s2/mm", "kN-s2/mm", "mm", "mm", "mm", "mm"],
            ["Story1", "D1", 1, 1, centers_xcm, 1200, 1, 1, centers_xcm, 1200, 1100, 1250],
        ),
        "Base Reactions": (
            ["Output Case", "Case Type", "Step Type", "Step Number", "FX", "FY"],
            ["", "", "", "", "kN", "kN"],
            ["EQX", "LinStatic", "", "", -100, 0],
        ),
        "Story Drifts": (
            ["Story", "Output Case", "Case Type", "Step Type", "Step Number", "Direction", "Drift", "Drift/", "Label", "X", "Y", "Z"],
            ["", "", "", "", "", "", "", "", "", "mm", "mm", "mm"],
            ["Story1", "EQX", "LinStatic", "", "", "X", 0.004, "1/250", "1", 0, 0, 3500],
        ),
        "Modal Participating Mass Ratios": (
            ["Case", "Mode", "Period", "UX", "UY", "UZ", "SumUX", "SumUY"],
            ["", "", "sec", "", "", "", "", ""],
            ["Modal", "1", 0.45, 0.8, 0.1, 0, 0.95, 0.91],
        ),
        "Story Max Over Avg Disp": (
            ["Story", "Output Case", "Case Type", "Step Type", "Step Number", "Direction", "Maximum", "Average", "Ratio"],
            ["", "", "", "", "", "", "mm", "mm", ""],
            ["Story1", "EQX", "LinStatic", "", "", "X", 2, 1, 2],
        ),
        "Story Forces": (
            ["Story", "Output Case", "Case Type", "Step Type", "Step Number", "Location", "P", "VX", "VY"],
            ["", "", "", "", "", "", "kN", "kN", "kN"],
            ["Story1", "EQX", "LinStatic", "", "", "Bottom", 100, -100, 0],
        ),
        "Diaphragm CM Displacements": (
            ["Story", "Diaphragm", "Output Case", "Case Type", "Step Type", "Step Number", "UX", "UY", "RZ", "Point", "X", "Y", "Z"],
            ["", "", "", "", "", "", "mm", "mm", "rad", "", "mm", "mm", "mm"],
            ["Story1", "D1", "EQX", "LinStatic", "", "", 10, 0, 0, "1", 0, 0, 3500],
        ),
    }
    if include_stiffness:
        sheets["Story Stiffness"] = (
            ["Story", "Output Case", "Case Type", "Step Type", "Step Number", "Shear X", "Drift X", "Stiff X", "Shear Y", "Drift Y", "Stiff Y"],
            ["", "", "", "", "", "kN", "mm", "kN/mm", "kN", "mm", "kN/mm"],
            ["Story1", "EQX", "LinStatic", "", 1, 100, 5, 20, 0, 0, 10],
        )
    if not include_pdelta_sheets:
        sheets.pop("Story Forces")
        sheets.pop("Diaphragm CM Displacements")
    if include_additional_sheets:
        sheets["Story Masses"] = (
            ["Story", "Diaphragm", "Mass X", "Mass Y", "XCM", "YCM"],
            ["", "", "kN-s2/mm", "kN-s2/mm", "mm", "mm"],
            ["Story1", "D1", 1.0, 1.2, 1000, 1200],
        )
        sheets["Joint Displacements"] = (
            ["Story", "Output Case", "Case Type", "U1", "U2", "U3", "X", "Y", "Z"],
            ["", "", "", "mm", "mm", "mm", "mm", "mm", "mm"],
            ["Story1", "EQX", "LinStatic", 1.5, 0.1, 0.0, 100, 200, 3500],
        )
        sheets["Diaphragm Max Over Avg Drifts"] = (
            ["Story", "Output Case", "Case Type", "Step Type", "Step Number", "Item", "Max Drift", "Avg Drift", "Ratio"],
            ["", "", "", "", "", "", "", "", ""],
            ["Story1", "EQX", "LinStatic", "", "", "Diaph D1 X", 0.0013, 0.001, 1.3],
        )
    if omit_base_fx:
        headers, units, row = sheets["Base Reactions"]
        sheets["Base Reactions"] = (
            headers[:4] + headers[5:],
            units[:4] + units[5:],
            row[:4] + row[5:],
        )
    for sheet_name, (headers, units, row) in sheets.items():
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([f"TABLE: {sheet_name}"])
        sheet.append(headers)
        sheet.append(units)
        sheet.append(row)
    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


class EtabsProcessingTests(unittest.TestCase):
    def setUp(self):
        self.workbook = load_etabs_workbook(build_workbook())

    def test_reads_etabs_headers_and_skips_title_and_units(self):
        centers = self.workbook.tables["centers"]
        self.assertEqual(centers.loc[0, "story"], "Story1")
        self.assertEqual(centers.loc[0, "xcm"], 1000)
        self.assertIn("Drift/", self.workbook.tables["story_drifts"].columns)
        self.assertEqual(self.workbook.tables["story_drifts"].loc[0, "drift"], 0.004)

    def test_missing_required_sheet_is_rejected(self):
        with self.assertRaisesRegex(WorkbookValidationError, "Story Stiffness"):
            load_etabs_workbook(build_workbook(include_stiffness=False))

    def test_missing_required_column_is_rejected(self):
        with self.assertRaises(WorkbookValidationError):
            load_etabs_workbook(build_workbook(omit_base_fx=True))

    def test_missing_sheet_list_names_every_required_table(self):
        missing = missing_required_sheets(["Base Reactions"])
        missing_aliases = [aliases for _key, aliases in missing]
        self.assertIn(("Story Drifts",), missing_aliases)
        self.assertIn(("Centers Of Mass And Rigidity",), missing_aliases)
        self.assertNotIn(("Base Reactions",), missing_aliases)

    def test_missing_pdelta_sheets_are_reported_but_do_not_block_other_checks(self):
        workbook = load_etabs_workbook(
            build_workbook(
                include_pdelta_sheets=False, include_additional_sheets=False
            )
        )
        self.assertIn("Story Forces", workbook.pdelta_source_issue)
        self.assertIn("Diaphragm CM Displacements", workbook.pdelta_source_issue)
        self.assertIn("base_reactions", workbook.tables)
        self.assertIn("Story Masses", workbook.missing_optional_sheets)
        self.assertIn("Joint Displacements", workbook.missing_optional_sheets)
        self.assertIn(
            "Diaphragm Max Over Avg Drifts", workbook.missing_optional_sheets
        )
        self.assertNotIn("diaphragm_drifts", workbook.tables)

    def test_reads_optional_etabs_tables_without_imposing_new_formulas(self):
        self.assertEqual(
            self.workbook.tables["story_mass"].loc[0, "Mass X"], 1.0
        )
        self.assertEqual(
            self.workbook.tables["joint_displacements"].loc[0, "U1"], 1.5
        )
        self.assertAlmostEqual(
            self.workbook.tables["diaphragm_drifts"].loc[0, "max_drift"], 0.0013
        )
        self.assertEqual(self.workbook.missing_optional_sheets, ())

    def test_optional_sheet_names_are_listed_when_missing(self):
        missing = missing_optional_sheets(
            ["Centers Of Mass And Rigidity", "Base Reactions"]
        )
        labels = [label for _key, label, _aliases in missing]
        self.assertIn("Story Forces", labels)
        self.assertIn("Story Masses", labels)
        self.assertIn("Joint Displacements", labels)
        self.assertIn("Diaphragm Max Over Avg Drifts", labels)

    def test_upload_must_use_results_xlsx_filename(self):
        validate_upload_filename("results.xlsx")
        validate_upload_filename("RESULTS.XLSX")
        with self.assertRaises(WorkbookValidationError):
            validate_upload_filename("etabs_export.xlsx")

    def test_session_cache_restores_previous_values_after_file_change(self):
        state = {}
        original_content = build_workbook(centers_xcm=1000)
        changed_content = build_workbook(centers_xcm=2000)

        original_hash, original = get_cached_workbook(state, original_content)
        changed_hash, changed = get_cached_workbook(state, changed_content)
        restored_hash, restored = get_cached_workbook(state, original_content)

        self.assertNotEqual(original_hash, changed_hash)
        self.assertEqual(restored_hash, original_hash)
        self.assertEqual(original.tables["centers"].loc[0, "xcm"], 1000)
        self.assertEqual(changed.tables["centers"].loc[0, "xcm"], 2000)
        self.assertEqual(restored.tables["centers"].loc[0, "xcm"], 1000)

    def test_session_cache_retains_only_the_latest_four_versions(self):
        state = {}
        for value in range(MAX_CACHED_WORKBOOK_VERSIONS + 1):
            get_cached_workbook(state, build_workbook(centers_xcm=1000 + value))
        self.assertEqual(
            len(state["_etabs_workbook_versions"]), MAX_CACHED_WORKBOOK_VERSIONS
        )

    def test_eccentricity_calculation(self):
        result, status = calculate_eccentricity(
            self.workbook.tables["centers"], dx=10000, dy=5000
        )
        self.assertEqual(status, "PASS")
        self.assertAlmostEqual(result.loc[0, "ex_dx_percent"], 1.0)
        self.assertAlmostEqual(result.loc[0, "ey_dy_percent"], 1.0)

    def test_base_reaction_check_and_scale_factor(self):
        reactions = pd.DataFrame(
            [
                {"output_case": "EQX", "fx": -100, "fy": 0},
                {"output_case": "EQY", "fx": 0, "fy": -200},
                {"output_case": "RSX1", "fx": -90, "fy": 0},
                {"output_case": "RSY1", "fx": 0, "fy": -180},
            ]
        )
        result, status = calculate_base_reactions(reactions, 0.8, 1.0, 5.0)
        self.assertEqual(status, "PASS")
        self.assertAlmostEqual(
            result.loc[0, "recommended_scale_factor"], (80 / 90) * (1 / 5)
        )

    def test_base_reaction_requires_both_directions(self):
        with self.assertRaises(WorkbookValidationError):
            calculate_base_reactions(
                self.workbook.tables["base_reactions"], 0.85, 1.0, 5.0
            )

    def test_drift_rule_uses_entered_periods_and_response_factor(self):
        drifts = pd.DataFrame(
            [
                {"story": "Story1", "output_case": "EQX", "direction": "X", "drift": 0.005},
                {"story": "Story1", "output_case": "EQY", "direction": "Y", "drift": 0.006},
            ]
        )
        result, status = calculate_story_drifts(drifts, 5.0, tx=0.8, ty=0.5)
        self.assertEqual(status, "PASS")
        self.assertAlmostEqual(result.loc[0, "limit_ratio"], 0.02 / (0.7 * 5.0))
        self.assertAlmostEqual(result.loc[1, "limit_ratio"], 0.025 / (0.7 * 5.0))

    def test_modal_mass_requires_ninety_percent_in_both_axes(self):
        modal = pd.DataFrame([{"sumux": 0.95, "sumuy": 0.89}])
        result, status = calculate_modal_mass(modal)
        self.assertEqual(status, "FAIL")
        self.assertEqual(result.loc[0, "check"], "PASS")
        self.assertEqual(result.loc[1, "check"], "FAIL")

    def test_modal_mass_keeps_cases_separate(self):
        modal = pd.DataFrame(
            [
                {"case": "ModalA", "sumux": 0.92, "sumuy": 0.83},
                {"case": "ModalB", "sumux": 0.88, "sumuy": 0.94},
            ]
        )
        result, status = calculate_modal_mass(modal)
        self.assertEqual(status, "FAIL")
        self.assertEqual(
            result[["case", "direction", "check"]].values.tolist(),
            [
                ["ModalA", "X", "PASS"],
                ["ModalA", "Y", "FAIL"],
                ["ModalB", "X", "FAIL"],
                ["ModalB", "Y", "PASS"],
            ],
        )

    def test_iraqi_drift_uses_risk_category_and_cd_over_ie(self):
        drifts = pd.DataFrame(
            [
                {"story": "Story1", "output_case": "EQX", "direction": "X", "drift": 0.003},
                {"story": "Story2", "output_case": "EQX", "direction": "X", "drift": 0.004},
            ]
        )
        result, status = calculate_iraqi_2019_story_drifts(
            drifts, "III", applicability_confirmed=True, cd=5.5, ie=1.25
        )
        self.assertEqual(status, "FAIL")
        self.assertAlmostEqual(
            result.loc[0, "design_drift_ratio"], 0.003 * 5.5 / 1.25
        )
        self.assertAlmostEqual(result.loc[0, "limit_ratio"], 0.015)
        self.assertEqual(result["check"].tolist(), ["PASS", "FAIL"])

    def test_iraqi_drift_limits_by_risk_category(self):
        drifts = pd.DataFrame(
            [{"story": "Story1", "output_case": "EQX", "direction": "X", "drift": 0.001}]
        )
        for category, expected_limit in (
            ("I", 0.020),
            ("II", 0.020),
            ("III", 0.015),
            ("IV", 0.010),
        ):
            with self.subTest(category=category):
                result, status = calculate_iraqi_2019_story_drifts(
                    drifts,
                    category,
                    applicability_confirmed=True,
                    cd=1.0,
                    ie=1.0,
                )
                self.assertEqual(status, "PASS")
                self.assertEqual(result.loc[0, "limit_ratio"], expected_limit)

    def test_iraqi_drift_stays_uncomputed_without_confirmed_applicability(self):
        drifts = pd.DataFrame(
            [{"story": "Story1", "output_case": "EQX", "direction": "X", "drift": 0.003}]
        )
        result, status = calculate_iraqi_2019_story_drifts(
            drifts, "II", applicability_confirmed=False, cd=5.5, ie=1.0
        )
        self.assertEqual(status, "NOT_EVALUATED")
        self.assertEqual(result.loc[0, "check"], "NOT_EVALUATED")
        self.assertTrue(pd.isna(result.loc[0, "design_drift_ratio"]))

    def test_iraqi_drift_stays_uncomputed_without_risk_category(self):
        drifts = pd.DataFrame(
            [{"story": "Story1", "output_case": "EQX", "direction": "X", "drift": 0.003}]
        )
        result, status = calculate_iraqi_2019_story_drifts(
            drifts, None, applicability_confirmed=True, cd=5.5, ie=1.0
        )
        self.assertEqual(status, "NOT_EVALUATED")
        self.assertEqual(result.loc[0, "check"], "NOT_EVALUATED")

    def test_iraqi_torsion_uses_perpendicular_dimension_for_accidental_eccentricity(self):
        result, status = calculate_iraqi_2019_torsion_eccentricities(
            self.workbook.tables["centers"],
            dx=20000,
            dy=10000,
            accidental_percent=5.0,
        )
        self.assertEqual(status, "INFORMATIONAL")
        x_direction = result[result["direction"] == "X"].iloc[0]
        y_direction = result[result["direction"] == "Y"].iloc[0]
        self.assertAlmostEqual(x_direction["inherent_signed"], -50)
        self.assertAlmostEqual(x_direction["accidental_magnitude"], 500)
        self.assertAlmostEqual(y_direction["inherent_signed"], -100)
        self.assertAlmostEqual(y_direction["accidental_magnitude"], 1000)

    def test_torsion_leaves_accidental_offsets_uncomputed_without_selected_percent(self):
        result, status = calculate_iraqi_2019_torsion_eccentricities(
            self.workbook.tables["centers"],
            dx=20000,
            dy=10000,
            accidental_percent=None,
        )
        self.assertEqual(status, "INFORMATIONAL")
        self.assertTrue(result["inherent_absolute"].notna().all())
        self.assertTrue(result["accidental_magnitude"].isna().all())

    def test_torsional_irregularity_and_ax_are_conditional_and_capped(self):
        drifts = pd.DataFrame(
            [
                {
                    "story": "Story1",
                    "output_case": "EQX",
                    "item": "Diaph D1 X",
                    "max_drift": 0.0024,
                    "avg_drift": 0.001,
                    "ratio": 2.4,
                }
            ]
        )
        diagnostic, diagnostic_status = calculate_iraqi_2019_torsional_irregularity(
            drifts,
            accidental_torsion_confirmed=False,
            accidental_percent=None,
            seismic_design_category=None,
            ax_applicability_confirmed=False,
        )
        self.assertEqual(diagnostic_status, "NOT_EVALUATED")
        self.assertEqual(diagnostic.loc[0, "candidate_classification"], "Type 1b")
        self.assertTrue(pd.isna(diagnostic.loc[0, "ax"]))

        no_percentage, no_percentage_status = calculate_iraqi_2019_torsional_irregularity(
            drifts,
            accidental_torsion_confirmed=True,
            accidental_percent=None,
            seismic_design_category="D",
            ax_applicability_confirmed=True,
        )
        self.assertEqual(no_percentage_status, "NOT_EVALUATED")
        self.assertTrue(pd.isna(no_percentage.loc[0, "ax"]))

        confirmed, status = calculate_iraqi_2019_torsional_irregularity(
            drifts,
            accidental_torsion_confirmed=True,
            accidental_percent=5.0,
            seismic_design_category="D",
            ax_applicability_confirmed=True,
        )
        self.assertEqual(status, "INFORMATIONAL")
        self.assertEqual(confirmed.loc[0, "candidate_classification"], "Type 1b")
        self.assertEqual(confirmed.loc[0, "ax"], 3.0)

    def test_ax_stays_uncomputed_when_applicability_is_not_confirmed(self):
        drifts = pd.DataFrame(
            [
                {
                    "story": "Story1",
                    "output_case": "EQX",
                    "item": "Diaph D1 X",
                    "max_drift": 0.0014,
                    "avg_drift": 0.001,
                    "ratio": 1.4,
                }
            ]
        )
        result, status = calculate_iraqi_2019_torsional_irregularity(
            drifts,
            accidental_torsion_confirmed=True,
            accidental_percent=5.0,
            seismic_design_category="C",
            ax_applicability_confirmed=False,
        )
        self.assertEqual(status, "INFORMATIONAL")
        self.assertEqual(result.loc[0, "candidate_classification"], "Type 1a")
        self.assertTrue(pd.isna(result.loc[0, "ax"]))

    def test_ax_stays_uncomputed_without_seismic_design_category(self):
        drifts = pd.DataFrame(
            [
                {
                    "story": "Story1",
                    "output_case": "EQX",
                    "item": "Diaph D1 X",
                    "max_drift": 0.0014,
                    "avg_drift": 0.001,
                    "ratio": 1.4,
                }
            ]
        )
        result, _status = calculate_iraqi_2019_torsional_irregularity(
            drifts,
            accidental_torsion_confirmed=True,
            accidental_percent=2.0,
            seismic_design_category=None,
            ax_applicability_confirmed=True,
        )
        self.assertTrue(pd.isna(result.loc[0, "ax"]))
        self.assertEqual(result.loc[0, "ax_status"], "NOT_EVALUATED")

    def test_pdelta_is_explicitly_not_evaluated(self):
        result = pdelta_assessment()
        self.assertEqual(result["status"], "NOT_EVALUATED")
        self.assertTrue(result["required_inputs"])
        self.assertIn("مراجع خارجية مكسورة", result["message"])


if __name__ == "__main__":
    unittest.main()
