import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest

from app.test_etabs_processing import build_workbook


class StreamlitAppNavigationTests(unittest.TestCase):
    def test_local_app_starts_without_auth_and_exposes_section_navigation(self):
        app = AppTest.from_file(str(Path(__file__).with_name("etabs_tables.py")))

        app.run()

        self.assertEqual(len(app.exception), 0)
        self.assertIn("selected_audit_section", app.session_state)
        self.assertEqual(len(app.get("file_uploader")), 6)

    def test_selecting_audit_section_hides_other_section_containers(self):
        app = AppTest.from_file(str(Path(__file__).with_name("etabs_tables.py"))).run()
        app.get("file_uploader")[0].upload(
            "results.xlsx",
            build_workbook(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        app.radio[0].set_value("pdelta").run()

        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.radio[0].value, "pdelta")
        rendered_styles = "\n".join(element.value for element in app.markdown)
        self.assertIn(".st-key-audit_section_centers", rendered_styles)
        self.assertNotIn(".st-key-audit_section_pdelta", rendered_styles)
        workbook_versions = app.session_state["_etabs_workbook_versions"]
        parsed_workbook = next(iter(workbook_versions.values()))
        self.assertIn("story_forces", parsed_workbook.tables)


if __name__ == "__main__":
    unittest.main()
