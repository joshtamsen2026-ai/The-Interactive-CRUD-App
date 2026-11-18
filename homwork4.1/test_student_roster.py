"""Regression tests using temporary databases and reports."""

import csv
import io
import sqlite3
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import call, patch

from docx import Document

import student_roster as roster


class StudentRosterTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(temporary_directory.cleanup)
        directory = Path(temporary_directory.name)
        paths = patch.multiple(
            roster,
            DATABASE_PATH=directory / "school.db",
            CSV_PATH=directory / "results.csv",
            REPORT_PATH=directory / "report.docx",
        )
        paths.start()
        self.addCleanup(paths.stop)
        open_reports = patch("os.startfile")
        self.open_reports = open_reports.start()
        self.addCleanup(open_reports.stop)
        roster.initialize_database()

    def test_crud_persists_changes_and_handles_missing_ids(self) -> None:
        student_id = roster.add_student("O'Connor", 10, None)
        self.assertEqual(
            roster.get_student_by_id(student_id),
            (student_id, "O'Connor", 10, None),
        )
        self.assertTrue(roster.update_student_gpa(student_id, 3.8))
        self.assertEqual(
            roster.get_all_students(), [(student_id, "O'Connor", 10, 3.8)]
        )
        self.assertTrue(roster.delete_student(student_id))
        self.assertIsNone(roster.get_student_by_id(student_id))
        self.assertFalse(roster.update_student_gpa(student_id, 4.0))
        self.assertFalse(roster.delete_student(student_id))
        self.assertEqual(roster.get_all_students(), [])
        self.assertGreater(roster.add_student("Next Student", 11, 3.0), student_id)

    def test_database_context_rolls_back_on_failure(self) -> None:
        def abort_transaction() -> None:
            with roster._database_connection() as connection:
                connection.execute(
                    "INSERT INTO students (name, grade, gpa) VALUES (?, ?, ?)",
                    ("Not Committed", 10, 3.0),
                )
                raise RuntimeError("Abort transaction")

        with self.assertRaisesRegex(RuntimeError, "Abort transaction"):
            abort_transaction()
        self.assertEqual(roster.get_all_students(), [])

    def test_database_context_closes_connection(self) -> None:
        with roster._database_connection() as connection:
            connection.execute("SELECT 1")
        with self.assertRaises(sqlite3.ProgrammingError):
            connection.execute("SELECT 1")

    def test_duplicate_cleanup_keeps_lowest_ids_and_distinct_records(self) -> None:
        first_id = roster.add_student("Alex", 10, None)
        roster.add_student("Alex", 10, None)
        other_grade_id = roster.add_student("Alex", 11, None)
        zero_gpa_id = roster.add_student("Alex", 10, 0.0)
        roster.add_student("Alex", 10, 0.0)
        self.assertEqual(roster.remove_duplicate_students(), 2)
        self.assertEqual(
            roster.get_all_students(),
            [
                (first_id, "Alex", 10, None),
                (other_grade_id, "Alex", 11, None),
                (zero_gpa_id, "Alex", 10, 0.0),
            ],
        )
        self.assertEqual(roster.remove_duplicate_students(), 0)

    def test_csv_preserves_comparison_details_and_quoted_names(self) -> None:
        before: list[roster.Student] = [
            (1, 'Zo\u00eb, "Z"', 10, 3.7),
            (2, "Removed Student", 9, None),
        ]
        after: list[roster.Student] = [(1, 'Zo\u00eb, "Z"', 10, 4.0)]
        report_path = roster.create_csv_report(before, after)
        with report_path.open(encoding="utf-8-sig", newline="") as file:
            rows = list(csv.DictReader(file))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Student Name"], 'Zo\u00eb, "Z"')
        self.assertEqual(rows[0]["GPA Before"], "3.70")
        self.assertEqual(rows[0]["GPA After"], "4.00")
        self.assertEqual(rows[0]["GPA Change"], "+0.30")
        self.assertEqual(rows[0]["Change Status"], "Updated")
        self.assertEqual(rows[1]["Change Status"], "Deleted")
        self.assertEqual(rows[1]["GPA Before"], "Not recorded")
        self.assertEqual(rows[1]["GPA After"], "")

    def test_main_exports_cleaned_roster_without_seeding_or_changing_grades(self) -> None:
        first_id = roster.add_student("Existing Student", 10, 3.7)
        duplicate_id = roster.add_student("Existing Student", 10, 3.7)
        for index in range(20):
            roster.add_student(
                f"Existing Student {index}", 9 + index % 4, None if index == 0 else 3.0
            )
        expected = [
            student for student in roster.get_all_students()
            if student[0] != duplicate_id
        ]

        for removed_count in (1, 0):
            with self.subTest(duplicates_removed=removed_count):
                with (
                    patch("os.startfile") as opener,
                    patch("builtins.input", return_value="5"),
                    redirect_stdout(io.StringIO()),
                ):
                    roster.main()
                self.assertEqual(roster.get_all_students(), expected)
                self.assertEqual(
                    opener.call_args_list,
                    [call(str(roster.REPORT_PATH)), call(str(roster.CSV_PATH))],
                )
                with roster.CSV_PATH.open(encoding="utf-8-sig", newline="") as file:
                    rows = list(csv.DictReader(file))
                self.assertEqual(len(rows), len(expected))
                self.assertTrue(all(row["Change Status"] == "Current" for row in rows))
                selected = next(row for row in rows if row["Student ID"] == str(first_id))
                self.assertEqual(selected["GPA After"], "3.70")
                document = Document(str(roster.REPORT_PATH))
                self.assertEqual(len(document.tables[0].rows), len(expected) + 1)
                self.assertEqual(len(document.element.xpath(".//wp:inline")), 4)
                text = "\n".join(paragraph.text for paragraph in document.paragraphs)
                self.assertIn(f"Duplicate records removed: {removed_count}.", text)
                self.assertIn("def add_student(", text)
                self.assertIn("def delete_student(", text)

    def test_csv_comparison_tracks_added_unchanged_and_missing_gpas(self) -> None:
        before: list[roster.Student] = [
            (1, "Old Name", 9, None),
            (2, "Unchanged Student", 11, 0.0),
            (3, "Missing GPA", 12, 3.2),
        ]
        after: list[roster.Student] = [
            (1, "New Name", 10, 3.5),
            (2, "Unchanged Student", 11, 0.0),
            (3, "Missing GPA", 12, None),
            (4, "Added Student", 9, None),
        ]
        report_path = roster.create_csv_report(before, after)
        with report_path.open(encoding="utf-8-sig", newline="") as file:
            rows = {row["Student ID"]: row for row in csv.DictReader(file)}
        self.assertEqual(len(rows), 4)
        self.assertEqual(
            rows["1"]["Change Details"],
            "Name changed from Old Name; Grade updated; GPA updated",
        )
        self.assertEqual(rows["1"]["GPA Change"], "")
        self.assertEqual(rows["2"]["Change Status"], "Unchanged")
        self.assertEqual(rows["2"]["GPA Change"], "+0.00")
        self.assertEqual(rows["3"]["GPA After"], "Not recorded")
        self.assertEqual(rows["3"]["GPA Change"], "")
        self.assertEqual(rows["4"]["Change Status"], "Added")
        self.assertEqual(rows["4"]["Grade Before"], "")
        self.assertEqual(rows["4"]["GPA After"], "Not recorded")

    def test_empty_roster_produces_reports_without_creating_students(self) -> None:
        with (
            patch("os.startfile"),
            patch("builtins.input", return_value="5"),
            redirect_stdout(io.StringIO()),
        ):
            roster.main()
        self.assertEqual(roster.get_all_students(), [])
        with roster.CSV_PATH.open(encoding="utf-8-sig", newline="") as file:
            self.assertEqual(list(csv.DictReader(file)), [])
        document = Document(str(roster.REPORT_PATH))
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("Average GPA: N/A (no recorded GPAs).", text)
        self.assertEqual(len(document.element.xpath(".//wp:inline")), 0)

    def test_menu_runs_all_crud_actions_until_quit(self) -> None:
        answers = [
            "1", "Alice", "10", "3.5",
            "2",
            "3", "1", "3.9",
            "2",
            "4", "1", "yes",
            "2",
            "5",
        ]
        output = io.StringIO()
        with patch("builtins.input", side_effect=answers), redirect_stdout(output):
            roster._run_menu()
        text = output.getvalue()
        self.assertIn("Added Alice with student ID 1.", text)
        self.assertIn("Updated student ID 1: GPA 3.90.", text)
        self.assertIn("3.90", text)
        self.assertIn("Deleted student ID 1.", text)
        self.assertIn("No students found.", text)
        self.assertEqual(text.count("Student Roster Manager"), 7)
        self.assertEqual(roster.get_all_students(), [])

    def test_menu_retries_invalid_inputs_and_reports_unknown_ids(self) -> None:
        answers = [
            "9",
            "1", " ", " Alice ", "bad", "10", "bad", "nan", "4.1", "3.5",
            "3", "999",
            "4", "bad", "999",
            "5",
        ]
        output = io.StringIO()
        with patch("builtins.input", side_effect=answers), redirect_stdout(output):
            roster._run_menu()
        text = output.getvalue()
        self.assertIn("Please select an option from 1 to 5.", text)
        self.assertIn("Student name cannot be empty.", text)
        self.assertIn("Please enter a whole number.", text)
        self.assertIn("Please enter a number, or leave the GPA blank.", text)
        self.assertIn("GPA must be between 0.0 and 4.0.", text)
        self.assertEqual(text.count("Student ID 999 was not found."), 2)
        self.assertEqual(roster.get_all_students(), [(1, "Alice", 10, 3.5)])

    def test_menu_accepts_missing_gpa_and_can_clear_an_existing_gpa(self) -> None:
        answers = [
            "1", "First Student", "10", "3.2",
            "3", "1", "",
            "1", "Second Student", "9", "",
            "5",
        ]
        with patch("builtins.input", side_effect=answers), redirect_stdout(io.StringIO()):
            roster._run_menu()
        self.assertEqual(
            roster.get_all_students(),
            [(1, "First Student", 10, None), (2, "Second Student", 9, None)],
        )

    def test_menu_reports_database_errors_and_continues(self) -> None:
        output = io.StringIO()
        with (
            patch("builtins.input", side_effect=["1", "2", "5"]),
            patch.object(
                roster,
                "_add_student_from_input",
                side_effect=sqlite3.OperationalError("database is locked"),
            ),
            redirect_stdout(output),
        ):
            roster._run_menu()
        self.assertIn("Database operation failed: database is locked", output.getvalue())
        self.assertIn("No students found.", output.getvalue())
        self.assertEqual(roster.get_all_students(), [])

    def test_main_exports_saved_roster_when_input_is_interrupted(self) -> None:
        student_id = roster.add_student("Existing Student", 10, 3.7)
        for interruption in (EOFError, KeyboardInterrupt):
            with self.subTest(interruption=interruption.__name__):
                with (
                    patch("builtins.input", side_effect=interruption),
                    patch("os.startfile"),
                    redirect_stdout(io.StringIO()),
                ):
                    roster.main()
                self.assertEqual(
                    roster.get_all_students(),
                    [(student_id, "Existing Student", 10, 3.7)],
                )
                self.assertTrue(roster.CSV_PATH.is_file())
                self.assertTrue(roster.REPORT_PATH.is_file())

    def test_menu_enforces_grade_id_and_gpa_boundaries(self) -> None:
        student_id = roster.add_student("Existing Student", 10, 3.5)
        answers = [
            "1", "New Student", "0", "13", "12", "nan", "-0.1", "4.1", "4.0",
            "3", "0", "-1", str(2**63), str(student_id), "0.0",
            "5",
        ]
        output = io.StringIO()
        with patch("builtins.input", side_effect=answers), redirect_stdout(output):
            roster._run_menu()
        self.assertEqual(
            roster.get_all_students(),
            [(student_id, "Existing Student", 10, 0.0), (2, "New Student", 12, 4.0)],
        )
        self.assertIn("Enter a whole number of at least 1.", output.getvalue())
        self.assertIn("Enter a whole number no greater than 12.", output.getvalue())
        self.assertIn("GPA must be between 0.0 and 4.0.", output.getvalue())

    def test_cancel_at_any_operation_prompt_preserves_the_roster(self) -> None:
        roster.add_student("Existing Student", 10, 3.5)
        expected = roster.get_all_students()
        operations = [
            ["1", " CANCEL "],
            ["1", "New Student", "cancel"],
            ["1", "New Student", "10", "cancel"],
            ["3", "cancel"],
            ["3", "1", "cancel"],
            ["4", "cancel"],
            ["4", "1", "cancel"],
        ]
        for operation in operations:
            with self.subTest(operation=operation):
                output = io.StringIO()
                with (
                    patch("builtins.input", side_effect=[*operation, "5"]),
                    redirect_stdout(output),
                ):
                    roster._run_menu()
                self.assertEqual(roster.get_all_students(), expected)
                self.assertIn(
                    "[INFO] Operation cancelled. No changes were made.", output.getvalue()
                )
                self.assertEqual(output.getvalue().count("Student Roster Manager"), 2)

    def test_deletion_displays_student_details_and_requires_confirmation(self) -> None:
        roster.add_student("Existing Student", 10, 3.5)
        expected = roster.get_all_students()
        output = io.StringIO()
        with (
            patch("builtins.input", side_effect=["4", "1", "maybe", "no", "5"]),
            redirect_stdout(output),
        ):
            roster._run_menu()
        self.assertEqual(roster.get_all_students(), expected)
        text = output.getvalue()
        self.assertIn("Selected Student", text)
        self.assertIn("Name:  Existing Student", text)
        self.assertIn("Grade: 10", text)
        self.assertIn("GPA:   3.50", text)
        self.assertIn("Please enter y or n.", text)
        self.assertIn("Deletion cancelled. No changes were made.", text)

    def test_roster_table_has_aligned_columns_and_excludes_missing_gpas_from_average(self) -> None:
        roster.add_student("A Student With A Name That Exceeds The Display Width", 10, 3.5)
        roster.add_student("Missing GPA", 11, None)
        roster.add_student("Zero GPA", 12, 0.0)
        expected = roster.get_all_students()
        output = io.StringIO()
        with redirect_stdout(output):
            roster._view_students()
        text = output.getvalue()
        table_lines = [line for line in text.splitlines() if line.startswith("|")]
        self.assertEqual(len(table_lines), 4)
        self.assertEqual(len({len(line) for line in table_lines}), 1)
        self.assertIn("...", text)
        self.assertIn("N/A", text)
        self.assertIn("Students: 3 | Recorded GPAs: 2 | Average GPA: 1.75", text)
        self.assertEqual(roster.get_all_students(), expected)

    def test_view_all_displays_every_student_and_opens_both_complete_reports(self) -> None:
        for index in range(25):
            roster.add_student(f"Student {index}", 9 + index % 4, 3.5)
        expected = roster.get_all_students()
        expected_ids = {student[0] for student in expected}
        output = io.StringIO()
        with redirect_stdout(output):
            roster._view_students()
        displayed_rows = [
            line for line in output.getvalue().splitlines()
            if line.startswith("|") and "Student Name" not in line
        ]
        self.assertEqual(len(displayed_rows), 25)
        self.assertEqual(
            {int(line.split("|")[1].strip()) for line in displayed_rows}, expected_ids
        )
        with roster.CSV_PATH.open(encoding="utf-8-sig", newline="") as file:
            csv_rows = list(csv.DictReader(file))
        self.assertEqual(len(csv_rows), 25)
        self.assertEqual({int(row["Student ID"]) for row in csv_rows}, expected_ids)
        document = Document(str(roster.REPORT_PATH))
        self.assertEqual(len(document.tables[0].rows), 26)
        self.assertEqual(
            {int(row.cells[0].text) for row in document.tables[0].rows[1:]}, expected_ids
        )
        self.assertEqual(len(document.element.xpath(".//wp:inline")), 4)
        self.assertEqual(
            self.open_reports.call_args_list,
            [call(str(roster.REPORT_PATH)), call(str(roster.CSV_PATH))],
        )
        self.assertEqual(roster.get_all_students(), expected)

    def test_view_empty_roster_still_opens_both_reports(self) -> None:
        output = io.StringIO()
        with redirect_stdout(output):
            roster._view_students()
        self.assertIn("No students found.", output.getvalue())
        with roster.CSV_PATH.open(encoding="utf-8-sig", newline="") as file:
            self.assertEqual(list(csv.DictReader(file)), [])
        document = Document(str(roster.REPORT_PATH))
        self.assertEqual(len(document.tables[0].rows), 1)
        self.assertEqual(
            self.open_reports.call_args_list,
            [call(str(roster.REPORT_PATH)), call(str(roster.CSV_PATH))],
        )

    def test_view_reports_file_errors_and_returns_to_the_menu(self) -> None:
        roster.add_student("Existing Student", 10, 3.5)
        expected = roster.get_all_students()
        output = io.StringIO()
        with (
            patch("builtins.input", side_effect=["2", "5"]),
            patch.object(
                roster, "create_csv_report", side_effect=PermissionError("report is open")
            ),
            redirect_stdout(output),
        ):
            roster._run_menu()
        self.assertIn(
            "[ERROR] Reports could not be saved or opened: report is open",
            output.getvalue(),
        )
        self.assertEqual(output.getvalue().count("Student Roster Manager"), 2)
        self.assertEqual(roster.get_all_students(), expected)
        self.open_reports.assert_not_called()


if __name__ == "__main__":
    unittest.main()
