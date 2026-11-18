"""Manage students through an interactive SQLite-backed command-line menu."""

import csv
import inspect
import os
import sqlite3
from collections.abc import Generator
from contextlib import closing, contextmanager
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.document import Document as WordDocument
from docx.shared import Inches, Pt
from matplotlib.figure import Figure


DATABASE_PATH = Path(__file__).resolve().with_name("school.db")
REPORT_PATH = Path(__file__).resolve().with_name("student_roster_report.docx")
CSV_PATH = Path(__file__).resolve().with_name("student_roster_results.csv")
STUDENTS_PER_CHART = 20
MAX_STUDENT_ID = 2**63 - 1
MENU_TEXT = """
========================================================
                 Student Roster Manager
========================================================
  1. Add a student
  2. View all students
  3. Update a student's GPA
  4. Delete a student
  5. Quit
--------------------------------------------------------
  Type 'cancel' during an operation to return to the menu.
"""
Student = tuple[int, str, int, float | None]


class _OperationCancelled(Exception):
    """Return to the menu without committing an unfinished operation."""


@contextmanager
def _database_connection() -> Generator[sqlite3.Connection, None, None]:
    """Commit successful operations, roll back failures, and always close."""
    with closing(sqlite3.connect(DATABASE_PATH)) as connection, connection:
        yield connection


def initialize_database() -> None:
    """Create the students table without changing existing records."""
    with _database_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                grade INTEGER NOT NULL,
                gpa REAL
            )
            """
        )


def add_student(name: str, grade: int, gpa: float | None) -> int:
    """Insert a student and return the generated ID."""
    with _database_connection() as connection:
        cursor = connection.execute(
            "INSERT INTO students (name, grade, gpa) VALUES (?, ?, ?)",
            (name, grade, gpa),
        )
        student_id = cursor.lastrowid
        if student_id is None:
            raise RuntimeError("The new student did not receive an ID.")
        return student_id


def get_all_students() -> list[Student]:
    """Return all students ordered by ID."""
    with _database_connection() as connection:
        return connection.execute(
            "SELECT id, name, grade, gpa FROM students ORDER BY id"
        ).fetchall()


def get_student_by_id(student_id: int) -> Student | None:
    """Return a student, or None if the ID does not exist."""
    with _database_connection() as connection:
        return connection.execute(
            "SELECT id, name, grade, gpa FROM students WHERE id = ?",
            (student_id,),
        ).fetchone()


def update_student_gpa(student_id: int, new_gpa: float | None) -> bool:
    """Update a GPA; return False if no student matches the ID."""
    with _database_connection() as connection:
        cursor = connection.execute(
            "UPDATE students SET gpa = ? WHERE id = ?",
            (new_gpa, student_id),
        )
        return cursor.rowcount == 1


def delete_student(student_id: int) -> bool:
    """Delete a student; return False if no student matches the ID."""
    with _database_connection() as connection:
        cursor = connection.execute(
            "DELETE FROM students WHERE id = ?", (student_id,)
        )
        return cursor.rowcount == 1


def remove_duplicate_students() -> int:
    """Remove identical name/grade/GPA records, keeping the lowest student ID."""
    with _database_connection() as connection:
        cursor = connection.execute(
            """
            DELETE FROM students
            WHERE id NOT IN (
                SELECT MIN(id)
                FROM students
                GROUP BY name, grade, gpa
            )
            """
        )
        return cursor.rowcount


def _describe_student_change(
    before: Student | None,
    after: Student | None,
    *,
    has_previous_snapshot: bool,
) -> tuple[str, str]:
    if not has_previous_snapshot:
        return "Current", "Current database snapshot; no prior snapshot provided"
    if before is None:
        return "Added", "Added after the first snapshot"
    if after is None:
        return "Deleted", "Removed from the final roster"

    changes = []
    if before[1] != after[1]:
        changes.append(f"Name changed from {before[1]}")
    if before[2] != after[2]:
        changes.append("Grade updated")
    if before[3] != after[3]:
        changes.append("GPA updated")
    if changes:
        return "Updated", "; ".join(changes)
    return "Unchanged", "No changes"


def _format_snapshot_gpa(snapshot: Student | None) -> str:
    if snapshot is None:
        return ""
    if snapshot[3] is None:
        return "Not recorded"
    return f"{snapshot[3]:.2f}"


def _gpa_change(before: Student | None, after: Student | None) -> str:
    if before is None or after is None:
        return ""
    if before[3] is None or after[3] is None:
        return ""
    return f"{after[3] - before[3]:+.2f}"


def create_csv_report(
    before_students: list[Student] | None, after_students: list[Student]
) -> Path:
    """Export sorted comparison rows, or a current-only roster when before is None."""
    before_by_id = {student[0]: student for student in (before_students or [])}
    after_by_id = {student[0]: student for student in after_students}
    combined = {**before_by_id, **after_by_id}
    students = sorted(
        combined.values(),
        key=lambda student: (
            student[0] not in after_by_id,
            student[2],
            student[1].casefold(),
            student[0],
        ),
    )
    with CSV_PATH.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                "Student ID",
                "Student Name",
                "Change Status",
                "Grade Before",
                "Grade After",
                "GPA Before",
                "GPA After",
                "GPA Change",
                "Change Details",
            )
        )
        for student_id, name, _, _ in students:
            before = before_by_id.get(student_id)
            after = after_by_id.get(student_id)
            status, details = _describe_student_change(
                before, after, has_previous_snapshot=before_students is not None
            )
            writer.writerow(
                (
                    student_id,
                    name,
                    status,
                    before[2] if before is not None else "",
                    after[2] if after is not None else "",
                    _format_snapshot_gpa(before),
                    _format_snapshot_gpa(after),
                    _gpa_change(before, after),
                    details,
                )
            )
    return CSV_PATH


def _add_roster_table(document: WordDocument, students: list[Student]) -> None:
    document.add_heading("Current student roster", level=1)
    table = document.add_table(rows=1, cols=4)
    table.style = "Light Shading Accent 1"
    for cell, label in zip(table.rows[0].cells, ("ID", "Name", "Grade", "GPA")):
        cell.text = label
    for student in students:
        for cell, value in zip(table.add_row().cells, student):
            cell.text = "N/A" if value is None else str(value)


def _add_code_explanations(document: WordDocument) -> None:
    document.add_page_break()
    document.add_heading("How the code works", level=1)
    document.add_paragraph(
        "sqlite3 connects to school.db beside the script. Path locates the files. "
        "Student describes each returned tuple: (id, name, grade, gpa). "
        "_database_connection() centralizes connection handling. closing() closes "
        "every connection, while 'with connection' commits successful writes and "
        "rolls back failed writes. SQL '?' placeholders "
        "bind values safely instead of inserting user input into SQL text. "
        "Database errors propagate rather than being hidden."
    )
    explanations = (
        (
            _database_connection,
            "A context manager shared by all database operations. It yields an open "
            "SQLite connection, commits on success, rolls back on an exception, and "
            "always closes the connection. Exceptions propagate to the caller.",
        ),
        (
            initialize_database,
            "Creates the table only if it does not exist. The ID is an automatically "
            "incremented primary key. Name and grade cannot be NULL; GPA may be NULL.",
        ),
        (
            add_student,
            "CREATE: INSERT stores the name, grade, and GPA. lastrowid returns the new "
            "student's ID; a missing generated ID raises an error.",
        ),
        (
            get_all_students,
            "READ: SELECT retrieves all four columns ordered by ID. fetchall() returns "
            "a list of student tuples, or an empty list when there are no students.",
        ),
        (
            get_student_by_id,
            "READ: WHERE id = ? limits the lookup to one student. fetchone() returns "
            "that student's tuple, or None when the ID is not found.",
        ),
        (
            update_student_gpa,
            "UPDATE: Changes only the GPA for the matching ID. rowcount == 1 returns "
            "True for a successful match and False when no student matches.",
        ),
        (
            delete_student,
            "DELETE: Removes only the row with the matching ID. The return value "
            "indicates whether a student was actually deleted.",
        ),
        (
            remove_duplicate_students,
            "Groups students by identical name, grade, and GPA, ignoring the ID. "
            "MIN(id) keeps the oldest ID in each group; other records are deleted "
            "in one transaction. Missing GPAs group together. Names are matched "
            "exactly, including case and whitespace. The return value counts the "
            "deleted duplicates. Same-name records with different grades or GPAs remain.",
        ),
    )
    for function, explanation in explanations:
        document.add_heading(function.__name__, level=2)
        document.add_paragraph(explanation)
        code = document.add_paragraph().add_run(inspect.getsource(function))
        code.font.name = "Consolas"
        code.font.size = Pt(8)

    document.add_heading("Interactive menu and reporting", level=2)
    document.add_paragraph(
        "The __main__ guard calls main() only when this file is executed directly. It "
        "ensures the students table exists, removes duplicate database records, "
        "and runs a menu for adding, viewing, updating, and deleting students. "
        "Input helpers require nonempty names, grades from 1 to 12, positive "
        "student IDs within SQLite's integer range, and GPAs from 0.0 to 4.0. "
        "A blank GPA means no recorded GPA. Typing 'cancel' at any operation "
        "prompt returns to the menu without saving that operation. Selected "
        "student details are displayed before an update or deletion, and "
        "deletion requires an explicit yes/no confirmation. "
        "View all students prints the entire roster and immediately generates "
        "and opens both Word and CSV reports, including when the roster is empty. "
        "Each successful change is committed immediately. On exit, exact duplicates "
        "are removed again and the current roster is exported to Word and CSV. "
        "Both reports and their charts use the same cleaned roster. No sample or "
        "random students are added automatically."
    )
    document.add_paragraph(
        "create_report() uses python-docx to build this document, including summary "
        "statistics, roster tables, and explanations with the actual CRUD source "
        "code obtained through inspect. Matplotlib builds horizontal bar charts "
        "in batches of 20 students. BytesIO holds each PNG in memory for embedding, "
        "so no temporary image files are needed. Missing GPAs are marked N/A and "
        "excluded from the GPA average. The document is saved before os.startfile() "
        "opens it in the default Windows Word-document application."
    )
    document.add_paragraph(
        "create_csv_report() uses Python's csv module to compare both roster "
        "snapshots in student_roster_results.csv, with one row per student. "
        "Columns show ID, name, change status, before/after grades and GPAs, "
        "signed GPA change, and change details. Current students are sorted by "
        "grade, name, and ID, with deleted students last. GPAs use two decimals. "
        "Small helpers classify changes and format snapshot GPAs and GPA differences. "
        "'Not recorded' indicates a missing GPA; blank before/after values mean "
        "the student was absent from that snapshot. 'Added' means added between "
        "snapshots. "
        "Passing None for the before snapshot exports only the current roster, "
        "marked 'Current', without claiming historical changes. This is how the "
        "application generates the CSV for View all students and when the menu closes. "
        "newline='' prevents extra blank lines on Windows, "
        "and UTF-8 with a byte-order mark supports names in spreadsheet applications. "
        "The CSV is replaced each run and opened with os.startfile()."
    )

def _add_student_charts(document: WordDocument, students: list[Student]) -> None:
    for start in range(0, len(students), STUDENTS_PER_CHART):
        batch = students[start : start + STUDENTS_PER_CHART]
        labels = [f"{student[0]}: {student[1]}" for student in batch]
        for title, column, maximum in (("Grade", 2, 12), ("GPA", 3, 4)):
            document.add_page_break()
            document.add_heading(f"Current roster: {title} by student", level=1)
            figure = Figure(figsize=(9, max(3, len(batch) * 0.3)), layout="constrained")
            axis = figure.subplots()
            values = [student[column] for student in batch]
            bars = axis.barh(
                labels, [value if value is not None else 0 for value in values]
            )
            axis.bar_label(
                bars,
                labels=["N/A" if value is None else str(value) for value in values],
                padding=3,
            )
            axis.invert_yaxis()
            axis.set_xlim(0, maximum * 1.15)
            if column == 2:
                axis.set_xticks(range(0, 13))
            axis.set_xlabel(title)
            with BytesIO() as image:
                figure.savefig(image, format="png", dpi=150)
                image.seek(0)
                document.add_picture(image, width=Inches(6))
            figure.clear()

def create_report(students: list[Student], duplicates_removed: int = 0) -> Path:
    """Save the current roster, CRUD explanations, and charts to Word."""
    document = Document()
    document.add_heading("Student Roster Report", 0)
    document.add_paragraph(f"Database: {DATABASE_PATH}")
    document.add_heading("Run summary", level=1)
    document.add_paragraph(
        f"Current students: {len(students)}. "
        f"Duplicate records removed: {duplicates_removed}. "
        "Reports are replaced each run and show the saved roster, including "
        "changes made through the interactive menu."
    )
    gpas = [student[3] for student in students if student[3] is not None]
    document.add_paragraph(
        f"Average GPA: {sum(gpas) / len(gpas):.2f}"
        if gpas
        else "Average GPA: N/A (no recorded GPAs)."
    )
    _add_roster_table(document, students)
    _add_code_explanations(document)
    _add_student_charts(document, students)
    document.save(str(REPORT_PATH))
    return REPORT_PATH


def _read_input(prompt: str) -> str:
    value = input(prompt).strip()
    if value.casefold() == "cancel":
        raise _OperationCancelled
    return value


def _read_integer(
    prompt: str, *, minimum: int = 1, maximum: int = MAX_STUDENT_ID
) -> int:
    while True:
        try:
            value = int(_read_input(prompt))
        except ValueError:
            print("[ERROR] Please enter a whole number.")
            continue
        if value < minimum:
            print(f"[ERROR] Enter a whole number of at least {minimum}.")
            continue
        if value > maximum:
            print(f"[ERROR] Enter a whole number no greater than {maximum}.")
            continue
        return value


def _read_gpa(prompt: str) -> float | None:
    while True:
        value = _read_input(prompt)
        if not value:
            return None
        try:
            gpa = float(value)
        except ValueError:
            print("[ERROR] Please enter a number, or leave the GPA blank.")
            continue
        if 0.0 <= gpa <= 4.0:
            return gpa
        print("[ERROR] GPA must be between 0.0 and 4.0.")


def _add_student_from_input() -> None:
    name = _read_input("Student name: ")
    while not name:
        print("[ERROR] Student name cannot be empty.")
        name = _read_input("Student name: ")
    grade = _read_integer("Grade (1-12): ", maximum=12)
    gpa = _read_gpa("GPA (0.0-4.0, blank if not recorded): ")
    student_id = add_student(name, grade, gpa)
    print(f"[SUCCESS] Added {name} with student ID {student_id}.")


def _print_students(current_students: list[Student]) -> None:
    if not current_students:
        print("[INFO] No students found.")
        return
    id_width = max([len("ID"), *(len(str(row[0])) for row in current_students)])
    name_width = min(
        32, max([len("Student Name"), *(len(row[1]) for row in current_students)])
    )
    separator = (
        f"+{'-' * (id_width + 2)}+{'-' * (name_width + 2)}+-------+-------+"
    )
    print(f"\nCurrent roster: {len(current_students)} students")
    print(separator)
    print(
        f"| {'ID':>{id_width}} | {'Student Name':<{name_width}} "
        f"| {'Grade':>5} | {'GPA':>5} |"
    )
    print(separator)
    for student_id, name, grade, gpa in current_students:
        display_name = name if len(name) <= 32 else f"{name[:29]}..."
        gpa_text = "N/A" if gpa is None else f"{gpa:.2f}"
        print(
            f"| {student_id:>{id_width}} | {display_name:<{name_width}} "
            f"| {grade:>5} | {gpa_text:>5} |"
        )
    print(separator)
    gpas = [student[3] for student in current_students if student[3] is not None]
    average = f"{sum(gpas) / len(gpas):.2f}" if gpas else "N/A"
    print(
        f"[INFO] Students: {len(current_students)} | "
        f"Recorded GPAs: {len(gpas)} | Average GPA: {average}"
    )


def _view_students() -> None:
    """Display every student and open reports of the same roster snapshot."""
    current_students = get_all_students()
    _print_students(current_students)
    try:
        _export_and_open_reports(current_students)
    except OSError as error:
        print(f"[ERROR] Reports could not be saved or opened: {error}")


def _select_student(prompt: str) -> Student | None:
    student_id = _read_integer(prompt)
    student = get_student_by_id(student_id)
    if student is None:
        print(f"[ERROR] Student ID {student_id} was not found.")
        return None
    _, name, grade, gpa = student
    gpa_text = "N/A" if gpa is None else f"{gpa:.2f}"
    print(
        f"\nSelected Student\n"
        f"  ID:    {student_id}\n"
        f"  Name:  {name}\n"
        f"  Grade: {grade}\n"
        f"  GPA:   {gpa_text}"
    )
    return student


def _update_student_from_input() -> None:
    student = _select_student("Student ID to update: ")
    if student is None:
        return
    student_id = student[0]
    gpa = _read_gpa("New GPA (0.0-4.0, blank to clear): ")
    if update_student_gpa(student_id, gpa):
        gpa_text = "not recorded" if gpa is None else f"{gpa:.2f}"
        print(f"[SUCCESS] Updated student ID {student_id}: GPA {gpa_text}.")
    else:
        print("[ERROR] The student no longer exists. No GPA was changed.")


def _delete_student_from_input() -> None:
    student = _select_student("Student ID to delete: ")
    if student is None:
        return
    confirmation = _read_input("Delete this student? (y/n): ").casefold()
    while confirmation not in {"y", "yes", "n", "no"}:
        print("[ERROR] Please enter y or n.")
        confirmation = _read_input("Delete this student? (y/n): ").casefold()
    if confirmation in {"n", "no"}:
        print("[INFO] Deletion cancelled. No changes were made.")
        return
    student_id = student[0]
    if delete_student(student_id):
        print(f"[SUCCESS] Deleted student ID {student_id}.")
    else:
        print("[ERROR] The student no longer exists. Nothing was deleted.")


def _run_menu() -> None:
    actions = {
        "1": _add_student_from_input,
        "2": _view_students,
        "3": _update_student_from_input,
        "4": _delete_student_from_input,
    }
    while True:
        print(MENU_TEXT)
        choice = input("Choose an option [1-5]: ").strip()
        if choice == "5":
            return
        action = actions.get(choice)
        if action is None:
            print("[ERROR] Please select an option from 1 to 5.")
            continue
        try:
            action()
        except _OperationCancelled:
            print("[INFO] Operation cancelled. No changes were made.")
        except sqlite3.Error as error:
            print(f"[ERROR] Database operation failed: {error}")


def _export_and_open_reports(
    students: list[Student], duplicates_removed: int = 0
) -> None:
    csv_path = create_csv_report(None, students)
    print(f"\nCSV results saved to: {csv_path}")
    report_path = create_report(students, duplicates_removed)
    print(f"\nWord report saved to: {report_path}")
    os.startfile(str(report_path))
    os.startfile(str(csv_path))


def main() -> None:
    """Run the CRUD menu, then export and open reports of the saved roster."""
    initialize_database()
    duplicates_removed = remove_duplicate_students()
    try:
        _run_menu()
    except (EOFError, KeyboardInterrupt):
        print("\nInput ended. Exporting the saved roster.")
    duplicates_removed += remove_duplicate_students()
    print(f"\nDuplicate records removed: {duplicates_removed}")
    _export_and_open_reports(get_all_students(), duplicates_removed)
    print("Goodbye!")


if __name__ == "__main__":
    main()
