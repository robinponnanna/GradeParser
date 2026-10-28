import os
import re
import sys
import shutil
import platform
import subprocess
from collections import defaultdict
import pymupdf
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def select_pdf_files() -> list[str]:
    """
    Open the native file application manager depending on the operating system
    (macOS, Windows, Linux with GNOME) to let the user select one or more PDF files.
    """
    current_os = platform.system()
    selected_files = []

    print(f"Opening native file manager for {current_os}...")

    if current_os == "Linux":
        # GNOME uses zenity for native GTK file selection dialogs
        if shutil.which("zenity"):
            try:
                result = subprocess.run(
                    [
                        "zenity",
                        "--file-selection",
                        "--multiple",
                        "--separator=|",
                        "--file-filter=PDF Files (pdf) | *.pdf",
                        "--title=Select PDF Grade Card(s)"
                    ],
                    capture_output=True,
                    text=True
                )
                if result.returncode == 0 and result.stdout.strip():
                    selected_files = [f.strip() for f in result.stdout.strip().split("|") if f.strip()]
            except Exception as err:
                print(f"Zenity dialog error: {err}")

    elif current_os == "Darwin":  # macOS
        # Native macOS NSOpenPanel via AppleScript
        script = """
        set selectedFiles to choose file with prompt "Select PDF Grade Card(s)" of type {"pdf"} with multiple selections allowed
        set posixPaths to {}
        repeat with aFile in selectedFiles
            copy POSIX path of aFile to end of posixPaths
        end repeat
        set AppleScript's text item delimiters to linefeed
        return posixPaths as text
        """
        try:
            result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                selected_files = [f.strip() for f in result.stdout.strip().splitlines() if f.strip()]
        except Exception as err:
            print(f"macOS dialog error: {err}")

    elif current_os == "Windows":
        # Native Windows Common File Dialog via PowerShell
        ps_command = (
            "[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null; "
            "$dialog = New-Object System.Windows.Forms.OpenFileDialog; "
            "$dialog.Filter = 'PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*'; "
            "$dialog.Multiselect = $true; "
            "$dialog.Title = 'Select PDF Grade Card(s)'; "
            "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { "
            "$dialog.FileNames -join '|' "
            "}"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_command],
                capture_output=True,
                text=True
            )
            if result.returncode == 0 and result.stdout.strip():
                selected_files = [f.strip() for f in result.stdout.strip().split("|") if f.strip()]
        except Exception as err:
            print(f"Windows dialog error: {err}")

    # Fallback: Tkinter file dialog if native tool did not return files
    if not selected_files:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            files = filedialog.askopenfilenames(
                title="Select PDF Grade Card(s)",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            root.destroy()
            if files:
                selected_files = list(files)
        except Exception:
            pass

    return selected_files


def normalize_semester(raw_semester: str) -> str:
    """Normalize semester names (e.g. 'First Semester' -> 'Semester 1')."""
    s = raw_semester.lower().strip()
    mapping = {
        "first": "Semester 1", "1st": "Semester 1", "i": "Semester 1",
        "second": "Semester 2", "2nd": "Semester 2", "ii": "Semester 2",
        "third": "Semester 3", "3rd": "Semester 3", "iii": "Semester 3",
        "fourth": "Semester 4", "4th": "Semester 4", "iv": "Semester 4",
        "fifth": "Semester 5", "5th": "Semester 5", "v": "Semester 5",
        "sixth": "Semester 6", "6th": "Semester 6", "vi": "Semester 6",
        "seventh": "Semester 7", "7th": "Semester 7", "vii": "Semester 7",
        "eighth": "Semester 8", "8th": "Semester 8", "viii": "Semester 8",
    }
    for key, val in mapping.items():
        if re.search(rf"\b{key}\b", s):
            return val
    return raw_semester.title()


def extract_semester_info(full_text: str, pdf_filename: str) -> str:
    """Extract semester details from PDF text or fallback to filename."""
    sem_match = re.search(
        r"(First|Second|Third|Fourth|Fifth|Sixth|Seventh|Eighth|I{1,3}|IV|V|VI{0,3}|VII|VIII|\d+(?:st|nd|rd|th)?)\s*Semester",
        full_text,
        re.IGNORECASE
    )
    if sem_match:
        return normalize_semester(sem_match.group(0))

    base = os.path.basename(pdf_filename)
    fn_match = re.search(r"(\d+(?:st|nd|rd|th)?\s*sem)", base, re.IGNORECASE)
    if fn_match:
        return normalize_semester(fn_match.group(1))

    return "Unknown Semester"


def extract_result_date(full_text: str) -> str:
    """Extract result publication date from PDF text."""
    date_match = re.search(r"Result Date:\s*([0-9\/\-]+)", full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r"Date:\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})", full_text)
    return date_match.group(1).strip() if date_match else "N/A"


def extract_courses_from_pdf(pdf_path: str) -> list[dict]:
    """
    Extract course details (Course Code, Course Name, Credits, Grade)
    along with semester and result date from a given PDF using PyMuPDF.
    """
    doc = pymupdf.open(pdf_path)
    full_text = "\n".join(page.get_text() for page in doc)

    semester = extract_semester_info(full_text, pdf_path)
    result_date = extract_result_date(full_text)

    records = []
    course_code_regex = re.compile(r"^[A-Z]{3,4}\d{4}$")
    valid_grade_regex = re.compile(r"^(?:O|A\+?|B\+?|C|D|F|S|U|I|NE|NP|RC|NC)$", re.IGNORECASE)

    for page in doc:
        words = page.get_text("words")
        if not words:
            continue

        # Group words by their vertical position (y-coordinate)
        lines_by_y = defaultdict(list)
        for w in words:
            y_mid = (w[1] + w[3]) / 2
            matched_key = None
            for key in lines_by_y:
                if abs(key - y_mid) < 4:
                    matched_key = key
                    break
            if matched_key is None:
                matched_key = y_mid
            lines_by_y[matched_key].append(w)

        # Sort lines from top to bottom
        for y in sorted(lines_by_y.keys()):
            # Sort words horizontally from left to right
            line_words = sorted(lines_by_y[y], key=lambda x: x[0])
            tokens = [x[4].strip() for x in line_words if x[4].strip()]

            # Look for course code token
            code_idx = -1
            for idx, token in enumerate(tokens):
                if course_code_regex.match(token):
                    code_idx = idx
                    break

            # A valid course row must have: Course Code, Course Name..., Credits, Grade
            if code_idx != -1 and len(tokens) >= code_idx + 4:
                course_code = tokens[code_idx]
                grade = tokens[-1]
                credits = tokens[-2]
                course_name = " ".join(tokens[code_idx + 1 : -2])

                # Validate credits (numeric) and grade
                if credits.isdigit() and valid_grade_regex.match(grade):
                    records.append({
                        "Semester": semester,
                        "Result Date": result_date,
                        "Course Code": course_code,
                        "Course Name": course_name,
                        "No. of Credits": int(credits),
                        "Grade": grade
                    })

    return records


def format_excel(excel_path: str):
    """Apply professional styling and auto-fit column widths to the generated Excel file."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # Style header row
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Style data rows
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if col == 4:  # Course Name column
                cell.alignment = left_align
            else:
                cell.alignment = center_align

    # Set column widths with extra padding
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Freeze header row
    ws.freeze_panes = "A2"
    wb.save(excel_path)


def main():
    # 1. Select PDF files via OS-native file dialog
    pdf_files = select_pdf_files()

    if not pdf_files:
        print("No files selected from the file manager.")
        # Fallback to Files/ directory if user cancelled or no GUI available
        default_folder = "Files"
        if os.path.exists(default_folder):
            fallback_files = sorted(
                os.path.join(default_folder, f)
                for f in os.listdir(default_folder)
                if f.lower().endswith(".pdf")
            )
            if fallback_files:
                use_fallback = input(
                    f"Would you like to process {len(fallback_files)} PDF(s) from '{default_folder}/'? (y/n): "
                ).strip().lower()
                if use_fallback in ("y", "yes"):
                    pdf_files = fallback_files

    if not pdf_files:
        print("Exiting: No PDF files to process.")
        return

    # 2. Ask user for output file name in the console (excluding .xlsx)
    output_input = input("\nEnter output Excel file name (excluding .xlsx): ").strip()
    if not output_input:
        output_input = "student_grades"
    elif output_input.lower().endswith(".xlsx"):
        output_input = output_input[:-5]

    output_excel = f"{output_input}.xlsx"

    # 3. Sort files to keep semester order where possible
    pdf_files = sorted(pdf_files)
    print(f"\nProcessing {len(pdf_files)} PDF file(s)...")

    all_records = []
    for pdf_file in pdf_files:
        records = extract_courses_from_pdf(pdf_file)
        print(f"  - {os.path.basename(pdf_file)}: extracted {len(records)} courses")
        all_records.extend(records)

    if not all_records:
        print("No course records extracted from the selected files.")
        return

    # 4. Save to Excel using pandas
    df = pd.DataFrame(all_records)
    df.to_excel(output_excel, index=False)

    # 5. Format the Excel workbook
    format_excel(output_excel)

    print(f"\nSuccessfully extracted {len(df)} course records.")
    print(f"Saved and formatted Excel file at: {output_excel}\n")
    print("Summary of extracted courses:")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()