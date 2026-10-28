# Grade Extractor (GradeParser)

A Python tool that extracts student grade details from university semester grade card PDFs and compiles them into a structured, styled Excel (`.xlsx`) sheet using **PyMuPDF**, **Pandas**, and **OpenPyXL**.

---

## Features

- **Native File Dialog**: Automatically launches your operating system's native file picker (macOS, Windows, Linux GNOME) to let you select one or more PDF grade sheets.
- **Robust Table Extraction**: Uses spatial coordinate grouping (`PyMuPDF`) to handle both bordered tables and floating text blocks without losing spaces in course names.
- **Key Fields Extracted**:
  - **Semester**: Normalized semester name (e.g., `Semester 1`, `Semester 2`, `Semester 3`, `Semester 4`).
  - **Result Date**: Extracted result publication date for each semester.
  - **Course Code**: e.g., `CSE1004`, `MAT1003`, `CCS2504`.
  - **Course Name**: Full course title without truncation.
  - **No. of Credits**: Numeric credits integer for easy summation.
  - **Grade**: Letter grade (e.g., `O`, `A+`, `A`, `B+`, `B`, `S`).
- **Custom Output Filename**: Prompts for a custom file name in the console (excluding `.xlsx`).
- **Formatted Excel Spreadsheets**: Auto-fits column widths, applies a modern dark blue header theme, centers text, and freezes top header row.

---

## How It Works

1. **OS Detection & File Selection**:
   - **Linux (GNOME)**: Uses `zenity` for native GTK file selection dialogs.
   - **macOS**: Uses AppleScript (`osascript`) calling the native `choose file` dialog (`NSOpenPanel`).
   - **Windows**: Uses PowerShell calling `System.Windows.Forms.OpenFileDialog`.
   - **Fallback**: Uses Python's `tkinter` or defaults to the local `Files/` directory.

2. **Text & Metadata Extraction**:
   - Parses the document text to find the **Semester** (matching roman numerals or ordinal numbers) and the **Result Date**.
   - Groups word bounding boxes horizontally by their vertical position (`y`-coordinate), sorting each line from left to right (`x`-coordinate).

3. **Pattern Matching & Validation**:
   - Identifies course codes matching university code patterns (e.g., `CSE1004`).
   - Parses trailing tokens as `Credits` and `Grade`.
   - Reconstructs intermediate words into the complete `Course Name`.

4. **Excel Export & Styling**:
   - Assembles records across all selected PDFs into a unified `pandas.DataFrame`.
   - Writes to `.xlsx` using `openpyxl` with custom styling (header fills, borders, cell alignments, auto-adjusted column widths).

---

## Installation

### 1. Prerequisites
- **Python 3.9+**
- **OS-specific Dialog Tools**:
  - **Linux**: `zenity` (preinstalled on GNOME; install via `sudo apt install zenity` or `sudo dnf install zenity` or `sudo pacman -S zenity` if missing).
  - **macOS**: Built-in `osascript`.
  - **Windows**: Built-in PowerShell.

### 2. Clone the Repository
```bash
git clone https://github.com/robinponnanna/GradeParser.git
cd GradeParser
```

### 3. Create and Activate a Virtual Environment
```bash
# On Linux / macOS:
python3 -m venv .venv
source .venv/bin/activate

# On Windows (PowerShell):
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## Usage

1. **Run the script**:
   ```bash
   python main.py
   ```

2. **Select PDF Files**:
   - The file selection dialog will appear.
   - Select one or more semester grade card PDFs and click **Open** / **Submit**.

3. **Enter Output Filename**:
   - In your console, you will be prompted:
     ```text
     Enter output Excel file name (excluding .xlsx): my_grades
     ```
   - Press **Enter** (defaults to `student_grades.xlsx` if left blank).

4. **View Results**:
   - The console will display the number of courses extracted per file and a tabular summary.
   - Your formatted Excel file (e.g., `my_grades.xlsx`) will be saved in the project root.

---

## Output Excel Format

| Semester | Result Date | Course Code | Course Name | No. of Credits | Grade |
| :--- | :---: | :---: | :--- | :---: | :---: |
| Semester 1 | 27 January 2025 | CSE1004 | Problem Solving Using C | 3 | O |
| Semester 1 | 27 January 2025 | DES1146 | Introduction to Design Thinking | 1 | O |
| Semester 2 | 14/07/2025 | ENG2001 | Advanced English | 2 | A+ |
| Semester 3 | 03/02/2026 | CSE2253 | Data Structures | 3 | A+ |
| Semester 4 | 01/07/2026 | CCS2504 | Ethical Hacking | 2 | A+ |

---

## Project Structure

```text
GradeExtractor/
├── Files/                 # (Optional) Folder for storing input PDF files
├── .gitignore             # Git ignore configuration
├── main.py                # Main extractor script
├── README.md              # Project documentation
└── requirements.txt       # Python package dependencies
```

---

## License

MIT License. Feel free to use and adapt this tool for your own academic institution.
