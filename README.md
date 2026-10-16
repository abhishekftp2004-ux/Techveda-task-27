# Task 27 — Reading and Writing Excel Data

## Objective
Use Python and Pandas to read an Excel workbook, inspect its structure, perform basic cleaning and transformation, and export processed results.

## Files
- `employee_dataset.xlsx` — synthetic source workbook with `Employees` and `Data_Dictionary` sheets.
- `Task_27_Reading_Writing_Excel.ipynb` — step-by-step notebook.
- `task27_excel_solution.py` — reusable Python script.
- `processed_employee_data.xlsx` — finished workbook with `Cleaned_Employees`, `Department_Summary`, and `Data_Profile` sheets.
- `processed_employees.csv` — cleaned records in CSV format.
- `Task_27_Explanation.pdf` — detailed concepts, code walkthrough, validation, interview questions, and exercises.

## Setup
Python 3.10+ is recommended.

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
```

## Run
Keep `employee_dataset.xlsx` in the same directory as the script, then run:

```bash
python task27_excel_solution.py
```

The script writes `processed_employee_data.xlsx` and `processed_employees.csv` to the current working directory.

## Cleaning decisions
- Trim whitespace and standardize employee names, departments, and locations.
- Convert salary strings such as `₹45,000` to numeric values.
- Parse joining dates and numeric age values.
- Replace a missing department with `Unknown`.
- Fill missing age with the median age.
- Convert missing email to a blank string.
- Remove exact duplicate rows.

These are teaching choices, not universal rules. For real business data, confirm data definitions and business rules before replacing missing values or deleting duplicates.

## Tools
Python, Pandas, OpenPyXL, Jupyter Notebook.
