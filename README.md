# Redstack Monthly Timesheet Automation

A Python-based automation tool that generates individual monthly employee timesheets by combining Google Form submissions with fingerprint attendance records.

The system automates attendance reconciliation, employee matching, timesheet generation, and exception reporting while preserving the original Excel timesheet template.

## Features

- **Automated Timesheet Generation** – Generates individual Excel timesheets for each employee.
- **Fingerprint Attendance Integration** – Uses fingerprint records as the authoritative source for Clock In and Clock Out times.
- **Google Form Integration** – Retrieves employee work details, including projects, tasks, and file names.
- **Employee Name Mapping** – Matches fingerprint employee names with Google Form employee names.
- **Attendance Reconciliation** – Identifies discrepancies between Google Form submissions and fingerprint records.
- **Exception Reporting** – Generates a separate `Exceptions.xlsx` report for records requiring attention.
- **Weekday Filtering** – Processes Monday to Friday records only.
- **Excel Template Preservation** – Uses the supplied timesheet template, retaining its formatting and logo.
- **Automated Working Hours Calculation** – Calculates working hours based on fingerprint attendance, deducting the configured one-hour break.
- **Graphical User Interface** – Uses file selection dialogs to simplify the automation process.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python 3.10+ | Main programming language |
| OpenPyXL | Excel workbook processing and generation |
| Pillow | Image handling for Excel templates |
| Tkinter | Graphical file selection and input dialogs |
| CSV | Employee name mapping |

## Project Structure

```text
Redstack-Timesheet-Automation/
│
├── timesheet_automation.py
├── Template.xlsx
├── Employee_Mapping.csv
├── .gitignore
├── README.md
│
├── input/
│   ├── Google_Form_Responses.xlsx
│   └── Fingerprint_Attendance.xlsx
│
└── output/
    └── 2026-10/
        ├── Employee_A_2026_10_Timesheet.xlsx
        ├── Employee_B_2026_10_Timesheet.xlsx
        └── Exceptions.xlsx
```

*Note: The input and output folder names are illustrative. The program allows users to select files and an output directory through its graphical interface.*

## Requirements

- Windows operating system
- Python 3.10 or newer
- Google Form responses exported as an Excel workbook
- Fingerprint attendance report exported as an Excel workbook
- Individual employee timesheet template (`Template.xlsx`)
- Employee mapping CSV file

## Installation

### 1. Install Python

Download and install Python 3.10 or newer:

https://www.python.org/downloads/

Ensure Python is added to the system PATH during installation.

### 2. Install Required Packages

Open Command Prompt and execute:

```bash
py -m pip install openpyxl Pillow
```

### 3. Clone the Repository

```bash
git clone YOUR_REPOSITORY_URL
```

Navigate into the project directory:

```bash
cd Redstack-Timesheet-Automation
```

### 4. Prepare the Required Files

Place the following files in an accessible location:

- Google Form responses workbook
- Fingerprint attendance report
- `Template.xlsx`
- `Employee_Mapping.csv` (if already created)

## How to Use

Run the automation:

```bash
py timesheet_automation.py
```

Follow the graphical prompts:

1. Select the Google Form responses workbook.
2. Select the fingerprint attendance report.
3. Select the individual timesheet template.
4. Enter the reporting year.
5. Enter the reporting month.
6. Select the output directory.

The program will process the records and generate individual employee timesheets and an exception report.

## Data Processing Logic

### 1. Google Form Data

Google Form submissions provide:

- Employee name
- Attendance date
- Project
- Work task
- File name
- Percentage of completion
- Overtime information and other submission details

Google Form clock-in and clock-out times are not treated as the authoritative attendance records.

### 2. Fingerprint Attendance Data

Fingerprint records are the authoritative source for:

- Time In
- Time Out

The system extracts the first and last fingerprint punches for each employee and date.

Records with missing or unusual fingerprint punches are flagged for review.

### 3. Employee Name Mapping

The system uses `Employee_Mapping.csv` to match fingerprint names with Google Form employee names.

Expected CSV headers:

```csv
Fingerprint Name,Google Form Name,Employee ID,Notes
```

Name matching ignores capitalization, spaces, and punctuation.

If the mapping file does not exist, the system generates a mapping template for review.

Every required employee mapping must be completed before processing.

### 4. Attendance Reconciliation

The automation compares fingerprint attendance records against Google Form submissions.

| Scenario | Result |
|---|---|
| Fingerprint and Form records exist | Generate timesheet entry |
| Fingerprint exists, Form missing | Generate timesheet entry and flag exception |
| Form exists, fingerprint missing | Flag exception; do not generate attendance entry |
| Only one fingerprint punch | Flag for review |
| Multiple fingerprint punches | Use first and last punches and flag for review |

Only Monday-to-Friday records are processed.

### 5. Working Hours Calculation

Working hours are calculated using fingerprint Time In and Time Out, with a one-hour break deduction.

The calculation supports overnight shifts.

## Output Files

The automation generates a separate Excel timesheet for each employee.

### Individual Employee Timesheets

Each timesheet contains:

- Employee name
- Reporting month and year
- Daily attendance dates
- Time In
- Time Out
- Calculated working hours
- Work tasks
- File names
- Review indicators where applicable

The original Excel template is used to maintain the intended layout and formatting.

### Exceptions.xlsx

A separate exception report is generated to identify attendance discrepancies.

It includes:

- Date
- Employee name
- Employee ID
- Issue type
- Details
- Raw fingerprint cell data

This report helps administrators identify missing submissions and attendance records requiring manual verification.

## Data Security

Employee attendance records and timesheets contain confidential company information.

The project uses `.gitignore` to prevent sensitive files from being committed to the Git repository.

The following files should not be uploaded to a public repository:

- Employee attendance workbooks
- Fingerprint reports
- Employee mapping CSV files
- Generated employee timesheets
- Exception reports
- Credentials and access tokens

Only source code, documentation, and sanitised example files should be committed.

## Troubleshooting

| Issue | Suggested Action |
|---|---|
| Missing Python packages | Run `py -m pip install openpyxl Pillow` |
| Employee mapping incomplete | Review and complete `Employee_Mapping.csv` |
| No employee timesheets generated | Check fingerprint parsing, reporting month, and employee mapping |
| Missing logo in generated timesheets | Verify the selected template and image handling |
| Reporting period mismatch | Confirm the selected month matches the fingerprint report |
| Missing attendance records | Check the source workbooks and exception report |

## Important Notes

- Fingerprint attendance is authoritative for Time In and Time Out.
- Google Form submissions supply work and task information.
- Form-only records are excluded from individual attendance timesheets.
- Weekends are excluded from processing.
- Always review `Exceptions.xlsx` before finalising employee timesheets.
- Keep the original template and source workbooks backed up.

## License

This project is intended for internal company use.

Do not redistribute employee attendance records, fingerprint data, or confidential company information.
