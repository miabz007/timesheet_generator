REDSTACK TIMESHEET AUTOMATION - STARTER

Files
- timesheet_automation.py: main program

Requirements
- Windows PC with Python 3.10 or newer
- Microsoft Excel is recommended for reviewing the workbook outputs

Install
1. Install Python from https://www.python.org/downloads/ if it is not installed.
2. Open Command Prompt in the folder containing this script.
3. Run: py -m pip install openpyxl
4. Run: py timesheet_automation.py

Workflow
1. Select the Google Form responses workbook.
2. Select the fingerprint attendance workbook.
3. Select the existing timesheet template.
4. Select reporting year and month.
5. Select an output folder.
6. On the first run, Employee_Mapping.csv is created in that output folder. Fill the Google Form Name column for every fingerprint name, save the CSV, then run again and select the SAME output folder.
7. Timesheets and Exceptions.xlsx are written into the selected output folder's year-month subfolder.

Employee mapping troubleshooting
- The program now tells you the exact Employee_Mapping.csv path it read.
- Select the same output folder used when you edited Employee_Mapping.csv. A mapping file in another folder will not be used.
- Headers should be Fingerprint Name and Google Form Name. Header capitalization/extra spaces are tolerated.
- Matching ignores capitalization, spaces and punctuation in the fingerprint-name key (e.g. AikZhang and Aik Zhang). You must still enter the correct full Google Form name in the Google Form Name column.
- Do not leave Google Form Name blank for any fingerprint name.

Important
- Fingerprint clock-in/out times override times entered in Google Form responses.
- Records found in only one source are listed in Exceptions.xlsx.
- Review employee mappings and exception reports before using timesheets officially.
- Working-hour calculations currently use elapsed time between first and last fingerprint punches. Confirm company rules for breaks, overtime, leave codes, public holidays, and overnight shifts before official use.
