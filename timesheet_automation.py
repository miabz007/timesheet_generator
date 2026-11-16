#!/usr/bin/env python3
"""
Redstack monthly timesheet automation
- Google Form submission supplies work/task/project/file details.
- Fingerprint data is authoritative for Time In / Time Out.
- Records found in only one source are flagged in Exceptions.xlsx.
- Uses the supplied Template.xlsx and preserves its formatting.
Requires: Python 3.10+, openpyxl
Install: py -m pip install openpyxl
Run: py timesheet_automation.py
"""
from __future__ import annotations
from email.mime import text
import calendar, csv, re, shutil, sys
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from copy import copy
from openpyxl.drawing.image import Image

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Missing dependency. Run: py -m pip install openpyxl")
    raise

def norm(value):
    if value is None: return ""
    s = str(value).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def name_key(value):
    """Loose key for matching fingerprint names despite spaces/punctuation/case."""
    return re.sub(r"[^a-z0-9]", "", norm(value))

def as_date(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try: return datetime.strptime(v.strip(), fmt).date()
            except ValueError: pass
    return None

def as_time(v):
    if isinstance(v, datetime): return v.time()
    if isinstance(v, time): return v
    if isinstance(v, timedelta):
        secs = int(v.total_seconds()) % 86400
        return time(secs//3600, (secs%3600)//60)
    if isinstance(v, (int,float)) and 0 <= v < 1:
        secs=round(v*86400); return time((secs//3600)%24,(secs%3600)//60)
    if isinstance(v,str):
        s=v.strip()
        for fmt in ("%H:%M:%S","%H:%M","%I:%M %p"):
            try: return datetime.strptime(s,fmt).time()
            except ValueError: pass
    return None

def time_to_fraction(t):
    return (t.hour*3600+t.minute*60+t.second)/86400 if t else None

def is_weekday(d):
    # Monday = 0, Sunday = 6
    return d.weekday() < 5


def find_header(ws, header_map):
    if ws is None:
        raise ValueError("The worksheet is None. Check which worksheet is being read.")

    max_col = ws.max_column or 0
    max_row = ws.max_row or 0

    if max_col == 0 or max_row == 0:
        raise ValueError(
            f"Worksheet '{ws.title}' is empty. "
            "Check the Google Form responses Excel file and selected worksheet."
        )

    found = {}

    for c in range(1, max_col + 1):
        value = ws.cell(row=1, column=c).value
        if value is None:
            continue

        header = str(value).strip().lower()

        for key, aliases in header_map.items():
            if any(header == alias.strip().lower() for alias in aliases):
                found[key] = c

    return found


def choose_file(title, filetypes):
    return filedialog.askopenfilename(title=title, filetypes=filetypes)

def read_form(path, report_month):
    wb=openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws=wb["Time Attendance"] if "Time Attendance" in wb.sheetnames else wb[wb.sheetnames[0]]
    headers=find_header(ws, {
        "name":["Employee Name","Name"], "date":["Date","Attendance Date"],
        "in":["Clock-In Time","Clock In","Time In"], "out":["Clock-Out Time","Clock Out","Time Out"],
        "project":["Project"], "task":["Work Task","Task","Description of Work"],
        "file":["File Name"], "pct":["Percentage of completion"],
        "breakflag":["Unpaid Break/Prayer Time?","Unpaid Break/Prayer Time"],
        "breakduration":["Duration"], "remark":["I am filing for:"],
        "ot":["Overtime?"], "otfrom":["Overtime From"], "otto":["Overtime To"],
        "otproject":["Project (Overtime)"], "ottask":["Work Tasks (Overtime)"],
        "otfile":["File Name (Overtime)"], "otpct":["Percentage of completion (Overtime)"]
    })
    required=["name","date"]
    if any(k not in headers for k in required):
        raise ValueError("Could not find Employee Name and Date headers in the Google Form workbook.")
    records=defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        name=row[headers["name"]-1]
        d=as_date(row[headers["date"]-1])
        if not name or not d or (d.year,d.month)!=(report_month.year,report_month.month): continue
        def get(k): return row[headers[k]-1] if k in headers and headers[k]-1<len(row) else None
        rec={"name":str(name).strip(),"date":d,"in":as_time(get("in")),"out":as_time(get("out")),
             "project":str(get("project") or "").strip(),"task":str(get("task") or "").strip(),
             "file":str(get("file") or "").strip(),"pct":get("pct"),"breakflag":str(get("breakflag") or "").strip(),
             "breakduration":get("breakduration"),"remark":str(get("remark") or "").strip(),
             "ot":str(get("ot") or "").strip(),"otfrom":as_time(get("otfrom")),"otto":as_time(get("otto")),
             "otproject":str(get("otproject") or "").strip(),"ottask":str(get("ottask") or "").strip(),
             "otfile":str(get("otfile") or "").strip(),"otpct":get("otpct")}
        records[(norm(name),d)].append(rec)
    wb.close()
    return records

def parse_fingerprint(path, report_month):
    wb=openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws=wb["Att.log report"] if "Att.log report" in wb.sheetnames else wb[wb.sheetnames[0]]
    # Report is in paired rows: employee metadata row, then 1 row of daily punch cells.
    days={}
    for c in range(1, min(31,ws.max_column)+1):
        v=ws.cell(4,c).value
        if isinstance(v,(int,float)) and 1 <= int(v) <= 31: days[c]=int(v)
    result={}
    issues=[]
    r=5
    while r <= ws.max_row:
        # Employee header row contains ID: in A and Name: in I; ID and name values are nearby.
        rowvals=[ws.cell(r,c).value for c in range(1, min(ws.max_column,30)+1)]
        if norm(ws.cell(r,1).value) in ("id:","id") or any(norm(x)=="id:" for x in rowvals):
            emp_id = ws.cell(r,3).value
            # Name commonly appears in L; search after "Name:" label.
            emp_name=None
            for c in range(1, ws.max_column):
                if norm(ws.cell(r,c).value) in ("name:","name"):
                    for cc in range(c+1, min(c+5,ws.max_column)+1):
                        if ws.cell(r,cc).value not in (None,""):
                            emp_name=ws.cell(r,cc).value; break
                    if emp_name: break
            punchrow=r+1
            if emp_name:
                for c,day in days.items():
                    d=date(report_month.year,report_month.month,day)
                    raw=ws.cell(punchrow,c).value
                    if raw is None or str(raw).strip()=="": continue
                    text=str(raw).strip()
                    
                    times = re.findall(
                        r"(?<!\d)(?:[01]?\d|2[0-3]):[0-5]\d(?!\d)",
                        text
                    )

                    parsed=[]
                    for t in times:
                        try: parsed.append(datetime.strptime(t,"%H:%M").time())
                        except ValueError: pass
                    status=""
                    if not parsed:
                        status=text
                        tin=tout=None
                    else:
                        tin=parsed[0]
                        tout=parsed[-1] if len(parsed)>=2 else None
                        if len(parsed)==1: status="Only one fingerprint time"
                        if len(parsed)>2: status=f"Multiple punches ({len(parsed)}); using first/last"
                    result[(norm(emp_name),d)]={"finger_name":str(emp_name).strip(),"emp_id":str(emp_id or "").strip(),
                                               "date":d,"in":tin,"out":tout,"raw":text,"status":status}
            r += 2
        else:
            r += 1
    wb.close()
    return result

def load_mapping(mapping_path):
    mapping={}
    if mapping_path and Path(mapping_path).exists():
        with open(mapping_path, newline="", encoding="utf-8-sig") as f:
            reader=csv.DictReader(f)
            # Tolerate extra spaces/capitalization in CSV header names.
            for row in reader:
                cleaned={norm(k).replace(" ", ""): v for k,v in row.items() if k is not None}
                short=str(cleaned.get("fingerprintname") or cleaned.get("fingerprint_name") or "").strip()
                full=str(cleaned.get("googleformname") or cleaned.get("google_form_name") or "").strip()
                if short and full:
                    mapping[name_key(short)]=full
    return mapping

def make_mapping_template(path, finger_records, form_records):
    form_names=sorted({r[0] for r in form_records.keys()})
    # fingerprint names distinct
    fp_names=sorted({v["finger_name"] for v in finger_records.values()}, key=str.lower)
    with open(path,"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f)
        w.writerow(["Fingerprint Name","Google Form Name","Employee ID","Notes"])
        for n in fp_names:
            candidates=[full for full in form_names if n.lower() in full.lower()]
            guess=""
            if len(candidates)==1:
                # Recover the original display name from the form rows.
                for key, rows in form_records.items():
                    if key[0]==candidates[0] and rows:
                        guess=rows[0]["name"]
                        break
            w.writerow([n,guess,"","Review mapping before running"])
    return path

def set_if_cell(ws, cell, value):
    ws[cell]=value


def populate_template(template_path, output_path, employee, month, daily, status_rows):
    # Load the original template so its formatting and supported embedded
    # images (including a normal inserted logo) are retained.
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Template"] if "Template" in wb.sheetnames else wb[wb.sheetnames[0]]

    print("Template worksheet:", ws.title)
    print("Images detected:", len(ws._images))

    for i, img in enumerate(ws._images, start=1):
        print(f"Image {i}: width={img.width}, height={img.height}")

    

    # Employee and reporting period.
    ws["D5"] = employee
    ws["P5"] = month.strftime("%B")
    ws["Q5"] = month.year

    days_in_month = calendar.monthrange(month.year, month.month)[1]

    for day in range(1, 33):
        rr = 8 + day

        if day > days_in_month:
            for col in ["B", "D", "E", "F", "H", "P"]:
                ws[f"{col}{rr}"] = None
            continue

        d = date(month.year, month.month, day)
        rec = daily.get(d)

        # Date and day of week.
        ws[f"B{rr}"] = d.strftime("%d/%m/%Y") + "\n" + d.strftime("%A")

        # Clear attendance/task values first to avoid retaining sample data.
        ws[f"D{rr}"] = None
        ws[f"E{rr}"] = None
        ws[f"H{rr}"] = None
        ws[f"P{rr}"] = None

        # Calculate hours automatically in Excel.
        # MOD also supports overnight shifts, e.g. 22:00 to 06:00 = 8 hours.
        ws[f"F{rr}"] = f'=IF(OR(D{rr}="",E{rr}=""),"",MIN(8,MOD(E{rr}-D{rr},1)*24-1))'
       
        ws[f"F{rr}"].number_format = "0.00"

        # No fingerprint record means no attendance or form task data.
        if not rec:
            continue

        tin = rec.get("in")
        tout = rec.get("out")

        # Fingerprint data is authoritative for Clock In / Clock Out.
        if tin:
            ws[f"D{rr}"] = time_to_fraction(tin)
            ws[f"D{rr}"].number_format = "hh:mm"

        if tout:
            ws[f"E{rr}"] = time_to_fraction(tout)
            ws[f"E{rr}"].number_format = "hh:mm"

        # Work/task details are retained only for fingerprint-backed days.
        ws[f"H{rr}"] = rec.get("task", "")
        ws[f"P{rr}"] = rec.get("file", "")

        # Highlight records needing review.
        if rec.get("flag"):
            for col in range(2, 20):
                ws.cell(rr, col).fill = PatternFill(
                    "solid", fgColor="FFF2CC"
                )

            existing = ws[f"H{rr}"].value or ""
            ws[f"H{rr}"] = (
                str(existing) + "\n[REVIEW: " + rec["flag"] + "]"
            ).strip()

    # Total hours: F is the No. of hour column.
    ws["F41"] = "=SUM(F9:F40)"
    ws["F41"].number_format = "0.00"

    # Request automatic formula recalculation when opened in Excel.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass

    # Reinsert template images
    
   # 1. Handle any existing images from the template first
    images_to_reinsert = []
    for img in ws._images:
        img.width = 150  # adjust size if needed
        img.height = 50
        images_to_reinsert.append(img)
    
    ws._images.clear()

    for img in images_to_reinsert:
        ws.add_image(img, img.anchor)

    # 2. Add your external logo file once at the very end
        
    wb.save(output_path)
    wb.close()


def main():
    root=tk.Tk(); root.withdraw()
    messagebox.showinfo("Redstack Timesheet Automation","Select the Google Form responses workbook, fingerprint report, and timesheet template.")
    form_path=choose_file("Select Google Form responses workbook",[("Excel files","*.xlsx *.xlsm")])
    if not form_path: return
    fp_path=choose_file("Select fingerprint attendance workbook",[("Excel files","*.xlsx *.xlsm")])
    if not fp_path: return
    template_path=choose_file("Select individual timesheet template",[("Excel files","*.xlsx *.xlsm")])

    if not template_path: return
    year=simpledialog.askinteger("Reporting month","Enter year (e.g. 2026):",initialvalue=2026,minvalue=2000,maxvalue=2100)
    if not year: return
    month_num=simpledialog.askinteger("Reporting month","Enter month number (1-12):",initialvalue=10,minvalue=1,maxvalue=12)
    if not month_num: return
    month=date(year,month_num,1)
    outdir=filedialog.askdirectory(title="Choose output folder")
    if not outdir: return
    outdir=Path(outdir); outdir.mkdir(parents=True,exist_ok=True)
    
    form=read_form(form_path,month)
    finger=parse_fingerprint(fp_path,month)

    # Verify the fingerprint report period before interpreting its day columns.
    fpwb=openpyxl.load_workbook(fp_path, data_only=True, read_only=True)
    title_text=str(fpwb.active["A1"].value or "")
    fpwb.close()
    month_names={name.lower(): idx for idx,name in enumerate(calendar.month_name) if name}
    title_match=re.search(r"(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\\s+(20\\d{2})",title_text,re.I)
    if title_match:
        source_month=month_names[title_match.group(1).lower()]
        source_year=int(title_match.group(2))
        if (source_year,source_month)!=(month.year,month.month):
            messagebox.showerror("Reporting period mismatch",
                f"The fingerprint report is for {calendar.month_name[source_month]} {source_year}, but you selected {month.strftime('%B %Y')}.\\n\\nChoose the matching month or select the correct fingerprint report. No timesheets were generated.")
            return
    mapping_path=outdir/"Employee_Mapping.csv"
    mapping_existed=mapping_path.exists()
    if not mapping_existed:
        make_mapping_template(mapping_path,finger,form)
        messagebox.showwarning("Employee mapping review required",
            f"Created {mapping_path}. Review every row and fill/correct the Google Form name for each fingerprint name. Save the CSV, then run the automation again.")
        return
    mapping=load_mapping(mapping_path)
    # Explicit mappings are required where fingerprint names differ from full Google Form names.
    unmapped=[v["finger_name"] for v in finger.values() if name_key(v["finger_name"]) not in mapping]
    if unmapped:
        messagebox.showerror("Employee mapping incomplete",
            "The automation read this mapping file:\n"+str(mapping_path)+"\n\nFill the Google Form Name column for these fingerprint names, save the CSV, and run again. Matching ignores capitalization, spaces, and punctuation.\n\nUnmapped: "+", ".join(sorted(set(unmapped))))
        return
    # Create form-keyed daily aggregate.
    combined=defaultdict(dict)
    exceptions=[]
    form_keys=set(form.keys())
    finger_form_keys=set()
    for (fpname, d), fp in finger.items():
        if not is_weekday(d):
            continue
        full=mapping.get(name_key(fp["finger_name"]),fp["finger_name"])
        key=(norm(full),d)
        finger_form_keys.add(key)
        formrows=form.get(key,[])
        if not formrows:
            exceptions.append([d.isoformat(),full,fp["emp_id"],"Fingerprint only","No Google Form submission for this employee/date",fp["raw"]])
        if fp["status"]:
            exceptions.append([d.isoformat(),full,fp["emp_id"],"Fingerprint review",fp["status"],fp["raw"]])
        combined[key]["finger"]=fp
    for key,rows in form.items():
        if key not in finger_form_keys:
            exceptions.append([key[1].isoformat(),rows[0]["name"],"","Form only","No fingerprint record for this employee/date",""])
        combined[key]["formrows"]=rows
    # Generate per employee, using all employees seen in either source.
    

    # Generate one timesheet per canonical employee name.
    # Fingerprint records determine attendance and which dates are included.
    employees = {}

    for key, info in combined.items():
        if not is_weekday(key[1]):
            continue
        d = key[1]
        rows = info.get("formrows", [])
        fp = info.get("finger")

        # Do not write form-only dates to timesheets.
        # They remain recorded in Exceptions.xlsx.
        if not fp:
            continue

        # Use the mapped employee name as the consistent grouping key.
        canonical_name = mapping.get(
            name_key(fp["finger_name"]),
            fp["finger_name"]
        ).strip()

        employee_key = norm(canonical_name)

        # Keep one display name for each canonical employee.
        if employee_key not in employees:
            employees[employee_key] = {
                "display_name": canonical_name,
                "daily": {}
            }

        # Combine form details for this employee and date.
        projects = []
        tasks = []
        files = []
        pct = []
        remarks = []
        otfrom = []
        otto = []
        otfiles = []

        for rec in rows:
            if rec.get("project") and rec["project"] not in projects:
                projects.append(rec["project"])
            if rec.get("task") and rec["task"] not in tasks:
                tasks.append(rec["task"])
            if rec.get("file") and rec["file"] not in files:
                files.append(rec["file"])
            if rec.get("pct") not in (None, "") and str(rec["pct"]) not in pct:
                pct.append(str(rec["pct"]))
            if rec.get("remark") and rec["remark"] not in remarks:
                remarks.append(rec["remark"])
            if rec.get("otfrom"):
                otfrom.append(rec["otfrom"])
            if rec.get("otto"):
                otto.append(rec["otto"])
            if rec.get("otfile") and rec["otfile"] not in otfiles:
                otfiles.append(rec["otfile"])

        tin = fp.get("in")
        tout = fp.get("out")
        flag = ""

        if not rows:
            flag = "Fingerprint record exists but Google Form submission is missing"

        if not tin or not tout:
            flag = (flag + "; " if flag else "") + "Fingerprint time incomplete"

        if len(rows) > 1:
            exceptions.append([
                d.isoformat(),
                canonical_name,
                fp.get("emp_id", ""),
                "Multiple form submissions",
                f"{len(rows)} form rows for same employee/date; task details combined",
                fp.get("raw", "")
            ])

        employees[employee_key]["daily"][d] = {
            "in": tin,
            "out": tout,
            "project": "; ".join(projects),
            "task": "; ".join(tasks),
            "file": "; ".join(files),
            "pct": "; ".join(pct),
            "remark": "; ".join(remarks),
            "otfrom": min(otfrom) if otfrom else None,
            "otto": max(otto) if otto else None,
            "otfile": "; ".join(otfiles),
            "flag": flag
        }

    # Write output files.
    month_dir=outdir/f"{month.year}-{month.month:02d}"
    month_dir.mkdir(exist_ok=True)
    
    for employee_data in employees.values():
        employee = employee_data["display_name"]
        daily = employee_data["daily"]

        safe = re.sub(r'[<>:"/\\\\|?*]+', "_", employee).strip() or "Employee"
        target = month_dir / f"{safe}_{month.strftime('%Y_%m')}_Timesheet.xlsx"

        populate_template(
            template_path,
            target,
            employee,
            month,
            daily,
            exceptions
        )

    # Exception report
    ex_path=month_dir/"Exceptions.xlsx"
    ewb=openpyxl.Workbook(); ews=ewb.active; ews.title="Exceptions"
    ews.append(["Date","Employee","Employee ID","Issue Type","Details","Raw Fingerprint Cell"])
    for row in exceptions: ews.append(row)
    for cell in ews[1]:
        cell.font=Font(bold=True,color="FFFFFF")
        cell.fill=PatternFill("solid",fgColor="1F4E78")
    ews.freeze_panes="A2"; ews.auto_filter.ref=ews.dimensions
    for col in range(1,7): ews.column_dimensions[get_column_letter(col)].width=[14,24,14,26,58,35][col-1]
    ewb.save(ex_path); ewb.close()
    messagebox.showinfo("Complete",f"Generated {len(employees)} employee timesheets.\nExceptions: {len(exceptions)}\nOutput folder:\n{month_dir}\n\nIMPORTANT: Check Employee_Mapping.csv and verify mapped names before using outputs.")
if __name__=="__main__":
    main()
