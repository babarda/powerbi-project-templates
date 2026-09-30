"""Build the four sample_data.xlsx workbooks for the Power BI project-controls templates.

All content is fictional: the "North Ridge Water Plant" project, its areas,
roles and activities are invented. Run with Python 3.10+ and openpyxl.
Author: Badr Abardazzou. License: MIT.
"""
import math
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
AUTHOR = "Badr Abardazzou"
PROJECT = "North Ridge Water Plant"
DATA_DATE = date(2026, 9, 25)  # a Friday, the weekly cut-off

HEAD_FILL = PatternFill("solid", fgColor="005F5F")
HEAD_FONT = Font(bold=True, color="FFFFFF", name="Segoe UI", size=10)
BODY_FONT = Font(name="Segoe UI", size=10)
TITLE_FONT = Font(bold=True, name="Segoe UI", size=14, color="005F5F")
NOTE_FONT = Font(name="Segoe UI", size=10, color="444444")


def new_workbook():
    wb = Workbook()
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.properties.company = AUTHOR
    wb.properties.title = f"{PROJECT} sample data"
    return wb


def write_table(ws, name, headers, rows, widths=None, formats=None):
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=len(rows) + 1):
        for cell in row:
            cell.font = BODY_FONT
    formats = formats or {}
    for col_idx, h in enumerate(headers, start=1):
        fmt = formats.get(h)
        if fmt:
            for r in range(2, len(rows) + 2):
                ws.cell(row=r, column=col_idx).number_format = fmt
        w = (widths or {}).get(h, max(12, min(40, len(h) + 4)))
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.freeze_panes = "A2"
    ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    t = Table(displayName=f"tbl{name}", ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)


def write_readme(ws, title, intro_lines, columns):
    """columns: list of (sheet, column, type, required, meaning, example)."""
    ws.title = "README"
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    r = 2
    for line in intro_lines:
        ws.cell(row=r, column=1, value=line).font = NOTE_FONT
        r += 1
    r += 1
    headers = ["Sheet", "Column", "Type", "Required", "What to enter", "Example"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
    for spec in columns:
        r += 1
        for c, v in enumerate(spec, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEF", (14, 22, 12, 10, 70, 24)):
        ws.column_dimensions[col].width = w


COMMON_STEPS = [
    "How to use: keep every sheet name and every column header exactly as they are. Replace the sample rows with your own.",
    "Then open the Power BI template, set the DataFile parameter to the full path of this workbook, and press Refresh.",
    "Sample content is fictional (the North Ridge Water Plant project). It identifies no real company or project.",
]


def settings_sheet(wb, extra_headers=(), extra_values=()):
    ws = wb.create_sheet("Settings")
    headers = ["Project Name", "Data Date", *extra_headers]
    values = [PROJECT, datetime.combine(DATA_DATE, datetime.min.time()), *extra_values]
    write_table(ws, "Settings", headers, [values], widths={"Project Name": 32, "Data Date": 14},
                formats={"Data Date": "yyyy-mm-dd"})
    return ws


SETTINGS_DOC = [
    ("Settings", "Project Name", "Text", "Yes", "Project name shown in the report header.", PROJECT),
    ("Settings", "Data Date", "Date", "Yes", "Status date of the report (cut-off). Actuals are counted up to and including this date. Leave empty to use today.", "2026-09-25"),
]


# ---------------------------------------------------------------- 1. S-curve
def bell(n):
    w = [math.sin(math.pi * (i + 0.5) / n) ** 2 for i in range(n)]
    s = sum(w)
    return [x / s for x in w]


def spread(total, weights):
    vals = [round(total * w) for w in weights]
    vals[-1] += total - sum(vals)
    return vals


def build_progress():
    rnd = random.Random(42)
    start = date(2026, 1, 9)
    dd_idx = (DATA_DATE - start).days // 7  # 37
    disciplines = [
        # name, budget hours, first week, last week, performance factor
        ("Civil", 18000, 0, 30, 1.02),
        ("Structural", 12000, 8, 38, 0.97),
        ("Mechanical", 15000, 18, 52, 0.93),
        ("Piping", 22000, 20, 56, 0.80),
        ("Electrical", 14000, 24, 60, 0.90),
        ("Instrumentation", 9000, 30, 63, 0.86),
    ]
    areas = [("Intake", 0.20, 0), ("Treatment", 0.40, 1), ("Pumping Station", 0.25, 2), ("Reservoir", 0.15, 3)]
    rows_by_key = {}
    max_week = 0
    for disc, budget, w0, w1, perf in disciplines:
        for area, share, off in areas:
            a0, a1 = w0 + off, w1 + off
            n = a1 - a0 + 1
            planned = spread(round(budget * share), bell(n))
            plan = {a0 + i: planned[i] for i in range(n)}
            earned = {}
            for wk in range(a0, min(a1, dd_idx) + 1):
                p = plan.get(wk, 0)
                earned[wk] = max(0, round(p * perf * rnd.uniform(0.85, 1.15)))
            area_budget = round(budget * share)
            total_earned = sum(earned.values())
            excess = total_earned - area_budget
            for wk in sorted(earned, reverse=True):  # never earn more than the budget
                if excess <= 0:
                    break
                cut = min(excess, earned[wk])
                earned[wk] -= cut
                excess -= cut
            total_earned = sum(earned.values())
            remaining = max(0, round(budget * share) - total_earned)
            forecast = {}
            if remaining > 0:
                tail_start = dd_idx + 1
                planned_tail = max(1, a1 - tail_start + 1) if a1 >= tail_start else 1
                stretch = 1 + (1 - perf) * 2.2
                m = max(2, math.ceil(planned_tail * stretch))
                if a1 < tail_start:
                    m = max(3, math.ceil(3 * stretch))
                # decreasing-then-tail profile: remaining work front-loaded after the data date
                wts = [math.sin(math.pi * (i + 0.5) / (2 * m)) ** 0.6 * (1 - i / (m + 1)) + 0.15 for i in range(m)]
                sw = sum(wts)
                vals = spread(remaining, [x / sw for x in wts])
                for i, v in enumerate(vals):
                    forecast[tail_start + i] = v
                max_week = max(max_week, tail_start + m - 1)
            max_week = max(max_week, a1)
            rows_by_key[(disc, area)] = (plan, earned, forecast)
    rows = []
    for (disc, area), (plan, earned, forecast) in rows_by_key.items():
        for wk in range(0, max_week + 1):
            we = start + timedelta(weeks=wk)
            p = plan.get(wk, 0)
            e = earned.get(wk, 0) if wk <= dd_idx else None
            f = forecast.get(wk, 0) if wk > dd_idx else None
            rows.append((datetime.combine(we, datetime.min.time()), disc, area, p, e, f))
    rows.sort(key=lambda r: (r[0], r[1], r[2]))

    wb = new_workbook()
    write_readme(wb.active, "Progress S-curve: sample data dictionary", COMMON_STEPS + [
        "One row per week, discipline and area. Planned Hours over all weeks of a discipline/area add up to its budget.",
        "Progress is measured in earned hours (budget hours x % complete). Any consistent unit works: hours, weighted points, cost.",
    ], SETTINGS_DOC + [
        ("Progress", "Week Ending", "Date", "Yes", "Last day of the reporting week. Use the same weekday for every row (the sample uses Fridays).", "2026-09-25"),
        ("Progress", "Discipline", "Text", "Yes", "Discipline or trade. Drives the slicer and the variance table.", "Piping"),
        ("Progress", "Area", "Text", "Yes", "Area, unit or work front. Drives the slicer and the variance table.", "Treatment"),
        ("Progress", "Planned Hours", "Number", "Yes", "Baseline hours planned to be earned in that week. The sum over all weeks is the budget of that discipline/area. Use 0 for weeks with no planned work.", "412"),
        ("Progress", "Earned Hours", "Number", "Up to Data Date", "Hours actually earned in that week (budget x progress gained). Fill weeks up to the Data Date, leave later weeks empty.", "338"),
        ("Progress", "Forecast Hours", "Number", "After Data Date", "Hours you expect to earn in that week. Fill weeks after the Data Date only. Planned + earned + forecast rows can extend past the baseline finish.", "295"),
    ])
    settings_sheet(wb)
    ws = wb.create_sheet("Progress")
    write_table(ws, "Progress", ["Week Ending", "Discipline", "Area", "Planned Hours", "Earned Hours", "Forecast Hours"], rows,
                widths={"Week Ending": 14, "Discipline": 18, "Area": 18}, formats={"Week Ending": "yyyy-mm-dd"})
    out = ROOT / "progress-s-curve" / "data" / "sample_data.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out, len(rows)


# ---------------------------------------------------------- 2. Milestones
def d(s):
    return datetime.strptime(s, "%Y-%m-%d") if s else None


MILESTONES = [
    # id, name, phase, area, owner, baseline, forecast, actual
    ("M01", "Contract award", "Engineering", "Site-wide", "Project Manager", "2026-01-12", "2026-01-12", "2026-01-12"),
    ("M02", "Design basis approved", "Engineering", "Site-wide", "Engineering Manager", "2026-02-13", "2026-02-20", "2026-02-20"),
    ("M03", "Process P&IDs issued for design", "Engineering", "Treatment", "Engineering Manager", "2026-03-13", "2026-03-27", "2026-03-27"),
    ("M04", "HAZOP closed out", "Engineering", "Treatment", "HSE Manager", "2026-04-24", "2026-05-15", "2026-05-15"),
    ("M05", "Long-lead pumps ordered", "Procurement", "Pumping Station", "Procurement Lead", "2026-04-03", "2026-04-01", "2026-04-01"),
    ("M06", "Site mobilisation complete", "Construction", "Site-wide", "Construction Manager", "2026-03-02", "2026-03-06", "2026-03-06"),
    ("M07", "Intake cofferdam closed", "Construction", "Intake", "Construction Manager", "2026-05-29", "2026-06-12", "2026-06-12"),
    ("M08", "Clarifier foundations poured", "Construction", "Treatment", "Construction Manager", "2026-06-26", "2026-07-10", "2026-07-17"),
    ("M09", "IFC drawings 90% issued", "Engineering", "Site-wide", "Engineering Manager", "2026-07-31", "2026-08-21", "2026-08-28"),
    ("M10", "Reservoir slab complete", "Construction", "Reservoir", "Construction Manager", "2026-09-04", "2026-09-18", "2026-09-18"),
    ("M11", "Pump factory acceptance test", "Procurement", "Pumping Station", "Procurement Lead", "2026-09-18", "2026-10-09", None),
    ("M12", "Main switchgear delivered", "Procurement", "Pumping Station", "Procurement Lead", "2026-10-16", "2026-10-16", None),
    ("M13", "Filter building roof closed", "Construction", "Treatment", "Construction Manager", "2026-10-30", "2026-11-27", None),
    ("M14", "Intake screens installed", "Construction", "Intake", "Construction Manager", "2026-11-13", "2026-11-13", None),
    ("M15", "Pumps set on foundations", "Construction", "Pumping Station", "Construction Manager", "2026-11-27", "2027-01-15", None),
    ("M16", "Energisation of main substation", "Construction", "Pumping Station", "Commissioning Manager", "2026-12-18", "2027-01-08", None),
    ("M17", "Piping hydrotests complete", "Construction", "Treatment", "Construction Manager", "2027-01-22", "2027-03-05", None),
    ("M18", "Reservoir leak test passed", "Construction", "Reservoir", "QA/QC Manager", "2026-12-04", "2026-12-04", None),
    ("M19", "Mechanical completion, Intake", "Commissioning", "Intake", "Commissioning Manager", "2027-02-05", "2027-02-12", None),
    ("M20", "Mechanical completion, Treatment", "Commissioning", "Treatment", "Commissioning Manager", "2027-03-05", "2027-04-16", None),
    ("M21", "Ready for commissioning, Pumping Station", "Commissioning", "Pumping Station", "Commissioning Manager", "2027-03-19", "2027-04-09", None),
    ("M22", "First water through plant", "Commissioning", "Site-wide", "Commissioning Manager", "2027-04-16", "2027-05-21", None),
    ("M23", "Performance test passed", "Handover", "Site-wide", "Project Manager", "2027-05-28", "2027-06-25", None),
    ("M24", "Handover to operations", "Handover", "Site-wide", "Project Manager", "2027-06-30", "2027-07-23", None),
    ("M25", "Vendor data books received", "Procurement", "Site-wide", "Procurement Lead", "2026-09-11", "2026-09-18", None),
]


def build_milestones():
    rows = [(i, n, ph, ar, ow, d(b), d(f), d(a)) for i, n, ph, ar, ow, b, f, a in MILESTONES]
    wb = new_workbook()
    write_readme(wb.active, "Milestone tracker: sample data dictionary", COMMON_STEPS + [
        "One row per milestone. Keep the baseline date frozen; move the forecast date every week; fill the actual date when it is achieved.",
        "Status rules (computed in Power BI): Complete = Actual Date filled. Late = not complete and Forecast Date before the Data Date, "
        "or slip above Late Threshold Days. At risk = slip above 0 days. On track = everything else.",
    ], SETTINGS_DOC + [
        ("Settings", "Late Threshold Days", "Whole number", "Yes", "Slip (in calendar days) above which an open milestone counts as Late instead of At risk.", "30"),
        ("Milestones", "Milestone ID", "Text", "Yes", "Unique code of the milestone. Also sets the default sort.", "M11"),
        ("Milestones", "Milestone", "Text", "Yes", "Short milestone name.", "Pump factory acceptance test"),
        ("Milestones", "Phase", "Text", "Yes", "Project phase (Engineering, Procurement, Construction, Commissioning, Handover or your own list).", "Procurement"),
        ("Milestones", "Area", "Text", "Yes", "Area or unit. Use Site-wide for project-level milestones.", "Pumping Station"),
        ("Milestones", "Owner", "Text", "Yes", "Role or person accountable for the milestone.", "Procurement Lead"),
        ("Milestones", "Baseline Date", "Date", "Yes", "Date in the approved baseline schedule. Do not change it after approval.", "2026-09-18"),
        ("Milestones", "Forecast Date", "Date", "Yes", "Current forecast date. For achieved milestones keep the last forecast or copy the actual.", "2026-10-09"),
        ("Milestones", "Actual Date", "Date", "When achieved", "Date the milestone was achieved. Leave empty while open.", ""),
    ])
    settings_sheet(wb, ["Late Threshold Days"], [30])
    ws = wb.create_sheet("Milestones")
    fm = {"Baseline Date": "yyyy-mm-dd", "Forecast Date": "yyyy-mm-dd", "Actual Date": "yyyy-mm-dd"}
    write_table(ws, "Milestones", ["Milestone ID", "Milestone", "Phase", "Area", "Owner", "Baseline Date", "Forecast Date", "Actual Date"],
                rows, widths={"Milestone": 40, "Owner": 22, "Area": 18, "Phase": 16}, formats=fm)
    out = ROOT / "milestone-tracker" / "data" / "sample_data.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out, len(rows)


# ---------------------------------------------------------------- 3. RACI
ROLES = ["Project Manager", "Project Controls", "Engineering Manager", "Procurement Lead", "Construction Manager",
         "HSE Manager", "QA/QC Manager", "Commissioning Manager", "Client Representative"]
PM, PC, EM, PL, CM, HS, QA, CO, CR = range(9)
RACI_ROWS = [
    # id, activity, phase, {role index: code}
    ("CTL-01", "Baseline schedule", "Project controls", {PM: "A", PC: "R", EM: "C", PL: "C", CM: "C", CO: "C", CR: "I"}),
    ("CTL-02", "Weekly progress report", "Project controls", {PM: "A", PC: "R", EM: "C", CM: "C", CR: "I"}),
    ("CTL-03", "Cost forecast", "Project controls", {PM: "A", PC: "R", PL: "C", CR: "I"}),
    ("CTL-04", "Change control", "Project controls", {PM: "A", PC: "R", EM: "C", CM: "C", CR: "C"}),
    ("CTL-05", "Risk register", "Project controls", {PM: "A", PC: "R", HS: "C", EM: "C", CM: "C", CR: "I"}),
    ("CTL-06", "Monthly client report", "Project controls", {PM: "R", PC: "R", CR: "I"}),  # no A on purpose
    ("ENG-01", "Design basis", "Engineering", {PM: "C", EM: "A/R", HS: "C", CR: "C"}),
    ("ENG-02", "P&ID approval", "Engineering", {EM: "A", HS: "C", CO: "C", CR: "C", PM: "I"}),  # no R on purpose
    ("ENG-03", "HAZOP study", "Engineering", {HS: "A", EM: "R", CO: "C", CM: "C", CR: "I"}),
    ("ENG-04", "IFC drawings issue", "Engineering", {EM: "A/R", CM: "I", QA: "I", PC: "I"}),
    ("ENG-05", "Technical queries", "Engineering", {EM: "A", CM: "R", QA: "C"}),
    ("PRC-01", "Vendor list", "Procurement", {PL: "A/R", EM: "C", QA: "C", CR: "I"}),
    ("PRC-02", "Long-lead purchase orders", "Procurement", {PM: "A", PL: "R", EM: "C", PC: "I"}),
    ("PRC-03", "Expediting", "Procurement", {PL: "A/R", PC: "I", CM: "I"}),
    ("PRC-04", "Factory acceptance tests", "Procurement", {PL: "A", QA: "R", EM: "C", CR: "I"}),
    ("PRC-05", "Material receipt on site", "Procurement", {PL: "A", CM: "R", QA: "C"}),
    ("CON-01", "Site mobilisation", "Construction", {CM: "A/R", HS: "C", PM: "I"}),
    ("CON-02", "Permit to work", "Construction", {HS: "A", CM: "R", CO: "C"}),
    ("CON-03", "Civil works", "Construction", {CM: "A/R", QA: "C", PC: "I"}),
    ("CON-04", "Mechanical installation", "Construction", {CM: "A/R", QA: "C", EM: "C", PC: "I"}),
    ("CON-05", "Daily toolbox talks", "Construction", {CM: "R", HS: "C"}),  # no A on purpose
    ("CON-06", "Punch list closure", "Construction", {CM: "A", QA: "R", CO: "C", CR: "I"}),
    ("QHS-01", "Inspection and test plans", "Quality and HSE", {QA: "A/R", EM: "C", CM: "C", CR: "C"}),
    ("QHS-02", "Incident investigation", "Quality and HSE", {HS: "A", CM: "R", PM: "A", CR: "I"}),  # two A on purpose
    ("QHS-03", "Quality audits", "Quality and HSE", {QA: "A/R", PM: "I"}),
    ("COM-01", "System walkdowns", "Commissioning", {CO: "A", CM: "R", QA: "C"}),
    ("COM-02", "Pre-commissioning checks", "Commissioning", {CO: "A/R", CM: "C", QA: "C"}),
    ("COM-03", "Energisation", "Commissioning", {CO: "A", CM: "R", HS: "C", CR: "I"}),
    ("COM-04", "Performance test", "Commissioning", {CO: "R", PM: "A", CR: "C", EM: "C"}),
    ("COM-05", "Handover dossier", "Commissioning", {PM: "A", CO: "R", QA: "R", CR: "I", PC: "I"}),
]


def build_raci():
    rows = []
    for aid, act, ph, codes in RACI_ROWS:
        rows.append([aid, act, ph] + [codes.get(i) for i in range(len(ROLES))])
    wb = new_workbook()
    write_readme(wb.active, "RACI and responsibility matrix: sample data dictionary", COMMON_STEPS + [
        "Fill the RACI sheet the way you would on paper: one row per activity, one column per role, a letter in each cell.",
        "Add or remove role columns freely (keep the first three columns). Power BI unpivots every column after Phase into a role.",
        "Codes: R = Responsible (does the work), A = Accountable (one owner who signs off), C = Consulted (two-way), I = Informed (one-way).",
        "A cell can hold two codes separated by a slash, for example A/R. Empty cell = no involvement.",
        "The report flags activities with no A, with more than one A, or with no R.",
    ], SETTINGS_DOC[:1] + [
        ("RACI", "Activity ID", "Text", "Yes", "Unique activity code. Also sets the sort order.", "CTL-01"),
        ("RACI", "Activity", "Text", "Yes", "Activity or deliverable name.", "Baseline schedule"),
        ("RACI", "Phase", "Text", "Yes", "Group the activity belongs to (phase, process, work package).", "Project controls"),
        ("RACI", "<one column per role>", "Text", "No", "Header = role name. Cell = R, A, C, I or a combination such as A/R. Leave empty for no involvement.", "A/R"),
    ])
    ws0 = wb.worksheets[0]
    settings_sheet(wb)
    ws = wb.create_sheet("RACI")
    write_table(ws, "RACI", ["Activity ID", "Activity", "Phase"] + ROLES, rows,
                widths={"Activity": 30, "Phase": 18, **{r: 13 for r in ROLES}})
    for c in range(4, 4 + len(ROLES)):
        for r in range(2, len(rows) + 2):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 32
    out = ROOT / "raci-matrix" / "data" / "sample_data.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out, len(rows)


# --------------------------------------------------------- 4. Look-ahead
LOOKAHEAD = [
    # id, activity, area, owner, start, finish, pct, constraint
    ("A1010", "Excavate intake pump pit", "Intake", "Civil crew A", "2026-09-14", "2026-10-02", 0.70, None),
    ("A1020", "Blinding and rebar, pump pit base", "Intake", "Civil crew A", "2026-10-01", "2026-10-08", 0.0, "Rebar delivery confirmed for 30 Sep"),
    ("A1030", "Pour pump pit base slab", "Intake", "Civil crew A", "2026-10-09", "2026-10-09", 0.0, "Concrete pour permit"),
    ("A1040", "Install intake screen guides", "Intake", "Mechanical crew", "2026-10-12", "2026-10-21", 0.0, "Screen guides at vendor, ETA 9 Oct"),
    ("A1050", "Backfill around intake pipe", "Intake", "Civil crew B", "2026-09-28", "2026-10-06", 0.0, None),
    ("A2010", "Clarifier wall pour, lift 3", "Treatment", "Civil crew B", "2026-09-21", "2026-09-30", 0.40, None),
    ("A2020", "Clarifier wall pour, lift 4", "Treatment", "Civil crew B", "2026-10-05", "2026-10-14", 0.0, "Formwork release from lift 3"),
    ("A2030", "Filter building steel erection", "Treatment", "Steel subcontractor", "2026-09-17", "2026-10-16", 0.35, "Crane booking weeks 41 and 42"),
    ("A2040", "Filter building roof sheeting", "Treatment", "Steel subcontractor", "2026-10-15", "2026-10-30", 0.0, "Roof sheets on site"),
    ("A2050", "Chemical dosing skid setting", "Treatment", "Mechanical crew", "2026-10-06", "2026-10-09", 0.0, None),
    ("A2060", "Process piping spools, rack 1", "Treatment", "Piping crew", "2026-09-28", "2026-10-23", 0.0, "IFC isometrics rev C"),
    ("A2070", "Underground drainage, filter area", "Treatment", "Civil crew C", "2026-10-19", "2026-10-28", 0.0, None),
    ("A2080", "Scaffold for clarifier walkway", "Treatment", "Scaffold team", "2026-10-13", "2026-10-15", 0.0, None),
    ("A3010", "Pump house block walls", "Pumping Station", "Civil crew C", "2026-09-10", "2026-10-02", 0.80, None),
    ("A3020", "Pump house roof slab pour", "Pumping Station", "Civil crew C", "2026-10-07", "2026-10-07", 0.0, "Concrete pour permit"),
    ("A3030", "Cable tray installation, pump house", "Pumping Station", "E&I subcontractor", "2026-10-12", "2026-10-23", 0.0, None),
    ("A3040", "Earthing grid, substation", "Pumping Station", "E&I subcontractor", "2026-09-29", "2026-10-09", 0.0, "Permit to dig"),
    ("A3050", "Transformer plinths", "Pumping Station", "Civil crew A", "2026-10-13", "2026-10-20", 0.0, None),
    ("A3060", "Receive main switchgear", "Pumping Station", "Material controller", "2026-10-16", "2026-10-16", 0.0, "Delivery slot and heavy lift plan"),
    ("A3070", "Overhead crane rails", "Pumping Station", "Mechanical crew", "2026-10-19", "2026-10-27", 0.0, None),
    ("A4010", "Reservoir wall waterproofing", "Reservoir", "Civil crew B", "2026-09-22", "2026-10-05", 0.30, None),
    ("A4020", "Reservoir roof formwork", "Reservoir", "Civil crew B", "2026-10-06", "2026-10-20", 0.0, "Props hired from week 41"),
    ("A4030", "Reservoir inlet valve chamber", "Reservoir", "Civil crew A", "2026-10-14", "2026-10-26", 0.0, None),
    ("A4040", "Leak test preparation", "Reservoir", "QA/QC", "2026-10-26", "2026-11-06", 0.0, "Water source agreed with client"),
    ("A5010", "Temporary power to filter area", "Site-wide", "E&I subcontractor", "2026-09-30", "2026-10-02", 0.0, None),
    ("A5020", "Access road resurfacing", "Site-wide", "Civil crew C", "2026-10-05", "2026-10-09", 0.0, None),
    ("A5030", "Monthly HSE audit", "Site-wide", "HSE", "2026-10-08", "2026-10-08", 0.0, None),
    ("A5040", "Survey of as-built foundations", "Site-wide", "Survey team", "2026-10-19", "2026-10-23", 0.0, None),
    ("A5050", "Laydown area extension", "Site-wide", "Civil crew C", "2026-10-26", "2026-10-30", 0.0, None),
    ("A1060", "Intake pipe tie-in (shutdown)", "Intake", "Piping crew", "2026-11-02", "2026-11-04", 0.0, "Client shutdown window"),
    ("A2090", "Clarifier mechanism delivery", "Treatment", "Material controller", "2026-09-18", "2026-09-18", 1.0, None),
]


def build_lookahead():
    rows = [(i, a, ar, ow, d(s), d(f), p, c) for i, a, ar, ow, s, f, p, c in LOOKAHEAD]
    wb = new_workbook()
    write_readme(wb.active, "3-week look-ahead: sample data dictionary", COMMON_STEPS + [
        "One row per activity. Export it from your schedule (Primavera P6, MS Project) or type it. Include at least the next six weeks of work.",
        "The look-ahead window starts on the Monday after the Data Date and runs for Look-ahead Weeks (3 by default).",
        "An activity shows when it is not complete and its Start-Finish span touches the window: it starts, finishes or runs through it.",
        "Write the open constraint in Constraint (material, permit, drawing, access, crane). Leave it empty once cleared.",
    ], SETTINGS_DOC + [
        ("Settings", "Look-ahead Weeks", "Whole number", "Yes", "Length of the window in weeks.", "3"),
        ("Activities", "Activity ID", "Text", "Yes", "Unique activity code, as in the schedule.", "A2030"),
        ("Activities", "Activity", "Text", "Yes", "Activity name.", "Filter building steel erection"),
        ("Activities", "Area", "Text", "Yes", "Area or work front. Drives the slicer and row grouping.", "Treatment"),
        ("Activities", "Owner", "Text", "Yes", "Crew, subcontractor or person who does the work.", "Steel subcontractor"),
        ("Activities", "Start Date", "Date", "Yes", "Planned or actual start.", "2026-09-17"),
        ("Activities", "Finish Date", "Date", "Yes", "Planned or forecast finish. Same as Start for a one-day task or a milestone.", "2026-10-16"),
        ("Activities", "% Complete", "Number 0 to 1", "Yes", "Physical progress as a fraction (0.35 = 35%). Activities at 1 drop out of the look-ahead.", "0.35"),
        ("Activities", "Constraint", "Text", "No", "What must be cleared before the work can start or continue. Empty = no constraint.", "Crane booking weeks 41 and 42"),
    ])
    settings_sheet(wb, ["Look-ahead Weeks"], [3])
    ws = wb.create_sheet("Activities")
    write_table(ws, "Activities", ["Activity ID", "Activity", "Area", "Owner", "Start Date", "Finish Date", "% Complete", "Constraint"], rows,
                widths={"Activity": 36, "Area": 18, "Owner": 22, "Constraint": 40},
                formats={"Start Date": "yyyy-mm-dd", "Finish Date": "yyyy-mm-dd", "% Complete": "0%"})
    out = ROOT / "three-week-lookahead" / "data" / "sample_data.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out, len(rows)


if __name__ == "__main__":
    for fn in (build_progress, build_milestones, build_raci, build_lookahead):
        path, n = fn()
        print(f"{path}  rows={n}")
