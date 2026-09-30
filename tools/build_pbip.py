"""Generate the four PBIP projects (report + semantic model).

Usage:
  python build_pbip.py            -> DataFile parameter = neutral default path (release)
  python build_pbip.py --local    -> DataFile parameter = this repo's data/sample_data.xlsx (for testing)

Author: Badr Abardazzou. License: MIT.
"""
import sys
from pathlib import Path

from pbir import (BAD, C, FONT, FONT_BOLD, GOOD, GREY, INK, LINE, M, MUTED, TEAL, TEAL_DARK, TEAL_LIGHT, WARN,
                  Table, b, categorical_filter, cell_format, chart_axes, color, container, excel_source, i,
                  kata_theme, kpi_card, lit, mcolor, n, proj, qref, s, slicer, subtitle_card, table_common, tag,
                  textbox, vco, write_model, write_pbip, write_report)

ROOT = Path(__file__).resolve().parent.parent
KATA = ROOT / "tools" / "kata.json"  # vendored from github.com/babarda/report-themes
BASE_THEME = ROOT / "tools" / "CY26SU02.json"
LOCAL = "--local" in sys.argv

SETTINGS_DATA_DATE = """VAR d = MAX ( Settings[Data Date] )
RETURN IF ( ISBLANK ( d ), TODAY (), d )"""


def settings_table(model, extra_cols=(), extra_m_types=""):
    t = Table(model, "Settings", "One row read from the Settings sheet: project name, data date and template options.")
    t.column("Project Name", "string", source="Project Name")
    t.column("Data Date", "dateTime", source="Data Date", fmt="dd mmm yyyy")
    for name, dtype in extra_cols:
        t.column(name, dtype, source=name, summarize="none")
    t.measure("Data Date Used", SETTINGS_DATA_DATE, fmt="dd mmm yyyy",
              desc="Status date of the report. Settings[Data Date], or today when the cell is empty.")
    t.measure("Report Subtitle",
              'MAX ( Settings[Project Name] ) & "   |   Data date " & FORMAT ( [Data Date Used], "dd mmm yyyy" )',
              desc="Header line: project name and data date.")
    types = '{"Project Name", type text}, {"Data Date", type date}' + extra_m_types
    t.m_partition(f"""let
    {excel_source("Settings")}
    Typed = Table.TransformColumnTypes(Headers, {{{types}}}),
    FirstRow = Table.FirstN(Table.SelectRows(Typed, each [Project Name] <> null), 1)
in
    FirstRow""")
    return t


def header(prefix, title, subtitle, sub_entity="Settings"):
    return [
        textbox(f"{prefix}_title", 20, 14, 800, 64, 1000, title, subtitle),
    ]


def data_file(slug):
    if LOCAL:
        return str(ROOT / slug / "data" / "sample_data.xlsx")
    return f"C:\\PowerBI-Templates\\{slug}\\sample_data.xlsx"


def ids(slug):
    return tag(slug, "report"), tag(slug, "model")


# =================================================================== 1
def build_progress():
    slug, name = "progress-s-curve", "Progress S-Curve"
    st = settings_table(name)
    p = Table(name, "Progress", "One row per week, discipline and area: planned, earned and forecast hours.")
    p.column("Week Ending", "dateTime", source="Week Ending", fmt="dd mmm yy")
    p.column("Discipline", "string", source="Discipline")
    p.column("Area", "string", source="Area")
    p.column("Planned Hours", "double", source="Planned Hours", fmt="#,0", summarize="sum")
    p.column("Earned Hours", "double", source="Earned Hours", fmt="#,0", summarize="sum")
    p.column("Forecast Hours", "double", source="Forecast Hours", fmt="#,0", summarize="sum")
    F = "Progress[Week Ending]"
    p.measure("Budget Hours", f"CALCULATE ( SUM ( Progress[Planned Hours] ), REMOVEFILTERS ( {F} ) )", fmt="#,0",
              folder="Base", desc="Budget at completion in hours: all planned hours of the selection, every week.")
    p.measure("Last Actual Week", f"""VAR dd = [Data Date Used]
RETURN CALCULATE ( MAX ( {F} ), REMOVEFILTERS ( {F} ), {F} <= dd )""", fmt="dd mmm yyyy", folder="Base",
              desc="Last week ending on or before the data date. Actuals stop here, the forecast starts here.")
    p.measure("Planned % Cum", f"""VAR w = MAX ( {F} )
RETURN DIVIDE ( CALCULATE ( SUM ( Progress[Planned Hours] ), REMOVEFILTERS ( {F} ), {F} <= w ), [Budget Hours] )""",
              fmt="0.0%", folder="S-curve", desc="Baseline cumulative progress: planned hours up to the week / budget hours.")
    p.measure("Actual % Cum", f"""VAR w = MAX ( {F} )
VAR lw = [Last Actual Week]
RETURN
    IF ( w <= lw,
        DIVIDE ( CALCULATE ( SUM ( Progress[Earned Hours] ), REMOVEFILTERS ( {F} ), {F} <= w ), [Budget Hours] ) )""",
              fmt="0.0%", folder="S-curve", desc="Actual cumulative progress: earned hours up to the week / budget hours. Blank after the data date.")
    p.measure("Forecast % Cum", f"""VAR w = MAX ( {F} )
VAR lw = [Last Actual Week]
VAR earned = CALCULATE ( SUM ( Progress[Earned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw )
VAR fc = CALCULATE ( SUM ( Progress[Forecast Hours] ), REMOVEFILTERS ( {F} ), {F} > lw, {F} <= w )
RETURN IF ( w >= lw, DIVIDE ( earned + fc, [Budget Hours] ) )""",
              fmt="0.0%", folder="S-curve", desc="Forecast cumulative progress: earned hours to the data date plus forecast hours after it, / budget hours. Starts at the data date.")
    p.measure("Planned % Week", "DIVIDE ( SUM ( Progress[Planned Hours] ), [Budget Hours] )", fmt="0.0%", folder="S-curve",
              desc="Share of the budget planned in the week.")
    p.measure("Actual % Week", f"""VAR w = MAX ( {F} )
RETURN IF ( w <= [Last Actual Week], DIVIDE ( SUM ( Progress[Earned Hours] ), [Budget Hours] ) )""",
              fmt="0.0%", folder="S-curve", desc="Share of the budget earned in the week. Blank after the data date.")
    p.measure("Planned % To Date", f"""VAR lw = [Last Actual Week]
RETURN DIVIDE ( CALCULATE ( SUM ( Progress[Planned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw ), [Budget Hours] )""",
              fmt="0.0%", folder="KPI", desc="Planned cumulative progress at the data date.")
    p.measure("Actual % To Date", f"""VAR lw = [Last Actual Week]
RETURN DIVIDE ( CALCULATE ( SUM ( Progress[Earned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw ), [Budget Hours] )""",
              fmt="0.0%", folder="KPI", desc="Actual cumulative progress at the data date.")
    p.measure("Variance To Date", "[Actual % To Date] - [Planned % To Date]", fmt="+0.0%;-0.0%;0.0%", folder="KPI",
              desc="Actual minus planned progress at the data date, in percentage points. Negative = behind.")
    p.measure("SPI", f"""VAR lw = [Last Actual Week]
VAR ev = CALCULATE ( SUM ( Progress[Earned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw )
VAR pv = CALCULATE ( SUM ( Progress[Planned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw )
RETURN DIVIDE ( ev, pv )""", fmt="0.00", folder="KPI",
              desc="Schedule performance index = earned hours / planned hours, both cumulative to the data date. 1.00 = on plan, below 1 = behind.")
    p.measure("SPI Color", """VAR v = [SPI]
RETURN SWITCH ( TRUE (), ISBLANK ( v ), "#AAAAAA", v >= 0.95, "#339966", v >= 0.90, "#E08A00", "#D63A2F" )""",
              folder="Format", desc="Traffic light for SPI: green from 0.95, amber from 0.90, red below.")
    p.measure("Variance Color", """VAR v = [Variance To Date]
RETURN SWITCH ( TRUE (), ISBLANK ( v ), "#AAAAAA", v >= -0.02, "#339966", v >= -0.05, "#E08A00", "#D63A2F" )""",
              folder="Format", desc="Traffic light for the variance: green down to -2 points, amber down to -5, red below.")
    p.measure("Planned Finish", f"CALCULATE ( MAX ( {F} ), REMOVEFILTERS ( {F} ), Progress[Planned Hours] > 0 )",
              fmt="dd mmm yyyy", folder="KPI", desc="Last week with planned hours in the baseline.")
    p.measure("Forecast Finish", f"""VAR lw = [Last Actual Week]
VAR bud = [Budget Hours]
VAR earned = CALCULATE ( SUM ( Progress[Earned Hours] ), REMOVEFILTERS ( {F} ), {F} <= lw )
VAR weeks =
    FILTER ( CALCULATETABLE ( VALUES ( {F} ), REMOVEFILTERS ( {F} ) ), {F} > lw )
VAR withCum =
    ADDCOLUMNS ( weeks, "@cum",
        VAR w = {F}
        RETURN earned + CALCULATE ( SUM ( Progress[Forecast Hours] ), REMOVEFILTERS ( {F} ), {F} > lw, {F} <= w ) )
RETURN
    IF ( earned >= bud * 0.999, lw, MINX ( FILTER ( withCum, [@cum] >= bud * 0.999 ), {F} ) )""",
              fmt="dd mmm yyyy", folder="KPI", desc="First week where the forecast curve reaches 99.9% of the budget.")
    p.measure("Finish Slip Weeks", "DIVIDE ( INT ( [Forecast Finish] - [Planned Finish] ), 7 )", fmt="+0;-0;0", folder="KPI",
              desc="Forecast finish minus planned finish, in weeks. Positive = late.")
    p.measure("Finish Slip Label", """VAR w = [Finish Slip Weeks]
RETURN FORMAT ( [Forecast Finish], "dd mmm yy" ) & IF ( ISBLANK ( w ), "", "  (" & FORMAT ( w, "+0;-0;0" ) & " wk)" )""",
              folder="KPI", desc="Forecast finish date and its slip against the baseline finish, in weeks.")
    p.measure("Finish Color", """VAR w = [Finish Slip Weeks]
RETURN SWITCH ( TRUE (), ISBLANK ( w ), "#AAAAAA", w <= 0, "#339966", w <= 4, "#E08A00", "#D63A2F" )""",
              folder="Format", desc="Green on or before the baseline finish, amber up to 4 weeks late, red beyond.")
    p.m_partition(f"""let
    {excel_source("Progress")}
    Typed = Table.TransformColumnTypes(Headers, {{{{"Week Ending", type date}}, {{"Discipline", type text}}, {{"Area", type text}}, {{"Planned Hours", type number}}, {{"Earned Hours", type number}}, {{"Forecast Hours", type number}}}}),
    Clean = Table.SelectRows(Typed, each [Week Ending] <> null and [Discipline] <> null)
in
    Clean""")
    tables = [st, p]

    E = "Progress"
    week = C(E, "Week Ending")
    curve = {
        "visualType": "lineClusteredColumnComboChart",
        "query": {"queryState": {
            "Category": {"projections": [proj(week, "Week ending", active=True)]},
            "Y": {"projections": [proj(M(E, "Planned % Week"), "Planned, weekly"), proj(M(E, "Actual % Week"), "Actual, weekly")]},
            "Y2": {"projections": [proj(M(E, "Planned % Cum"), "Planned, cumulative"),
                                   proj(M(E, "Actual % Cum"), "Actual, cumulative"),
                                   proj(M(E, "Forecast % Cum"), "Forecast, cumulative")]}},
            "sortDefinition": {"sort": [{"field": week, "direction": "Ascending"}], "isDefaultSort": True}},
        "objects": {
            **chart_axes(),
            "valueAxis": [{"properties": {"fontSize": n(8), "labelColor": color(MUTED), "showAxisTitle": b(False),
                                          "gridlineShow": b(False), "secShow": b(True), "secFontSize": n(8),
                                          "secLabelColor": color(MUTED), "secShowAxisTitle": b(False),
                                          "secStart": n(0), "secEnd": n(1)}}],
            "categoryAxis": [{"properties": {"fontSize": n(8), "labelColor": color(MUTED), "showAxisTitle": b(False),
                                             "axisType": s("Scalar")}}],
            "dataPoint": [
                {"properties": {"fill": color("#D5DEDE")}, "selector": {"metadata": "Progress.Planned % Week"}},
                {"properties": {"fill": color("#7FBFBF")}, "selector": {"metadata": "Progress.Actual % Week"}},
                {"properties": {"fill": color("#6B7777")}, "selector": {"metadata": "Progress.Planned % Cum"}},
                {"properties": {"fill": color(TEAL)}, "selector": {"metadata": "Progress.Actual % Cum"}},
                {"properties": {"fill": color(WARN)}, "selector": {"metadata": "Progress.Forecast % Cum"}}],
            "lineStyles": [
                {"properties": {"strokeWidth": n(2), "lineStyle": s("solid"), "showMarker": b(False)}},
                {"properties": {"strokeWidth": n(2), "lineStyle": s("dashed")}, "selector": {"metadata": "Progress.Planned % Cum"}},
                {"properties": {"strokeWidth": n(3), "lineStyle": s("solid")}, "selector": {"metadata": "Progress.Actual % Cum"}},
                {"properties": {"strokeWidth": n(2), "lineStyle": s("dotted")}, "selector": {"metadata": "Progress.Forecast % Cum"}}],
        },
        "visualContainerObjects": vco("Cumulative progress: planned, actual and forecast", subtitle_measure=("Settings", "Report Subtitle")),
    }
    var_tbl = {
        "visualType": "pivotTable",
        "query": {"queryState": {
            "Rows": {"projections": [proj(C(E, "Discipline")), proj(C(E, "Area"))]},
            "Values": {"projections": [proj(M(E, "Budget Hours"), "Budget h"), proj(M(E, "Planned % To Date"), "Planned"),
                                       proj(M(E, "Actual % To Date"), "Actual"), proj(M(E, "Variance To Date"), "Var."),
                                       proj(M(E, "SPI"))]}}},
        "objects": {**table_common(),
                    "values": table_common()["values"] + [
                        cell_format("fontColor", E, "SPI Color", "Progress.SPI"),
                        cell_format("fontColor", E, "Variance Color", "Progress.Variance To Date")],
                    "rowHeaders": [{"properties": {"fontSize": n(9), "fontFamily": s(FONT), "fontColor": color(INK),
                                                   "stepped": b(True), "showExpandCollapseButtons": b(True)}}],
                    "total": [{"properties": {"fontFamily": s(FONT_BOLD), "fontSize": n(9), "backColor": color("#EEF4F4")}}],
                    "subTotals": [{"properties": {"rowSubtotals": b(True)}}]},
        "visualContainerObjects": {**vco("Variance at the data date, by discipline (expand for areas)"),
                                   "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    visuals = header("p1", "Progress S-curve", "Planned vs actual vs forecast, cumulative % of budget hours")
    visuals += [
        slicer("p1_sl_disc", 860, 12, 196, 78, 1200, C(E, "Discipline"), "Discipline"),
        slicer("p1_sl_area", 1064, 12, 196, 78, 1300, C(E, "Area"), "Area"),
    ]
    cards = [("Planned % To Date", "Planned to date", TEAL_LIGHT, None),
             ("Actual % To Date", "Actual to date", TEAL, None),
             ("Variance To Date", "Variance (points)", None, (E, "Variance Color")),
             ("SPI", "SPI (earned / planned)", None, (E, "SPI Color")),
             ("Planned Finish", "Baseline finish", GREY, None),
             ("Finish Slip Label", "Forecast finish (slip vs baseline)", None, (E, "Finish Color"))]
    x = 20
    for k, (meas, label, acc, accm) in enumerate(cards):
        w = 196 if k < 5 else 220
        visuals.append(kpi_card(f"p1_k{k}", x, 96, w, 78, 2000 + k * 100, E, meas, label, accent=acc or TEAL,
                                accent_measure=accm, value_size=15 if k == 5 else 20))
        x += w + 8
    bars = {
        "visualType": "clusteredBarChart",
        "query": {"queryState": {
            "Category": {"projections": [proj(C(E, "Discipline"), active=True)]},
            "Y": {"projections": [proj(M(E, "Planned % To Date"), "Planned to date"), proj(M(E, "Actual % To Date"), "Actual to date")]}},
            "sortDefinition": {"sort": [{"field": C(E, "Discipline"), "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {
            **chart_axes(),
            "valueAxis": [{"properties": {"show": b(False), "gridlineShow": b(False), "showAxisTitle": b(False)}}],
            "labels": [{"properties": {"show": b(True), "fontSize": n(8), "color": color(INK)}}],
            "dataPoint": [
                {"properties": {"fill": color("#C9D3D3")}, "selector": {"metadata": "Progress.Planned % To Date"}},
                {"properties": {"fill": color(TEAL)}, "selector": {"metadata": "Progress.Actual % To Date"}}]},
        "visualContainerObjects": vco("Planned vs actual to date, by discipline"),
    }
    visuals += [container("p1_curve", 20, 180, 780, 528, curve, 3000),
                container("p1_var", 808, 180, 452, 250, var_tbl, 3100),
                container("p1_bars", 808, 436, 452, 272, bars, 3200)]
    pages = [{"name": "scurve", "displayName": "S-curve", "visuals": visuals}]
    return slug, name, tables, [], pages, ["DataFile", "Settings", "Progress"]


# =================================================================== 2
def build_milestones():
    slug, name = "milestone-tracker", "Milestone Tracker"
    st = settings_table(name, [("Late Threshold Days", "int64")], ', {"Late Threshold Days", Int64.Type}')
    st.measure("Late Threshold", "VAR t = MAX ( Settings[Late Threshold Days] ) RETURN IF ( ISBLANK ( t ), 30, t )", fmt="0",
               desc="Slip in days above which an open milestone is Late. Default 30.")
    m = Table(name, "Milestones", "One row per milestone: baseline, forecast and actual dates.")
    m.column("Milestone ID", "string", source="Milestone ID")
    m.column("Milestone", "string", source="Milestone")
    m.column("Phase", "string", source="Phase")
    m.column("Area", "string", source="Area")
    m.column("Owner", "string", source="Owner")
    m.column("Baseline Date", "dateTime", source="Baseline Date", fmt="dd mmm yy")
    m.column("Forecast Date", "dateTime", source="Forecast Date", fmt="dd mmm yy")
    m.column("Actual Date", "dateTime", source="Actual Date", fmt="dd mmm yy")
    m.column("Finish Date", "dateTime", fmt="dd mmm yy",
             expr="IF ( NOT ISBLANK ( Milestones[Actual Date] ), Milestones[Actual Date], Milestones[Forecast Date] )",
             desc="Actual date when achieved, otherwise the forecast date.")
    m.column("Slip Days", "int64", fmt="+0;-0;0", summarize="none",
             expr="INT ( Milestones[Finish Date] - Milestones[Baseline Date] )",
             desc="Finish date minus baseline date, in calendar days. Positive = later than baseline.")
    m.column("Status", "string", expr="""VAR dd = MAX ( Settings[Data Date] )
VAR ddUsed = IF ( ISBLANK ( dd ), TODAY (), dd )
VAR th = MAX ( Settings[Late Threshold Days] )
VAR thUsed = IF ( ISBLANK ( th ), 30, th )
RETURN
    SWITCH ( TRUE (),
        NOT ISBLANK ( Milestones[Actual Date] ), "Complete",
        Milestones[Forecast Date] < ddUsed, "Late",
        Milestones[Slip Days] > thUsed, "Late",
        Milestones[Slip Days] > 0, "At risk",
        "On track" )""",
             desc="Complete = actual date filled. Late = open and forecast before the data date, or slip above the threshold. At risk = slip above 0. On track = the rest.")
    m.column("Status Order", "int64", hidden=True,
             expr='SWITCH ( Milestones[Status], "Late", 1, "At risk", 2, "On track", 3, "Complete", 4, 5 )')
    m.column("Status Color", "string", hidden=True,
             expr='SWITCH ( Milestones[Status], "Late", "#D63A2F", "At risk", "#E08A00", "On track", "#339966", "Complete", "#005F5F", "#AAAAAA" )')
    m.column("Status Tint", "string", hidden=True,
             expr='SWITCH ( Milestones[Status], "Late", "#F8D8D5", "At risk", "#FBE8C8", "On track", "#D8EEE1", "Complete", "#DCEBEB", "#FFFFFF" )')
    m.column("Baseline Order", "int64", hidden=True,
             expr="RANKX ( ALL ( Milestones ), Milestones[Baseline Date], , ASC, DENSE )")
    m.column("Milestone Label", "string", expr='Milestones[Milestone ID] & "  " & Milestones[Milestone]', sort_by="Baseline Order",
             desc="Milestone ID and name, sorted by baseline date.")
    m.measure("Milestone Count", "COUNTROWS ( Milestones )", fmt="0", folder="KPI", desc="Number of milestones.")
    for st_name in ("Complete", "On track", "At risk", "Late"):
        m.measure(st_name if st_name != "Complete" else "Completed",
                  f'CALCULATE ( COUNTROWS ( Milestones ), Milestones[Status] = "{st_name}" ) + 0', fmt="0", folder="KPI",
                  desc=f"Number of milestones with status {st_name}.")
    m.measure("Average Slip Open", "AVERAGEX ( FILTER ( Milestones, ISBLANK ( Milestones[Actual Date] ) ), Milestones[Slip Days] )",
              fmt="+0.0;-0.0;0.0", folder="KPI", desc="Average slip in days of the milestones not yet achieved.")
    m.measure("Slip (days)", "MAX ( Milestones[Slip Days] )", fmt="+0;-0;0", folder="Chart",
              desc="Slip of the milestone in context, in days.")
    m.measure("Slip Bar Color", "MAX ( Milestones[Status Color] )", folder="Format", desc="Status color of the milestone in context.")
    m.measure("Status Tint Color", "MAX ( Milestones[Status Tint] )", folder="Format", desc="Light status color for table cells.")
    m.measure("Status Font Color", "MAX ( Milestones[Status Color] )", folder="Format", desc="Status color for text.")
    tl_vars = """VAR ms = MIN ( Months[Month Start] )
VAR me = EOMONTH ( ms, 0 )
VAR bl = MAX ( Milestones[Baseline Date] )
VAR fi = MAX ( Milestones[Finish Date] )
VAR inB = bl >= ms && bl <= me
VAR inF = fi >= ms && fi <= me
VAR lo = MIN ( bl, fi )
VAR hi = MAX ( bl, fi )
VAR between = ms > lo && me < hi"""
    m.measure("Timeline Mark", tl_vars + """
RETURN
    IF ( HASONEVALUE ( Milestones[Milestone ID] ),
        SWITCH ( TRUE (), inF, UNICHAR ( 9670 ), inB, UNICHAR ( 9671 ), between, UNICHAR ( 9644 ), "" ) )""",
              folder="Chart", desc="Milestone chart cell: filled diamond = forecast or actual month, hollow diamond = baseline month, bar = months of slip in between.")
    m.measure("Timeline Color", tl_vars + """
RETURN
    SWITCH ( TRUE (),
        inF, MAX ( Milestones[Status Color] ),
        inB, "#8A9494",
        between && fi > bl, "#EBA8A1",
        between, "#9CCFB0",
        "#FFFFFF" )""", folder="Format", desc="Font color of the milestone chart cell.")
    m.measure("Timeline Back", """VAR ms = MIN ( Months[Month Start] )
VAR dd = [Data Date Used]
RETURN IF ( dd >= ms && dd <= EOMONTH ( ms, 0 ), "#E6F2F2", "#FFFFFF" )""", folder="Format",
              desc="Highlights the month that contains the data date.")
    m.m_partition(f"""let
    {excel_source("Milestones")}
    Typed = Table.TransformColumnTypes(Headers, {{{{"Milestone ID", type text}}, {{"Milestone", type text}}, {{"Phase", type text}}, {{"Area", type text}}, {{"Owner", type text}}, {{"Baseline Date", type date}}, {{"Forecast Date", type date}}, {{"Actual Date", type date}}}}),
    Clean = Table.SelectRows(Typed, each [Milestone ID] <> null and [Milestone ID] <> "")
in
    Clean""")
    mo = Table(name, "Months", "Calculated calendar of months spanning every baseline and finish date.")
    mo.column("Month Start", "dateTime", source="[Month Start]", fmt="mmm yyyy", calc_table_col=True)
    mo.column("Month", "string", source="[Month]", sort_by="Month Start", calc_table_col=True)
    mo.calc_partition("""VAR mn0 = MIN ( MIN ( Milestones[Baseline Date] ), MIN ( Milestones[Finish Date] ) )
VAR mx0 = MAX ( MAX ( Milestones[Baseline Date] ), MAX ( Milestones[Finish Date] ) )
VAR mn = IF ( ISBLANK ( mn0 ), TODAY (), mn0 )
VAR mx = IF ( ISBLANK ( mx0 ), TODAY (), mx0 )
RETURN
    SELECTCOLUMNS (
        FILTER ( CALENDAR ( DATE ( YEAR ( mn ), MONTH ( mn ), 1 ), EOMONTH ( mx, 0 ) ), DAY ( [Date] ) = 1 ),
        "Month Start", [Date],
        "Month", FORMAT ( [Date], "mmm" ) & " " & FORMAT ( [Date], "yy" )
    )""")
    tables = [st, m, mo]
    E = "Milestones"
    tl = {
        "visualType": "pivotTable",
        "query": {"queryState": {
            "Rows": {"projections": [proj(C(E, "Milestone Label"), "Milestone")]},
            "Columns": {"projections": [proj(C("Months", "Month"))]},
            "Values": {"projections": [proj(M(E, "Timeline Mark"), " ")]}},
            "sortDefinition": {"sort": [{"field": C(E, "Milestone Label"), "direction": "Ascending"}], "isDefaultSort": True}},
        "objects": {
            **{k: v for k, v in table_common(grow=False).items() if k != "total"},
            "columnHeaders": [{"properties": {"fontColor": color("#FFFFFF"), "backColor": color(TEAL_DARK), "fontSize": n(8),
                                              "fontFamily": s(FONT_BOLD), "wordWrap": b(True), "alignment": s("Center"),
                                              "columnAdjustment": s("fixedWidth"), "defaultColumnWidth": n(29)}}],
            "values": [{"properties": {"fontSize": n(11), "fontFamily": s(FONT), "backColorPrimary": color("#FFFFFF"),
                                       "backColorSecondary": color("#FFFFFF")}},
                       cell_format("fontColor", E, "Timeline Color", "Milestones.Timeline Mark"),
                       cell_format("backColor", E, "Timeline Back", "Milestones.Timeline Mark")],
            "rowHeaders": [{"properties": {"fontSize": n(9), "fontFamily": s(FONT), "fontColor": color(INK), "wordWrap": b(False)}}],
            "columnWidth": [{"properties": {"value": n(290)}, "selector": {"metadata": "Milestones.Milestone Label"}}],
            "subTotals": [{"properties": {"rowSubtotals": b(False), "columnSubtotals": b(False)}}],
            "grid": [{"properties": {"gridVertical": b(True), "gridVerticalColor": color("#F0F3F3"), "gridHorizontal": b(True),
                                     "gridHorizontalColor": color("#EEF1F1"), "rowPadding": n(2)}}]},
        "visualContainerObjects": {**vco("Milestone chart: hollow diamond = baseline, filled diamond = forecast or actual (color = status)", subtitle_measure=("Settings", "Report Subtitle")),
                                   "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    slip = {
        "visualType": "barChart",
        "query": {"queryState": {
            "Category": {"projections": [proj(C(E, "Milestone Label"), "Milestone", active=True)]},
            "Y": {"projections": [proj(M(E, "Slip (days)"), "Slip (days)")]},
            "Tooltips": {"projections": [proj(M(E, "Timeline Mark"), "Mark")]}},
            "sortDefinition": {"sort": [{"field": M(E, "Slip (days)"), "direction": "Descending"}], "isDefaultSort": False}},
        "objects": {
            "categoryAxis": [{"properties": {"fontSize": n(8), "labelColor": color(INK), "showAxisTitle": b(False),
                                             "maxMarginFactor": n(45)}}],
            "valueAxis": [{"properties": {"show": b(False), "gridlineShow": b(False), "showAxisTitle": b(False)}}],
            "legend": [{"properties": {"show": b(False)}}],
            "labels": [{"properties": {"show": b(True), "fontSize": n(8), "color": color(INK)}}],
            "dataPoint": [{"properties": {"fill": mcolor(E, "Slip Bar Color")},
                           "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]},
        "visualContainerObjects": vco("Slip ranking, days against baseline"),
    }
    visuals = header("m1", "Milestone tracker", "Baseline vs forecast vs actual, slip and status per milestone")
    visuals += [slicer("m1_sl_phase", 860, 12, 196, 78, 1200, C(E, "Phase"), "Phase"),
                slicer("m1_sl_area", 1064, 12, 196, 78, 1300, C(E, "Area"), "Area")]
    cards = [("Milestone Count", "Milestones", GREY), ("Completed", "Complete", TEAL_DARK), ("On track", "On track", GOOD),
             ("At risk", "At risk", WARN), ("Late", "Late", BAD), ("Average Slip Open", "Average slip, open (days)", GREY)]
    x = 20
    for k, (meas, label, acc) in enumerate(cards):
        visuals.append(kpi_card(f"m1_k{k}", x, 96, 199, 78, 2000 + k * 100, E, meas, label, accent=acc))
        x += 207
    visuals += [container("m1_tl", 20, 180, 880, 528, tl, 3000),
                container("m1_slip", 908, 180, 352, 528, slip, 3100)]

    reg_cols = [("Milestone ID", "ID"), ("Milestone", None), ("Phase", None), ("Area", None), ("Owner", None),
                ("Baseline Date", "Baseline"), ("Forecast Date", "Forecast"), ("Actual Date", "Actual"),
                ("Slip Days", "Slip (d)"), ("Status", None)]
    reg = {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(C(E, c), d) for c, d in reg_cols]}},
                  "sortDefinition": {"sort": [{"field": C(E, "Baseline Date"), "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {**table_common(font_size=10),
                    "values": table_common(font_size=10)["values"] + [
                        cell_format("backColor", E, "Status Tint Color", "Milestones.Status"),
                        cell_format("fontColor", E, "Status Font Color", "Milestones.Slip Days")],
                    "total": [{"properties": {"totals": b(False)}}]},
        "visualContainerObjects": {**vco("Milestone register", subtitle_measure=("Settings", "Report Subtitle")), "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    v2 = header("m2", "Milestone register", "Every milestone with its dates, slip and status")
    v2 += [slicer("m2_sl_phase", 860, 12, 196, 78, 1200, C(E, "Phase"), "Phase"),
           slicer("m2_sl_area", 1064, 12, 196, 78, 1300, C(E, "Area"), "Area"),
           container("m2_reg", 20, 96, 1240, 612, reg, 3000)]
    pages = [{"name": "milestones", "displayName": "Milestones", "visuals": visuals},
             {"name": "register", "displayName": "Register", "visuals": v2}]
    return slug, name, tables, [], pages, ["DataFile", "Settings", "Milestones"]


# =================================================================== 3
def build_raci():
    slug, name = "raci-matrix", "RACI Matrix"
    st = Table(name, "Settings", "One row read from the Settings sheet: project name.")
    st.column("Project Name", "string", source="Project Name")
    st.measure("Report Subtitle", 'MAX ( Settings[Project Name] ) & "   |   R = Responsible, A = Accountable, C = Consulted, I = Informed"',
               desc="Header line: project name and code legend.")
    st.m_partition(f"""let
    {excel_source("Settings")}
    Typed = Table.TransformColumnTypes(Headers, {{{{"Project Name", type text}}}}),
    FirstRow = Table.FirstN(Table.SelectRows(Typed, each [Project Name] <> null), 1)
in
    FirstRow""")
    act = Table(name, "Activities", "One row per activity: the first three columns of the RACI sheet.")
    act.column("Activity ID", "string", source="Activity ID")
    act.column("Activity", "string", source="Activity")
    act.column("Phase", "string", source="Phase")
    act.column("Activity Order", "int64", source="Activity Order", hidden=True)
    act.column("Activity Label", "string", expr='Activities[Activity ID] & "  " & Activities[Activity]', sort_by="Activity Order",
               desc="Activity ID and name, in sheet order.")
    act.column("A Count", "int64", expr='CALCULATE ( COUNTROWS ( RACI ), RACI[Code] = "A" ) + 0', summarize="sum",
               desc="Number of roles marked Accountable on the activity.")
    act.column("R Count", "int64", expr='CALCULATE ( COUNTROWS ( RACI ), RACI[Code] = "R" ) + 0', summarize="sum",
               desc="Number of roles marked Responsible on the activity.")
    act.column("Check", "string", expr="""SWITCH ( TRUE (),
    Activities[A Count] = 0, "No Accountable",
    Activities[A Count] > 1, "More than one Accountable",
    Activities[R Count] = 0, "No Responsible",
    "OK" )""", desc="Rule check: exactly one A and at least one R per activity.")
    act.column("Has Issue", "string", expr='IF ( Activities[Check] = "OK", "No", "Yes" )')
    act.measure("Activities Count", "COUNTROWS ( Activities )", fmt="0", folder="KPI", desc="Number of activities.")
    act.measure("Issues", 'CALCULATE ( COUNTROWS ( Activities ), Activities[Has Issue] = "Yes" ) + 0', fmt="0", folder="KPI",
                desc="Activities that break the rule: no A, several A, or no R.")
    act.measure("Issues Color", 'IF ( [Issues] > 0, "#D63A2F", "#339966" )', folder="Format", desc="Red when any activity fails the check.")
    act.measure("Check Color", 'IF ( SELECTEDVALUE ( Activities[Check] ) = "OK", "#339966", "#D63A2F" )', folder="Format",
                desc="Font color for the check column.")
    act.m_partition(f"""let
    {excel_source("RACI")}
    Keep = Table.SelectColumns(Headers, {{"Activity ID", "Activity", "Phase"}}),
    Typed = Table.TransformColumnTypes(Keep, {{{{"Activity ID", type text}}, {{"Activity", type text}}, {{"Phase", type text}}}}),
    Clean = Table.SelectRows(Typed, each [Activity ID] <> null and [Activity ID] <> ""),
    Indexed = Table.AddIndexColumn(Clean, "Activity Order", 1, 1, Int64.Type)
in
    Indexed""")
    ro = Table(name, "Roles", "One row per role: every column header after Phase on the RACI sheet.")
    ro.column("Role", "string", source="Role", sort_by="Role Order")
    ro.column("Role Order", "int64", source="Role Order", hidden=True)
    ro.measure("Roles Count", "COUNTROWS ( Roles )", fmt="0", folder="KPI", desc="Number of roles (columns) in the matrix.")
    ro.m_partition(f"""let
    {excel_source("RACI")}
    Names = List.Select(List.Skip(Table.ColumnNames(Headers), 3), each not Text.StartsWith(_, "Column")),
    AsTable = Table.FromColumns({{Names}}, type table [Role = text]),
    Indexed = Table.AddIndexColumn(AsTable, "Role Order", 1, 1, Int64.Type)
in
    Indexed""")
    r = Table(name, "RACI", "One row per activity, role and code: the RACI sheet unpivoted.")
    r.column("Activity ID", "string", source="Activity ID")
    r.column("Role", "string", source="Role")
    r.column("Code", "string", source="Code")
    r.column("Code Order", "int64", hidden=True, expr='SWITCH ( RACI[Code], "A", 1, "R", 2, "C", 3, "I", 4, 9 )')
    r.measure("Assignments", "COUNTROWS ( RACI )", fmt="0", folder="KPI", desc="Number of R, A, C and I entries.")
    r.measure("RACI Code", 'CONCATENATEX ( SUMMARIZE ( RACI, RACI[Code], RACI[Code Order] ), RACI[Code], "/", RACI[Code Order], ASC )',
              folder="Matrix", desc="Codes of the role on the activity, for example A/R.")
    for code, word in (("R", "Responsible"), ("A", "Accountable"), ("C", "Consulted"), ("I", "Informed")):
        r.measure(f"{code} ({word})", f'CALCULATE ( COUNTROWS ( RACI ), RACI[Code] = "{code}" )', fmt="0", folder="Workload",
                  desc=f"Number of activities where the role is {word}.")
    r.measure("RACI Cell Color", """VAR c = [RACI Code]
RETURN
    SWITCH ( TRUE (),
        ISBLANK ( c ), "#FFFFFF",
        CONTAINSSTRING ( c, "A" ), "#005F5F",
        CONTAINSSTRING ( c, "R" ), "#008080",
        CONTAINSSTRING ( c, "C" ), "#B7DCDC",
        "#E9ECEC" )""", folder="Format", desc="Cell fill: A dark teal, R teal, C light teal, I grey.")
    r.measure("RACI Font Color", """VAR c = [RACI Code]
RETURN IF ( CONTAINSSTRING ( c, "A" ) || CONTAINSSTRING ( c, "R" ), "#FFFFFF", "#222222" )""", folder="Format",
              desc="White text on A and R cells, dark text elsewhere.")
    r.m_partition(f"""let
    {excel_source("RACI")}
    Clean = Table.SelectRows(Headers, each [Activity ID] <> null and [Activity ID] <> ""),
    Roles = List.Select(List.Skip(Table.ColumnNames(Clean), 3), each not Text.StartsWith(_, "Column")),
    Slim = Table.SelectColumns(Clean, {{"Activity ID"}} & Roles),
    AsText = Table.TransformColumns(Slim, List.Transform({{"Activity ID"}} & Roles, (c) => {{c, each if _ = null then null else Text.From(_), type nullable text}})),
    Long = Table.UnpivotOtherColumns(AsText, {{"Activity ID"}}, "Role", "Codes"),
    Lists = Table.TransformColumns(Long, {{{{"Codes", each List.Select(List.Transform(Text.SplitAny(_, "/,;+ "), (x) => Text.Upper(Text.Trim(x))), (x) => x <> ""), type list}}}}),
    Expanded = Table.ExpandListColumn(Lists, "Codes"),
    Renamed = Table.RenameColumns(Expanded, {{{{"Codes", "Code"}}}}),
    Valid = Table.SelectRows(Renamed, each List.Contains({{"R", "A", "C", "I"}}, [Code])),
    Typed = Table.TransformColumnTypes(Valid, {{{{"Activity ID", type text}}, {{"Role", type text}}, {{"Code", type text}}}})
in
    Typed""")
    tables = [st, act, ro, r]
    rels = [("RACI.'Activity ID'", "Activities.'Activity ID'"), ("RACI.Role", "Roles.Role")]

    mx = {
        "visualType": "pivotTable",
        "query": {"queryState": {
            "Rows": {"projections": [proj(C("Activities", "Activity Label"), "Activity")]},
            "Columns": {"projections": [proj(C("Roles", "Role"))]},
            "Values": {"projections": [proj(M("RACI", "RACI Code"), " ")]}},
            "sortDefinition": {"sort": [{"field": C("Activities", "Activity Label"), "direction": "Ascending"}], "isDefaultSort": True}},
        "objects": {
            "columnHeaders": [{"properties": {"fontColor": color("#FFFFFF"), "backColor": color(TEAL_DARK), "fontSize": n(8),
                                              "fontFamily": s(FONT_BOLD), "wordWrap": b(True), "alignment": s("Center"),
                                              "columnAdjustment": s("fixedWidth"), "defaultColumnWidth": n(58)}}],
            "values": [{"properties": {"fontSize": n(10), "fontFamily": s(FONT_BOLD), "backColorPrimary": color("#FFFFFF"),
                                       "backColorSecondary": color("#FFFFFF")}},
                       cell_format("backColor", "RACI", "RACI Cell Color", "RACI.RACI Code"),
                       cell_format("fontColor", "RACI", "RACI Font Color", "RACI.RACI Code")],
            "columnWidth": [{"properties": {"value": n(210)}, "selector": {"metadata": "Activities.Activity Label"}}],
            "rowHeaders": [{"properties": {"fontSize": n(9), "fontFamily": s(FONT), "fontColor": color(INK), "wordWrap": b(False)}}],
            "subTotals": [{"properties": {"rowSubtotals": b(False), "columnSubtotals": b(False)}}],
            "grid": [{"properties": {"gridVertical": b(True), "gridVerticalColor": color("#FFFFFF"), "gridVerticalWeight": n(2),
                                     "gridHorizontal": b(True), "gridHorizontalColor": color("#FFFFFF"), "gridHorizontalWeight": n(2),
                                     "rowPadding": n(3)}}]},
        "visualContainerObjects": {**vco("Responsibility matrix: activities by role", subtitle_measure=("Settings", "Report Subtitle")), "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    load = {
        "visualType": "barChart",
        "query": {"queryState": {
            "Category": {"projections": [proj(C("Roles", "Role"), active=True)]},
            "Y": {"projections": [proj(M("RACI", "A (Accountable)"), "A"), proj(M("RACI", "R (Responsible)"), "R"),
                                  proj(M("RACI", "C (Consulted)"), "C"), proj(M("RACI", "I (Informed)"), "I")]}},
            "sortDefinition": {"sort": [{"field": C("Roles", "Role"), "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {
            **chart_axes(),
            "valueAxis": [{"properties": {"show": b(False), "gridlineShow": b(False), "showAxisTitle": b(False)}}],
            "labels": [{"properties": {"show": b(True), "fontSize": n(8), "labelPosition": s("InsideCenter"), "color": color(INK)}},
                       {"properties": {"color": color("#FFFFFF")}, "selector": {"metadata": "RACI.A (Accountable)"}},
                       {"properties": {"color": color("#FFFFFF")}, "selector": {"metadata": "RACI.R (Responsible)"}}],
            "dataPoint": [
                {"properties": {"fill": color(TEAL_DARK)}, "selector": {"metadata": "RACI.A (Accountable)"}},
                {"properties": {"fill": color(TEAL)}, "selector": {"metadata": "RACI.R (Responsible)"}},
                {"properties": {"fill": color(TEAL_LIGHT)}, "selector": {"metadata": "RACI.C (Consulted)"}},
                {"properties": {"fill": color("#D5DADA")}, "selector": {"metadata": "RACI.I (Informed)"}}]},
        "visualContainerObjects": vco("Workload per role"),
    }
    chk = {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(C("Activities", "Activity ID"), "ID"),
                                                            proj(C("Activities", "Activity")),
                                                            proj(C("Activities", "Check"))]}}},
        "objects": {**table_common(),
                    "values": table_common()["values"] + [cell_format("fontColor", "Activities", "Check Color", "Activities.Check")]},
        "visualContainerObjects": {**vco("Rule check: one A and at least one R per activity"),
                                   "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    visuals = header("r1", "RACI and responsibility matrix", "Who does, owns, advises and hears about each activity")
    visuals += [slicer("r1_sl_phase", 860, 12, 196, 78, 1200, C("Activities", "Phase"), "Phase"),
                slicer("r1_sl_role", 1064, 12, 196, 78, 1300, C("Roles", "Role"), "Role")]
    cards = [("Activities", "Activities Count", "Activities", GREY, None), ("Roles", "Roles Count", "Roles", GREY, None),
             ("RACI", "Assignments", "Assignments", TEAL, None),
             ("Activities", "Issues", "Activities failing the check", None, ("Activities", "Issues Color"))]
    x = 20
    for k, (ent, meas, label, acc, accm) in enumerate(cards):
        visuals.append(kpi_card(f"r1_k{k}", x, 96, 304, 78, 2000 + k * 100, ent, meas, label, accent=acc or TEAL, accent_measure=accm))
        x += 312
    visuals += [container("r1_mx", 20, 180, 812, 528, mx, 3000),
                container("r1_load", 840, 180, 420, 300, load, 3100),
                container("r1_chk", 840, 486, 420, 222, chk, 3200,
                          filters=[categorical_filter("Filter" + tag(slug, "chk")[:24].replace("-", "0"), "Activities", "Has Issue", ["Yes"])])]
    pages = [{"name": "raci", "displayName": "RACI", "visuals": visuals}]
    return slug, name, tables, rels, pages, ["DataFile", "Settings", "Activities", "Roles", "RACI"]


# =================================================================== 4
def build_lookahead():
    slug, name = "three-week-lookahead", "3-Week Look-Ahead"
    st = settings_table(name, [("Look-ahead Weeks", "int64")], ', {"Look-ahead Weeks", Int64.Type}')
    st.measure("Window Start", """VAR dd = [Data Date Used]
RETURN dd + MOD ( 8 - WEEKDAY ( dd, 2 ), 7 )""", fmt="ddd dd mmm yyyy",
               desc="First day of the look-ahead: the Monday after the data date (the data date itself when it is a Monday).")
    st.measure("Window End", """VAR w = MAX ( Settings[Look-ahead Weeks] )
RETURN [Window Start] + 7 * IF ( ISBLANK ( w ), 3, w ) - 1""", fmt="ddd dd mmm yyyy",
               desc="Last day of the look-ahead window.")
    st.measure("Window Subtitle", 'MAX ( Settings[Project Name] ) & "   |   Data date " & FORMAT ( [Data Date Used], "dd mmm yyyy" ) & "   |   Window " & FORMAT ( [Window Start], "dd mmm" ) & " to " & FORMAT ( [Window End], "dd mmm yyyy" )',
               desc="Header line: project, data date and window.")
    a = Table(name, "Activities", "One row per schedule activity.")
    a.column("Activity ID", "string", source="Activity ID")
    a.column("Activity", "string", source="Activity")
    a.column("Area", "string", source="Area")
    a.column("Owner", "string", source="Owner")
    a.column("Start Date", "dateTime", source="Start Date", fmt="ddd dd mmm")
    a.column("Finish Date", "dateTime", source="Finish Date", fmt="ddd dd mmm")
    a.column("% Complete", "double", source="% Complete", fmt="0%")
    a.column("Constraint", "string", source="Constraint")
    win = """VAR dd0 = MAX ( Settings[Data Date] )
VAR dd = IF ( ISBLANK ( dd0 ), TODAY (), dd0 )
VAR wk0 = MAX ( Settings[Look-ahead Weeks] )
VAR ws = dd + MOD ( 8 - WEEKDAY ( dd, 2 ), 7 )
VAR we = ws + 7 * IF ( ISBLANK ( wk0 ), 3, wk0 ) - 1"""
    a.column("In Window", "string", expr=win + """
RETURN IF ( Activities[% Complete] < 1 && Activities[Start Date] <= we, "Yes", "No" )""",
             desc="Yes when the activity is not complete and starts before the window ends (includes overdue work).")
    a.column("Event", "string", expr=win + """
RETURN
    SWITCH ( TRUE (),
        Activities[Finish Date] < ws, "Overdue",
        Activities[Start Date] >= ws && Activities[Finish Date] <= we, "Starts and finishes",
        Activities[Start Date] >= ws, "Starts",
        Activities[Finish Date] <= we, "Finishes",
        "Runs through" )""", desc="What the activity does in the window.")
    a.column("Has Constraint", "string", expr='IF ( LEN ( TRIM ( Activities[Constraint] ) ) > 0, "Yes", "No" )')
    a.column("Start Order", "int64", hidden=True, expr="RANKX ( ALL ( Activities ), Activities[Start Date], , ASC, DENSE )")
    a.column("Activity Label", "string", expr='Activities[Activity ID] & "  " & Activities[Activity]', sort_by="Start Order",
             desc="Activity ID and name, sorted by start date.")
    a.measure("In Window Count", 'CALCULATE ( COUNTROWS ( Activities ), Activities[In Window] = "Yes" ) + 0', fmt="0", folder="KPI",
              desc="Open activities that touch the window, overdue ones included.")
    a.measure("Starting", 'CALCULATE ( COUNTROWS ( Activities ), Activities[In Window] = "Yes", Activities[Event] IN { "Starts", "Starts and finishes" } ) + 0',
              fmt="0", folder="KPI", desc="Activities that start inside the window.")
    a.measure("Finishing", 'CALCULATE ( COUNTROWS ( Activities ), Activities[In Window] = "Yes", Activities[Event] IN { "Finishes", "Starts and finishes" } ) + 0',
              fmt="0", folder="KPI", desc="Activities that finish inside the window.")
    a.measure("Open Constraints", 'CALCULATE ( COUNTROWS ( Activities ), Activities[In Window] = "Yes", Activities[Has Constraint] = "Yes" ) + 0',
              fmt="0", folder="KPI", desc="Activities in the window that still carry a constraint.")
    a.measure("Overdue", 'CALCULATE ( COUNTROWS ( Activities ), Activities[In Window] = "Yes", Activities[Event] = "Overdue" ) + 0',
              fmt="0", folder="KPI", desc="Open activities whose finish date is before the window.")
    a.measure("Constraint Color", 'IF ( [Open Constraints] > 0, "#E08A00", "#339966" )', folder="Format", desc="Amber when constraints are open.")
    a.measure("Overdue Color", 'IF ( [Overdue] > 0, "#D63A2F", "#339966" )', folder="Format", desc="Red when work is overdue.")
    gvars = """VAR d = MAX ( Days[Date] )
VAR s = MAX ( Activities[Start Date] )
VAR f = MAX ( Activities[Finish Date] )
VAR ws = MIN ( Days[Date] )
VAR onBar = d >= s && d <= f
VAR ws0 = CALCULATE ( MIN ( Days[Date] ), ALL ( Days ) )
VAR overdueCell = f < ws0 && d = ws0"""
    a.measure("Gantt Cell", gvars + """
RETURN
    IF ( HASONEVALUE ( Activities[Activity ID] ),
        SWITCH ( TRUE (), overdueCell, "!", onBar && s = f, UNICHAR ( 9670 ), "" ) )""", folder="Gantt",
              desc="Gantt cell text: diamond for a one-day activity, ! on the first day for overdue work, empty otherwise.")
    a.measure("Gantt Color", gvars + """
VAR weekend = WEEKDAY ( d, 2 ) > 5
VAR hasC = MAX ( Activities[Has Constraint] ) = "Yes"
VAR pct = MAX ( Activities[% Complete] )
VAR doneTo = s + pct * ( f - s + 1 )
RETURN
    SWITCH ( TRUE (),
        overdueCell, "#D63A2F",
        onBar && s = f, IF ( weekend, "#F1F3F3", "#FFFFFF" ),
        onBar && hasC, "#E08A00",
        onBar && d < doneTo, "#005F5F",
        onBar, "#46A3A3",
        weekend, "#F1F3F3",
        "#FFFFFF" )""", folder="Format", desc="Gantt cell fill: teal bar, dark teal for the part done, amber when a constraint is open, red flag for overdue.")
    a.measure("Gantt Font", """VAR hasC = MAX ( Activities[Has Constraint] ) = "Yes"
RETURN IF ( [Gantt Cell] = "!", "#FFFFFF", IF ( hasC, "#E08A00", "#005F5F" ) )""", folder="Format",
              desc="Gantt text color: white on the overdue flag, amber or teal diamonds.")
    a.m_partition(f"""let
    {excel_source("Activities")}
    Typed = Table.TransformColumnTypes(Headers, {{{{"Activity ID", type text}}, {{"Activity", type text}}, {{"Area", type text}}, {{"Owner", type text}}, {{"Start Date", type date}}, {{"Finish Date", type date}}, {{"% Complete", type number}}, {{"Constraint", type text}}}}),
    Clean = Table.SelectRows(Typed, each [Activity ID] <> null and [Activity ID] <> ""),
    NoNullPct = Table.ReplaceValue(Clean, null, 0, Replacer.ReplaceValue, {{"% Complete"}})
in
    NoNullPct""")
    dy = Table(name, "Days", "Calculated calendar of the look-ahead window, one row per day.")
    dy.column("Date", "dateTime", source="[Date]", fmt="ddd dd mmm", calc_table_col=True)
    dy.column("Day", "string", source="[Day]", sort_by="Date", calc_table_col=True)
    dy.column("Week", "string", source="[Week]", calc_table_col=True)
    dy.calc_partition("""VAR dd0 = MAX ( Settings[Data Date] )
VAR dd = IF ( ISBLANK ( dd0 ), TODAY (), dd0 )
VAR wk0 = MAX ( Settings[Look-ahead Weeks] )
VAR ws = dd + MOD ( 8 - WEEKDAY ( dd, 2 ), 7 )
VAR nDays = 7 * IF ( ISBLANK ( wk0 ), 3, wk0 )
RETURN
    ADDCOLUMNS (
        CALENDAR ( ws, ws + nDays - 1 ),
        "Day", LEFT ( FORMAT ( [Date], "ddd" ), 1 ) & UNICHAR ( 10 ) & FORMAT ( [Date], "d" ),
        "Week", "Week " & ( INT ( ( [Date] - ws ) / 7 ) + 1 )
    )""")
    tables = [st, a, dy]
    E = "Activities"
    in_window = lambda k: [categorical_filter("Filter" + tag(slug, k)[:24].replace("-", "0"), E, "In Window", ["Yes"])]
    gantt = {
        "visualType": "pivotTable",
        "query": {"queryState": {
            "Rows": {"projections": [proj(C(E, "Activity Label"), "Activity")]},
            "Columns": {"projections": [proj(C("Days", "Day"))]},
            "Values": {"projections": [proj(M(E, "Gantt Cell"), " ")]}},
            "sortDefinition": {"sort": [{"field": C(E, "Activity Label"), "direction": "Ascending"}], "isDefaultSort": True}},
        "objects": {
            "columnHeaders": [{"properties": {"fontColor": color("#FFFFFF"), "backColor": color(TEAL_DARK), "fontSize": n(8),
                                              "fontFamily": s(FONT_BOLD), "wordWrap": b(True), "alignment": s("Center"),
                                              "columnAdjustment": s("fixedWidth"), "defaultColumnWidth": n(28)}}],
            "values": [{"properties": {"fontSize": n(9), "fontFamily": s(FONT_BOLD), "backColorPrimary": color("#FFFFFF"),
                                       "backColorSecondary": color("#FFFFFF")}},
                       cell_format("backColor", E, "Gantt Color", "Activities.Gantt Cell"),
                       cell_format("fontColor", E, "Gantt Font", "Activities.Gantt Cell")],
            "columnWidth": [{"properties": {"value": n(250)}, "selector": {"metadata": "Activities.Activity Label"}}],
            "rowHeaders": [{"properties": {"fontSize": n(9), "fontFamily": s(FONT), "fontColor": color(INK), "wordWrap": b(False)}}],
            "subTotals": [{"properties": {"rowSubtotals": b(False), "columnSubtotals": b(False)}}],
            "grid": [{"properties": {"gridVertical": b(True), "gridVerticalColor": color("#FFFFFF"), "gridVerticalWeight": n(1),
                                     "gridHorizontal": b(True), "gridHorizontalColor": color("#FFFFFF"), "gridHorizontalWeight": n(3),
                                     "rowPadding": n(2)}}]},
        "visualContainerObjects": {**vco("Look-ahead bars: teal = planned, dark = done, amber = constraint open, red ! = overdue",
                                         subtitle_measure=("Settings", "Window Subtitle")),
                                   "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    cons = {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(C(E, "Activity ID"), "ID"), proj(C(E, "Constraint")),
                                                            proj(C(E, "Start Date"), "Need by")]}},
                  "sortDefinition": {"sort": [{"field": C(E, "Start Date"), "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {**table_common()},
        "visualContainerObjects": {**vco("Constraints to clear, by need-by date"), "stylePreset": [{"properties": {"name": s("None")}}]},
    }
    cons_filters = in_window("cons") + [categorical_filter("Filter" + tag(slug, "cons2")[:24].replace("-", "0"), E, "Has Constraint", ["Yes"])]
    visuals = header("l1", "3-week look-ahead", "Work that starts, finishes or runs in the next three weeks, with its constraints")
    visuals += [slicer("l1_sl_area", 860, 12, 196, 78, 1200, C(E, "Area"), "Area"),
                slicer("l1_sl_owner", 1064, 12, 196, 78, 1300, C(E, "Owner"), "Owner")]
    cards = [("In Window Count", "Activities in window", TEAL, None), ("Starting", "Starting", TEAL_LIGHT, None),
             ("Finishing", "Finishing", TEAL_DARK, None), ("Open Constraints", "Open constraints", None, (E, "Constraint Color")),
             ("Overdue", "Overdue", None, (E, "Overdue Color"))]
    x = 20
    for k, (meas, label, acc, accm) in enumerate(cards):
        visuals.append(kpi_card(f"l1_k{k}", x, 96, 241, 78, 2000 + k * 100, E, meas, label, accent=acc or TEAL, accent_measure=accm))
        x += 249
    visuals += [container("l1_gantt", 20, 180, 876, 528, gantt, 3000, filters=in_window("gantt")),
                container("l1_cons", 904, 180, 356, 528, cons, 3100, filters=cons_filters)]
    pages = [{"name": "lookahead", "displayName": "Look-ahead", "visuals": visuals}]
    return slug, name, tables, [], pages, ["DataFile", "Settings", "Activities"]


def main():
    theme = kata_theme(KATA)
    only = [a for a in sys.argv[1:] if not a.startswith("--")]
    for fn in (build_progress, build_milestones, build_raci, build_lookahead):
        slug, name, tables, rels, pages, qorder = fn()
        if only and slug not in only:
            continue
        root = ROOT / slug
        rid, mid = ids(slug)
        write_model(root, name, tables, rels, data_file(slug), mid, qorder)
        write_report(root, name, pages, theme, BASE_THEME, [rid])
        write_pbip(root, name)
        print("built", root / f"{name}.pbip")


if __name__ == "__main__":
    main()
