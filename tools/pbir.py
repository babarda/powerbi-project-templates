"""Small helpers that write PBIP projects (PBIR report + TMDL semantic model).

Author: Badr Abardazzou. License: MIT.
"""
import json
import shutil
import uuid
from pathlib import Path

VC_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.11.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
NS = uuid.UUID("6b1f9e52-3c1d-4a8e-9f11-2f0b7c6d9a01")

INK = "#222222"
MUTED = "#6B7777"
LINE = "#E2E8E8"
TEAL = "#008080"
TEAL_DARK = "#005F5F"
TEAL_LIGHT = "#B7DCDC"
GREY = "#AAAAAA"
GOOD = "#339966"
BAD = "#D63A2F"
WARN = "#E08A00"
PAGE_BG = "#F4F6F6"
FONT = "Segoe UI"
THEME_FILE = "KATA-5a1c0e2b.json"
FONT_BOLD = "Segoe UI Semibold"


def tag(*parts):
    return str(uuid.uuid5(NS, "/".join(parts)))


# ------------------------------------------------------------ literals
def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def s(v):
    return lit("'" + str(v).replace("'", "''") + "'")


def b(v):
    return lit("true" if v else "false")


def n(v):
    return lit(f"{v}D")


def i(v):
    return lit(f"{int(v)}L")


def color(hexv):
    return {"solid": {"color": lit(f"'{hexv}'")}}


def measure_expr(entity, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def mcolor(entity, prop):
    return {"solid": {"color": {"expr": measure_expr(entity, prop)}}}


# ------------------------------------------------------------ fields
def C(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def M(entity, prop):
    return measure_expr(entity, prop)


def proj(field, display=None, active=None):
    kind = "Column" if "Column" in field else "Measure"
    ent = field[kind]["Expression"]["SourceRef"]["Entity"]
    prop = field[kind]["Property"]
    p = {"field": field, "queryRef": f"{ent}.{prop}", "nativeQueryRef": display or prop}
    if display:
        p["displayName"] = display
    if active is not None:
        p["active"] = active
    return p


def qref(field):
    kind = "Column" if "Column" in field else "Measure"
    return f"{field[kind]['Expression']['SourceRef']['Entity']}.{field[kind]['Property']}"


# ------------------------------------------------------------ chrome
def vco(title=None, subtitle_measure=None, background=True, border=True, padding=8):
    o = {}
    if title:
        o["title"] = [{"properties": {
            "show": b(True), "text": s(title), "fontSize": n(11), "bold": b(True),
            "fontColor": color(INK), "fontFamily": s(FONT_BOLD), "alignment": s("left"), "titleWrap": b(False)}}]
    else:
        o["title"] = [{"properties": {"show": b(False)}}]
    if subtitle_measure:
        o["subTitle"] = [{"properties": {
            "show": b(True), "fontSize": n(9), "fontColor": color(MUTED), "alignment": s("left"),
            "text": {"expr": measure_expr(*subtitle_measure)}}}]
    if background:
        o["background"] = [{"properties": {"show": b(True), "color": color("#FFFFFF"), "transparency": n(0)}}]
    if border:
        o["border"] = [{"properties": {"show": b(True), "color": color(LINE), "radius": n(6)}}]
    o["padding"] = [{"properties": {"top": n(padding), "bottom": n(padding), "left": n(padding), "right": n(padding)}}]
    o["visualHeader"] = [{"properties": {"show": b(True), "showPinButton": b(False), "showVisualInformationButton": b(False)}}]
    return o


def container(name, x, y, w, h, visual, z, filters=None):
    d = {"$schema": VC_SCHEMA, "name": name,
         "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
         "visual": visual}
    if filters:
        d["filterConfig"] = {"filters": filters}
    return d


def textbox(name, x, y, w, h, z, title, subtitle=None):
    paras = [{"textRuns": [{"value": title, "textStyle": {
        "fontWeight": "bold", "fontFamily": FONT_BOLD, "fontSize": "18pt", "color": TEAL_DARK}}]}]
    if subtitle:
        paras.append({"textRuns": [{"value": subtitle, "textStyle": {
            "fontFamily": FONT, "fontSize": "10pt", "color": MUTED}}]})
    vis = {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paras}}]},
           "visualContainerObjects": {"background": [{"properties": {"show": b(False)}}],
                                      "padding": [{"properties": {"top": n(0), "bottom": n(0), "left": n(0), "right": n(0)}}]},
           "drillFilterOtherVisuals": True}
    return container(name, x, y, w, h, vis, z)


def subtitle_card(name, x, y, w, h, z, entity, meas):
    vis = {"visualType": "cardVisual",
           "query": {"queryState": {"Data": {"projections": [proj(M(entity, meas))]}}},
           "objects": {
               "value": [{"properties": {"fontSize": n(10), "bold": b(False), "fontFamily": s(FONT),
                                         "fontColor": color(INK), "horizontalAlignment": s("left")},
                          "selector": {"id": "default"}}],
               "label": [{"properties": {"show": b(False)}, "selector": {"id": "default"}}],
               "outline": [{"properties": {"show": b(False)}, "selector": {"id": "default"}}],
               "layout": [{"properties": {"alignment": s("left")}, "selector": {"id": "default"}}]},
           "visualContainerObjects": {
               "title": [{"properties": {"show": b(False)}}],
               "background": [{"properties": {"show": b(False)}}],
               "border": [{"properties": {"show": b(False)}}],
               "padding": [{"properties": {"top": n(0), "bottom": n(0), "left": n(0), "right": n(0)}}],
               "visualHeader": [{"properties": {"show": b(False)}}]}}
    return container(name, x, y, w, h, vis, z)


def kpi_card(name, x, y, w, h, z, entity, meas, label, accent=TEAL, accent_measure=None, value_size=22):
    acc_color = mcolor(*accent_measure) if accent_measure else color(accent)
    vis = {"visualType": "cardVisual",
           "query": {"queryState": {"Data": {"projections": [proj(M(entity, meas))]}}},
           "objects": {
               "value": [{"properties": {"fontSize": n(value_size), "bold": b(True), "fontFamily": s(FONT_BOLD),
                                         "fontColor": color(INK), "horizontalAlignment": s("left")},
                          "selector": {"id": "default"}}],
               "label": [{"properties": {"show": b(True), "text": s(label), "fontSize": n(9), "bold": b(False),
                                         "fontColor": color(MUTED), "position": s("aboveValue"),
                                         "horizontalAlignment": s("left")},
                          "selector": {"id": "default"}}],
               "outline": [{"properties": {"show": b(False)}, "selector": {"id": "default"}}],
               "accentBar": [{"properties": {"show": b(True), "position": s("Left"), "width": n(4), "color": acc_color},
                              "selector": {"id": "default"}}],
               "padding": [{"properties": {"leftMargin": i(12), "rightMargin": i(8), "topMargin": i(0), "bottomMargin": i(0)},
                            "selector": {"id": "default"}}],
               "layout": [{"properties": {"topOuterMargin": i(0), "bottomOuterMargin": i(0), "leftOuterMargin": i(0),
                                          "rightOuterMargin": i(0), "alignment": s("left")},
                           "selector": {"id": "default"}}]},
           "visualContainerObjects": {
               "title": [{"properties": {"show": b(False)}}],
               "background": [{"properties": {"show": b(True), "color": color("#FFFFFF"), "transparency": n(0)}}],
               "border": [{"properties": {"show": b(True), "color": color(LINE), "radius": n(6)}}],
               "padding": [{"properties": {"top": n(0), "bottom": n(0), "left": n(0), "right": n(0)}}],
               "visualHeader": [{"properties": {"show": b(False)}}]}}
    return container(name, x, y, w, h, vis, z)


def slicer(name, x, y, w, h, z, field, header):
    vis = {"visualType": "slicer",
           "query": {"queryState": {"Values": {"projections": [proj(field)]}}},
           "objects": {"data": [{"properties": {"mode": s("Dropdown")}}],
                       "header": [{"properties": {"show": b(True), "text": s(header), "textSize": n(9),
                                                  "fontColor": color(MUTED)}}]},
           "visualContainerObjects": {
               "background": [{"properties": {"show": b(True), "color": color("#FFFFFF"), "transparency": n(0)}}],
               "border": [{"properties": {"show": b(True), "color": color(LINE), "radius": n(6)}}],
               "padding": [{"properties": {"top": n(4), "bottom": n(4), "left": n(8), "right": n(8)}}]}}
    return container(name, x, y, w, h, vis, z)


def categorical_filter(fid, entity, prop, values):
    return {"name": fid, "field": C(entity, prop), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "f", "Entity": entity, "Type": 0}],
                       "Where": [{"Condition": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "f"}}, "Property": prop}}],
                           "Values": [[{"Literal": {"Value": f"'{v}'"}}] for v in values]}}}]}}


def table_common(header_bg=TEAL_DARK, font_size=9, grow=True, wrap_headers=True):
    ch = {"fontColor": color("#FFFFFF"), "backColor": color(header_bg), "fontSize": n(font_size),
          "fontFamily": s(FONT_BOLD), "wordWrap": b(wrap_headers)}
    if grow:
        ch["columnAdjustment"] = s("growToFit")
        ch["autoSizeColumnWidth"] = b(True)
    return {
        "columnHeaders": [{"properties": ch}],
        "values": [{"properties": {"fontSize": n(font_size), "fontFamily": s(FONT),
                                   "backColorPrimary": color("#FFFFFF"), "backColorSecondary": color("#F7F9F9"),
                                   "fontColorPrimary": color(INK), "fontColorSecondary": color(INK)}}],
        "total": [{"properties": {"totals": b(False)}}],
        "grid": [{"properties": {"gridVertical": b(False), "gridHorizontal": b(True),
                                 "gridHorizontalColor": color("#EEF1F1"), "rowPadding": n(3)}}],
    }


def cell_format(prop, entity, meas, target_queryref):
    """Field-value conditional formatting of a table/matrix cell (backColor or fontColor)."""
    return {"properties": {prop: mcolor(entity, meas)},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": target_queryref}}


def chart_axes(value_format_show=True):
    return {
        "categoryAxis": [{"properties": {"fontSize": n(8), "labelColor": color(MUTED), "showAxisTitle": b(False)}}],
        "valueAxis": [{"properties": {"fontSize": n(8), "labelColor": color(MUTED), "showAxisTitle": b(False),
                                      "gridlineShow": b(True), "gridlineColor": color("#EEF1F1")}}],
        "legend": [{"properties": {"show": b(True), "position": s("Top"), "fontSize": n(8), "labelColor": color(MUTED)}}],
    }


# ------------------------------------------------------------ files
def write_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def write_text(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def kata_theme(src):
    t = json.loads(Path(src).read_text(encoding="utf-8"))
    t["name"] = THEME_FILE
    for cls in t.get("textClasses", {}).values():
        cls["fontFace"] = FONT
    t["textClasses"]["title"]["fontFace"] = FONT_BOLD
    t["textClasses"]["callout"]["fontFace"] = FONT_BOLD
    star = t["visualStyles"]["*"]["*"]
    star["title"][0]["fontFamily"] = FONT_BOLD
    t["visualStyles"]["page"]["*"]["background"][0]["color"]["solid"]["color"] = PAGE_BG
    return t


def write_report(root, name, pages, theme, base_theme_file, logical_ids):
    rep = root / f"{name}.Report"
    if rep.exists():
        shutil.rmtree(rep)
    write_json(rep / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": name},
        "config": {"version": "2.0", "logicalId": logical_ids[0]}})
    write_json(rep / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{name}.SemanticModel"}}})
    d = rep / "definition"
    write_json(d / "version.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0"})
    rv = {"visual": "2.6.0", "report": "3.1.0", "page": "2.3.0"}
    write_json(d / "report.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": "CY26SU02", "reportVersionAtImport": rv, "type": "SharedResources"},
            "customTheme": {"name": THEME_FILE, "reportVersionAtImport": rv, "type": "RegisteredResources"}},
        "objects": {"outspacePane": [{"properties": {"expanded": b(False)}}]},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY26SU02", "path": "BaseThemes/CY26SU02.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": THEME_FILE, "path": THEME_FILE, "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                     "useDefaultAggregateDisplayName": True}})
    shutil.copy(base_theme_file, _mk(rep / "StaticResources" / "SharedResources" / "BaseThemes") / "CY26SU02.json")
    write_json(rep / "StaticResources" / "RegisteredResources" / THEME_FILE, theme)
    write_json(d / "pages" / "pages.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": [p["name"] for p in pages], "activePageName": pages[0]["name"]})
    for p in pages:
        pd = d / "pages" / p["name"]
        write_json(pd / "page.json", {
            "$schema": PAGE_SCHEMA, "name": p["name"], "displayName": p["displayName"],
            "displayOption": "FitToPage", "height": 720, "width": 1280,
            "objects": {"background": [{"properties": {"color": color(PAGE_BG), "transparency": n(0)}}]}})
        for v in p["visuals"]:
            write_json(pd / "visuals" / v["name"] / "visual.json", v)


def _mk(p):
    p.mkdir(parents=True, exist_ok=True)
    return p


def write_pbip(root, name):
    write_json(root / f"{name}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{name}.Report"}}],
        "settings": {"enableAutoRecovery": True}})


# ------------------------------------------------------------ TMDL
def q(name):
    """Quote a TMDL object name when needed."""
    if all(ch.isalnum() or ch == "_" for ch in name) and not name[0].isdigit():
        return name
    return "'" + name.replace("'", "''") + "'"


def _desc(desc, indent):
    if not desc:
        return ""
    return "".join(f"{indent}/// {line}\n" for line in desc.split("\n"))


def _expr_block(expr, indent_props):
    expr = expr.strip("\n")
    if "\n" not in expr:
        return f" = {expr}\n"
    body = "".join(f"{indent_props}\t{line}\n" if line.strip() else "\n" for line in expr.split("\n"))
    return " =\n" + body


class Table:
    def __init__(self, model, name, desc=None):
        self.model, self.name, self.desc = model, name, desc
        self.items = []
        self.partition = None

    def column(self, name, dtype, source=None, expr=None, fmt=None, summarize="none", sort_by=None,
               hidden=False, desc=None, calc_table_col=False):
        t = "\t"
        out = _desc(desc, t)
        out += f"{t}column {q(name)}" + (_expr_block(expr, t * 2) if expr else "\n")
        out += f"{t*2}dataType: {dtype}\n"
        if fmt:
            out += f"{t*2}formatString: {fmt}\n"
        if hidden:
            out += f"{t*2}isHidden\n"
        if calc_table_col:
            out += f"{t*2}isNameInferred\n"
        out += f"{t*2}lineageTag: {tag(self.model, self.name, 'col', name)}\n"
        out += f"{t*2}summarizeBy: {summarize}\n"
        if source:
            out += f"{t*2}sourceColumn: {source}\n"
        if sort_by:
            out += f"{t*2}sortByColumn: {q(sort_by)}\n"
        if dtype == "dateTime" or fmt:
            pass
        self.items.append(out)
        return self

    def measure(self, name, expr, fmt=None, folder=None, desc=None):
        t = "\t"
        out = _desc(desc, t)
        out += f"{t}measure {q(name)}" + _expr_block(expr, t * 2)
        if fmt:
            out += f"{t*2}formatString: {fmt}\n"
        if folder:
            out += f"{t*2}displayFolder: {folder}\n"
        out += f"{t*2}lineageTag: {tag(self.model, self.name, 'measure', name)}\n"
        self.items.append(out)
        return self

    def m_partition(self, m):
        body = "".join(f"\t\t\t\t{line}\n" for line in m.strip("\n").split("\n"))
        self.partition = f"\tpartition {q(self.name)} = m\n\t\tmode: import\n\t\tsource =\n{body}"
        return self

    def calc_partition(self, dax):
        body = "".join(f"\t\t\t\t{line}\n" for line in dax.strip("\n").split("\n"))
        self.partition = f"\tpartition {q(self.name)} = calculated\n\t\tmode: import\n\t\tsource =\n{body}"
        return self

    def render(self):
        out = _desc(self.desc, "")
        out += f"table {q(self.name)}\n\tlineageTag: {tag(self.model, self.name)}\n\n"
        out += "\n".join(self.items) + "\n"
        out += self.partition + "\n"
        out += f"\tannotation PBI_Id = {self.name.replace(' ', '')}\n"
        return out


def write_model(root, name, tables, relationships, data_file_default, logical_id, query_order):
    sm = root / f"{name}.SemanticModel"
    if sm.exists():
        shutil.rmtree(sm)
    write_json(sm / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": name},
        "config": {"version": "2.0", "logicalId": logical_id}})
    write_json(sm / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}})
    d = sm / "definition"
    write_text(d / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")
    write_text(d / "expressions.tmdl",
               "/// Full path of the Excel workbook that feeds this report (sample_data.xlsx or your own copy).\n"
               f"expression DataFile = \"{data_file_default}\" meta [IsParameterQuery=true, Type=\"Text\", IsParameterQueryRequired=true]\n"
               f"\tlineageTag: {tag(name, 'DataFile')}\n\n"
               "\tannotation PBI_ResultType = Text\n\n")
    refs = "".join(f"ref table {q(t.name)}\n" for t in tables)
    qo = json.dumps(query_order)
    write_text(d / "model.tmdl",
               "model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
               "\tsourceQueryCulture: en-US\n\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
               "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
               f"annotation PBI_QueryOrder = {qo}\n\n"
               "annotation PBI_ProTooling = [\"DevMode\"]\n\n" + refs + "\n")
    for t in tables:
        write_text(d / "tables" / f"{t.name}.tmdl", t.render())
    if relationships:
        out = ""
        for frm, to in relationships:
            out += f"relationship {tag(name, 'rel', frm, to)}\n\tfromColumn: {frm}\n\ttoColumn: {to}\n\n"
        write_text(d / "relationships.tmdl", out)


def excel_source(sheet):
    return (f'Source = Excel.Workbook(File.Contents(DataFile), null, true),\n'
            f'    Sheet = Source{{[Item="{sheet}",Kind="Sheet"]}}[Data],\n'
            f'    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),')
