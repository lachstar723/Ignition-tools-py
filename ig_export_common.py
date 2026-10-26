#!/usr/bin/env python3
"""
ig_export_common.py
Shared helpers for turning Ignition tag / UDT JSON exports into
stakeholder-friendly CSV files and a multi-sheet .xlsx workbook.

Standard library only (json, csv, zipfile, xml) - no pandas / numpy / openpyxl.
Python 3.6+.  Keep this file in the same folder as:
    udt_export_to_sheets.py
    tag_export_to_sheets.py
"""
import copy
import csv
import datetime
import json
import math
import os
import re
import zipfile
from xml.sax.saxutils import escape

# --------------------------------------------------------------------------
# Column ordering. Columns are generated dynamically from whatever keys exist
# in the export; these lists only control the ORDER of the common ones.
# Anything not listed here still appears, sorted alphabetically after these.
# --------------------------------------------------------------------------
PREFERRED_TAG_COLS = [
    "typeId", "enabled", "valueSource", "dataType", "opcServer", "opcItemPath",
    "expression", "sourceTagPath", "value", "engUnit", "engLow", "engHigh",
    "engLimitMode", "formatString", "documentation", "tooltip", "readOnly",
    "scaleMode", "rawLow", "rawHigh", "scaledLow", "scaledHigh", "clampMode",
    "deadband", "deadbandMode", "tagGroup", "historyEnabled", "historyProvider",
    "historicalDeadband", "historicalDeadbandMode", "historicalDeadbandStyle",
    "sampleMode", "historySampleRate", "historySampleRateUnits",
    "historyMaxAge", "historyMaxAgeUnits", "parameters",
]
PREFERRED_ALARM_COLS = [
    "enabled", "mode", "setpointA", "setpointB", "priority", "label",
    "displayPath", "notes", "ackMode", "ackNotesReqd", "shelvingAllowed",
    "deadband", "deadbandMode", "timeOnDelaySeconds", "timeOffDelaySeconds",
    "activePipeline", "clearPipeline", "ackPipeline",
]
REVIEW_COLS = ["Review Status", "Reviewer", "Comments"]

# Keys handled structurally (not dumped as plain columns)
INTERNAL_KEYS = {"__overrides__", "__type_missing__", "__parent_missing__"}
STRUCTURE_KEYS = {"name", "tagType", "tags", "alarms", "parameters"} | INTERNAL_KEYS


# ==========================================================================
# Loading
# ==========================================================================
def load_json(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def top_nodes(data):
    """Return the list of top-level tag nodes from any Ignition export shape:
    a whole provider, a single folder/UDT, or a bare list."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        if "tags" in data and (data.get("tagType") in (None, "", "Provider")
                               or not data.get("name")):
            return data["tags"] or []
        return [data]
    raise ValueError("Unrecognised export format (expected a JSON object or list)")


# ==========================================================================
# Value formatting
# ==========================================================================
_PARAM_REF = re.compile(r"\{([^{}\[\]]+)\}")


def is_binding(v):
    return isinstance(v, dict) and "bindType" in v


def _binding_text(v):
    for k in ("binding", "value", "expression"):
        if k in v:
            return v[k]
    return ""


def substitute_params(text, params):
    """Replace {ParamName} references with the instance's parameter values."""
    if not params or not isinstance(text, str):
        return text

    def rep(m):
        key = m.group(1).strip()
        return str(params[key]) if key in params else m.group(0)

    return _PARAM_REF.sub(rep, text)


def fmt(v, params=None):
    """Turn an Ignition property value into something readable in a cell.

    - bools -> true/false, numbers stay numbers
    - parameter bindings -> resolved text if params given, else 'parameter: ...'
    - other bindings -> 'Expression: ...', 'Tag: ...' etc.
    - anything else structured -> compact JSON
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        return v
    if is_binding(v):
        bind_type = str(v.get("bindType"))
        text = _binding_text(v)
        if not isinstance(text, str):
            text = json.dumps(text, ensure_ascii=False)
        if bind_type.lower() == "parameter" and params is not None:
            return substitute_params(text, params)
        return "%s: %s" % (bind_type, text)
    return json.dumps(v, ensure_ascii=False, separators=(", ", ": "))


def param_value(p):
    """Parameters export as {"dataType": .., "value": ..} or a bare value."""
    if isinstance(p, dict) and not is_binding(p):
        return p.get("value")
    return p


# ==========================================================================
# Merging (UDT inheritance + instance overrides)
# ==========================================================================
def _name(n):
    return n.get("name") if isinstance(n, dict) else None


def strip_markers(obj):
    if isinstance(obj, dict):
        return {k: strip_markers(v) for k, v in obj.items() if k not in INTERNAL_KEYS}
    if isinstance(obj, list):
        return [strip_markers(x) for x in obj]
    return obj


def _merge_named_lists(base, over, merge_fn, track):
    result = [copy.deepcopy(b) for b in (base or [])]
    index = {_name(b): i for i, b in enumerate(result) if _name(b) is not None}
    for o in over or []:
        n = _name(o)
        if n is not None and n in index:
            result[index[n]] = merge_fn(result[index[n]], o)
        else:
            item = copy.deepcopy(o)
            if track and isinstance(item, dict):
                item["__overrides__"] = ["(added - not in UDT definition)"]
            result.append(item)
    return result


def merge_alarm(base, over, track=False):
    result = copy.deepcopy(base)
    changed = []
    for k, v in over.items():
        result[k] = copy.deepcopy(v)
        if k != "name" and k not in INTERNAL_KEYS:
            changed.append(k)
    if track and changed:
        result["__overrides__"] = changed
    return result


def merge_tag(base, over, track=False):
    """Overlay `over` (instance or child UDT) on top of `base` (definition)."""
    result = copy.deepcopy(base)
    changed = []
    for k, v in over.items():
        if k == "tags":
            result["tags"] = _merge_named_lists(
                base.get("tags"), v, lambda b, o: merge_tag(b, o, track), track)
        elif k == "alarms":
            result["alarms"] = _merge_named_lists(
                base.get("alarms"), v, lambda b, o: merge_alarm(b, o, track), track)
        elif k == "parameters" and isinstance(v, dict) and isinstance(base.get(k), dict):
            params = copy.deepcopy(base[k])
            for pk, pv in v.items():
                if (isinstance(pv, dict) and not is_binding(pv)
                        and isinstance(params.get(pk), dict) and not is_binding(params[pk])):
                    params[pk].update(copy.deepcopy(pv))
                else:
                    params[pk] = copy.deepcopy(pv)
            result[k] = params
        elif k in INTERNAL_KEYS:
            continue
        else:
            result[k] = copy.deepcopy(v)
            if k not in ("name", "tagType", "typeId"):
                changed.append(k)
    if track and changed:
        result["__overrides__"] = changed
    return result


class UdtLibrary(object):
    """Index of UDT definitions, with inheritance + nested-instance resolution."""

    def __init__(self):
        self.types = {}      # path relative to _types_ -> raw UdtType node
        self._cache = {}

    def add_export(self, data):
        for node in top_nodes(data):
            self._collect(node, [])

    def _collect(self, node, parts):
        if not isinstance(node, dict):
            return
        name = str(node.get("name", ""))
        ttype = node.get("tagType")
        if ttype == "UdtType":
            self.types["/".join(parts + [name])] = node
            return
        if ttype in ("Folder", "Provider", None):
            if name == "_types_":
                child_parts = []
            elif name:
                child_parts = parts + [name]
            else:
                child_parts = parts
            for c in node.get("tags") or []:
                self._collect(c, child_parts)

    def lookup(self, type_id):
        """Find a definition path from a typeId, tolerating folder-prefix differences."""
        if not type_id or not isinstance(type_id, str):
            return None
        tid = re.sub(r"^\[[^\]]*\]", "", type_id.strip()).strip("/")
        if tid.startswith("_types_/"):
            tid = tid[len("_types_/"):]
        if tid in self.types:
            return tid
        cands = [p for p in self.types
                 if p.endswith("/" + tid) or tid.endswith("/" + p)]
        return cands[0] if len(cands) == 1 else None

    def resolve_type(self, path, _stack=frozenset()):
        """Definition with parent-UDT inheritance and nested UDT members expanded."""
        if path in self._cache:
            return self._cache[path]
        node = self.types[path]
        stack = _stack | {path}
        parent_id = node.get("typeId")
        merged = None
        if parent_id:
            ppath = self.lookup(parent_id)
            if ppath and ppath not in stack:
                merged = merge_tag(self.resolve_type(ppath, stack), node)
            else:
                merged = copy.deepcopy(node)
                merged["__parent_missing__"] = True
        else:
            merged = copy.deepcopy(node)
        merged["tags"] = self.expand_tags(merged.get("tags"), stack)
        self._cache[path] = merged
        return merged

    def expand_tags(self, tags, _stack=frozenset()):
        out = []
        for t in tags or []:
            if not isinstance(t, dict):
                continue
            if t.get("tagType") == "UdtInstance":
                out.append(self.expand_instance(t, _stack))
            elif t.get("tags"):
                c = dict(t)
                c["tags"] = self.expand_tags(t["tags"], _stack)
                out.append(c)
            else:
                out.append(t)
        return out

    def expand_instance(self, node, _stack=frozenset()):
        path = self.lookup(node.get("typeId"))
        if path is None or path in _stack:
            c = copy.deepcopy(node)
            c["__type_missing__"] = True
            c["tags"] = self.expand_tags(c.get("tags"), _stack)
            return c
        base = strip_markers(self.resolve_type(path, _stack))
        return merge_tag(base, node, track=True)


# ==========================================================================
# Flattening a (resolved) tag tree into rows
# ==========================================================================
class Flattener(object):
    def __init__(self, path_col, subst_params):
        self.path_col = path_col
        self.subst = subst_params
        self.tag_rows = []
        self.alarm_rows = []
        self.instance_rows = []
        self.param_rows = []

    def _instance_params(self, node, parent_params):
        out = {}
        params = node.get("parameters") or {}
        if isinstance(params, dict):
            for k, p in params.items():
                out[k] = fmt(param_value(p), parent_params if self.subst else None)
        return out

    def walk(self, tags, prefix, ctx, params):
        for t in tags or []:
            if not isinstance(t, dict):
                continue
            name = str(t.get("name", ""))
            path = prefix + "/" + name if prefix else name
            ttype = t.get("tagType", "")
            use_params = params if self.subst else None

            if ttype in ("Folder", "Provider"):
                self.walk(t.get("tags"), path, ctx, params)
                continue

            row = dict(ctx)
            row[self.path_col] = path
            row["name"] = name
            row["tagType"] = ttype
            for k, v in t.items():
                if k not in STRUCTURE_KEYS:
                    row[k] = fmt(v, use_params)
            if t.get("__overrides__"):
                row["overridden_props"] = ", ".join(t["__overrides__"])

            child_ctx, child_params = ctx, params
            if ttype == "UdtInstance":
                inst_params = self._instance_params(t, params)
                row["parameters"] = "; ".join("%s=%s" % kv for kv in inst_params.items())
                found = "no" if t.get("__type_missing__") else "yes"
                irow = dict(ctx)
                irow.update({
                    self.path_col: path,
                    "typeId": fmt(t.get("typeId")),
                    "definition_found": found,
                    "parameters": row["parameters"],
                    "members_with_overrides": _count_overrides(t.get("tags")),
                })
                self.instance_rows.append(irow)
                for pk, pv in inst_params.items():
                    prow = dict(ctx)
                    prow.update({"instance": path, "typeId": fmt(t.get("typeId")),
                                 "parameter": pk, "value": pv})
                    self.param_rows.append(prow)
                child_ctx = dict(ctx)
                child_ctx["parent_instance"] = path
                child_ctx["parent_instance_type"] = fmt(t.get("typeId"))
                child_params = dict(inst_params)
                child_params.setdefault("InstanceName", name)
                child_params.setdefault("PathToParentFolder", prefix)

            self.tag_rows.append(row)

            for a in t.get("alarms") or []:
                if not isinstance(a, dict):
                    continue
                ar = dict(ctx)
                ar[self.path_col] = path
                ar["tag_dataType"] = fmt(t.get("dataType"))
                ar["alarm"] = str(a.get("name", ""))
                for k, v in a.items():
                    if k != "name" and k not in INTERNAL_KEYS:
                        ar[k] = fmt(v, use_params)
                if a.get("__overrides__"):
                    ar["overridden_props"] = ", ".join(a["__overrides__"])
                self.alarm_rows.append(ar)

            self.walk(t.get("tags"), path, child_ctx, child_params)


def _count_overrides(tags):
    n = 0
    for t in tags or []:
        if isinstance(t, dict):
            if t.get("__overrides__") or any(
                    isinstance(a, dict) and a.get("__overrides__") for a in t.get("alarms") or []):
                n += 1
            n += _count_overrides(t.get("tags"))
    return n


# ==========================================================================
# Output: columns, CSV, XLSX (hand-built - no third-party libs)
# ==========================================================================
def build_columns(rows, leading=(), preferred=(), trailing=(), always=(), required=()):
    """leading/preferred/trailing appear only if present in the data.
    `required` columns always appear (after preferred); `always` go last."""
    present = set()
    for r in rows:
        present.update(r.keys())
    if not rows:
        return list(leading) + [c for c in required if c not in leading] + list(always)
    fixed = set(leading) | set(preferred) | set(trailing) | set(always) | set(required)
    cols = [c for c in leading if c in present]
    cols += [c for c in preferred if c in present and c not in cols]
    cols += [c for c in required if c not in cols]
    cols += sorted(c for c in present if c not in fixed)
    cols += [c for c in trailing if c in present]
    cols += [c for c in always if c not in cols]
    return cols


class Sheet(object):
    def __init__(self, name, columns, rows, highlight_cols=(), highlight_blank_cols=()):
        self.name = name
        self.columns = columns
        self.rows = rows
        self.highlight_cols = set(highlight_cols)              # always yellow
        self.highlight_blank_cols = set(highlight_blank_cols)  # yellow only when empty


def write_csv(path, sheet):
    # utf-8-sig so Excel opens it with the right encoding
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(sheet.columns)
        for r in sheet.rows:
            w.writerow([r.get(c, "") for c in sheet.columns])


_ILLEGAL_XML = re.compile(u"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# style ids: 0 normal, 1 header, 2 review/input cell
_STYLES = _XML_HEAD + (
    '<styleSheet xmlns="%s">'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="4"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>' % _MAIN_NS)


def _col_letter(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(ref, v, style=0):
    s = ' s="%d"' % style if style else ""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, v)
    text = _ILLEGAL_XML.sub("", "" if v is None else str(v))[:32767]
    if text == "":
        return '<c r="%s"%s/>' % (ref, s) if style else ""
    return '<c r="%s" t="inlineStr"%s><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, s, escape(text))


def _sheet_xml(sheet):
    cols = sheet.columns or ["(empty)"]
    last_col = _col_letter(len(cols) - 1)
    last_row = len(sheet.rows) + 1

    widths = []
    for c in cols:
        w = len(str(c))
        for r in sheet.rows[:500]:
            w = max(w, len(str(r.get(c, ""))))
        widths.append(min(max(w + 2, 8), 60))

    out = [_XML_HEAD, '<worksheet xmlns="%s" xmlns:r="%s">' % (_MAIN_NS, _REL_NS)]
    out.append('<sheetViews><sheetView workbookViewId="0">'
               '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
               '</sheetView></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="15"/><cols>')
    for i, w in enumerate(widths):
        out.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, w))
    out.append('</cols><sheetData>')

    out.append('<row r="1">')
    for i, c in enumerate(cols):
        out.append(_cell("%s1" % _col_letter(i), c, 1))
    out.append('</row>')
    for ri, r in enumerate(sheet.rows, start=2):
        out.append('<row r="%d">' % ri)
        for ci, c in enumerate(cols):
            v = r.get(c, "")
            style = 2 if (c in sheet.highlight_cols or
                          (c in sheet.highlight_blank_cols and v in ("", None))) else 0
            out.append(_cell("%s%d" % (_col_letter(ci), ri), v, style))
        out.append('</row>')
    out.append('</sheetData>')
    out.append('<autoFilter ref="A1:%s%d"/>' % (last_col, last_row))
    out.append('</worksheet>')
    return "".join(out), "$A$1:$%s$%d" % (last_col, last_row)


def _safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "_", name).strip("'")[:31] or "Sheet"
    candidate, n = base, 2
    while candidate.lower() in used:
        suffix = " (%d)" % n
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def write_xlsx(path, sheets):
    used = set()
    names = [_safe_sheet_name(s.name, used) for s in sheets]
    n = len(sheets)

    content_types = [_XML_HEAD,
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(n):
        content_types.append(
            '<Override PartName="/xl/worksheets/sheet%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
    content_types.append('</Types>')

    root_rels = (_XML_HEAD + '<Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>' % (_PKG_REL_NS, _REL_NS))

    wb_rels = [_XML_HEAD, '<Relationships xmlns="%s">' % _PKG_REL_NS]
    for i in range(n):
        wb_rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
                       % (i + 1, _REL_NS, i + 1))
    wb_rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (n + 1, _REL_NS))
    wb_rels.append('</Relationships>')

    sheet_xml, ranges = [], []
    for s in sheets:
        x, rng = _sheet_xml(s)
        sheet_xml.append(x)
        ranges.append(rng)

    wb = [_XML_HEAD, '<workbook xmlns="%s" xmlns:r="%s"><sheets>' % (_MAIN_NS, _REL_NS)]
    for i, nm in enumerate(names):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                  % (escape(nm, {'"': "&quot;"}), i + 1, i + 1))
    wb.append('</sheets><definedNames>')
    for i, nm in enumerate(names):
        wb.append('<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s</definedName>'
                  % (i, escape("'%s'!%s" % (nm.replace("'", "''"), ranges[i]))))
    wb.append('</definedNames></workbook>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(content_types))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        z.writestr("xl/styles.xml", _STYLES)
        for i, x in enumerate(sheet_xml):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), x)


# ==========================================================================
# Shared output driver
# ==========================================================================
COMMON_NOTES = [
    ("Blank cells", "Property not set in the export, so the Ignition default applies."),
    ("Bindings", "Values shown as 'Expression: ...', 'Tag: ...', 'parameter: ...' are bindings, not literal values."),
    ("Review columns", "Review Status / Reviewer / Comments (yellow) are for stakeholders to fill in."),
    ("Filtering", "Every sheet has filters on the header row - filter by UDT / path to focus on one item."),
]


def info_sheet(title, sources, extra_rows):
    rows = [{"Item": "Report", "Value": title},
            {"Item": "Generated", "Value": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}]
    for s in sources:
        rows.append({"Item": "Source file", "Value": os.path.abspath(s)})
    for k, v in list(extra_rows) + COMMON_NOTES:
        rows.append({"Item": k, "Value": v})
    return Sheet("Info", ["Item", "Value"], rows)


def write_outputs(sheets, out_dir, base_name, csv_out=True, xlsx_out=True):
    os.makedirs(out_dir, exist_ok=True)
    written = []
    if csv_out:
        for s in sheets:
            if s.name == "Info":
                continue
            slug = re.sub(r"[^A-Za-z0-9]+", "_", s.name).strip("_").lower()
            p = os.path.join(out_dir, "%s_%s.csv" % (base_name, slug))
            write_csv(p, s)
            written.append(p)
    if xlsx_out:
        p = os.path.join(out_dir, "%s.xlsx" % base_name)
        write_xlsx(p, sheets)
        written.append(p)
    return written
