#!/usr/bin/env python3
"""
udt_export_to_sheets.py
Flatten an Ignition UDT *definitions* JSON export into CSVs + one .xlsx workbook
so operators / electrical engineers can review tag and alarm config per UDT.

Export from Designer: Tag Browser -> UDT Definitions -> right-click the UDT
(or a folder, or _types_) -> Export Tags -> JSON.

Usage:
    python udt_export_to_sheets.py udts.json
    python udt_export_to_sheets.py udts.json -o review --filter Charger
    python udt_export_to_sheets.py udts.json --xlsx-only
    python udt_export_to_sheets.py udts.json --assoc-fields "Area,Equipment,Response"

Alarm associated data is exported as extra keys on each alarm, so it is picked
up automatically as columns. --assoc-fields forces the listed fields to exist
as columns on every alarm (blank + yellow when not set) so they can be filled in.

Sheets / CSVs produced:
    UDT Definitions  one row per UDT (parent type, parameters, member/alarm counts)
    UDT Parameters   one row per UDT parameter with its default value
    Members          one row per member tag (all configured tag properties)
    Alarms           one row per alarm on every member (all configured alarm properties)
    Nested UDTs      (only if a UDT contains other UDT instances)

Requires ig_export_common.py in the same folder. Standard library only.
"""
import argparse
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ig_export_common import (  # noqa: E402
    PREFERRED_ALARM_COLS, PREFERRED_TAG_COLS, REVIEW_COLS, STRUCTURE_KEYS,
    Flattener, Sheet, UdtLibrary, build_columns, fmt, info_sheet, is_binding,
    load_json, param_value, write_outputs,
)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Ignition UDT definitions JSON -> CSV/XLSX for stakeholder review")
    ap.add_argument("udt_json", nargs="+", help="UDT definitions export file(s) (.json)")
    ap.add_argument("-o", "--out-dir", help="output folder (default: <input name>_sheets next to the input)")
    ap.add_argument("--name", help="base name for output files (default: input file name)")
    ap.add_argument("--filter", action="append", default=[],
                    help="only include UDTs whose path contains this text (case-insensitive, repeatable)")
    ap.add_argument("--no-resolve", action="store_true",
                    help="show definitions exactly as exported (don't merge parent-UDT inheritance or expand nested UDTs)")
    ap.add_argument("--assoc-fields", default="",
                    help='comma-separated alarm associated data fields that must appear as columns even if no alarm '
                         'has them yet, e.g. --assoc-fields "Area,Equipment,Response". Blank cells are highlighted '
                         'and a missing_associated_data column lists what each alarm lacks. Names are case-sensitive.')
    ap.add_argument("--no-review-cols", action="store_true", help="omit Review Status / Reviewer / Comments columns")
    grp = ap.add_mutually_exclusive_group()
    grp.add_argument("--csv-only", action="store_true")
    grp.add_argument("--xlsx-only", action="store_true")
    args = ap.parse_args(argv)

    lib = UdtLibrary()
    for f in args.udt_json:
        lib.add_export(load_json(f))
    if not lib.types:
        sys.exit("No UDT definitions (tagType 'UdtType') found. Export from the 'UDT Definitions' tab, "
                 "or use tag_export_to_sheets.py for tag instances.")

    filters = [f.lower() for f in args.filter]
    paths = sorted(p for p in lib.types if not filters or any(f in p.lower() for f in filters))
    if not paths:
        sys.exit("No UDTs matched the filter(s): %s" % ", ".join(args.filter))

    fl = Flattener("member_path", subst_params=False)
    def_rows, param_rows = [], []

    for p in paths:
        raw = lib.types[p]
        node = raw if args.no_resolve else lib.resolve_type(p)
        n_tags, n_alarms = len(fl.tag_rows), len(fl.alarm_rows)
        fl.walk(node.get("tags"), "", {"udt": p}, None)

        params = node.get("parameters") or {}
        for pk, pv in (params.items() if isinstance(params, dict) else []):
            param_rows.append({
                "udt": p, "parameter": pk,
                "dataType": pv.get("dataType", "") if isinstance(pv, dict) and not is_binding(pv) else "",
                "default_value": fmt(param_value(pv)),
            })

        drow = {
            "udt": p,
            "name": raw.get("name", ""),
            "parent_type": fmt(raw.get("typeId")),
            "parameters": "; ".join("%s=%s" % (k, fmt(param_value(v))) for k, v in
                                    (params.items() if isinstance(params, dict) else [])),
            "member_tags": len(fl.tag_rows) - n_tags,
            "alarms": len(fl.alarm_rows) - n_alarms,
        }
        if node.get("__parent_missing__"):
            drow["note"] = "Parent UDT not found in export - inherited members not shown"
        for k, v in raw.items():
            if k not in STRUCTURE_KEYS and k != "typeId":
                drow[k] = fmt(v)
        def_rows.append(drow)

    assoc = parse_field_list(args.assoc_fields)
    if assoc:
        apply_required_assoc(fl.alarm_rows, assoc)

    review = () if args.no_review_cols else REVIEW_COLS
    sheets = [
        info_sheet("Ignition UDT definitions review", args.udt_json, [
            ("UDTs", len(def_rows)), ("Member tags", len(fl.tag_rows)), ("Alarms", len(fl.alarm_rows)),
            ("Inheritance", "Not resolved (--no-resolve)" if args.no_resolve
             else "Parent-UDT members and nested UDT members are merged in"),
        ] + ([("Required associated data", ", ".join(assoc) +
               " (yellow = not set yet; see missing_associated_data on the Alarms sheet)")] if assoc else [])),
        Sheet("UDT Definitions",
              build_columns(def_rows, ["udt", "name", "parent_type", "parameters", "member_tags", "alarms"],
                            ["documentation"], ["note"]),
              def_rows),
        Sheet("UDT Parameters", ["udt", "parameter", "dataType", "default_value"], param_rows),
        Sheet("Members",
              build_columns(fl.tag_rows, ["udt", "member_path", "name", "tagType"], PREFERRED_TAG_COLS,
                            ["parent_instance", "parent_instance_type", "overridden_props"], review),
              fl.tag_rows, review),
        Sheet("Alarms",
              build_columns(fl.alarm_rows, ["udt", "member_path", "tag_dataType", "alarm"], PREFERRED_ALARM_COLS,
                            ["parent_instance", "parent_instance_type", "overridden_props",
                             "missing_associated_data"], review, required=assoc),
              fl.alarm_rows, review, highlight_blank_cols=assoc),
    ]
    if fl.instance_rows:
        sheets.append(Sheet("Nested UDTs",
                            build_columns(fl.instance_rows, ["udt", "member_path", "typeId", "definition_found",
                                                             "parameters", "members_with_overrides"]),
                            fl.instance_rows))

    in0 = args.udt_json[0]
    base = args.name or os.path.splitext(os.path.basename(in0))[0]
    out_dir = args.out_dir or os.path.join(os.path.dirname(os.path.abspath(in0)), base + "_sheets")
    written = write_outputs(sheets, out_dir, base, not args.xlsx_only, not args.csv_only)

    print("UDTs: %d | member tags: %d | alarms: %d" % (len(def_rows), len(fl.tag_rows), len(fl.alarm_rows)))
    for w in written:
        print("  wrote", w)


def parse_field_list(text):
    """Parse a comma-separated list; quote names containing commas: 'Area,"Unit, Bay"'."""
    if not text or not text.strip():
        return []
    fields = []
    for f in next(csv.reader([text], skipinitialspace=True)):
        f = f.strip()
        if f and f not in fields:
            fields.append(f)
    return fields


def apply_required_assoc(alarm_rows, fields):
    """Ensure each required associated data field exists on every alarm row and
    record which ones each alarm is missing."""
    seen_keys = set()
    for r in alarm_rows:
        seen_keys.update(r.keys())
    for f in fields:
        if f not in seen_keys:
            near = sorted(k for k in seen_keys if k.lower() == f.lower())
            if near:
                print("WARNING: associated data '%s' not found, but '%s' exists - names are case-sensitive."
                      % (f, "', '".join(near)), file=sys.stderr)
    for r in alarm_rows:
        missing = [f for f in fields if r.get(f, "") in ("", None)]
        for f in missing:
            r.setdefault(f, "")
        r["missing_associated_data"] = "; ".join(missing)


if __name__ == "__main__":
    main()
