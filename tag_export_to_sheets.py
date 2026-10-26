#!/usr/bin/env python3
"""
tag_export_to_sheets.py
Flatten an Ignition *tag* JSON export (folders, standard tags, UDT instances)
into CSVs + one .xlsx workbook for stakeholder review of tag and alarm config.

Export from Designer: Tag Browser -> Tags tab -> right-click a folder or the
provider -> Export Tags -> JSON.

IMPORTANT - UDT instances only export the properties OVERRIDDEN on the
instance. To see every member tag/alarm with its effective config, also pass
the UDT definitions export with --types (or export the whole provider, which
includes _types_):

    python tag_export_to_sheets.py tags.json --types udts.json
    python tag_export_to_sheets.py tags.json --types udts.json --filter "Unit1/" --filter "Unit2/"
    python tag_export_to_sheets.py provider_export.json          (contains _types_ already)

Sheets / CSVs produced:
    Tags                 one row per tag (UDT members included), full tag path
    Alarms               one row per alarm
    UDT Instances        one row per UDT instance (type, parameters, # overridden members)
    Instance Parameters  one row per instance parameter value

Parameter bindings (e.g. opcItemPath "[{DeviceName}]HR100") are resolved with
each instance's parameter values. Use --no-param-subst to show them raw.

Requires ig_export_common.py in the same folder. Standard library only.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ig_export_common import (  # noqa: E402
    PREFERRED_ALARM_COLS, PREFERRED_TAG_COLS, REVIEW_COLS,
    Flattener, Sheet, UdtLibrary, build_columns, info_sheet, load_json,
    top_nodes, write_outputs,
)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Ignition tag JSON export -> CSV/XLSX for stakeholder review")
    ap.add_argument("tags_json", help="tag export file (.json)")
    ap.add_argument("-t", "--types", action="append", default=[],
                    help="UDT definitions export (.json) used to expand instances (repeatable)")
    ap.add_argument("-o", "--out-dir", help="output folder (default: <input name>_sheets next to the input)")
    ap.add_argument("--name", help="base name for output files (default: input file name)")
    ap.add_argument("--filter", action="append", default=[],
                    help="only include tags whose path contains this text (case-insensitive, repeatable)")
    ap.add_argument("--no-param-subst", action="store_true",
                    help="show parameter bindings as written instead of resolving instance parameter values")
    ap.add_argument("--no-review-cols", action="store_true", help="omit Review Status / Reviewer / Comments columns")
    grp = ap.add_mutually_exclusive_group()
    grp.add_argument("--csv-only", action="store_true")
    grp.add_argument("--xlsx-only", action="store_true")
    args = ap.parse_args(argv)

    data = load_json(args.tags_json)
    lib = UdtLibrary()
    lib.add_export(data)                 # picks up _types_ if the export has it
    for f in args.types:
        lib.add_export(load_json(f))

    nodes = [n for n in top_nodes(data) if not (isinstance(n, dict) and n.get("name") == "_types_")]
    has_instances = _has_instances(nodes)
    if lib.types:
        nodes = lib.expand_tags(nodes)
    elif has_instances:
        print("WARNING: export contains UDT instances but no UDT definitions were supplied.\n"
              "         Instance members will only show properties overridden on the instance.\n"
              "         Re-run with --types <udt_definitions.json> for the full picture.", file=sys.stderr)

    fl = Flattener("path", subst_params=not args.no_param_subst)
    fl.walk(nodes, "", {}, None)

    tag_rows, alarm_rows = fl.tag_rows, fl.alarm_rows
    inst_rows, param_rows = fl.instance_rows, fl.param_rows
    if args.filter:
        fs = [f.lower() for f in args.filter]
        keep = lambda p: any(f in (p or "").lower() for f in fs)  # noqa: E731
        tag_rows = [r for r in tag_rows if keep(r["path"])]
        alarm_rows = [r for r in alarm_rows if keep(r["path"])]
        inst_rows = [r for r in inst_rows if keep(r["path"])]
        param_rows = [r for r in param_rows if keep(r["instance"])]

    missing = sorted({r["typeId"] for r in inst_rows if r["definition_found"] == "no"})
    if missing and lib.types:
        print("WARNING: UDT definition not found for: %s" % ", ".join(missing), file=sys.stderr)

    review = () if args.no_review_cols else REVIEW_COLS
    expanded = "Expanded using %d UDT definitions" % len(lib.types) if lib.types else \
        "NOT expanded - no UDT definitions supplied (members show instance overrides only)"
    sheets = [
        info_sheet("Ignition tag configuration review", [args.tags_json] + args.types, [
            ("Tags", len(tag_rows)), ("Alarms", len(alarm_rows)), ("UDT instances", len(inst_rows)),
            ("UDT instances", expanded),
            ("overridden_props", "Properties set on the instance that differ from the UDT definition."),
        ] + ([("Definitions missing", ", ".join(missing))] if missing else [])),
        Sheet("Tags",
              build_columns(tag_rows, ["path", "name", "tagType", "parent_instance", "parent_instance_type"],
                            PREFERRED_TAG_COLS, ["overridden_props"], review),
              tag_rows, review),
        Sheet("Alarms",
              build_columns(alarm_rows, ["path", "tag_dataType", "alarm", "parent_instance",
                                         "parent_instance_type"],
                            PREFERRED_ALARM_COLS, ["overridden_props"], review),
              alarm_rows, review),
        Sheet("UDT Instances",
              build_columns(inst_rows, ["path", "typeId", "definition_found", "parameters",
                                        "members_with_overrides", "parent_instance"]),
              inst_rows),
        Sheet("Instance Parameters",
              build_columns(param_rows, ["instance", "typeId", "parameter", "value"]),
              param_rows),
    ]

    base = args.name or os.path.splitext(os.path.basename(args.tags_json))[0]
    out_dir = args.out_dir or os.path.join(os.path.dirname(os.path.abspath(args.tags_json)), base + "_sheets")
    written = write_outputs(sheets, out_dir, base, not args.xlsx_only, not args.csv_only)

    print("tags: %d | alarms: %d | UDT instances: %d" % (len(tag_rows), len(alarm_rows), len(inst_rows)))
    for w in written:
        print("  wrote", w)


def _has_instances(tags):
    for t in tags or []:
        if isinstance(t, dict) and (t.get("tagType") == "UdtInstance" or _has_instances(t.get("tags"))):
            return True
    return False


if __name__ == "__main__":
    main()
