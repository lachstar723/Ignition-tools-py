# Ignition-tools-py

Small, dependency-free Python tools for working with Inductive Automation **Ignition** tag exports.

## Tag / UDT export → CSV + Excel

Turn Ignition JSON tag exports into flat CSVs and a multi-sheet `.xlsx` so operators and
electrical engineers can review tag and alarm configuration without reading JSON.

- Standard library only (`json`, `csv`, `zipfile`) – no pandas, numpy or openpyxl. Python 3.6+.
- Columns are generated dynamically from whatever properties exist in the export.
- Workbook has a frozen header row, filters on every sheet, and yellow
  **Review Status / Reviewer / Comments** columns for stakeholders to fill in.

| File | Purpose |
|---|---|
| `udt_export_to_sheets.py` | UDT **definitions** export → UDT Definitions, UDT Parameters, Members, Alarms, Nested UDTs |
| `tag_export_to_sheets.py` | **Tag** export (folders, tags, UDT instances) → Tags, Alarms, UDT Instances, Instance Parameters |
| `ig_export_common.py` | Shared helpers (must sit in the same folder as the two scripts) |

### Exporting from Ignition

Designer → Tag Browser → right-click a UDT, folder or provider → **Export Tags** → JSON.
Use the *UDT Definitions* tab for UDTs and the *Tags* tab for instances.

### Usage

```bash
# UDT definitions
python udt_export_to_sheets.py udts.json
python udt_export_to_sheets.py udts.json --filter Charger -o review

# Tags / instances (pass the UDT export so instances show their full config)
python tag_export_to_sheets.py tags.json --types udts.json
python tag_export_to_sheets.py tags.json --types udts.json --filter "GT1/" --xlsx-only
```

#### Required alarm associated data

Associated data on alarms is picked up automatically as columns. To make sure specific
fields appear on **every** alarm, even ones that don't have them yet, list them with
`--assoc-fields`:

```bash
python udt_export_to_sheets.py udts.json --assoc-fields "Area,Equipment,Response"
```

Blank required fields are highlighted yellow in the workbook, and a
`missing_associated_data` column lists what each alarm still needs, so you can filter on it.
Names are case-sensitive (as in Ignition); the script warns if it finds a near-match such
as `equipment` vs `Equipment`. Quote a name that contains a comma: `'Area,"Unit, Bay"'`.

Run either script with `-h` for all options (`--no-resolve`, `--no-param-subst`,
`--no-review-cols`, `--csv-only`, `--xlsx-only`, `--name`).

### How it reads the export

- **UDT instances** only export properties overridden on the instance. With `--types`
  (or a full provider export containing `_types_`) the definition is merged in so every
  member and alarm shows its effective config; `overridden_props` flags what differs.
- **Parent UDT inheritance** and **nested UDT instances** are resolved.
- **Parameter bindings** are substituted per instance in the tag script
  (e.g. `[{DeviceName}]HR100` → `[GT1_CHG1]HR100`). Other bindings are shown as
  `Expression: ...`, `Tag: ...`.
- **Blank cells** mean the property was not in the export, so the Ignition default applies.

Tested against Ignition 8.1-format JSON exports.
