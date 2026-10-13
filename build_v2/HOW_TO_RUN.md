# IRIS Job Architecture v2.2: build code

One pipeline builds all three workbooks from the source files:
- **Master:** `IRIS_JA_Master_v2.2.xlsx`
- **Simplified:** `IRIS_JA_Simplified_v2.2.xlsx`
- **Hierarchy Audit:** `IRIS_JA_Hierarchy_Audit_v2.2.xlsx`

## What you need

- Python 3.11 or later, with:
  - `openpyxl` 3.1.5
  - `lxml` (only for the optional schema check)
- LibreOffice 7 or later on the PATH as `soffice`. It recalculates the formulas so that every cell opens with its value.
- **The source files.** These hold personal data, so they are not in this package.
  - Into the folder you pass to `build_model.py`:
    - the five talent-mapping files (CARBON, iFile, iDEAL, Gopal, Shilpa)
    - `IRIS_Employee_Raw_Data_All_Sources.xlsx`
  - Into the folder named by `JA_V1`:
    - the six v1 `IRIS_<Business>_Talent_Mapping.xlsx` workbooks
    - the Employee Masters
    - the v1.15 HR Super Master

## Run

```
cd code
export JA_V1=<v1 folder>            # default /home/user/job-arc
export JA_WORK=<a scratch folder>   # for the LibreOffice recalculation, default /tmp

python3 build_model.py <source folder>     # writes the Master and the Simplified, about 12 minutes
python3 hier_audit.py  <source folder>     # writes the Hierarchy Audit from the Master

# checks: 49 on the Master, 45 on the Simplified, including live edits recalculated in LibreOffice
JA_FILE=<workbook> JA_SRC=<source folder> python3 verify_model.py

# the cross-sheet QA sweep: counts, levels, roles, titles, managers, privacy, text glitches
python3 qa_sweep.py <workbook>

# the scan for anything that makes Excel offer to repair a file
python3 xlsx_check.py <workbook> [<workbook> …]
```

**Optional schema check.** Set `JA_XSD` to a folder holding the ECMA-376 SpreadsheetML schemas. `xlsx_check.py` and `verify_model.py` then also check every sheet, the styles and the workbook against the schema.

The schemas are in `schemaorg_apache_xmlbeans/src` inside Apache POI's `ooxml-schemas-1.4.jar`, on Maven Central.

## Files

| File | What it does |
|---|---|
| `build_model.py` | Builds the Master and the Simplified. It holds the role framework, the placement, the level rule, every sheet and the live formulas. |
| `build_bh.py` | Shared pieces: parameters, weights, market tilt, cell and header styles. Also builds the v1.14 business-head workbook. |
| `build_super.py` | Person values and links shared with the HR Super Master builder. |
| `hier_audit.py` | The Hierarchy Audit: every reporting line checked against the level rule, plus every source file that names a different manager. |
| `verify_model.py` | The check suite, with live tests recalculated in LibreOffice. |
| `qa_sweep.py` | The cross-sheet QA sweep: every sheet checked against every other, plus privacy and text glitches. |
| `xlsx_check.py` | XML-level scan for Excel repair triggers, plus the optional schema validation. |
| `ja_common.py` | Text, name and ID helpers. |
| `ja_person.py` | Reads the per-person source sheets. |
| `ja_raw.py` | Joins every source into Raw Data and blanks the placeholders. |
| `ja_scale.py` | The five-step weightage and rating scales and the scoring rule. |
| `ja_market.py` | Market evidence, references and the method sources. |
| `ja_style.py` | Banners, colours and fonts. |
| `ja_v114.py` | Levels, bands, titles and ladders. |
| `ja_xlsx.py` | Writes the workbook, recalculates it in LibreOffice and stores the cached values. |
