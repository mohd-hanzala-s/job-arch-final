# IRIS Job Architecture v2.2 (13 October 2026)

Two workbooks, built by one pipeline from the same data:

| File | For | What it holds |
|---|---|---|
| `IRIS_JA_Master_v2.2.xlsx` | HR | Everything: 9 numbered sheets plus Backend Data and References & Benchmarks. Raw Data holds every bit of data received plus the analysis on file from every earlier version. |
| `IRIS_JA_Simplified_v2.2.xlsx` | Business heads | The same live model in 7 numbered sheets plus Backend Data and References. Raw Data holds only what was received. The market reference columns on Talent Mapping are folded away; click + to open them. |
| `IRIS_JA_Hierarchy_Audit_v2.2.xlsx` | HR | Every reporting line checked against the level rule, and every person whose manager in another source file differs from the Employee Master. |

Both files have no formula errors and pass every check in `build_v2/verify_model.py`: CHECKS_M checks on the Master and CHECKS_S on the Simplified. One check scans the file's XML for anything that makes Excel offer to repair it.

## What changed in v2.2

### The framework is role-based
- **A role is a sub-family**, such as "Front-End / UI Development". There are 64 roles, held or not yet held.
- Every role runs on one shared ladder from A1 Intern to E2.
  - In Bands B and C a role has two lanes, IC and Manager. This is the dual career path, and the market weighs the lanes differently.
  - Band A, the Player-Coach steps (B+, C+) and Bands D and E have one shared path.
- **Why:** job architecture keeps families and sub-families separate from levels, and the levels mean the same in every family (S78 Payscale, S80 Mercer). Parallel IC and manager tracks are the standard dual ladder (S79 WorldatWork).
- **Weightage is set per role, level and lane** on sheet 3. A leader's entry on a role's ROLE row re-weights that role everywhere in IRIS, and an entry on a level row wins over it.

### No overlap in a team: the level rule
- **Every reportee sits at a lower level than their manager.** A "band" here means each of the 14 levels: A1 A2 A3 · B1 B2 B3 B+ · C1 C2 C+ · D1 D2 · E1 E2. So a B1 can report to a B2 or a B3, but never to another B1.
- **Everyone reports to their original manager.** No reporting line is changed.
- **The deepest line fits.** It is CARBON: CEO (E2) → Dixit Jasani (E1) → Anuradha RK (D2) → Suman Adhikary (D1) → K Ashok Kumar (C2) → Vijay Parulekar (C1) → Janardan Bale (B2) → Tejas Kat (B1) → Dhanraj Patil (A3).
- **Placement uses the fewest moves.** 48 reportees were at or above their manager's level in the evidence placement. The model moves the fewest people needed, mostly by one level, and moves up and down count the same. People who move down keep their pay and title for one cycle.
- **Result:** no structure exceptions.
- **The rule also works live.** If a leader's input moves a manager down, their reportees are held one level below, and the "Level rule · live" column says why. The effect carries down the line.

### Raw Data
- **Added:** the All-Sources collation (`IRIS_Employee_Raw_Data_All_Sources.xlsx`). It brings in:
  - the six-theme survey summary
  - the survey follow-up and excluded lists
  - outreach status
  - the CARBON 9-box grid (box, star, leadership documents naming the person)
  - extra unnamed columns from the manager files
  - its 22 people outside the roster
- **Already joined:** the four talent-mapping files uploaded again were byte-identical to the ones already in.
- **Checked:** all 71,000+ non-blank, non-personal cells of the All-Sources file, and all 10,394 survey answers, are on the right person's row.
- **"No input" stays blank:** placeholder values such as "Not provided", "Not rated" or "Not placed" mean nothing was received, so the cell is blank.

### Glitches fixed
- The "◂ Start" link on every sheet now goes to "How it works".
- Column-group titles are merged across their columns, so they are never cut off.
- The colour legend now matches this model: yellow is yours, blue is the IRIS base, red is a warning, orange is a title to align, plus the band colours.
- On shared rows (Band A, the Player-Coach steps, Bands D and E), the title is the one people actually hold, not two titles joined.
- The References tables lost their empty spacer columns.
- The check suite now includes a glitch scan:
  - links and named ranges point to real sheets
  - formulas are within Excel's limits
  - row heights are valid
  - no stray "None" text
  - every column has a title
  - drop-downs are valid

### Excel repair message fixed
- Excel said "We found a problem with some content" when opening the Master.
- **Cause:** four named ranges (W5_KEY, W5_MID, R5_KEY, R5_MOVE) began with a cell address (W5, R5), which Excel does not allow in a name.
- **Fix:**
  - The names are now WT_KEY, WT_MID, LIST_WT, RATE_KEY, RATE_MOVE and LIST_RATE.
  - Links between sheets are now written as internal links, not as web addresses starting with "#".
  - A new check, `build_v2/xlsx_check.py`, scans the XML of both files for every known repair trigger.

### Hierarchy audit
- **Every reporting line follows the level rule.** Each reportee sits below their manager, both at the IRIS level and at the level in force.
- **No reporting problems found:** no loops, no self-reports, and every named manager is in the roster.
- **The model uses the managers in the Employee Master of 23 September.** It matches the Master for every person.
- **57 people have a different manager in another source file.** The files are the older Employee Master, the CARBON files and the iFile and iDEAL talent mapping. Nothing was changed; the audit lists each one for HR to confirm.
- **The CARBON Mapping Master proposes 39 new managers.** These are listed with a level-rule check, but not applied.

## Sheets (Master)

| # | Sheet |
|---|---|
| 1 | How it works: the story, the scales, the role framework, the level rule |
| 2 | Ladder: 12 levels, 5 bands and 2 Player-Coach steps, each with its own experience range |
| 3 | Role Ladder & Weightage: the base, 64 roles × levels × lanes, with IRIS base (HR may edit), Market, Leader, In use and Share |
| 4 | Role Definitions: what ratings 1, 3 and 5 mean per role and level |
| 5 | Role Journey: two rows per role, IRIS and Market benchmark, A1 to E2 |
| 6 | Market Benchmark: published factor shares in the five steps |
| 7 | Talent Mapping: every person, the survey in full, and the live proposal with the level rule |
| 8 | Person Card |
| 9 | Raw Data: about 410 columns |
| — | Backend Data |
| — | References & Benchmarks, last |

The Simplified file drops sheets 4 and 6 and the "on file" analysis, and renumbers the rest.

## Scales and rule

- **Weightage** comes in five 20-point importance steps: VL, L, M, H and VH, counted at 10, 30, 50, 70 and 90. You can type a number from 0 to 100 instead.
- **Ratings** run from 1 to 5 in half-level steps.
- **The shift** is the live weighted rating minus IRIS's anchor. Half a level or more moves the person one level, up to two levels.
- **The level rule then applies**, keeping every reportee below their manager's level.

See How it works in either file for the full explanation.

## Rebuild and check

```
cd build_v2
JA_V1=<folder with the six v1 workbooks and the Employee Masters> \
python3 build_model.py <folder with the source files: the five talent-mapping files and the All-Sources file>
JA_FILE=<the workbook> JA_SRC=<the same source folder> python3 verify_model.py
JA_V1=<as above> python3 hier_audit.py <the same source folder>
```

The source files and the v1 workbooks hold personal data, so they are never committed. The workbooks themselves are shared directly, not committed, while this repository is public.
