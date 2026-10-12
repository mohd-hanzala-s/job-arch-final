# IRIS Job Architecture — Final v1.14 (12 October 2026)

Three files, **HR and business heads only**. v1.14 is v1.13 plus HR's decisions of 12 October 2026 and the new
Business-Head Workbook.

| File | For | What it is |
|---|---|---|
| `IRIS_JA_2_HR_Super_Master_v1.14.xlsx` | HR | Every v1.12.2 sheet, the v1.13 additions, and the v1.14 decisions below |
| `IRIS_JA_2S_HR_Master_Simplified_v1.14.xlsx` | HR | Five sheets, one line from method to person to role, now with the HR proposed level |
| `IRIS_JA_7_Business_Head_Workbook_v1.14.xlsx` | Business heads | One file for all businesses, top down, where every Low / Medium / High is live |

## HR's edits, incorporated

Your two files differed from v1.13 in six places:

- **Ladder titles.**
  - D2 is now Senior Vice President / Business Head.
  - D1 is now Associate Vice President / Distinguished Principal, so the D1 specialist title is decided: **Distinguished Principal**.
  - B1 is now Team Lead / Specialist / Asst Manager.
  - These titles are applied wherever an IRIS title appears: 306 cells across the ladder, catalogue, requirement and reference sheets. Raw data and people's designations are left as received.
- **Hemant Karale:** sub-family is now Business Head.
- **Sales sketch (your `Sheet1`):** carried as an HR draft into the Simplified file, sheet 1 section 5f, and onto the Sales matrix sheet of the Business-Head Workbook. **The unit, and how L1–L6 map to IRIS levels, still need confirming.**
- **A filter left on Talent Mapping:** removed, so all rows are visible again.
- **"59%" became 0.59:** this was Excel converting the cell when you saved. It's kept as 59%.

The Super Master had no value changes.

## HR decisions applied

- **Department and function heads.** The top person of each department who leads a team at C1 or above moves to job family **Executive Leadership**, sub-family **Business Head**. That is 18 people: the CEO, 4 CXOs, the business heads, VP – HR, and the heads and associate heads of departments. Their earlier family is kept on the row, and 6 Executive Leadership · Business Head roles are added to the catalogue.
- **Top-down manager rule, hybrid.** Applied from the CEO down:
  - Under a head at C1 or above, a reportee sits in a **lower band**.
  - Below C1, a reportee sits at least **one level below** the manager.
  - A Band A lead whose reportees sit at or above the lead's own level becomes B1 Team Lead.
  - The CEO with the CXOs is the agreed exception.
- **Result:** 68 people move down, with pay and title protected for one cycle, and 8 move up. Every manager–reportee pair now passes. The v1.13 model level stays in FINAL LEVEL, and the HR proposed level sits beside it.
- **Why not the strict band rule:** it was tested first. Reporting chains run up to 9 layers but there are only five bands, so capping by band cascades: 88 people moved down and 172 had no band left below their manager.

## Business-Head Workbook — how it works

- **Sheets, in order:**
  1. How to use
  2. Talent Mapping
  3. Raw Data
  4. Market Benchmark
  5. Department Ladder
  6. One "RB · family" sheet per job family (20)
  7. Backend Data, last
- **Talent Mapping columns, left to right:**
  1. Core employee details
  2. Summary of the responses received, with a link to the person's full raw data
  3. **Weightage L/M/H** for seven parameters: Complexity, Autonomy, Accountability, Leadership, Business Impact, **Experience**, **Education** — Experience and Education are now separate
  4. **Assessment L/M/H** against the band
  5. **Market benchmark** for the role group
  6. **HR proposed band**, **Input from benchmarking**, and the **Final proposed band**, which is live, plus an override, the manager check and a comment
- **How Low / Medium / High work:**
  - *Weightage:* L = 1, M = 3, H = 9 points, turned into shares of the score. The defaults come from the published market weights (Korn Ferry Hay, Mercer IPE, five public plans) with the IRIS type-of-work adjustment.
  - *Assessment:* L = works at the level below, M = meets the band, H = works at the level above. The weighted net of the seven moves the person one level from the HR proposed band once it reaches ±0.5.
- **The live chain:** RB family sheet weightage → Department Ladder weightage → each person's weightage → their final proposed band → the manager check. Each step can be overwritten, and every change recalculates at once.
- **Untouched**, the final proposed band equals the HR proposed band for everyone.

## Numbers

| | |
|---|---|
| People | 586, top down from the CEO |
| Department and function heads moved to Executive Leadership · Business Head | 18 |
| HR proposed vs the model | 68 down, protected · 8 up · every pair passes the manager rule |
| Department-ladder roles, each with editable weightage | 291 |
| Role & band matrix sheets | 20 families × 19 level rows |
| Benchmarking input differs from the HR proposed band | 220 people — the business heads' first look |

## Rebuild and check

The source talent-mapping files are not in this repository because they hold personal data.

```
cd build_v2
python3 build_super.py <folder with the five source files>
python3 build_simplified.py <folder with the five source files>
python3 build_bh.py <folder with the five source files>
python3 verify_v114.py        # 42 checks, including live edits recalculated in LibreOffice; exit 0 = all pass
```
