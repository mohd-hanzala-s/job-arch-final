# IRIS Job Architecture — Final v1.15 (12 October 2026)

Three files, **HR and business heads only**. v1.15 is v1.14 with Band E set apart from the Business Heads in its own job
family. v1.14 was v1.13 plus HR's decisions of 12 October 2026 and the new Business-Head Workbook.

## v1.15 — Band E in its own job family

| Level | Job family | Sub-family | People |
|---|---|---|---|
| E2 | Enterprise Leadership | Chief Executive | Balachandran K, CEO |
| E1 | Enterprise Leadership | Chief Officers · CXO | Dixit Jasani, CRO · Vineet Kandoi, CFO · Deepta Rangarajan, COO · P K X Thomas, CTO |
| C1–D2 heads | Executive Leadership | Business Head | the other 13 department and function heads |

- **Why a separate family:** it mirrors the separate executive stream in the Mercer and Radford levelling systems, which use E1–E6 codes (S01, S03).
- **New catalogue roles:** RC-EN-CE-E2 and RC-EN-CO-E1.
- **Business-Head Workbook:** gains an `RB · Enterprise Leadership` matrix sheet, so the families now total 21.
- **No levels change.** All 43 checks pass.

| File | For | What it is |
|---|---|---|
| `IRIS_JA_2_HR_Super_Master_v1.15.xlsx` | HR | Every v1.12.2 sheet, the v1.13 additions, and the v1.14 decisions below |
| `IRIS_JA_2S_HR_Master_Simplified_v1.15.xlsx` | HR | Five sheets, one line from method to person to role, now with the HR proposed level |
| `IRIS_JA_7_Business_Head_Workbook_v1.15.xlsx` | Business heads | One file for all businesses, top down, where every Low / Medium / High is live |

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

- **Department and function heads.** The top person of each department who leads a team at C1 or above: 18 people. From v1.15, Band E (the CEO and the 4 CXOs) sits in **Enterprise Leadership**, and the other 13 sit in **Executive Leadership · Business Head**. Their earlier family is kept on the row, and 6 Executive Leadership · Business Head roles are added to the catalogue.
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
  6. One "RB · family" sheet per job family (21)
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
| Department and function heads moved | 18 · 13 to Executive Leadership · Business Head, 5 to Enterprise Leadership |
| HR proposed vs the model | 68 down, protected · 8 up · every pair passes the manager rule |
| Department-ladder roles, each with editable weightage | 291 |
| Role & band matrix sheets | 21 families × 19 level rows |
| Benchmarking input differs from the HR proposed band | 220 people — the business heads' first look |

## Rebuild and check

The source talent-mapping files are not in this repository because they hold personal data.

```
cd build_v2
python3 build_super.py <folder with the five source files>
python3 build_simplified.py <folder with the five source files>
python3 build_bh.py <folder with the five source files>
python3 verify_v115.py        # 43 checks, including live edits recalculated in LibreOffice; exit 0 = all pass
```
