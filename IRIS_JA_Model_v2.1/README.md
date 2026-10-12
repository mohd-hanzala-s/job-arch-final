# IRIS Job Architecture Model v2.1 (12 October 2026)

One workbook for HR and business leaders: `IRIS_JA_Model_v2.1.xlsx`.

## The story the workbook follows

1. **We asked.** The How We Grow survey (20 questions, 521 responses) and the managers' talent mapping (KRAs, rating, potential, 9-box).
2. **We built the ladder.** There was no level framework before this. Each role in each department now has a ladder from A1 Intern to the top of the role.
3. **We placed everyone.** Each person has a job family, sub-family, role and level from the evidence.
4. **We checked the market.** Every weight is set against published job-evaluation methods, and every role's journey is compared with competitors, regtech and SaaS leaders.
5. **Leaders give their view.** Any change to a weightage or a rating re-scores people live, and their level and band move with it.

## Sheets

| # | Sheet | What it does |
|---|---|---|
| 1 | How it works | The five steps, the scales and the model in one line |
| 2 | Ladder | 12 levels plus the two Player-Coach steps; each level has its own experience range, scope and span |
| 3 | Role Ladder & Weightage | **The base of the model.** 193 department-role ladders, 150 held today and 43 not yet held, with 2,196 level rows, 1,949 of them vacant. Each level row has IRIS base, Market, Leader, In use and Share of the decision, plus a written justification. |
| 4 | Role Definitions | What rating 1, 3 and 5 mean for every role at every level |
| 5 | Role Journey | Two rows per role, IRIS and Market benchmark, across A1 to E2: titles, experience, Radford/Mercer, SFIA and the market experience range, plus the weightage IRIS sets against the market's |
| 6 | Market Benchmark | The published factor shares (Korn Ferry Hay, Mercer IPE, five public plans) turned into the five steps, with the share each step gives back |
| 7 | Talent Mapping | Every person, top down: who they are → everything they told us in the survey → IRIS placement → market → weightage (IRIS base, Market, theirs, in use) → ratings (IRIS, Leader) → the live proposal |
| 8 | Person Card | One person on one page |
| 9 | Raw Data | Every bit of data held on each person, in 377 columns |
| — | Backend Data | The step values and lookups |
| — | References & Benchmarks | **Last sheet.** Every source with its link and every market benchmark |

## The scales

**Weightage: how much a parameter matters for the role at that level.** It is scored as importance from 0 to 100, in five 20-point steps:

| Step | Importance | Counts as |
|---|---|---|
| VL · Very low | 0–20 | 10 |
| L · Low | 20–40 | 30 |
| M · Medium | 40–60 | 50 |
| H · High | 60–80 | 70 |
| VH · Very high | 80–100 | 90 |

- Instead of a step, you can type an exact number from 0 to 100; it is used as typed.
- Each parameter's share of the decision is its importance divided by the sum of all seven.
- The midpoints stand in the ratio 1:3:5:7:9, the odd steps of Saaty's relative-importance scale. Turning importance into shares this way is direct-rating weighting (SMART · Edwards & Barron).
- Published shares become steps by importance = a factor's share as a percentage of the largest factor's share. The steps reproduce the published shares to within 1–3 points on average.
- 0 means a parameter is not weighed at that level, for example Leadership in Band A.

**Rating: how the person works against the level's definition.** Each step is half a level:

| Rating | Meaning | Move |
|---|---|---|
| 1 · Well below | works at the level below | −1 |
| 2 · Below | part of the way there | −½ |
| 3 · Meets | as the level is defined | 0 |
| 4 · Above | part of the way to the level above | +½ |
| 5 · Well above | already works at the level above | +1 |

## How a change moves people

- **Weighted rating** = Σ share × rating move.
- **IRIS's placement is the anchor.** It is the weighted rating using IRIS's ratings and IRIS's weightage as built.
- **The shift** is the new weighted rating minus the anchor. Half a level or more moves the person one level, and one and a half levels moves them two. The move is capped at two levels either way and follows the person's own ladder. The band follows the level.

**Weightage in use for a person** is the first of these that has an entry:

1. Their own weightage, typed on Talent Mapping.
2. The leader's entry on their level row in Role Ladder & Weightage.
3. The leader's entry on their role's ROLE row in that department.
4. The IRIS base.

**What re-scores people live:**
- **HR edits the IRIS base** (blue cells on sheet 3). Everyone at that department, role and level is re-scored.
- **A leader edits a ROLE row** (yellow). The whole role in that department is re-scored.
- **A leader edits a level row** (yellow). That level is re-scored, and the level row wins over the ROLE row.
- **A leader types weightage or ratings for one person** on Talent Mapping. Only that person is re-scored.
- **An override** sets the level directly.

The Leader proposal columns stay blank until someone gives an input. Before any input, everyone's live score equals the anchor, so no one moves by accident.

## Raw Data: every bit of data we hold

**As received:**
- **HR record:** the Employee Master (HRIS) of 23 Sep 2026, including the US & Singapore and consultant sheets.
- **How We Grow survey:** Q1 to Q20 one by one, the survey status, and the earlier or alternate survey round. That is 10,394 answers.
- **CARBON Mapping Master:** every column.
- **Managers' talent mapping:**
  - KRAs (self, manager and quarterly)
  - rating, potential, 9-box, promotion, succession, training and actual JD
- **CARBON extras:** reporting changes, the FY27 leadership scorecard and open hiring items.
- **The older Employee Master:** only where it differs from the current one.

**On file from every earlier version**, from the v1 workbooks through v1.12.2, v1.13 (including HR's edited files), v1.14 and v1.15, in six groups:
- placement history
- scoring
- role profile, hiring and market standard
- survey and KRA readings
- talent, competency, learning and mobility
- data quality, inputs and open decisions

**How the data is shown:**
- Where files disagree, the newest value comes first and the others follow after ‖, so nothing is lost.
- People in the files but not in the architecture are listed at the bottom.

**What is left out:**
- **Personal fields:** gender, date of birth, age, e-mail, mobile, and resignation and relieving dates.
- **Processing notes:** which file or row an answer was matched from.
- **"No input" placeholder lines:** these are kept blank.

**Checked:**
- All 85,381 non-blank source cells, other than the excluded personal fields, are on the right person's row.
- A cell-by-cell audit of every per-person sheet in every earlier workbook finds nothing missing beyond the notes left out on purpose.

## Decisions built in

- **Placement uses the fewest moves.**
  - No reportee sits at or above their manager.
  - A shared band appears only where a reporting chain is deeper than the bands allow; those 22 pairs are flagged.
  - 36 people move down and 39 move up.
- **One title per role and level.** 103 title changes, 35 of them protected.
- **Ladders run as far as the role does.**
  - IC roles run from A1 to D1.
  - Manager roles run from A1 to E2.
  - Business Heads run from C1 to D2.
  - Band E roles are one level each.
  - Consultants stay outside the ladder.
- **The market weights Experience and Education VL.** Every published method folds them into knowledge.

## Rebuild and check

```
cd build_v2
JA_V1=<folder with the six v1 IRIS_<Business>_Talent_Mapping.xlsx and the Employee Masters> \
python3 build_model.py <folder with the five source talent-mapping files>
python3 verify_model.py      # 41 checks, including live edits recalculated in LibreOffice and compared with the rule recomputed independently
```

The source files hold personal data, so they are never committed. The v1 workbooks and Employee Masters are in the private `job-arc` repository.
