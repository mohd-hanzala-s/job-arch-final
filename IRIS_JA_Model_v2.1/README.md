# IRIS Job Architecture Model v2.1 (12 October 2026)

One workbook for HR and business leaders: `IRIS_JA_Model_v2.1.xlsx`.

## What is new in v2.1

- **Role Ladder & Weightage (sheet 3).** Every department has every role, and every role has every level. Levels run from A1 Intern to the top of the role, including roles and levels nobody holds yet.
  - Each level row shows the weightage IRIS → Market → Leader for the seven parameters, with a written justification. The justification gives the published share for each method, the Experience / Education split, the type-of-work tilt and the references.
  - A ROLE row re-weights every level of that role in that department. A level row's own entry wins over the ROLE row.
  - Talent Mapping reads each person's weightage from their own department × role × level row, so every edit flows through live.
- **Inputs received in the survey** (was "What they shared") now holds every answer for each person, question by question (Q1 to Q20). It then adds the CARBON role mapping, the earlier survey round, how the rater read the answers by parameter, and the learning they asked for. It is blank if nothing was received.
- **Raw Data (sheet 8) now holds every bit of data we have on each person** — 160 columns, joined from every source, with values exactly as received:
  - **HR record** from the Employee Master (HRIS) of 23 Sep 2026, including the US & Singapore and consultant sheets.
  - **The How We Grow survey, question by question:** all 20 questions, the survey status and how the answers came in, plus the earlier / alternate survey round. That is 10,394 answers from 521 people.
  - **The CARBON Mapping Master**, every column.
  - **The managers' talent mapping:** KRAs (self, manager, quarterly), ratings, potential, 9-box, promotion, succession, training needs, actual JD, projects.
  - **CARBON extras:** proposed reporting changes, the FY27 leadership scorecard and open hiring items, set against the leaders they belong to.
  - **The older Employee Master**, only where it differs.
  - Where two files disagree, the newest comes first and the others follow after ‖, so nothing is dropped.
  - People in the files but not in the architecture are listed at the bottom.
  - Personal fields (gender, date of birth, age, e-mail, mobile, resignation and relieving dates) are left out.
  - A check confirms that all 85,381 non-blank, non-personal cells in the source files appear on the right person's row.
- **Role Definitions (sheet 4)** says what Low / Medium / High mean for every role at every level, including vacant ones.
- Weightage and the market level now follow each person's own level and department, not the role's most common level. This changes the IRIS Business Impact weight for 9 people and the market level for 7. Placement and titles are unchanged.

## The story the workbook follows

1. **We asked.** The How We Grow survey, plus the managers' talent mapping (KRAs, rating, potential, 9-box).
2. **We built the ladder.** There was no level framework before this. The ladder has 12 levels in five bands, plus two Player-Coach steps.
3. **We placed everyone.** Each person gets a job family, sub-family, role and level from the evidence.
4. **We checked the market.** Every weight is set against published job-evaluation methods.
5. **Leaders give their view.** Their proposed level and band appear once they give an input.

## Sheets

| # | Sheet | What it does |
|---|---|---|
| 1 | How it works | The five steps and the model in one line |
| 2 | Ladder | Each level has a non-overlapping experience range, plus its scope, span, education norm, market codes and headcount |
| 3 | Role Ladder & Weightage | 193 department-role ladders: 150 held today plus 43 roles nobody holds yet, under IRIS · Not yet held. Together they have 2,196 level rows, 1,949 of them vacant. Each level row has IRIS / Market / Leader / In-use weightage and the justification. Rows are grouped, so you can collapse a role to its ROLE row. |
| 4 | Role Definitions | One block per role (sub-family × IC / Manager), held or vacant, with every level of its ladder. Each level says what Low / Medium / High mean on each parameter. |
| 5 | Market Benchmark | Low / Medium / High for each parameter and role group, each with the published numbers behind it (Korn Ferry Hay, Mercer IPE, five public plans) and the level benchmarks for experience, span, Radford and SFIA |
| 6 | Talent Mapping | Every person, top down: who they are → inputs received in the survey and KRAs → IRIS placement → market → weightage (IRIS, Market, Leader) → ratings (IRIS, Leader) → Leader proposal |
| 7 | Person Card | One person on one page; pick a name or type an Emp ID |
| 8 | Raw Data | Every bit of data held, 160 columns: HR record, all 20 survey questions and the earlier round, CARBON Mapping Master, managers' talent mapping, CARBON scorecard, older Employee Master where it differs |
| 9 | Backend Data | The lookups |

## The model in one line

Each parameter is rated Low / Medium / High against the level, and each role weights the parameters Low / Medium / High.

- If a leader's weighted rating is **half a level or more above** IRIS's, the person moves up one level.
- If it is **half a level or more below**, they move down one level.
- A leader can also set **any level** directly with the override.

There are seven parameters: Complexity & Expertise, Autonomy & Decision, Accountability & Impact, Leadership & Strategy, Business Impact, Experience and Education.

## Decisions built in

- **Blank means nothing was received.** This includes ratings and evidence that would only have repeated a starting point or a colleague's answers.
- **Placement uses the fewest moves.**
  - No reportee sits at or above their manager's level.
  - No reportee shares their manager's band, except where the reporting chain is deeper than the bands allow. Those **22 pairs** are flagged as structure decisions, not forced.
  - Band E with Band E (the CEO and CXOs) is the agreed exception.
  - Moving up counts twice as much as moving down, so the model prefers moving down.
  - **Result:** 36 people move down (pay and title protected for one cycle) and 39 move up — mostly managers of Band B staff going to C1, and leads of Band A staff to B1.
- **One title per role and level.**
  - The title before the dash is aligned to what people who stayed at that level use; each person's own function and POD Lead marker are kept.
  - People who moved down keep their title for one cycle.
  - Leadership titles change only on a move up.
  - 103 title changes are listed per person, 35 of them protected.
- **Ladders run as far as the role does.**
  - IC roles run from A1 to D1, where the IC track ends.
  - Manager roles run from A1 to E2.
  - Business Heads run from C1 to D2.
  - The two Band E roles are one level each.
  - Consultants stay outside the ladder and carry no weightage.
- **Weightage is set per department, role and level.** The level's role group (for example Band B · People Manager) sets the market weights. The department's type of work tilts IRIS's weights by ±5 in Bands A to C.
- **Leader weightage in use** is the level row if given, otherwise the ROLE row, otherwise IRIS.
- **Leader columns stay blank** until a leader enters a rating, a weightage or an override.
- **Experience and Education are weighted Low by the market.** None of the published methods weights them on their own; all fold them into knowledge.

## Not changed by this model

The v1.15 HR Super Master and Simplified files still show the v1.14 HR proposed level. This workbook's IRIS proposed level replaces it once you approve it.

## Rebuild and check

```
cd build_v2
JA_V1=<folder with the six v1 IRIS_<Business>_Talent_Mapping.xlsx and the Employee Masters> \
python3 build_model.py <folder with the five source talent-mapping files>
python3 verify_model.py      # 33 checks, including live ROLE-row, level-row and vacant-level edits recalculated in LibreOffice
```
