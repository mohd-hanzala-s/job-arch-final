# IRIS Job Architecture Model v2.0 (12 October 2026)

One workbook, rebuilt from scratch on HR's brief: `IRIS_JA_Model_v2.0.xlsx`. It is for HR and business leaders.

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
| 3 | Role Model | One block per role (sub-family × IC / Manager), with its levels listed below it. The role row holds the weightage: IRIS → Market → Leader. The level rows say what Low / Medium / High mean at each level. Filter by role. |
| 4 | Market Benchmark | Low / Medium / High for each parameter and role group, each with the published numbers behind it (Korn Ferry Hay, Mercer IPE, five public plans) and the level benchmarks for experience, span, Radford and SFIA |
| 5 | Talent Mapping | Every person, top down: who they are → what they told us → IRIS placement → market → weightage (IRIS, Market, Leader) → ratings (IRIS, Leader) → Leader proposal |
| 6 | Person Card | One person on one page; pick a name or type an Emp ID |
| 7 | Raw Data | 43 core columns: HR record, experience and education, everything shared in the survey, performance and talent, KRAs and work |
| 8 | Backend Data | The lookups |

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
- **Roles are role-based.** A leader's weightage on the Role Model applies to everyone in the role.
- **Leader columns stay blank** until a leader enters a rating, a weightage or an override.
- **Experience and Education are weighted Low by the market.** None of the published methods weights them on their own; all fold them into knowledge.

## Not changed by this model

The v1.15 HR Super Master and Simplified files still show the v1.14 HR proposed level. This workbook's IRIS proposed level replaces it once you approve it.

## Rebuild and check

```
cd build_v2
python3 build_model.py <folder with the five source talent-mapping files>
python3 verify_model.py      # 21 checks, including live leader edits recalculated in LibreOffice
```
