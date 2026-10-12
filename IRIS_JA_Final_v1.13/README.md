# IRIS Job Architecture — HR masters, Final v1.13 (12 October 2026)

Two files, both **HR only**. Built from the v1.12.2 HR Super Master plus the five
manager talent-mapping files (CARBON, iFile, iDEAL, Database Content, Datatech).

| File | What it is |
|---|---|
| `IRIS_JA_2_HR_Super_Master_v1.13.xlsx` | The v1.12.2 Super Master with every sheet kept, plus the new columns, sections and sheet below |
| `IRIS_JA_2S_HR_Master_Simplified_v1.13.xlsx` | Five sheets, one line from method to person to role, with no people data left out |

## What v1.13 adds to the Super Master

- **Sheet 2 Talent Mapping:** three new column groups at the right.
  - *KRAs received from managers:* annual manager-set KRAs, quarterly KRAs, self-reported KRAs, Q1 tentative rating, actual JD, KRA status and source.
  - *What the employee shared in the survey:* questions answered, role mandate, deliverables and accountability, capabilities and tools, development needs, challenges, the evidence used for each rating, role themes and learning topics.
  - *Links & data quality:* Role ID, reporting manager's Emp ID and level, Emp ID where the source files differ, a data-quality line, weight sensitivity.
- **Sheet 3 Raw Data:**
  - 86 values in 25 people's rows restored from the source files. v1.12.2 had these blank or cut short; for example, text inside brackets had been deleted.
  - The CARBON Mapping Master's comments and five survey-summary columns, which were not carried before.
- **1b Person Card:** the same three sections.
- **20. Not Yet Mapped:** 20 people who appear in the source files but not in the job architecture. That is 19 to decide on, plus Shreyash Pawar, removed by HR on 9 October.
- **Start and Contents:** a "What v1.13 adds" note and the version stamp.

No placement, rating, weight or formula changed.

## The Simplified master — five sheets

1. **Framework & Model**
   - How to use the file.
   - The model the simple way.
   - The Master Framework M0–M17.
   - The weight benchmark.
   - The **live Weightage Matrix**, with the named ranges WM_*.
   - How each number is calculated and how each open item is tracked.
   - References S01–S72.
   - The decision log and glossary.
   - A map showing where every Super Master sheet lives in this file.
2. **Talent Mapping:** one row per person, 151 columns, with live weights, points, score and indicated level. Validation and **HR UPDATE** columns are in yellow.
3. **Raw Data:** every raw field, the restored values, and the not-yet-mapped people at the bottom.
4. **Person Card:** type an Emp ID, or pick Product → Name. Shows every sheet-2 column for the person, their role profile and a ratings chart.
5. **Role Profiles:** 175 catalogue roles.
   - The profile, requirements and competencies.
   - **Weightage**, live, with the holders' typical ratings and score.
   - **Ladder:** the step below, the step above and the full sub-family ladder.
   - Market equivalents, learning, research detail, holders and references.

## Numbers

| | |
|---|---|
| People | 586 |
| Scored people | 540 · Simplified scores = Super Master scores for all 540 |
| KRAs on file | 456 people, of whom 443 have manager-set annual KRAs. Missing: Corporate 46, iFile 32, CARBON 29, Datatech 19, iDEAL 4 |
| Per-person survey summary | 205 CARBON people |
| Survey answered, full text not in the files | 316 people · needs the How We Grow export |
| No survey response | 65 people |
| Emp ID differs between files | 27 people · Datatech, and interns who now hold IRIS IDs |
| Not yet mapped | 19 + 1 removed by HR |

## Rules kept and changed

- **Changed by HR on 12 October 2026:** per-person survey content is now shown against names. Before, only counted themes were shown, and only where 3 or more role-holders had answered.
- **Kept:**
  - Personal fields (gender, date of birth, age, e-mail, mobile, resignation and relieving dates) are never written.
  - Pay stays blank.
  - No round brackets in the Simplified file: brackets are written as " · ", and list numbers such as "1)" become "1.".

## Rebuild and check

The source talent-mapping files are not in this repository because they hold personal data.

```
cd build_v2
python3 build_super.py <folder with the five source files>
python3 build_simplified.py <folder with the five source files>
python3 verify_v113.py        # 23 checks; exit code 0 = all pass
```

Requires Python 3 with openpyxl, and LibreOffice, which stores the formula results.
