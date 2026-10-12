"""HR Master — Simplified, Final v1.13.

Five sheets, read in order, with nothing from the HR Super Master's people data left out:
  1. Framework & Model — the method end to end, the live Weightage Matrix, how to
     calculate and how to track, references; and where every Super Master sheet lives
  2. Talent Mapping   — one row per person, every column: placement, live score,
     reasons, KRAs, survey, role profile, talent, competencies, learning, mobility,
     recruitment, market, inputs, decisions, validation, HR updates
  3. Raw Data         — every raw field as received, values restored, not-yet-mapped people
  4. Person Card      — everything on one person, found by Emp ID or Product → Name
  5. Role Profiles    — one row per catalogue role: profile, requirements, weightage
     (live), ladder below and above, market, references

Usage: python3 build_simplified.py <folder with the five talent-mapping source files>
"""
import collections
import os
import re
import statistics
import sys

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402
import ja_person as P  # noqa: E402
import ja_style as Y  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_super import LINK_COLS, person_values  # noqa: E402

OUTFILE = os.path.join(J.OUT, "IRIS_JA_2S_HR_Master_Simplified_v1.13.xlsx")
FW, TMN, RAWN, PCN, RPN = "1. Framework & Model", "2. Talent Mapping", "3. Raw Data", "4. Person Card", "5. Role Profiles"
TMQ, RPQ = f"'{TMN}'", f"'{RPN}'"
LEVELS = '{"A1","A2","A3","B1","B2","B3","C1","C2","D1","D2","E1","E2"}'

CITE = {
    "scale": "S01, S02 Aon Radford and Pave level codes · S08 SFIA levels of responsibility",
    "factors": "S05 Korn Ferry Hay · S04 Mercer IPE · S07 WTW GGS · S08 SFIA",
    "weights": "S46–S51 Korn Ferry Hay evaluations · S04, S52 Mercer IPE · S53–S57 NHS, UK NJC/PSC, GLPC, SJC and US OPM point-factor plans",
    "exp": "S05, S49 Hay know-how includes experience however acquired · S04, S52 Mercer IPE knowledge · S07 WTW GGS · S58 McKinsey · S59 Deloitte · S60 WEF · S61 EY India 2026 · the 10% / 5% level is an IRIS leadership decision",
    "bi": "S05 Hay accountability, nature of impact · S04 Mercer IPE impact · S07 WTW GGS · S53–S57 public plans' financial or consequences factor",
    "lead": "S05, S46, S47 Hay managerial know-how by level · S04 Mercer IPE communication · S09 McKinsey spans of control · the 0 / 5 / 10 / 15 / 20% steps are an IRIS design rule",
    "tilt": "S06 Hay short profile · the ±5 size is an IRIS design rule",
    "rules": "S03 Mercer career levels · S07 WTW GGS · S09 McKinsey spans · S44 Google Tech Lead Manager · S45 SaaS player-coach leaders · IRIS rules R1–R13",
    "titles": "S63 iPleaders, Companies Act 2013 section 2, clause 34 · S64 Kaanoon · D1 options: S72 PayScale India, S33 Infosys · S65 Wipro, S66 Salesforce, S67 LeadDev, S32",
    "comp": "S08 SFIA 9 · S12 WEF Future of Jobs 2025 · S13 PMI · S14 SHL Great Eight · S15 Korn Ferry Leadership Architect · S16 Google Project Oxygen · S17–S25 professional standards",
    "data": "S40 Employee Master HRIS · S41 How We Grow survey · S42 manager talent-mapping files · S43 IRIS talent-mapping workbooks",
    "family": "S26 ILO ISCO-08 / NCO-2015 · S27 Deloitte skills-based organisation",
    "pay": "S11 WorldatWork · S01 Radford · ranges are IRIS working assumptions until a pay survey is bought",
}


# ================================================================ helpers
def shift_formula(f, drow):
    """Move every same-sheet cell reference in a formula down by drow rows (absolute or relative)."""
    if not isinstance(f, str) or not f.startswith("="):
        return f
    out = "="
    for t in Tokenizer(f).items:
        v = t.value
        if t.type == "OPERAND" and t.subtype == "RANGE" and "!" not in v:
            v = re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)\b",
                       lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}{int(m.group(4)) + drow}", v)
        out += v
    return out


def put(ws, r, c, v, tint=None, first=False, wrap=True):
    cell = ws.cell(r, c, None if v == "" else v)
    Y.body(cell, tint, first, wrap)
    if P.MISSING_SURVEY in str(v) or str(v).startswith((P.MISSING_KRA, "No survey response")):
        cell.fill, cell.font = Y.F_RED, Y.T_RED
    return cell


def write_headers(ws, groups, header_row=6):
    """groups: [(title, dark, [(header, width)])] → returns list of (col, header) and group spans."""
    c = 1
    cols = []
    for title, dark, heads in groups:
        g = ws.cell(header_row - 1, c, title)
        for j, (h, w) in enumerate(heads):
            Y.header(ws.cell(header_row - 1, c + j), dark, 10 if j == 0 else 9)
            Y.header(ws.cell(header_row, c + j, h), dark)
            ws.column_dimensions[L(c + j)].width = w
            cols.append((c + j, h))
        ws.cell(header_row - 1, c).value = title
        c += len(heads)
    ws.row_dimensions[header_row].height = 54
    ws.row_dimensions[header_row - 1].height = 30
    return cols


# ================================================================ sheet 1
FW_SECTIONS = [
    ("0", "How this exercise works — the journey from a base framework to today", "Start — How It Works", 5, CITE["data"]),
    ("1", "The model the simple way — the 1–12 scale, six factors, weights, points, score → level, worked examples", "1c. Scoring Made Simple", 4, CITE["scale"] + " · " + CITE["factors"] + " · " + CITE["weights"]),
    ("2", "The Master Framework M0–M17 — principles, architecture, families, levels, tracks, titles, factors, weights, rules, competencies, crosswalk, pay principles, governance, sources", "F14. Master Framework", 4, "Each element carries its own source IDs in the last column · " + CITE["family"] + " · " + CITE["titles"]),
    ("3", "Where the default weights come from — the market benchmark and its derivation", "1e. Weight Benchmark", 4, CITE["weights"] + " · " + CITE["exp"] + " · " + CITE["bi"]),
    ("4", "Weightage Matrix — LIVE: change a yellow weight and sheets 2, 4 and 5 recalculate", "1d. Weightage Matrix", 5, CITE["weights"] + " · " + CITE["tilt"] + " · " + CITE["lead"]),
    ("5", "How to calculate and how to track", None, None, CITE["rules"] + " · " + CITE["data"]),
    ("5b", "The end-to-end workflow and the review loop", "FB. End-to-End Workflow", 4, CITE["data"]),
    ("5c", "Governance, data rules, validation and pay principles", "F12. Governance & Change", 4, CITE["pay"]),
    ("6", "References & benchmarks — evidence grades, peer universe, professional standards, department × level benchmarks, source register S01–S72", "11. References & Benchmarks", 4, "Every source with its link and evidence grade"),
    ("7", "Decision log — every decision, its version and its effect", "FF. Decision Log", 4, "Owner decisions, October 2026"),
    ("7b", "Glossary", "F13. Glossary", 4, ""),
    ("8", "Where everything in the HR Super Master lives in this file", None, None, ""),
]


def framework_sheet(wb, src_wb, M, stats):
    ws = wb.create_sheet(FW)
    Y.banner(ws, "HR Master — Simplified · Framework & Model",
             "Sheet 1 of 5 · The whole method on one page, top to bottom: how the job architecture works, how every number is "
             "calculated, how to track open items, and the references for each part. Then sheet 2 Talent Mapping, 3 Raw Data, "
             "4 Person Card and 5 Role Profiles.", J.STAMP, back=False, width_cols=14)
    r = 5
    # how to use
    intro = [
        ("This file", "HR Master — Simplified: five sheets that hold every person, every number and every reason from the HR Super Master v1.13, laid out in one line from the method to the person to the role. HR edits it; the other files are rebuilt from it."),
        ("Sheet 2 · Talent Mapping", f"One row per person · {stats['people']} people · every column grouped left to right: who → role before and after → the live score → how the level was set → the reasons for each rating → KRAs received from the manager → what the employee shared in the survey → role profile and business impact → talent, competencies, learning, mobility → recruitment → market → inputs, decisions and data quality → pay, blank → validation → HR update."),
        ("Sheet 3 · Raw Data", f"Every raw field as received from HRIS, the survey count, the manager talent-mapping files and the CARBON Mapping Master · values the earlier clean-up lost are restored from the source files · {stats['unmapped']} people found in the source files but not yet mapped are added at the bottom, marked."),
        ("Sheet 4 · Person Card", "Type or paste an Emp ID, or pick Product, then Name: every column of sheet 2 for that person, with their role profile from sheet 5 and a ratings chart."),
        ("Sheet 5 · Role Profiles", f"One row per role in the IRIS role catalogue · {stats['roles']} roles · what the role is for, does, decides and needs; its weightage, live from the matrix in section 4; the step below and above on the ladder; market equivalents; references."),
        ("Live numbers", "Weights, points, the score and 'Level the score indicates' are Excel formulas. Change a yellow weight in section 4 or a Business Impact rating on sheet 2 and they recalculate. FINAL LEVEL is a decision, not a formula: it moves through the rules R1–R13 or an approved leadership decision."),
        ("How to update", "Write changes in the yellow HR UPDATE columns at the right of sheet 2 · field, new value, reason, source, who, date · or in the validation columns. Each change is read back on the next rebuild, so every number stays traceable."),
        ("Privacy", "HR only. Survey content is shown against names by HR's decision of 12 October 2026. Gender, date of birth, age, e-mail, mobile, resignation and relieving dates are never written. Pay columns stay blank until a pay survey is bought."),
    ]
    c = ws.cell(r, 1, "HOW TO USE THIS FILE")
    Y.header(c, Y.NAVY, 11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    r += 1
    for k, v in intro:
        a = ws.cell(r, 1, k)
        a.font = Y.font(9, True)
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b = ws.cell(r, 2, v)
        b.font = Y.font(9)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
        ws.row_dimensions[r].height = Y.est_height([(v, 150)])
        r += 1
    r += 1
    c = ws.cell(r, 1, "CONTENTS OF THIS SHEET — click a section")
    Y.header(c, Y.NAVY, 11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    r += 1
    toc_rows = {}
    for sid, title, *_ in FW_SECTIONS:
        ws.cell(r, 1, sid).font = Y.font(9, True)
        t = ws.cell(r, 2, title)
        t.font = Y.font(9, False, "0563C1", underline="single")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
        toc_rows[sid] = r
        r += 1
    r += 1

    widths = collections.defaultdict(float)
    for col, w in ((1, 18), (2, 22)):
        widths[col] = w
    hidden_cols = set()
    named = {}
    for sid, title, src_name, start, cite in FW_SECTIONS:
        sec_row = r
        c = ws.cell(r, 1, f"{sid} · {title}")
        Y.header(c, "2E75B6", 12)
        for k in range(2, 15):
            ws.cell(r, k).fill = PatternFill("solid", fgColor="2E75B6")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        ws.row_dimensions[r].height = 22
        r += 1
        if cite:
            c = ws.cell(r, 1, "References: " + cite)
            c.font = Y.font(8, True, "375623", italic=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
            ws.row_dimensions[r].height = Y.est_height([(c.value, 200)])
            r += 1
        toc = ws.cell(toc_rows[sid], 2)
        toc.hyperlink = f"#'{FW}'!A{sec_row}"
        if src_name:
            src = src_wb[src_name]
            drow = r - start
            maxc = src.max_column
            for sr in range(start, src.max_row + 1):
                for sc in range(1, maxc + 1):
                    s = src.cell(sr, sc)
                    if s.value is None and not s.has_style:
                        continue
                    v = s.value
                    if isinstance(v, str) and v.strip() == "◂ Contents":
                        continue
                    d = ws.cell(sr + drow, sc, shift_formula(v, drow))
                    Y.copy_style(s, d)
                    if s.hyperlink is not None and s.hyperlink.target and not str(s.hyperlink.target).startswith("#"):
                        d.hyperlink = s.hyperlink.target
                h = src.row_dimensions[sr].height
                if h:
                    ws.row_dimensions[sr + drow].height = h
            for mr in src.merged_cells.ranges:
                if mr.min_row >= start:
                    ws.merge_cells(start_row=mr.min_row + drow, start_column=mr.min_col,
                                   end_row=mr.max_row + drow, end_column=mr.max_col)
            for sc in range(1, maxc + 1):
                cd = src.column_dimensions[L(sc)]
                if cd.hidden:
                    hidden_cols.add(sc)
                elif cd.width:
                    widths[sc] = max(widths[sc], min(60, cd.width))
            if src_name == "1d. Weightage Matrix":
                named = {"WM_EXP_AB": f"$B${6 + drow}", "WM_EXP_CE": f"$B${7 + drow}"}
                for nm, colL in (("WM_KEY", "A"), ("WM_C", "N"), ("WM_AU", "O"), ("WM_AC", "P"), ("WM_L", "Q"), ("WM_E", "R"), ("WM_BI", "S")):
                    named[nm] = f"${colL}${17 + drow}:${colL}${123 + drow}"
                stats["matrix_rows"] = (17 + drow, 123 + drow)
            r = src.max_row + drow + 2
        elif sid == "5":
            r = calc_and_track(ws, r, M, stats)
        elif sid == "8":
            r = where_everything(ws, r, src_wb)
    fixed = {1: 30, 2: 22, 3: 40, 4: 34, 5: 16, 6: 18, 7: 12, 8: 34, 20: 30}
    for col in range(1, 27):
        ws.column_dimensions[L(col)].width = fixed.get(col, 11)
    fit_heights(ws, 5)
    for col in hidden_cols:
        ws.column_dimensions[L(col)].hidden = True
    for nm, ref in named.items():
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{FW}'!{ref}")
    ws.freeze_panes = "A5"
    ws.sheet_view.zoomScale = 90
    return ws


def fit_heights(ws, first_row):
    """Row heights for wrapped text at this sheet's column widths, merged cells counted at their full width."""
    width = {c: (ws.column_dimensions[L(c)].width or 9) for c in range(1, ws.max_column + 1)}
    span = {}
    for mr in ws.merged_cells.ranges:
        span[(mr.min_row, mr.min_col)] = sum(width.get(c, 9) for c in range(mr.min_col, mr.max_col + 1))
    for row in ws.iter_rows(min_row=first_row):
        items = []
        for c in row:
            if c.value is None or (isinstance(c.value, str) and c.value.startswith("=")):
                continue
            items.append((c.value, span.get((c.row, c.column), width.get(c.column, 9))))
        if items:
            size = max((c.font.sz or 9) for c in row if c.value is not None)
            ws.row_dimensions[row[0].row].height = Y.est_height(items, base=max(12.5, size * 1.4))


def fit_rows(ws, first, last, cap=300):
    width = {c: (ws.column_dimensions[L(c)].width or 9) for c in range(1, ws.max_column + 1)}
    for row in ws.iter_rows(min_row=first, max_row=last):
        items = [(c.value, width.get(c.column, 9)) for c in row
                 if c.value is not None and not (isinstance(c.value, str) and c.value.startswith("="))]
        ws.row_dimensions[row[0].row].height = Y.est_height(items, maximum=cap, minimum=30)


def col_lengths(ws, first, last):
    """90th-percentile text length of each column — sizes the person card's rows."""
    out = {}
    for col in ws.iter_cols(min_row=first, max_row=last):
        ls = sorted(len(str(c.value)) for c in col if c.value is not None and not str(c.value).startswith("="))
        out[col[0].column] = ls[int(len(ls) * 0.9)] if ls else 0
    return out


def table(ws, r, heads, rows, dark="1F3864", widths=None):
    n = len(heads)
    for j, h in enumerate(heads, 1):
        Y.header(ws.cell(r, j, h), dark)
    for k in range(n + 1, 15):
        ws.cell(r, k).fill = PatternFill("solid", fgColor=dark)
    ws.merge_cells(start_row=r, start_column=n, end_row=r, end_column=14)
    r += 1
    for row in rows:
        for j, v in enumerate(row, 1):
            put(ws, r, j, v, "FFFFFF", j == 1)
        ws.merge_cells(start_row=r, start_column=n, end_row=r, end_column=14)
        ws.row_dimensions[r].height = Y.est_height([(v, 90 if k == n - 1 else 24) for k, v in enumerate(row)])
        r += 1
    return r + 1


def calc_and_track(ws, r, M, stats):
    def sub(title):
        nonlocal r
        c = ws.cell(r, 1, title)
        c.font = Y.font(11, True, Y.NAVY)
        r += 1
    sub("5a · How every number is calculated — follow one person across sheet 2")
    calc = [
        ("1", "Rate the six factors", "Complexity & Expertise, Autonomy & Decision, Accountability & Impact, Leadership & Strategy, Experience & Qualification, Business Impact — each on the same 1–12 scale, where the number is the ladder position: 4 = B1-level, 8 = C2-level. Half points allowed. Start from the designation and reporting line; move toward the survey and KRA evidence by up to 1 level if specific, ½ if brief, 0 if generic.", "Sheet 2 · the six 'rating' columns · reasons in JUSTIFICATION", CITE["factors"]),
        ("2", "Experience & Qualification rating", "Experience scale: 0 yrs = 1, 1 = 2, 3 = 3, 5.5 = 4, 7.5 = 5, 10 = 6, 12 = 7, 14.5 = 8, 17.5 = 9, 21 = 10, 24 = 11, 28 = 12, in proportion between points; then +1.0 relevant professional qualification · CA, CMA, CS, LLB · +0.5 relevant postgraduate, −0.5 below graduate.", "Sheet 2 · Exp + Qual rating and its justification", CITE["exp"]),
        ("3", "Business Impact rating", "Starts at the Accountability rating; +0.5 each if the survey or manager mapping marks the role business-critical or unique, at most +1. Managers and heads overwrite it · yellow.", "Sheet 2 · Bus. Impact rating · yellow", CITE["bi"]),
        ("4", "Look up the weights", "Each person's weight row is department × role group · band and track. Excel finds that row in the matrix by its key and reads the factor's weight from the Used column: INDEX on WM_C, matched on WM_KEY — and the same for WM_AU, WM_AC, WM_L, WM_E and WM_BI. The weights live in section 4 of this sheet.", "Sheet 2 · 'weight' columns · section 4", CITE["weights"] + " · " + CITE["tilt"]),
        ("5", "Points", "Points = rating × weight, for each factor · blank when the rating is blank.", "Sheet 2 · 'points' columns", ""),
        ("6", "Weighted score", "Score = the sum of the six points; weights add to 100% · blank unless all six points are present, as for people placed by a category rule.", "Sheet 2 · WEIGHTED SCORE", ""),
        ("7", "Level the score indicates", "The score rounds half up to the nearest whole level: 1 = A1, 2 = A2, 3 = A3, 4 = B1, 5 = B2, 6 = B3, 7 = C1, 8 = C2, 9 = D1, 10 = D2, 11 = E1, 12 = E2.", "Sheet 2 · Level the score indicates", CITE["scale"]),
        ("8", "Rules R1–R13 → FINAL LEVEL", "The fair-play rules can hold a level back or keep the structure in order: category, provisional, band weights, role substance, structural gate, top of band, Band A experience, experience limit, track minimum, Player-Coach step, leadership decision, reporting-line cap, reports to the top. Every rule applied is written with its numbers. FINAL LEVEL is a value, changed only by evidence or an approved leadership decision.", "Sheet 2 · HOW THE LEVEL WAS SET · section 2 M8", CITE["rules"]),
        ("9", "Flags", "Close call: score within ±0.05 of a rounding point where the score decided. Near a gate: the deciding number within 0.25 of flipping. Player-Coach candidate or watch list. Weight sensitivity: how many of 30 weight tests move the level.", "Sheet 2 · Flags, Player-Coach / calibration flag, Weight sensitivity", ""),
        ("10", "Role weightage and typical score", "Each role's main weight row is looked up live; its holders' average ratings × those weights give a typical score and the level it indicates.", "Sheet 5 · WEIGHTAGE", CITE["weights"]),
    ]
    r = table(ws, r, ["Step", "What", "How it is worked out", "Where", "References"], calc)
    sub("5b · How to track — every open item, where it lives, who closes it")
    t3 = collections.Counter(J.text(x[6]) for x in M["t3"] if x[0])
    t2 = collections.Counter(J.text(x[7]) for x in M["t2"] if x[0])
    track = [
        ("Critical inputs · red", "Sheet 2 · ADDITIONAL INPUTS NEEDED and Data requests · red cells say what and who", f"{sum(t3.values())} open requests", "The person named in the cell", "Supply the value; HR enters it in HR UPDATE with reason, source and date; rebuild"),
        ("Decisions", "Sheet 2 · Open decisions · one line per decision with type, priority and owner", f"{sum(t2.values())} decisions", "Business head, panel or HR as named", "Record in HR decision / Approved level / Approved designation; rule R11 applies the approved level"),
        ("Manager validation", "Sheet 2 · VALIDATION · yellow", "Agree / Change / Discuss", "Manager → HR", "A change needs evidence and a rule ID · V2"),
        ("KRAs received", "Sheet 2 · KRA status · red when not received", f"{stats['kra']} of {stats['people']} received", "Manager", "Send the KRAs; HR pastes them in HR UPDATE · field 'KRAs set by manager'"),
        ("Survey text", "Sheet 2 · Survey status", f"{stats['survey']} per-person summaries · the rest need the How We Grow export", "HR", "Load the survey export; the five 'as shared' columns fill on the next build"),
        ("Data quality", "Sheet 2 · Data quality · one line per person", "Survey · KRAs · experience · provisional · Emp ID differences", "HR", "Fix the record in HRIS or HR UPDATE"),
        ("Business impact", "Sheet 2 · Business impact — from manager / head · yellow; Bus. Impact rating · yellow", "Blank until the manager writes", "Manager / head", "Write the view; change the rating if needed — the score updates at once"),
        ("Weights", "Section 4 · yellow weights and 'Why changed'", "Defaults from the market benchmark", "Department leaders", "Change in steps of 5%; HR imports; CEO approves changes to weights"),
        ("Not yet mapped", f"Sheet 3 · bottom rows marked NOT YET MAPPED · {stats['unmapped']} people", "Listed, not placed", "HR", "Decide: add and place, confirm left, or out of scope"),
        ("Emp ID differences", "Sheet 2 · Emp ID in the source files · where different", f"{stats['id_diff']} people", "HR", "Settle one Emp ID per person in HRIS"),
        ("HR updates", "Sheet 2 · HR UPDATE · yellow, at the far right", "Field · New value · Reason · Source · Updated by · Date", "HR", "Read back on the next rebuild into the model's inputs; every change keeps its reason"),
    ]
    r = table(ws, r, ["What to track", "Where in this file", "Today", "Who", "How it closes"], track)
    sub("5c · Critical inputs open today · from the data-request tracker")
    r = table(ws, r, ["What is missing", "Open requests"], sorted(t3.items(), key=lambda x: -x[1]))
    sub("5d · Decisions open today · by type")
    r = table(ws, r, ["Decision type", "Open"], sorted(t2.items(), key=lambda x: -x[1]))
    sub("5e · Validation, in order · V1–V6")
    val = [("1", "Manager validation", "Sheet 2 VALIDATION: agree, change or discuss each placement. A change needs evidence and a rule ID."),
           ("2", "HR inputs", "Red cells: experience corrections, provisional placements R2, missing HRIS fields, reporting lines."),
           ("3", "Calibration panel", "Player-Coach cases and candidates, close calls, moves down, title checks, long-tenure A3, senior titles without teams, all C+ placements."),
           ("4", "Open inputs", "Chase each owner until the red cells are cleared."),
           ("5", "Approval", "HR decides; the CEO approves any change to levels, weights or rules. Moves down are protected for one cycle."),
           ("6", "Record", "Every change keeps its reason, source and date; the version number rises when the model changes.")]
    r = table(ws, r, ["#", "Step", "What happens"], val)
    return r


def where_everything(ws, r, src_wb):
    plan = {
        "Start — How It Works": "Sheet 1 · section 0", "Contents": "Sheet 1 · contents of this sheet",
        "0. Dashboard": "Counts: sheet 2 filters · charts stay in the Super Master", "0b. Org Health & Insights": "Sheet 2 · MOBILITY & STRUCTURE · Org health & A3 notes",
        "1. Overview": "Sheet 1 · sections 0–2 · counts by filtering sheet 2", "1a. Placements": "Sheet 2 · every 1a column, incl. What set the level, Flags, Confidence",
        "1b. Person Card": "Sheet 4", "1c. Scoring Made Simple": "Sheet 1 · section 1", "1d. Weightage Matrix": "Sheet 1 · section 4 · live",
        "1e. Weight Benchmark": "Sheet 1 · section 3", "2. Talent Mapping": "Sheet 2 · every column", "3. Raw Data & Market Standards": "Sheet 3 · every column · market standards also on sheet 2 MARKET",
        "4. Role Competency Framework": "Sheet 5 · REQUIREMENTS and LEARNING", "5. Capability Framework": "Sheet 5 · competencies by role · department view stays in the Super Master",
        "6. Role Catalogue": "Sheet 5 · every catalogue column", "7. Role Requirements": "Sheet 5 · REQUIREMENTS, WEIGHTAGE and MARKET",
        "8. Talent, Learning & Mobility": "Sheet 2 · TALENT, COMPETENCY, LEARNING, MOBILITY", "9. Calibration & Decisions": "Sheet 2 · Open decisions · section 5",
        "10. Framework & Logic": "Sheet 1 · section 2 Master Framework", "11. References & Benchmarks": "Sheet 1 · section 6",
        "12. Org Health": "Sheet 2 · Org health & A3 notes", "13. Weight Sensitivity": "Sheet 2 · Weight sensitivity", "14. A3 Deep-Dive": "Sheet 2 · Org health & A3 notes",
        "15. Succession": "Sheet 2 · TALENT · Succession", "15a. Critical Roles at Risk": "Sheet 2 · TALENT · critical-role risk, cover, action",
        "16. Role Definitions": "Sheet 5 · Designations used at IRIS", "17. Data Gaps": "Sheet 2 · Data gaps", "18. Career Ladder": "Sheet 5 · LADDER",
        "18a. Families & Designations": "Sheet 5", "18b. Family Summary": "Sheet 5 · filter by family", "18c. Ladder — How to Read": "Sheet 5 · LADDER",
        "19. Competency Library": "Sheet 5 · competencies by role · the full library stays in the Super Master", "20. Not Yet Mapped": "Sheet 3 · bottom rows",
        "T1. Tracker Dashboard": "Sheet 1 · sections 5b–5d", "T2. Decisions": "Sheet 2 · Open decisions", "T3. Data Requests": "Sheet 2 · Data requests",
        "T4. Feedback Log": "Sheet 2 · VALIDATION and HR UPDATE", "T5. Question Review": "Sheet 2 · Interview questions",
    }
    rows = []
    for name in src_wb.sheetnames:
        if src_wb[name].sheet_state != "visible":
            continue
        where = plan.get(name)
        if where is None:
            if name.startswith("L"):
                where = "Counts by product and level: filter sheet 2 · the leadership summaries stay in the Super Master"
            elif name.startswith(("F", "FA", "FB")):
                where = "Sheet 1 · section 2 Master Framework covers it · the guide text stays in the Super Master"
            elif name.startswith("CP"):
                where = "Calibration panel pack: stays in the Super Master and its own file"
            else:
                where = "Super Master"
        rows.append((name, where))
    return table(ws, r, ["HR Super Master sheet", "Where it is in this file"], rows)


# ================================================================ sheet 2
def tm_spec(M):
    """Column groups of sheet 2. Each column: (header, width, source). source: ('tm'|'raw'|'s8', header) or ('x', key) or ('f', kind)."""
    T = lambda h: ("tm", h)
    R = lambda h: ("raw", h)
    S8 = lambda h: ("s8", h)
    F6 = [("Complexity", "WM_C"), ("Autonomy", "WM_AU"), ("Accountab.", "WM_AC"), ("Leadership", "WM_L"), ("Exp + Qual", "WM_E"), ("Bus. Impact", "WM_BI")]
    score = [("Weight row — Weightage Matrix", 26, T("Weight row — Weightage Matrix"))]
    for f, nm in F6:
        score += [(f"{f} rating", 8, T(f"{f} rating")), (f"{f} weight", 8, ("f", "w:" + nm)), (f"{f} points", 8, ("f", "p:" + f))]
    score += [("WEIGHTED SCORE = Σ points", 10, ("f", "score")), ("Level the score indicates", 9, ("f", "ind")),
              ("FINAL LEVEL · after the rules", 9, T("FINAL LEVEL"))]
    return [
        ("PERSON & KEYS", "1F3864", [
            ("Business", 10, T("Business")), ("Emp ID", 12, T("Emp ID")), ("Employee name", 22, T("Employee name")),
            ("Department", 18, T("Department")), ("Role ID · catalogue", 10, ("x", "role")),
            ("Reporting manager", 20, T("Reporting manager")), ("Reporting manager · Emp ID", 12, ("x", "mgr_id")),
            ("Reporting manager · final level", 9, ("x", "mgr_lv")), ("Direct reports", 8, S8("Direct reports")),
            ("Total exp. · yrs", 8, T("Total exp. · yrs")), ("Exp. at IRIS · yrs", 8, T("Exp. at IRIS · yrs")),
            ("Qualification", 18, T("Qualification")), ("Emp ID in the source files · where different", 14, ("x", "src_id"))]),
        ("ROLE — BEFORE → AFTER", "2E75B6", [
            ("Current designation", 24, T("Current designation")), ("Final designation", 26, T("Final designation")),
            ("Job family", 20, T("Job family")), ("Sub-family · what they do", 24, T("Sub-family · what they do")),
            ("Current level", 8, T("Current level")), ("FINAL LEVEL", 8, T("FINAL LEVEL")), ("Level title", 16, T("Level title")),
            ("Band", 14, T("Band")), ("Track", 12, T("Track")), ("Change vs current", 14, T("Change vs current")),
            ("Radford / Mercer code", 10, R("Radford / Mercer code")), ("Pave level", 12, R("Pave level")), ("SFIA 9", 12, R("SFIA 9"))]),
        ("SCORE — rating × weight = points · weights live from sheet 1 section 4", "806000", score),
        ("HOW THE LEVEL WAS SET", "375623", [
            ("What set the level", 30, ("x", "set")), ("Why these weights", 40, T("Why these weights")),
            ("Calculation", 40, T("Calculation")),
            ("Gates, limits & Player-Coach transition steps applied — with the numbers", 50, T("Gates, limits & Player-Coach transition steps applied — with the numbers")),
            ("Flags", 24, ("x", "flags")), ("Player-Coach / calibration flag", 26, T("Player-Coach / calibration flag")),
            ("Confidence", 14, T("Confidence")), ("Points to confirm", 30, T("Points to confirm")), ("Weight sensitivity", 24, ("x", "sens"))]),
        ("JUSTIFICATION FOR EACH RATING — start point, evidence, market standard", "843C0C", [
            (h, 42, T(h)) for h in ("Complexity & Expertise — justification", "Autonomy & Decision — justification",
                                    "Accountability & Impact — justification", "Leadership & Strategy — justification",
                                    "Experience & Qualification — justification", "Business Impact — justification")]),
        ("KRAs RECEIVED FROM MANAGERS — as submitted", "375623", [(h, w, ("kra", i)) for i, (h, w) in enumerate(P.KRA_COLS)]),
        ("WHAT THE EMPLOYEE SHARED IN THE SURVEY — per person", "7030A0", [(h, w, ("sv", i)) for i, (h, w) in enumerate(P.SURVEY_COLS)]),
        ("ROLE PROFILE & BUSINESS IMPACT", "548235", [
            ("Role profile — what they do, responsibilities & competencies required", 60, T("Role profile — what they do, responsibilities & competencies required")),
            ("Business impact — from survey and data", 45, T("Business impact — from survey and data")),
            ("Business impact — from manager / head · to complete", 30, T("Business impact — from manager / head · to complete")),
            ("Bus. Impact — model start", 9, T("Bus. Impact — model start"))]),
        ("TALENT · manager mapping, succession, critical roles", "C55A11", [
            (h, w, S8(h)) for h, w in (("Latest rating", 9), ("Potential", 9), ("9-box", 18), ("Readiness", 14), ("Top performer", 9),
                                        ("Critical role · Y/N", 9), ("Successor named", 14))] + [
            ("Succession · C–E roles", 40, ("x", "succ")), ("Critical-role risk", 30, ("x", "risk")),
            ("Possible cover · same sub-family", 34, ("x", "cover")), ("Risk action", 34, ("x", "action"))]),
        ("COMPETENCY RESULT", "806000", [
            (h, w, S8(h)) for h, w in (("Competency index %", 9), ("Fit", 12), ("Core %", 8), ("Future-ready %", 8), ("Functional %", 8),
                                        ("Leadership %", 8), ("Top gaps", 30), ("Readiness for next level", 24),
                                        ("Ratings by competency · rating / required", 45))]),
        ("LEARNING NEEDS", "2E75B6", [
            (h, w, S8(h)) for h, w in (("Self-identified — topics · survey; not quoted", 30), ("Manager-identified", 30),
                                        ("Role-based priorities", 30), ("To move to the next role — topics · survey; not quoted", 30),
                                        ("Training for next role", 30))]),
        ("MOBILITY, OVERLAP & STRUCTURE", "595959", [
            (h, w, S8(h)) for h, w in (("Lateral move options", 30), ("Duplication — same business", 30),
                                        ("Alternate talent pool — other businesses", 30), ("Span status", 14), ("Span action", 30))] + [
            ("Org health & A3 notes", 36, ("x", "org"))]),
        ("RECRUITMENT — REPLACEMENT PROFILE", "C55A11", [
            (h, w, T(h)) for h, w in (("Replacement — what they do · simple", 40), ("Hire equivalent · external market", 40),
                                       ("Competencies & skill level to hire for", 40), ("Interview questions for the perfect fit", 50))]),
        ("MARKET STANDARD", "3A3838", [
            ("Market standard for this role and level", 40, T("Market standard for this role and level"))] + [
            (h, w, R(h)) for h, w in (("IRIS experience guide", 14), ("Market experience · Accenture · Infosys", 18),
                                       ("Typical market title based on competitor/market benchmark", 24),
                                       ("Competitors, regtech & SaaS leaders", 36), ("Professional standard", 24),
                                       ("Required capability · Core · Future · Functional · Leadership", 30))]),
        ("INPUTS, DECISIONS & DATA QUALITY", "C00000", [
            ("ADDITIONAL INPUTS NEEDED — what, and who provides it", 36, T("ADDITIONAL INPUTS NEEDED — what, and who provides it")),
            ("Data used for this person", 30, T("Data used for this person")), ("Open decisions", 50, ("x", "t2")),
            ("Data requests", 36, ("x", "t3")), ("Data gaps", 30, ("x", "gaps")), ("Data quality", 40, ("x", "dq"))]),
        ("PAY — set later from the pay survey · left blank", "595959", [
            (h, 12, T(h)) for h in ("Current fixed pay · annual CTC", "Proposed range — minimum", "Proposed range — midpoint",
                                    "Proposed range — maximum", "Compa-ratio")]),
        ("VALIDATION — manager & HR", "1F3864", [
            (h, w, T(h)) for h, w in (("Manager validation", 14), ("Manager comment / proposed change", 30), ("HR decision", 16),
                                       ("Approved level", 9), ("Approved designation", 22), ("Pay / title protection · Y/N", 10),
                                       ("Reviewer", 14), ("Decision date", 11))]),
        ("HR UPDATE — write a change here; it is read back on the next rebuild", "BF8F00", [
            ("Field to update", 22, ("y", "")), ("New value", 22, ("y", "")), ("Reason", 30, ("y", "")),
            ("Source", 18, ("y", "")), ("Updated by", 14, ("y", "")), ("Date", 11, ("y", ""))]),
    ]


TINTS = {"1F3864": "EEF2F9", "2E75B6": "EAF3FB", "806000": "F8F4E6", "375623": "EAF2E6", "843C0C": "FBEEE6", "7030A0": "F4EEF9",
         "548235": "EEF6EA", "C55A11": "FCF1EA", "595959": "F2F2F2", "3A3838": "F2F2F2", "C00000": "FDECEC", "BF8F00": "FFF2CC"}


def talent_sheet(wb, src_wb, M, F, people, extras):
    ws = wb.create_sheet(TMN)
    Y.banner(ws, "Talent Mapping — every person, every column, in one line",
             "Sheet 2 of 5 · One row per person. Read left to right: who → role before and after → the live score → how the level was set "
             "→ the reasons → KRAs from the manager → what the employee shared in the survey → role profile and business impact → talent, "
             "competencies, learning, mobility → recruitment → market → inputs, decisions and data quality → pay, blank → validation → HR update. "
             "Filter any column; the weights are live from sheet 1 section 4.", J.STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    spec = tm_spec(M)
    cols = write_headers(ws, [(t, d, [(h, w) for h, w, _ in c]) for t, d, c in spec])
    flat = [(t, d, h, w, s) for t, d, c in spec for h, w, s in c]
    col_of = {}
    for i, (_, _, h, _, _) in enumerate(flat, 1):
        col_of.setdefault(h, L(i))
    ws2 = src_wb["2. Talent Mapping"]
    ws3 = src_wb["3. Raw Data & Market Standards"]
    ws8 = src_wb["8. Talent, Learning & Mobility"]
    th = {h: i + 1 for i, h in enumerate(M["tm_hdr"])}
    rh = {h: i + 1 for i, h in enumerate(M["raw_hdr"])}
    s8h = {h: i + 1 for i, h in enumerate(M["s8_hdr"])}
    s8row = {}
    for i, r8 in enumerate(M["s8"]):
        s8row[(J.key(r8[1]), J.nname(r8[2]))] = 7 + i
    first_data = 7
    for i, p in enumerate(people):
        r = first_data + i
        sr = 7 + p["i"]
        e = extras[p["i"]]
        r8 = s8row.get((J.key(p["emp"]), J.nname(p["name"])))
        prev_group = None
        for ci, (gt, dark, h, w, src) in enumerate(flat, 1):
            first = gt != prev_group
            prev_group = gt
            tint = TINTS.get(dark, "FFFFFF")
            kind, k = src
            v, src_cell = None, None
            if kind == "tm":
                src_cell = ws2.cell(sr, th[k])
                v = src_cell.value
            elif kind == "raw":
                src_cell = ws3.cell(sr, rh[k])
                v = src_cell.value
            elif kind == "s8":
                if r8:
                    src_cell = ws8.cell(r8, s8h[k])
                    v = src_cell.value
            elif kind == "kra":
                v = e["kra"][k]
            elif kind == "sv":
                v = e["sv"][k]
            elif kind == "x":
                v = e[k]
            elif kind == "f":
                v = formula(k, r, col_of)
            elif kind == "y":
                v = None
            if isinstance(v, str) and not v.startswith("="):
                v = J.debracket(v)
            cell = put(ws, r, ci, v, tint, first)
            if src_cell is not None and src_cell.has_style:
                fill = src_cell.fill.fgColor.rgb if src_cell.fill and src_cell.fill.fill_type == "solid" else None
                fill = fill if isinstance(fill, str) else None
                if fill and fill[-6:] in ("FFC7CE", "F8CBAD", "C6EFCE", "FFEB9C", "DDEBF7", "F2F2F2", "C6E0B4", "A9D08E", "FFF2CC",
                                          "E2EFDA", "D0E8E5", "E4DFEC", "EDE7DD", "EDEDED"):
                    cell.fill = PatternFill("solid", fgColor=fill[-6:])
                    cell.font = Y.font(9, bool(src_cell.font and src_cell.font.b),
                                       (src_cell.font.color.rgb[-6:] if src_cell.font and src_cell.font.color and isinstance(src_cell.font.color.rgb, str) else "000000"))
            if kind == "y" or h.startswith("Bus. Impact rating") or h == "Business impact — from manager / head · to complete":
                cell.fill = Y.F_YEL
            if gt.startswith("PAY"):
                cell.fill = Y.F_GREY
            if gt.startswith("VALIDATION"):
                cell.fill = Y.F_YEL
            if h in ("FINAL LEVEL", "FINAL LEVEL · after the rules"):
                cell.font = Y.font(10, True)
                cell.alignment = Alignment(horizontal="center", vertical="top")
            if kind == "f" or h.endswith((" rating", " weight", " points")):
                cell.number_format = "0.00" if not h.endswith("rating") else "0.0"
                if h.endswith(" weight"):
                    cell.number_format = "0%"
    last = first_data + len(people) - 1
    fit_rows(ws, first_data, last, cap=300)
    ws.freeze_panes = "D7"
    ws.auto_filter.ref = f"A6:{L(len(flat))}{last}"
    ws.sheet_view.zoomScale = 90
    return col_of, last


def formula(kind, r, col_of):
    if kind.startswith("w:"):
        nm = kind[2:]
        return f'=IFERROR(INDEX({nm},MATCH(${col_of["Weight row — Weightage Matrix"]}{r},WM_KEY,0)),"")'
    if kind.startswith("p:"):
        f = kind[2:]
        a, b = col_of[f"{f} rating"], col_of[f"{f} weight"]
        return f'=IF(OR({a}{r}="",{b}{r}=""),"",{a}{r}*{b}{r})'
    pts = [col_of[f"{f} points"] for f in ("Complexity", "Autonomy", "Accountab.", "Leadership", "Exp + Qual", "Bus. Impact")]
    if kind == "score":
        refs = ",".join(f"{c}{r}" for c in pts)
        return f'=IF(COUNT({refs})<6,"",SUM({refs}))'
    if kind == "ind":
        s = col_of["WEIGHTED SCORE = Σ points"]
        return f'=IF({s}{r}="","",INDEX({LEVELS},MAX(1,MIN(12,ROUND({s}{r},0)))))'
    raise ValueError(kind)


# ================================================================ sheet 3
def raw_sheet(wb, src_wb, M, people, unmapped):
    ws = wb.create_sheet(RAWN)
    Y.banner(ws, "Raw Data — every field as received",
             "Sheet 3 of 5 · One row per person: HRIS, the survey count, the managers' talent-mapping files and the CARBON Mapping Master, "
             "with the source file in brackets-free labels · [source] field. Values the v1.12.2 clean-up lost are restored from the source files "
             "and listed per person. People found in the source files but not yet in the job architecture are at the bottom, marked "
             "NOT YET MAPPED. Personal fields are never carried.", J.STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    hdr = M["raw_hdr"]
    grp = M["raw_grp"]
    groups = []
    cur = None
    for j, h in enumerate(hdr):
        if grp[j]:
            cur = [grp[j], ["1F3864", "806000", "548235"][min(2, len(groups))], []]
            groups.append(cur)
        cur[2].append((h, 18 if j > 21 else 14))
    groups[0][2].insert(0, ("Status", 14))
    groups.append(["ADDED IN v1.13 — CARBON Mapping Master columns not carried before · values restored · Emp ID differences", "7030A0",
                   [(h, 45) for _, _, h in P.EXTRA_RAW] + [(P.RAW_LOG, 36), (P.RAW_IDS, 14)]])
    write_headers(ws, [(t, d, h) for t, d, h in groups])
    ws3 = src_wb["3. Raw Data & Market Standards"]
    r = 7
    for p in people:
        row, extra, restored = P.restore_raw(M, p)
        put(ws, r, 1, "Mapped", "EEF2F9", True)
        for j, v in enumerate(row):
            c = put(ws, r, j + 2, v if not isinstance(v, str) else J.debracket(v), "FFFFFF" if j > 21 else "EEF2F9", grp[j] != "")
            s = ws3.cell(7 + p["i"], j + 1)
            if s.fill is not None and s.fill.fill_type == "solid" and s.fill.fgColor.rgb and s.fill.fgColor.rgb[-6:] in ("FFC7CE", "F8CBAD", "C6EFCE", "FFEB9C", "DDEBF7"):
                c.fill = PatternFill("solid", fgColor=s.fill.fgColor.rgb[-6:])
            if restored and any(hdr[j].endswith(x) for x in restored) and v != p["raw"][j]:
                c.fill = PatternFill("solid", fgColor="E2EFDA")
        for k, v in enumerate(extra):
            put(ws, r, len(hdr) + 2 + k, v, "F4EEF9", k == 0)
        r += 1
    # not yet mapped: slot each source field into the matching raw column
    keyed = {}
    for u in unmapped:
        k = J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"])
        keyed.setdefault(k, []).append(u)
    for k, us in keyed.items():
        vals = {}
        for u in us:
            s, rec = u["source"], u["rec"]
            pre = J.PREFIX[s["kind"]]
            for h, v in rec["row"].items():
                if not J.has(v):
                    continue
                for j, mh in enumerate(hdr):
                    if mh.startswith(pre) and J.hkey(mh) == J.hkey(h):
                        vals.setdefault(j, J.clean(v))
            for (kind, sh, _), kk in zip(P.EXTRA_RAW, range(len(P.EXTRA_RAW))):
                if kind == s["kind"]:
                    v = rec["row"].get(sh)
                    if J.has(v):
                        vals.setdefault(("x", kk), J.clean(v))
        u0 = us[0]
        nm = u0["rec"]["name"]
        removed = J.nname(nm) == "SHREYASH PAWAR"
        put(ws, r, 1, "Removed by HR · 9 Oct 2026" if removed else "NOT YET MAPPED", "FFC7CE", True)
        base = {0: {"MM": "CARBON", "CRAW": "CARBON"}.get(u0["source"]["kind"], u0["source"]["label"].split("· ")[-1]),
                1: u0["rec"]["id"] or "—", 2: nm,
                11: "Not in the job architecture · HR to decide: add and place, confirm left, or out of scope"}
        def g(*hs):
            for u in us:
                for h, v in u["rec"]["row"].items():
                    if J.hkey(h) in [J.hkey(x) for x in hs] and J.has(v):
                        return J.clean(v)
            return None
        base[3] = g("Department")
        base[4] = g("Designation", "Current Designation")
        for j in range(len(hdr)):
            v = base.get(j, vals.get(j))
            c = put(ws, r, j + 2, v, "FDECEC", grp[j] != "")
        for kk in range(len(P.EXTRA_RAW)):
            put(ws, r, len(hdr) + 2 + kk, vals.get(("x", kk)), "FDECEC", kk == 0)
        put(ws, r, len(hdr) + 2 + len(P.EXTRA_RAW), "Source: " + " + ".join(sorted({f"{u['source']['file']} · {u['source']['sheet']}" for u in us})), "FDECEC")
        r += 1
    fit_rows(ws, 7, r - 1, cap=250)
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"A6:{L(len(hdr) + 3 + len(P.EXTRA_RAW))}{r - 1}"
    ws.sheet_view.zoomScale = 90
    return len(keyed)


# ================================================================ sheet 5
def role_sheet(wb, src_wb, M, F, people, stats):
    ws = wb.create_sheet(RPN)
    Y.banner(ws, "Role Profiles — every role in the IRIS role catalogue",
             "Sheet 5 of 5 · One row per role · job family × sub-family × level × track, pooled across IRIS. What the role is for, does, decides "
             "and needs; its weightage, live from sheet 1 section 4, with the holders' typical ratings and score; the step below and above on "
             "the ladder; requirements, market equivalents, learning and the research detail; and the references behind it.", J.STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    ch = M["cat_hdr"]
    H = F.H
    holders = collections.defaultdict(list)
    for p in people:
        c = F.cat_of[p["i"]]
        if c:
            holders[c[0]].append(p)
    req_by = collections.defaultdict(list)
    for r_ in M["req"]:
        req_by[(J.text(r_[0]), J.text(r_[1]), J.text(r_[4]), J.text(r_[5]))].append(r_)
    comp_by = collections.defaultdict(list)
    for r_ in M["comp"]:
        comp_by[(J.text(r_[3]), J.text(r_[1]), J.text(r_[4]))].append(r_)
    rq, cp = M["req_hdr"], M["comp_hdr"]

    def agg(rows, hdr, h):
        if h not in hdr:
            return ""
        j = hdr.index(h)
        vals, seen = [], set()
        for x in rows:
            v = J.text(x[j])
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        return "\n".join(vals) if len(vals) <= 1 else "\n".join("• " + v for v in vals)

    order = {lv: i for i, lv in enumerate(["X1", "X2", "X3", "A1", "A2", "A3", "B1", "B2", "B3", "B+", "C1", "C2", "C+", "D1", "D2", "E1", "E2"])}
    lane = collections.defaultdict(list)
    for c in M["cat"]:
        lane[(J.text(c[1]), J.text(c[2]), J.text(c[5]))].append(c)
    for k in lane:
        lane[k].sort(key=lambda c: order.get(J.text(c[3]), 99))
    fam_all = collections.defaultdict(list)
    for c in M["cat"]:
        fam_all[(J.text(c[1]), J.text(c[2]))].append(c)

    cat_cols = [(h, 30 if j > 12 else 14) for j, h in enumerate(ch) if h]
    groups = [
        ("ROLE", "1F3864", [(h, w) for h, w in cat_cols[:13]]),
        ("LADDER — where the role sits", "2E75B6", [cat_cols[37], ("Step below · same sub-family and track", 34), ("Step above · same sub-family and track", 34),
                                                  ("The ladder in this sub-family · ★ = this role", 50), ("Other track at this level", 30)]),
        ("WEIGHTAGE — live from sheet 1 section 4 · the role's main weight row", "806000",
         [("Main weight row · most holders", 30), ("Other weight rows used by holders", 34), ("Role group default weights · C / Au / Ac / L / E / BI", 22),
          ("Complexity weight", 8), ("Autonomy weight", 8), ("Accountability weight", 8), ("Leadership weight", 8), ("Experience weight", 8), ("Business Impact weight", 8),
          ("Typical Complexity rating", 8), ("Typical Autonomy rating", 8), ("Typical Accountability rating", 8), ("Typical Leadership rating", 8),
          ("Typical Experience rating", 8), ("Typical Business Impact rating", 8), ("Typical score · live", 9), ("Level the typical score indicates", 9)]),
        ("WHAT THE ROLE IS — catalogue · role-holders' counted themes", "548235", [(h, w) for h, w in cat_cols[13:29]]),
        ("REQUIREMENTS — competencies, gates, assessment", "843C0C", [(h, w) for h, w in cat_cols[29:33]] + [
            ("Department technical competencies", 30), ("Gate requirements", 30), ("How to assess", 30), ("Mandatory & gate requirements", 30),
            ("Role-holders today · avg index · meet level", 24), ("Most common gaps", 30)]),
        ("MARKET", "3A3838", [(h, w) for h, w in cat_cols[33:37]] + [
            ("Pave level", 12), ("Typical market title based on competitor/market benchmark", 24), ("Product companies · Google · MS · Workiva", 26),
            ("Indian IT services · Infosys · TCS · Accenture", 26), ("Consulting & XBRL services · Big Four · DataTracks", 26),
            ("Competitors, regtech & SaaS leaders — this family", 30), ("Hiring route", 24), ("Future skills", 30)]),
        ("LEARNING & VOICE OF THE ROLE", "C55A11", [(h, w) for h, w in cat_cols[38:44]] + [("Learning priorities for this role", 30)]),
        ("RESEARCH DETAIL — Career Architecture V2", "595959", [(h, w) for h, w in cat_cols[44:50]]),
        ("EVIDENCE & REFERENCES", "7030A0", [(h, w) for h, w in cat_cols[50:]] + [("Holders · name · Emp ID · business", 40), ("References", 50)]),
    ]
    flat = [(g, d, h) for g, d, hs in groups for h, _ in hs]
    write_headers(ws, groups)
    col_of = {h: L(i) for i, (_, _, h) in enumerate(flat, 1)}
    cat_idx = {h: j for j, h in enumerate(ch)}
    r = 7
    wm_rows = {}
    for c in M["cat"]:
        cid = c[0]
        hs = holders[cid]
        fam, sub, lv, tr = J.text(c[1]), J.text(c[2]), J.text(c[3]), J.text(c[5])
        ln = lane[(fam, sub, tr)]
        pos = [J.text(x[0]) for x in ln].index(cid)
        fmt = lambda x: f"{J.text(x[0])} · {J.text(x[3])} · {J.text(x[6]).split(chr(10))[0][:60]}"
        below = fmt(ln[pos - 1]) if pos > 0 else "— entry step in this sub-family"
        above = fmt(ln[pos + 1]) if pos + 1 < len(ln) else "— top step in this sub-family"
        ladder = " → ".join(("★ " if x is c else "") + f"{J.text(x[3])} {J.text(x[6]).split(chr(10))[0][:40]}" for x in ln)
        other = [x for x in fam_all[(fam, sub)] if J.text(x[3]) == lv and x is not c]
        wrows = collections.Counter(J.text(p["tm"][H["Weight row — Weightage Matrix"]]) for p in hs if J.has(p["tm"][H["Weight row — Weightage Matrix"]]))
        main = wrows.most_common(1)[0][0] if wrows else ""
        others = "\n".join(f"{k} · {n}" for k, n in wrows.most_common()[1:])
        ratings = []
        for f in ("Complexity rating", "Autonomy rating", "Accountab. rating", "Leadership rating", "Exp + Qual rating", "Bus. Impact rating"):
            xs = []
            for p in hs:
                try:
                    xs.append(float(p["tm"][H[f]]))
                except (TypeError, ValueError):
                    pass
            ratings.append(round(statistics.mean(xs), 2) if xs else "")
        reqs = [x for p in hs for x in req_by.get((p["biz"], J.text(p["tm"][H["Department"]]), J.text(p["tm"][H["Final designation"]]), lv), [])]
        comps = [x for p in hs for x in comp_by.get((J.text(p["tm"][H["Final designation"]]), lv, sub), [])]
        dflt = ""
        if reqs:
            dflt = " / ".join(J.text(reqs[0][rq.index(h)]) for h in ("Complexity %", "Autonomy %", "Accountability %", "Leadership %", "Exp. & Qual. %", "Bus. Impact %"))
        peer_ids = sorted(set(re.findall(r"\bS\d{2}\b|\bC\d{2}\b", " ".join(J.text(c[j]) for j in range(len(c))))))
        refs = ("Level & codes: " + CITE["scale"] + " · factors: " + CITE["factors"] + " · weights: " + CITE["weights"] +
                " · competencies: " + CITE["comp"] + " · family: " + CITE["family"] + " · evidence: " + CITE["data"] +
                (" · cited in this row: " + ", ".join(peer_ids) if peer_ids else ""))
        vals = {}
        for h, _ in cat_cols:
            vals[h] = c[cat_idx[h]]
        vals.update({
            "Step below · same sub-family and track": below, "Step above · same sub-family and track": above,
            "The ladder in this sub-family · ★ = this role": ladder,
            "Other track at this level": "\n".join(f"{J.text(x[0])} · {J.text(x[5])} · {J.text(x[6]).split(chr(10))[0][:50]}" for x in other) or "—",
            "Main weight row · most holders": main, "Other weight rows used by holders": others or "—",
            "Role group default weights · C / Au / Ac / L / E / BI": dflt,
            "Department technical competencies": agg(reqs, rq, "Department technical competencies"),
            "Gate requirements": agg(reqs, rq, "Gate requirements"), "How to assess": agg(comps, cp, "How to assess"),
            "Mandatory & gate requirements": agg(comps, cp, "Mandatory & gate requirements"),
            "Role-holders today · avg index · meet level": agg(comps, cp, "Role-holders today · avg index · meet level"),
            "Most common gaps": agg(comps, cp, "Most common gaps"), "Learning priorities for this role": agg(comps, cp, "Learning priorities for this role"),
            "Pave level": agg(reqs, rq, "Pave level"),
            "Typical market title based on competitor/market benchmark": agg(reqs, rq, "Typical market title based on competitor/market benchmark"),
            "Product companies · Google · MS · Workiva": agg(reqs, rq, "Product companies · Google · MS · Workiva"),
            "Indian IT services · Infosys · TCS · Accenture": agg(reqs, rq, "Indian IT services · Infosys · TCS · Accenture"),
            "Consulting & XBRL services · Big Four · DataTracks": agg(reqs, rq, "Consulting & XBRL services · Big Four · DataTracks"),
            "Competitors, regtech & SaaS leaders — this family": agg(reqs, rq, next((h for h in rq if h.startswith("Competitors, regtech")), "")),
            "Hiring route": agg(reqs, rq, "Hiring route"), "Future skills": agg(reqs, rq, "Future skills"),
            "Holders · name · Emp ID · business": "\n".join(f"{p['name']} · {p['emp']} · {p['biz']}" for p in hs),
            "References": refs,
        })
        for k, f in enumerate(("Complexity", "Autonomy", "Accountability", "Leadership", "Experience", "Business Impact")):
            vals[f"Typical {f} rating"] = ratings[k]
        wm = col_of["Main weight row · most holders"]
        for f, nm in (("Complexity", "WM_C"), ("Autonomy", "WM_AU"), ("Accountability", "WM_AC"), ("Leadership", "WM_L"),
                      ("Experience", "WM_E"), ("Business Impact", "WM_BI")):
            vals[f"{f} weight"] = f'=IFERROR(INDEX({nm},MATCH(${wm}{r},WM_KEY,0)),"")'
        pairs = [(col_of[f"Typical {f} rating"], col_of[f"{f} weight"]) for f in ("Complexity", "Autonomy", "Accountability", "Leadership", "Experience", "Business Impact")]
        prod = ",".join(f"{a}{r}*{b}{r}" for a, b in pairs)
        cond = ",".join(f"{a}{r},{b}{r}" for a, b in pairs)
        vals["Typical score · live"] = f'=IFERROR(IF(COUNT({cond})<12,"",SUM({prod})),"")'
        sc = col_of["Typical score · live"]
        vals["Level the typical score indicates"] = f'=IF({sc}{r}="","",INDEX({LEVELS},MAX(1,MIN(12,ROUND({sc}{r},0)))))'
        prev = None
        for ci, (g, d, h) in enumerate(flat, 1):
            v = vals.get(h, "")
            if isinstance(v, str) and not v.startswith("="):
                v = J.debracket(J.text(v))
            cell = put(ws, r, ci, v, TINTS.get(d, "FFFFFF"), g != prev)
            prev = g
            if h.endswith(" weight"):
                cell.number_format = "0%"
            if h == "Level" or h == "Level the typical score indicates":
                b = Y.band_of(J.text(v) if h == "Level" else "")
                if b:
                    cell.fill = PatternFill("solid", fgColor=Y.BAND[b])
        r += 1
    fit_rows(ws, 7, r - 1, cap=350)
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"A6:{L(len(flat))}{r - 1}"
    ws.sheet_view.zoomScale = 90
    stats["roles"] = len(M["cat"])
    return col_of


# ================================================================ sheet 4
def person_card(wb, M, people, tm_cols, tm_last, rp_cols, tm_len):
    ws = wb.create_sheet(PCN)
    lists = wb.create_sheet("Lists")
    lists.sheet_state = "hidden"
    Y.banner(ws, "Person Card — one person, everything, on one page",
             "Sheet 4 of 5 · Type or paste an Emp ID in the yellow box, any case — or pick a Product, then a Name. The card shows every column of "
             "sheet 2 for that person, grouped as on sheet 2, with their role profile from sheet 5 and a chart of the six ratings.", J.STAMP, width_cols=7)
    ws.column_dimensions["A"].width = 38
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 22
    ws.column_dimensions["H"].width = 4
    ws.column_dimensions["J"].width = 10
    prods = sorted({p["biz"] for p in people}, key=lambda b: ["CARBON", "iFile", "iDEAL", "Datatech", "Consume", "Corporate"].index(b) if b in ["CARBON", "iFile", "iDEAL", "Datatech", "Consume", "Corporate"] else 9)
    for i, b in enumerate(prods, 1):
        lists.cell(i, 1, b)
    wb.defined_names["PC_Products"] = DefinedName("PC_Products", attr_text=f"Lists!$A$1:$A${len(prods)}")
    r = 1
    for b in prods:
        start = r
        for i, p in enumerate(people):
            if p["biz"] != b:
                continue
            label = f"{p['name']} · {J.text(p['tm'][M['tm_hdr'].index('Department')])} · {p['emp']}"
            lists.cell(r, 4, label)
            lists.cell(r, 5, i + 1)
            r += 1
        wb.defined_names[f"L_{b}"] = DefinedName(f"L_{b}", attr_text=f"Lists!$D${start}:$D${r - 1}")
    n_lab = r - 1
    labels = [("Product", "B5"), ("Name", "B6"), ("Emp ID — paste or type · wins over the name", "B7")]
    first = next(p for p in people if p["biz"] == prods[0])
    defaults = {"B5": prods[0], "B6": f"{first['name']} · {J.text(first['tm'][M['tm_hdr'].index('Department')])} · {first['emp']}", "B7": None}
    for lab, ref in labels:
        rr = int(ref[1:])
        a = ws.cell(rr, 1, lab)
        a.font = Y.font(10, True)
        c = ws[ref]
        c.value = defaults[ref]
        c.fill = Y.F_YEL
        c.font = Y.font(10, True)
        c.border = Y.Border(left=Y.THIN, right=Y.THIN, top=Y.THIN, bottom=Y.THIN)
        ws.merge_cells(f"{ref}:F{rr}")
    dv1 = DataValidation(type="list", formula1="PC_Products", allow_blank=True)
    dv2 = DataValidation(type="list", formula1='INDIRECT("L_"&$B$5)', allow_blank=True)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add("B5")
    dv2.add("B6")
    ws["J5"] = "Row found"
    ws["J6"] = (f'=IFERROR(IF(TRIM($B$7)<>"",MATCH(TRIM($B$7),{TMQ}!$B$7:$B${tm_last},0),'
                f'INDEX(Lists!$E$1:$E${n_lab},MATCH($B$6,Lists!$D$1:$D${n_lab},0))),"")')
    ws["A8"] = "Showing"
    ws["A8"].font = Y.font(10, True)
    ws["B8"] = (f'=IF($J$6="","No match — check the Emp ID, or clear it and pick a name",'
                f'INDEX({TMQ}!$C$7:$C${tm_last},$J$6)&" · "&INDEX({TMQ}!$B$7:$B${tm_last},$J$6)&" · "&INDEX({TMQ}!$A$7:$A${tm_last},$J$6))')
    ws["B8"].font = Y.font(11, True, Y.NAVY)
    ws.merge_cells("B8:F8")
    for c in ("J5", "J6"):
        ws[c].font = Y.font(8, False, "808080")
    r = 10
    rating_rows = []
    group = None
    long_cols = {"Role profile — what they do, responsibilities & competencies required", "KRAs set by manager · annual", "KRAs · quarterly",
                 "KRAs · self-reported annual", "Interview questions for the perfect fit"}
    for (gtitle, dark, h, col) in tm_cols:
        if gtitle != group:
            group = gtitle
            c = ws.cell(r, 1, gtitle)
            Y.header(c, dark, 10)
            for k in range(2, 8):
                ws.cell(r, k).fill = PatternFill("solid", fgColor=dark)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            r += 1
        a = ws.cell(r, 1, h)
        a.font = Y.font(9, True)
        a.fill = PatternFill("solid", fgColor="F2F2F2")
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b = ws.cell(r, 2, f'=IFERROR(IF(INDEX({TMQ}!${col}$7:${col}${tm_last},$J$6)="","—",INDEX({TMQ}!${col}$7:${col}${tm_last},$J$6)),"")')
        b.font = Y.font(9)
        b.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        if h.endswith(" weight"):
            b.number_format = "0%"
        elif h.endswith((" points", "Σ points")):
            b.number_format = "0.00"
        if h.endswith(" rating") and gtitle.startswith("SCORE"):
            rating_rows.append(r)
        n = tm_len.get(openpyxl.utils.column_index_from_string(col), 0)
        ws.row_dimensions[r].height = Y.est_height([("x" * n, 128), (h, 36)], maximum=260, minimum=16)
        r += 1
    # role profile from sheet 5
    c = ws.cell(r, 1, "ROLE PROFILE — from sheet 5, by Role ID")
    Y.header(c, "548235", 10)
    for k in range(2, 8):
        ws.cell(r, k).fill = PatternFill("solid", fgColor="548235")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    rid_col = next(col for g, d, h, col in tm_cols if h == "Role ID · catalogue")
    rp_last = 6 + 175
    for h in ("Purpose", "Key responsibilities", "Decision rights", "Step below · same sub-family and track", "Step above · same sub-family and track",
              "The ladder in this sub-family · ★ = this role", "Main weight row · most holders", "Typical score · live",
              "★ Must-have competencies", "Next step", "References"):
        if h not in rp_cols:
            continue
        col = rp_cols[h]
        a = ws.cell(r, 1, h)
        a.font = Y.font(9, True)
        a.fill = PatternFill("solid", fgColor="F2F2F2")
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b = ws.cell(r, 2, f'=IFERROR(INDEX({RPQ}!${col}$7:${col}${rp_last},MATCH(INDEX({TMQ}!${rid_col}$7:${rid_col}${tm_last},$J$6),{RPQ}!$A$7:$A${rp_last},0)),"")')
        b.font = Y.font(9)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        if h == "Typical score · live":
            b.number_format = "0.00"
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 90 if h not in ("Typical score · live", "Main weight row · most holders") else 18
        r += 1
    # ratings chart, beside the card
    if rating_rows:
        k0 = 2
        for k, rr in enumerate(rating_rows):
            ws.cell(k0 + k, 26, ["Complexity", "Autonomy", "Accountability", "Leadership", "Exp & Qual", "Bus. Impact"][k])
            ws.cell(k0 + k, 27, f"=IFERROR(VALUE(B{rr}),0)")
            ws.cell(k0 + k, 26).font = Y.font(8, False, "FFFFFF")
            ws.cell(k0 + k, 27).font = Y.font(8, False, "FFFFFF")
        ch = BarChart()
        ch.type = "bar"
        ch.title = "The six ratings · 1–12"
        ch.y_axis.scaling.min = 0
        ch.y_axis.scaling.max = 12
        ch.legend = None
        ch.y_axis.majorUnit = 2
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.add_data(Reference(ws, min_col=27, min_row=k0, max_row=k0 + 5), titles_from_data=False)
        ch.set_categories(Reference(ws, min_col=26, min_row=k0, max_row=k0 + 5))
        ch.height, ch.width = 8, 16
        ws.add_chart(ch, "I9")
    ws.freeze_panes = "A9"
    ws.print_area = f"A1:G{r - 1}"
    return ws


# ================================================================ main
def main(srcdir):
    os.makedirs(J.OUT, exist_ok=True)
    M, S, people, unmapped = J.build(srcdir)
    F = P.Folds(M, people)
    src_wb = openpyxl.load_workbook(os.path.join(J.OUT, "IRIS_JA_2_HR_Super_Master_v1.13.xlsx"))
    extras = {}
    stats = {"people": len(people), "kra": 0, "survey": 0, "id_diff": 0, "roles": len(M["cat"])}
    for p in people:
        k, sv, links = person_values(M, F, p)
        succ = F.succession(p)
        extras[p["i"]] = {
            "kra": k, "sv": sv, "role": links[0], "mgr_id": links[1], "mgr_lv": links[2], "src_id": links[3], "dq": links[4],
            "sens": links[5], "set": F.placement(p)[0], "flags": F.placement(p)[1], "org": F.org(p),
            "succ": succ[0], "risk": succ[1], "cover": succ[2], "action": succ[3],
            "t2": F.decisions(p), "t3": F.requests(p), "gaps": F.gaps(p)}
        stats["kra"] += k[5].startswith("Received")
        stats["survey"] += sv[-1].startswith("Per-person")
        stats["id_diff"] += bool(p["src_ids"])
    stats["unmapped"] = len({J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"]) for u in unmapped})

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    framework_sheet(wb, src_wb, M, stats)
    col_of, tm_last = talent_sheet(wb, src_wb, M, F, people, extras)
    raw_sheet(wb, src_wb, M, people, unmapped)
    rp_cols = role_sheet(wb, src_wb, M, F, people, stats)
    tm_cols = []
    for gt, dark, cols in tm_spec(M):
        for h, w, s in cols:
            tm_cols.append((gt, dark, h, col_of[h]))
    person_card(wb, M, people, tm_cols, tm_last, rp_cols, col_lengths(wb[TMN], 7, tm_last))
    # order: 1, 2, 3, 4, 5, Lists
    order = [FW, TMN, RAWN, PCN, RPN, "Lists"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTFILE)
    work = os.path.join(os.environ.get("JA_WORK", "/tmp"), "ja_simpl")
    os.makedirs(work, exist_ok=True)
    n_f, errors = X.cache_values(OUTFILE, work)
    print(f"Wrote {OUTFILE}: people {stats['people']} · roles {stats['roles']} · KRAs {stats['kra']} · CARBON survey {stats['survey']} · "
          f"not yet mapped {stats['unmapped']} · formulas {n_f} · formula errors {len(errors)}")
    if errors:
        print(errors[:20])


if __name__ == "__main__":
    main(sys.argv[1])
