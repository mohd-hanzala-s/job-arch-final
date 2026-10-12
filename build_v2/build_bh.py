"""Business-Head Workbook — Final v1.14 (HR decisions of 12 October 2026).

One file for all businesses, read top down: the CEO first, then each head and their teams.

  1. How to use
  2. Talent Mapping   — core details → summary of responses, linked to the raw data → weightage
                         Low / Medium / High → assessment Low / Medium / High → market benchmark →
                         HR proposed band → input from benchmarking → final proposed band, live
  3. Raw Data         — every field for every person, one click from Talent Mapping
  4. Market Benchmark — published job-evaluation weights by role group, the level crosswalk,
                         experience, education and span norms, with references
  5. Department Ladder — every role in every department, top down, with its editable weightage
  6. RB · <family>     — one sheet per job family: what Low, Medium and High mean for each
                         parameter at each level, and the default weightage
  7. Backend Data      — the lookups that make it live

Live chain: RB family default weightage → Department Ladder weightage → Talent Mapping weightage
(each can be overwritten) → points → the net assessment → steps from the HR proposed level →
final proposed band → the manager check. Change any L/M/H anywhere and the bands move at once.

Usage: python3 build_bh.py <folder with the five talent-mapping source files>
"""
import collections
import os
import re
import statistics
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402
import ja_person as P  # noqa: E402
import ja_style as Y  # noqa: E402
import ja_v114 as V  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_super import person_values  # noqa: E402

OUTFILE = os.path.join(J.OUT, f"IRIS_JA_7_Business_Head_Workbook_{J.VER}.xlsx")
SUPER = os.path.join(J.OUT, f"IRIS_JA_2_HR_Super_Master_{J.VER}.xlsx")
TMN, RAWN, MBN, DLN, BEN = "2. Talent Mapping", "3. Raw Data", "4. Market Benchmark", "5. Department Ladder", "Backend Data"
Q = lambda n: f"'{n}'"

PARAMS = [("C", "Complexity & Expertise"), ("Au", "Autonomy & Decision"), ("Ac", "Accountability & Impact"),
          ("L", "Leadership & Strategy"), ("BI", "Business Impact"), ("Exp", "Experience"), ("Edu", "Education")]
PKEYS = [k for k, _ in PARAMS]
LMH = [("L", "Low", 1, -1, "Below what the matrix describes for this band · works at the level below on this parameter"),
       ("M", "Medium", 3, 0, "Meets what the matrix describes for this band"),
       ("H", "High", 9, 1, "Above what the matrix describes for this band · already works at the level above on this parameter")]
WEIGHT_H, WEIGHT_M = 30.0, 9.0   # weight % at or above → High; at or above → Medium; else Low
EXP_SPLIT = {"AB": (0.6, 0.4), "CE": (0.8, 0.2)}   # Experience / Education share of the Experience & Qualification weight

TILT = {  # Build Guide 5.6 · S06 Hay short profile; the ±5 size is an IRIS rule. Applied in Bands A–C and B+.
    "Technology & engineering": {"C": 5, "Ac": -5},
    "Product & programme": {"C": -5, "Au": 5},
    "Regulatory & domain consulting": {},
    "Service delivery & data operations": {"C": -5, "Ac": 5},
    "Customer support & success": {"C": -5, "BI": 5},
    "Sales & marketing": {"C": -5, "Au": 5, "Ac": -5, "BI": 5},
    "Corporate functions": {"C": -5, "Edu": 5},
    "General management": {},
}
RG_1E = {  # role group → label used in 1e. Weight Benchmark
    "A": "Band A", "B-IC": "Band B · IC / POD Lead", "B-M": "Band B · People Manager", "B+": "Player-Coach · B+",
    "C-IC": "Band C · Principal spe", "C-M": "Band C · Head / manage", "C+": "Player-Coach · C+", "DE": "Bands D–E",
}
RG_NAME = {"A": "Band A · all tracks", "B-IC": "Band B · Specialist / POD Lead", "B-M": "Band B · People Manager",
           "B+": "B+ · Player-Coach", "C-IC": "Band C · Principal specialist", "C-M": "Band C · Head / manager",
           "C+": "C+ · Player-Coach", "DE": "Bands D–E · Leadership"}
LEVEL_ROWS = [("E2", "DE"), ("E1", "DE"), ("D2", "DE"), ("D1", "DE"), ("C+", "C+"), ("C2", "C-M"), ("C2", "C-IC"),
              ("C1", "C-M"), ("C1", "C-IC"), ("B+", "B+"), ("B3", "B-M"), ("B3", "B-IC"), ("B2", "B-M"), ("B2", "B-IC"),
              ("B1", "B-M"), ("B1", "B-IC"), ("A3", "A"), ("A2", "A"), ("A1", "A")]
REFS = {
    "weights": "S46–S51 Korn Ferry Hay evaluations · S04, S52 Mercer IPE · S53–S57 NHS, UK NJC/PSC, GLPC, SJC, US OPM point-factor plans",
    "split": "S58 McKinsey skills-based hiring · S60 WEF skills-first · S61 EY India 2026 · the Experience / Education split is an IRIS rule",
    "tilt": "S06 Hay short profile · ±5 is an IRIS rule",
    "levels": "S01 Aon Radford · S02 Pave · S08 SFIA 9 · S03 Mercer career levels",
    "exp": "S32 Google, Microsoft, Amazon · S33 Infosys · S35 Accenture India · IRIS experience guide",
    "span": "S09 McKinsey spans of control · S37 GitLab layers · S44 Google Tech Lead Manager",
    "edu": "IRIS level standard F4 · S24 ICSI, S20 IMA, S18 SHRM for licence-to-practise roles",
}


def rg_of(level, track):
    b = V.band(level)
    mgr = J.text(track) in ("People Manager",)
    if b == "A":
        return "A"
    if b == "B":
        return "B-M" if mgr else "B-IC"
    if b == "C":
        return "C-M" if mgr else "C-IC"
    if b in ("B+", "C+"):
        return b
    if b in ("D", "E"):
        return "DE"
    return None


def short(s, n=360):
    s = J.text(s).replace("\n\n", "\n")
    return s if len(s) <= n else s[:n].rsplit(" ", 1)[0] + " …"


# ================================================================ data from the masters
def load_reference(src):
    """Published weights by role group, the level ladder, the rating grid, crosswalk, experience and spans."""
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    R = {}
    rows = list(wb["1e. Weight Benchmark"].iter_rows(min_row=8, max_row=39, values_only=True))
    pct = lambda v: float(str(v).rstrip("%")) if v not in (None, "") else 0.0
    W = collections.defaultdict(dict)
    for r in rows:
        if not r[0]:
            continue
        lab = J.text(r[0])
        for key, pre in RG_1E.items():
            if lab.startswith(pre):
                src_name = lab.split("·")[-1].strip()
                W[key][src_name] = [pct(x) for x in r[1:7]]   # C, Au, Ac, L, Exp&Qual, BI
    R["weights"] = W
    R["plans"] = [[J.text(x) for x in r] for r in wb["1e. Weight Benchmark"].iter_rows(min_row=79, max_row=88, values_only=True)]
    f4 = list(wb["F4. Levels & Bands"].iter_rows(min_row=17, max_row=31, values_only=True))
    hdr = [J.text(h) for h in f4[0]]
    R["f4"] = {J.text(r[1]): dict(zip(hdr, [J.text(x) for x in r])) for r in f4[1:]}
    grid = list(wb["1d. Weightage Matrix"].iter_rows(min_row=166, max_row=179, values_only=True))
    R["grid"] = {}
    for r in grid:
        lv = J.text(r[0]).split("·")[-1].strip()
        R["grid"][lv] = {"C": J.text(r[2]), "Au": J.text(r[5]), "Ac": J.text(r[8]), "L": J.text(r[11]), "Exp": J.text(r[14]), "BI": J.text(r[17])}
    assert all(all(v for v in d.values()) for d in R["grid"].values()), "rating grid columns moved"
    fm = list(wb["F14. Master Framework"].iter_rows(min_row=473, max_row=498, values_only=True))
    ch = [J.text(h) for h in fm[0]]
    R["cross_hdr"] = ch
    R["cross"] = {J.text(r[0]): [J.text(x) for x in r] for r in fm[1:15] if r[0]}
    eh = [J.text(h) for h in fm[17]]
    R["exp_hdr"] = eh
    ecol = {J.text(h): j for j, h in enumerate(fm[17]) if h}
    R["exp"] = {J.text(r[0]): {h: J.text(r[j]) for h, j in ecol.items()} for r in fm[18:26] if r[0]}
    assert all(d.get("Market range") for d in R["exp"].values()), "experience table columns moved"
    R["span"] = [[J.text(x) for x in r] for r in wb["F14. Master Framework"].iter_rows(min_row=394, max_row=403, values_only=True)]
    mx = list(wb["1d. Weightage Matrix"].iter_rows(min_row=17, max_row=123, values_only=True))
    R["dept_type"] = {(J.text(r[1]), J.text(r[2])): J.text(r[5]) for r in mx}
    wb.close()
    return R


def market_weights(R, rg):
    """Average of Korn Ferry Hay, Mercer IPE and the public plans for a role group, with Experience split from Education."""
    src = {k: v for k, v in R["weights"][rg].items() if "IRIS" not in k}
    avg = [round(statistics.mean(v[i] for v in src.values()), 1) for i in range(6)]
    split = EXP_SPLIT["AB"] if rg in ("A", "B-IC", "B-M", "B+") else EXP_SPLIT["CE"]
    c, au, ac, l_, e, bi = avg
    return {"C": c, "Au": au, "Ac": ac, "L": l_, "BI": bi, "Exp": round(e * split[0], 1), "Edu": round(e * split[1], 1)}, src


def default_weights(R, rg, work):
    w, _ = market_weights(R, rg)
    w = dict(w)
    if rg in ("A", "B-IC", "B-M", "B+", "C-IC", "C-M"):
        for k, d in TILT.get(work, {}).items():
            w[k] = max(0.0, w[k] + d)
    tot = sum(w.values())
    return {k: v * 100.0 / tot for k, v in w.items()}


def to_lmh(pct):
    return "H" if pct >= WEIGHT_H else ("M" if pct >= WEIGHT_M else "L")


def exp_range(R, level):
    s = R["f4"].get(level, {}).get("Experience", "")
    m = re.findall(r"\d+(?:\.\d+)?", s)
    if "+" in s and m:
        return float(m[0]), 99.0
    if len(m) >= 2:
        return float(m[0]), float(m[1])
    if m:
        return float(m[0]), float(m[0])
    return None


def qual_rank(q):
    s = J.text(q).upper()
    if not s or "NOT PROVIDED" in s:
        return None
    if re.search(r"\b(CA|CMA|CWA|ICWA|CS|LLB|ACCA|CFA|COST ACCOUNTANT|ICSI)\b", s):
        return 3
    if re.search(r"\b(MBA|M\.?COM|MCA|M\.?TECH|M\.?SC|MSC|PGDM|PGDBM|PGDBA|MMS|M\.?E\b|MA\b|MASTER|POST GRAD|PG)", s):
        return 2
    if re.search(r"\b(B\.?|BE\b|BTECH|BCOM|BSC|BA\b|BBA|BCA|BMS|BFM|BAF|BACHELOR|GRADUATE|ENGINEERING)", s) and "UNDER" not in s:
        return 1
    return 0


def edu_norm(level):
    v = V.VAL.get(level, 0)
    return 1 if v < 6 else 2


def data_shows(M, p, hr_level):
    """What the data says per parameter against the HR proposed level — guidance, not the input."""
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    if V.band(hr_level) == "X":
        return "Consultant · outside the ladder"
    base = V.VAL[hr_level]
    out = []
    for k, col in (("C", "Complexity rating"), ("Au", "Autonomy rating"), ("Ac", "Accountab. rating"),
                   ("L", "Leadership rating"), ("BI", "Bus. Impact rating")):
        try:
            d = float(p["tm"][H[col]]) - base
            out.append(f"{k} {'H' if d >= 0.75 else 'L' if d <= -0.75 else 'M'}")
        except (TypeError, ValueError):
            out.append(f"{k} —")
    return " · ".join(out)


def exp_edu_shows(R, M, p, hr_level):
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    out = []
    try:
        yrs = float(p["tm"][H["Total exp. · yrs"]])
        rng = exp_range(R, hr_level)
        if rng:
            tag = "L" if yrs < rng[0] else ("H" if yrs > rng[1] else "M")
            out.append(f"Exp {tag} · {yrs:g} yrs vs {R['f4'][hr_level]['Experience']}")
    except (TypeError, ValueError, KeyError):
        out.append("Exp — · not on file")
    q = qual_rank(p["tm"][H["Qualification"]])
    if q is None:
        out.append("Edu — · not on file")
    elif V.band(hr_level) != "X":
        n = edu_norm(hr_level)
        out.append(f"Edu {'L' if q < n else 'H' if q > n else 'M'} · {J.text(p['tm'][H['Qualification']])}")
    return " · ".join(out)


# ================================================================ helpers
TINT = {"1F3864": "EEF2F9", "7030A0": "F4EEF9", "BF8F00": "FFF2CC", "806000": "F8F4E6", "3A3838": "F2F2F2",
        "C00000": "FDECEC", "595959": "F2F2F2", "375623": "EAF2E6", "2E75B6": "EAF3FB"}


def headers(ws, groups, hrow=6):
    c = 1
    out = []
    for title, dark, cols in groups:
        for j, (h, w) in enumerate(cols):
            Y.header(ws.cell(hrow - 1, c + j, title if j == 0 else None), dark, 10)
            Y.header(ws.cell(hrow, c + j, h), dark)
            ws.column_dimensions[L(c + j)].width = w
            out.append((c + j, h, dark, title))
        c += len(cols)
    ws.row_dimensions[hrow].height = 60
    ws.row_dimensions[hrow - 1].height = 30
    return out


def cell(ws, r, c, v, tint=None, bold=False, fmt=None, first=False):
    x = ws.cell(r, c, None if v == "" else v)
    Y.body(x, tint, first)
    if bold:
        x.font = Y.font(9, True)
    if fmt:
        x.number_format = fmt
    s = J.text(v)
    if s.startswith(("KRAs not received", "Full survey answers not", "No survey response")):
        x.fill, x.font = Y.F_RED, Y.T_RED
    return x


def name_range(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=f"{Q(sheet)}!{ref}")


# ================================================================ main
def main(srcdir):
    os.makedirs(J.OUT, exist_ok=True)
    M, S, people, unmapped = J.build(srcdir)
    F_prev = P.Folds(M, people)
    org = V.Org(M, people, F_prev)
    heads = V.department_heads(M, people, org)
    changed = V.apply_heads(M, people, F_prev, heads)
    F = P.Folds(M, people)
    pr = V.propose(org)
    R = load_reference(J.MASTER)
    sup = openpyxl.load_workbook(SUPER, read_only=True, data_only=True)
    s2 = list(sup["2. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    s2h = {J.text(h): i for i, h in enumerate(s2[0])}
    s2rows = [r for r in s2[1:] if r[1] or r[2]]
    s3 = list(sup["3. Raw Data & Market Standards"].iter_rows(min_row=5, values_only=True))
    s3grp, s3h, s3rows = s3[0], s3[1], [r for r in s3[2:] if r[1] or r[2]]
    sup.close()
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    order = [i for i, d in org.top_down()]
    by_i = {p["i"]: p for p in people}

    # family → main type of work (most common department type among its holders)
    fam_work = collections.defaultdict(collections.Counter)
    for p in people:
        fam = J.text(p["tm"][H["Job family"]])
        w = R["dept_type"].get((p["biz"], J.text(p["tm"][H["Department"]])))
        if w:
            fam_work[fam][w] += 1
    families = sorted({J.text(p["tm"][H["Job family"]]) for p in people if J.text(p["tm"][H["Job family"]])})
    fam_type = {f: (fam_work[f].most_common(1)[0][0] if fam_work[f] else "General management") for f in families}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_use = wb.create_sheet("1. How to use")
    ws_tm = wb.create_sheet(TMN)
    ws_raw = wb.create_sheet(RAWN)
    ws_mb = wb.create_sheet(MBN)
    ws_dl = wb.create_sheet(DLN)
    fam_sheet = {}
    used = set()
    for f in families:
        nm = ("RB · " + f)[:31].rstrip(" &·")
        while nm in used:
            nm = nm[:29] + str(len(used))
        used.add(nm)
        fam_sheet[f] = nm
    for f in families:
        wb.create_sheet(fam_sheet[f])
    ws_be = wb.create_sheet(BEN)

    # ------------------------------------------------ Backend Data (built first: everything looks it up)
    Y.banner(ws_be, "Backend Data — the lookups that make the workbook live",
             "Last sheet · Low / Medium / High in points and steps, the level ladders, the band order, the default weightage of every "
             "family × level × role group (linked to the RB sheets), and the drop-down lists. Edit the yellow cells only if HR agrees "
             "to change the model.", J.STAMP, width_cols=10)
    r = 5
    Y.header(ws_be.cell(r, 1, "LOW · MEDIUM · HIGH — what each choice is worth"), Y.NAVY, 10)
    r += 1
    for j, h in enumerate(["Code", "Means", "Weightage points", "Assessment step", "Explained"], 1):
        Y.header(ws_be.cell(r, j, h), "595959")
    r += 1
    lmh_top = r
    for code, nm, pts, step, expl in LMH:
        cell(ws_be, r, 1, code, "FFFFFF", True)
        cell(ws_be, r, 2, nm, "FFFFFF")
        cell(ws_be, r, 3, pts, "FFF2CC")
        cell(ws_be, r, 4, step, "FFF2CC")
        cell(ws_be, r, 5, expl, "FFFFFF")
        r += 1
    name_range(wb, "LMH_KEY", BEN, f"$A${lmh_top}:$A${r - 1}")
    name_range(wb, "LMH_PTS", BEN, f"$C${lmh_top}:$C${r - 1}")
    name_range(wb, "LMH_STEP", BEN, f"$D${lmh_top}:$D${r - 1}")
    cell(ws_be, r, 1, "Weightage: a parameter's share = its points ÷ the points of all seven. With L 1 · M 3 · H 9 the shares follow the "
                      "published job-evaluation weights within a few points — see 4. Market Benchmark, section 4. Assessment: the "
                      "weighted net of the seven steps runs from −1 to +1; at +0.5 or more the person moves one level up from the HR "
                      "proposed level, at −0.5 or less one level down.", "FFFFFF")
    ws_be.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws_be.row_dimensions[r].height = 45
    r += 2
    Y.header(ws_be.cell(r, 1, "LEVELS — value, band and band order"), Y.NAVY, 10)
    r += 1
    for j, h in enumerate(["Level", "Value", "Band", "Band rank", "Standard ladder", "Player-Coach ladder"], 1):
        Y.header(ws_be.cell(r, j, h), "595959")
    r += 1
    lv_top = r
    lvls = V.LEVELS_PC + ["X1", "X2", "X3"]
    for k, lv in enumerate(lvls):
        b = V.band(lv)
        cell(ws_be, r, 1, lv, "FFFFFF", True)
        cell(ws_be, r, 2, V.VAL.get(lv, 0), "FFFFFF")
        cell(ws_be, r, 3, b, "FFFFFF")
        cell(ws_be, r, 4, V.BAND_ORDER.index(b) + 1 if b in V.BAND_ORDER else 0, "FFFFFF")
        cell(ws_be, r, 5, V.LEVELS_STD[k] if k < len(V.LEVELS_STD) else None, "FFFFFF")
        cell(ws_be, r, 6, V.LEVELS_PC[k] if k < len(V.LEVELS_PC) else None, "FFFFFF")
        r += 1
    name_range(wb, "B_LVL", BEN, f"$A${lv_top}:$A${r - 1}")
    name_range(wb, "B_VAL", BEN, f"$B${lv_top}:$B${r - 1}")
    name_range(wb, "B_RANK", BEN, f"$D${lv_top}:$D${r - 1}")
    name_range(wb, "LAD_STD", BEN, f"$E${lv_top}:$E${lv_top + len(V.LEVELS_STD) - 1}")
    name_range(wb, "LAD_PC", BEN, f"$F${lv_top}:$F${lv_top + len(V.LEVELS_PC) - 1}")
    name_range(wb, "LIST_LEVELS", BEN, f"$A${lv_top}:$A${r - 1}")
    r += 1
    Y.header(ws_be.cell(r, 1, "DEFAULT WEIGHTAGE — family × level × role group · linked to each RB sheet"), Y.NAVY, 10)
    r += 1
    for j, h in enumerate(["Key", "Job family", "Level", "Role group"] + [n for _, n in PARAMS], 1):
        Y.header(ws_be.cell(r, j, h), "595959")
    r += 1
    rbd_top = r
    rb_rows = {}       # (family, level, rg) → backend row; filled after RB sheets are laid out
    for f in families:
        for lv, rg in LEVEL_ROWS:
            rb_rows[(f, lv, rg)] = r
            r += 1
    rbd_end = r - 1
    name_range(wb, "RBD_KEY", BEN, f"$A${rbd_top}:$A${rbd_end}")
    for k, (pk, _) in enumerate(PARAMS):
        name_range(wb, f"RBD_{pk}", BEN, f"${L(5 + k)}${rbd_top}:${L(5 + k)}${rbd_end}")
    r += 1
    Y.header(ws_be.cell(r, 1, "LISTS"), Y.NAVY, 10)
    r += 1
    for k, (code, *_rest) in enumerate(LMH):
        cell(ws_be, r + k, 1, code, "FFFFFF")
    name_range(wb, "LIST_LMH", BEN, f"$A${r}:$A${r + 2}")
    for col, w in zip("ABCDEFGHIJK", (34, 22, 12, 22, 14, 14, 14, 14, 14, 14, 14)):
        ws_be.column_dimensions[col].width = w

    # ------------------------------------------------ RB family sheets
    cat_by_fl = collections.defaultdict(list)
    cj = {h: j for j, h in enumerate(M["cat_hdr"])}
    for c in M["cat"]:
        cat_by_fl[(J.text(c[1]), J.text(c[3]))].append(c)
    dv_lmh_all = []
    for f in families:
        ws = wb[fam_sheet[f]]
        work = fam_type[f]
        Y.banner(ws, f"Role & band matrix — {f}",
                 f"What Low, Medium and High mean for each of the seven parameters at every level of {f}, and how much each "
                 f"parameter counts · main type of work: {work}. Medium is the band as described; Low is the level below; High is "
                 "the level above. Change a yellow weightage and every role, department ladder and person in this family "
                 "recalculates.", J.STAMP, width_cols=14)
        r = 4
        if f.startswith("Sales"):
            Y.header(ws.cell(r, 1, "HR draft — Sales anchors by level · HR's sketch of 12 October 2026 · confirm the unit and how L1–L6 map to IRIS levels"), "C00000", 10)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
            r += 1
            for row in (("Parameter", "L1", "L2", "L3", "L4", "L5", "L6"), ("Accountability", "20–30", "30–40", "50", "60", "80", "100"),
                        ("Impact", "5", "25", "Revenue numbers", "", "", ""), ("Complexity", "", "", "", "", "", ""),
                        ("Autonomy", "", "", "", "", "", ""), ("Experience", "", "", "", "", "", ""), ("Education", "", "", "", "", "", "")):
                for j, v in enumerate(row, 1):
                    cell(ws, r, j, v, "FDECEC", j == 1 or row[0] == "Parameter")
                r += 1
            r += 1
        groups = [("LEVEL", "1F3864", [("Level", 7), ("Band", 16), ("Role group", 18), ("Level title", 20), ("Roles at this level in this family", 34)])]
        for pk, pn in PARAMS:
            groups.append((pn.upper(), "806000" if pk in ("C", "Au", "Ac") else ("375623" if pk in ("L", "BI") else "2E75B6"),
                           [("Weightage · L/M/H", 9), ("Low", 30), ("Medium", 34), ("High", 30)]))
        groups.append(("MARKET BENCHMARK & REFERENCES", "3A3838", [("Market weight % · C · Au · Ac · L · BI · Exp · Edu", 26),
                                                                   ("Market experience · yrs", 16), ("Span norm", 22),
                                                                   ("Radford / Mercer · SFIA 9", 18), ("References", 40)]))
        hdr_row = r + 1
        cols = headers(ws, groups, hdr_row)
        r = hdr_row + 1
        first_row = r
        for lv, rg in LEVEL_ROWS:
            f4 = R["f4"].get(lv, {})
            g = R["grid"].get(lv, {})
            pos = V.LEVELS_PC.index(lv)
            below = V.LEVELS_PC[pos - 1] if pos > 0 else None
            above = V.LEVELS_PC[pos + 1] if pos + 1 < len(V.LEVELS_PC) else None
            if lv not in ("B+", "C+"):
                below = next((x for x in reversed(V.LEVELS_PC[:pos]) if x not in ("B+", "C+")), None)
                above = next((x for x in V.LEVELS_PC[pos + 1:] if x not in ("B+", "C+")), None)
            cats = cat_by_fl.get((f, lv), [])
            roles = "\n".join(sorted({f"{J.text(c[2])} · {J.text(c[6]).splitlines()[0] if J.text(c[6]) else ''}" for c in cats})) or "— no role at this level today"
            dw = default_weights(R, rg, work if rg not in ("C+", "DE") else "General management")
            mw, _ = market_weights(R, rg)
            vals = [lv, V.BAND_NAME.get(V.band(lv), ""), RG_NAME[rg], V.level_title(lv, "People Manager" if rg.endswith("M") or rg == "DE" else "IC"), roles]
            theme = lambda col: next((short(J.text(c[cj[col]]), 160) for c in cats if J.has(c[cj[col]])), "")
            for pk, pn in PARAMS:
                if pk == "Exp":
                    rng = f4.get("Experience", "")
                    lo = f"Less experience than the band's typical {rng}"
                    md = f"Typical for the band: {rng} · market {R['exp'].get(lv, {}).get('Market range') or '—'} yrs"
                    hi = f"More experience than the band's typical {rng}"
                elif pk == "Edu":
                    md = f4.get("Education", "")
                    lo = ("Below: " + R["f4"].get(below, {}).get("Education", "")) if below else "Below graduation · still studying"
                    hi = ("Above: " + R["f4"].get(above, {}).get("Education", "")) if above else "Board-level credentials beyond the role's need"
                    if work == "Corporate functions":
                        md += " · licence-to-practise roles: CA, CS, CMA or LLB as the role needs"
                else:
                    gk = pk
                    md = g.get(gk, "")
                    if pk == "L" and V.band(lv) == "A":
                        md = "Helps peers and newer people informally · " + f4.get("Span of control", "")
                    if pk == "C" and theme("Skills & knowledge · counts"):
                        md += f" · in {f}: {theme('Skills & knowledge · counts')}"
                    if pk == "Ac" and theme("Manager-set KRAs · common themes"):
                        md += f" · KRA themes: {theme('Manager-set KRAs · common themes')}"
                    if pk == "BI" and theme("Purpose"):
                        md += f" · purpose: {theme('Purpose')}"
                    lo = ("Works at " + below + ": " + R["grid"].get(below, {}).get(gk, "")) if below else "Not yet at the entry standard · in induction"
                    hi = ("Already works at " + above + ": " + R["grid"].get(above, {}).get(gk, "")) if above else "Beyond the role · a board-level mandate"
                    if pk == "L" and below and V.band(below) == "A":
                        lo = "Works at " + below + ": contributes alone; guides others only informally"
                vals += [to_lmh(dw[pk]), lo, md, hi]
            cross = R["cross"].get(lv, [])
            span = f4.get("Span of control", "")
            vals += [" · ".join(f"{mw[k]:g}" for k in PKEYS), R["exp"].get(lv, {}).get("Market range") or f4.get("Experience", ""),
                     span, f"{cross[1] if len(cross) > 1 else ''} · SFIA {cross[3] if len(cross) > 3 else ''}",
                     f"Weights: {REFS['weights']} · Exp/Edu: {REFS['split']} · tilt: {REFS['tilt']} · levels: {REFS['levels']} · experience: {REFS['exp']} · span: {REFS['span']}"]
            prev = None
            for (c_, h, dark, title), v in zip(cols, vals):
                x = cell(ws, r, c_, v, TINT.get(dark, "FFFFFF"), first=title != prev)
                prev = title
                if h.startswith("Weightage"):
                    x.fill = Y.F_YEL
                    x.font = Y.font(10, True)
                    x.alignment = Alignment(horizontal="center", vertical="top")
                    dv_lmh_all.append((ws, f"{L(c_)}{r}"))
                if h == "Level":
                    b = Y.band_of(v)
                    if b:
                        x.fill = PatternFill("solid", fgColor=Y.BAND[b])
                    x.font = Y.font(10, True)
            # backend default row → this RB row
            br = rb_rows[(f, lv, rg)]
            cell(ws_be, br, 1, f"{f}|{lv}|{rg}", "FFFFFF")
            cell(ws_be, br, 2, f, "FFFFFF")
            cell(ws_be, br, 3, lv, "FFFFFF")
            cell(ws_be, br, 4, RG_NAME[rg], "FFFFFF")
            wcols = [c_ for c_, h, _, _ in cols if h.startswith("Weightage")]
            for k, wc in enumerate(wcols):
                cell(ws_be, br, 5 + k, f"={Q(fam_sheet[f])}!{L(wc)}{r}", "F2F2F2")
            ws.row_dimensions[r].height = 120
            r += 1
        ws.freeze_panes = ws.cell(first_row, 6)
        ws.auto_filter.ref = f"A{hdr_row}:{L(len(cols))}{r - 1}"
        ws.sheet_view.zoomScale = 85

    # ------------------------------------------------ people, top down
    tm_rows = {}
    for n, i in enumerate(order):
        tm_rows[i] = 7 + n
    raw_row = {i: 7 + n for n, i in enumerate(order)}

    # department ladder rows: business × department × designation × HR proposed level × role group
    dl = collections.OrderedDict()
    for i in order:
        p = by_i[i]
        lv = pr[i]["level"]
        rg = rg_of(lv, p["tm"][H["Track"]])
        key = "|".join([p["biz"], J.text(p["tm"][H["Department"]]), J.text(p["tm"][H["Final designation"]]), lv, rg or "X"])
        dl.setdefault(key, {"biz": p["biz"], "dept": J.text(p["tm"][H["Department"]]), "des": J.text(p["tm"][H["Final designation"]]),
                            "lv": lv, "rg": rg, "fam": collections.Counter(), "holders": []})
        dl[key]["fam"][J.text(p["tm"][H["Job family"]])] += 1
        dl[key]["holders"].append(p)
        p["dl_key"] = key
    biz_order = ["Corporate", "CARBON", "iFile", "iDEAL", "Datatech", "Consume"]
    dl_sorted = sorted(dl.values(), key=lambda d: (biz_order.index(d["biz"]) if d["biz"] in biz_order else 9, d["dept"],
                                                   -V.VAL.get(d["lv"], 0), d["des"]))

    # ---- Department Ladder sheet
    ws = ws_dl
    Y.banner(ws, "Department Ladder — every role in every department, top down, with its weightage",
             "Sheet 5 · One row per role: business → department → designation at its HR proposed level, highest first. The seven "
             "weightage cells start from the family's role & band matrix · RB sheets; overwrite any of them with L, M or H to set "
             "this role's own weightage. Talent Mapping reads it at once, so every holder's band recalculates.", J.STAMP, width_cols=14)
    Y.colour_key(ws, 4)
    groups = [("ROLE", "1F3864", [("Business", 10), ("Department", 18), ("Level · HR proposed", 9), ("Band", 16),
                                  ("Designation", 30), ("Role group", 18), ("Job family", 22), ("Holders", 8), ("Who", 34)]),
              ("WEIGHTAGE · L / M / H — edit to set this role's weightage", "BF8F00", [(n, 11) for _, n in PARAMS]),
              ("SHARE OF THE SCORE · %, live", "806000", [(n + " %", 9) for _, n in PARAMS]),
              ("NOTES", "595959", [("Why changed · business head", 30), ("Ladder key", 30)])]
    cols = headers(ws, groups)
    r = 7
    dl_top = r
    prev_dept = None
    for d in dl_sorted:
        fam = d["fam"].most_common(1)[0][0]
        vals = [d["biz"], d["dept"], d["lv"], V.BAND_NAME.get(V.band(d["lv"]), ""), d["des"], RG_NAME.get(d["rg"], "Consultant"), fam,
                len(d["holders"]), ", ".join(p["name"] for p in d["holders"])]
        for k, (pk, _) in enumerate(PARAMS):
            if d["rg"]:
                vals.append(f'=IFERROR(INDEX(RBD_{pk},MATCH("{fam}|{d["lv"]}|{d["rg"]}",RBD_KEY,0)),"M")')
            else:
                vals.append("")
        wc0 = 10
        for k in range(7):
            a = L(wc0 + k)
            allw = ",".join(f"IFERROR(INDEX(LMH_PTS,MATCH({L(wc0 + t)}{r},LMH_KEY,0)),0)" for t in range(7))
            vals.append(f'=IF({a}{r}="","",IFERROR(INDEX(LMH_PTS,MATCH({a}{r},LMH_KEY,0))/SUM({allw}),""))')
        vals += ["", "|".join([d["biz"], d["dept"], d["des"], d["lv"], d["rg"] or "X"])]
        prev = None
        for (c_, h, dark, title), v in zip(cols, vals):
            x = cell(ws, r, c_, v, TINT.get(dark, "FFFFFF"), first=title != prev)
            prev = title
            if title.startswith("WEIGHTAGE"):
                x.fill = Y.F_YEL
                x.font = Y.font(10, True)
                x.alignment = Alignment(horizontal="center", vertical="top")
                dv_lmh_all.append((ws, f"{L(c_)}{r}"))
            if title.startswith("SHARE"):
                x.number_format = "0%"
            if h == "Level · HR proposed":
                b = Y.band_of(v)
                if b:
                    x.fill = PatternFill("solid", fgColor=Y.BAND[b])
                x.font = Y.font(10, True)
        if d["dept"] != prev_dept:
            ws.cell(r, 2).font = Y.font(9, True)
        prev_dept = d["dept"]
        ws.row_dimensions[r].height = Y.est_height([(vals[8], 34), (vals[4], 30)], maximum=120, minimum=18)
        r += 1
    dl_end = r - 1
    keycol = L(len(cols))
    name_range(wb, "DL_KEY", DLN, f"${keycol}${dl_top}:${keycol}${dl_end}")
    for k, (pk, _) in enumerate(PARAMS):
        name_range(wb, f"DL_{pk}", DLN, f"${L(10 + k)}${dl_top}:${L(10 + k)}${dl_end}")
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:{keycol}{dl_end}"
    ws.sheet_view.zoomScale = 90

    # ---- Talent Mapping
    ws = ws_tm
    Y.banner(ws, "Talent Mapping — business heads: review, weigh and propose the band",
             "Sheet 2 · One row per person, top down from the CEO. Read left to right: who → what they and their manager told us, with a "
             "link to all their data → how much each parameter counts · L/M/H → your assessment against the band · L/M/H → the market "
             "benchmark → HR's proposed band → what benchmarking says → the final proposed band, which updates the moment you change any "
             "L, M or H. Yellow = yours to edit.", J.STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    groups = [
        ("CORE EMPLOYEE DETAILS", "1F3864", [("Business", 10), ("Department", 18), ("Layer from the CEO", 7), ("Emp ID", 12), ("Employee name", 22),
                                             ("Current designation", 24), ("Proposed designation", 26), ("Job family · v1.14", 20),
                                             ("Sub-family", 22), ("Track", 12), ("Role ID", 10), ("Reporting manager", 20),
                                             ("Manager Emp ID", 12), ("Total exp. · yrs", 8), ("Exp. at IRIS · yrs", 8), ("Qualification", 16),
                                             ("Current level · as is", 9), ("Model level · v1.13", 9)]),
        ("SUMMARY OF RESPONSES RECEIVED — full detail in 3. Raw Data", "7030A0", [
            ("Survey · questions answered", 9), ("What they shared · summary", 48), ("KRAs received · summary", 48),
            ("Manager talent mapping · rating · potential · 9-box · readiness", 26), ("What the data shows · vs HR proposed band", 34),
            ("All their data ▸", 12)]),
        ("WEIGHTAGE · L / M / H — from 5. Department Ladder; overwrite for this person", "BF8F00", [(f"{n} · weightage", 10) for _, n in PARAMS]),
        ("YOUR ASSESSMENT · L / M / H against the band — Medium = meets it", "C00000", [(f"{n} · assessment", 10) for _, n in PARAMS]),
        ("MARKET BENCHMARK — this role group at the HR proposed level", "3A3838",
         [(f"Market weight · {n} %", 9) for _, n in PARAMS] + [("Market experience · yrs", 12), ("Education norm", 24),
                                                               ("Span norm", 20), ("Radford / Mercer · SFIA 9", 16), ("References", 30)]),
        ("PROPOSAL", "C00000", [("HR proposed band · market benchmark and people inputs", 10), ("HR proposed · band and title", 26),
                                ("How HR set it", 40), ("Input from benchmarking", 10), ("Benchmarking · how", 40),
                                ("Weighted net · −1 to +1, live", 10), ("Final proposed band · business inputs, live", 11),
                                ("Business head override · optional", 11), ("FINAL PROPOSED BAND", 11), ("Band", 16),
                                ("Move vs HR proposed", 14), ("Manager's final proposed band", 11), ("Manager rule check · live", 22),
                                ("Business head comment", 30)]),
        ("ENGINE — points and steps, live", "595959", [(f"{k} points", 6) for k in PKEYS] + [(f"{k} step", 6) for k in PKEYS] + [("Ladder key", 30)]),
    ]
    cols = headers(ws, groups)
    col = {h: L(c_) for c_, h, _, _ in cols}
    wcols = [col[f"{n} · weightage"] for _, n in PARAMS]
    acols = [L(c_) for c_, h, d, t in cols if t.startswith("YOUR ASSESSMENT")]
    pcols = [col[f"{k} points"] for k in PKEYS]
    scols = [col[f"{k} step"] for k in PKEYS]
    last = 6 + len(order)
    E = col["Emp ID"]
    FIN = col["FINAL PROPOSED BAND"]
    for i in order:
        p = by_i[i]
        r = tm_rows[i]
        row2 = s2rows[i]
        g = lambda h: row2[s2h[h]] if h in s2h else None
        lv = pr[i]["level"]
        rg = rg_of(lv, p["tm"][H["Track"]])
        k, sv, links = person_values(M, F, p)
        survey = sv[1] if not J.text(sv[-1]).startswith(("Full", "No survey")) else (sv[6] or sv[-1])
        talent = " · ".join(J.text(g(h)) for h in ("Latest rating", "Potential", "9-box", "Readiness") if J.has(g(h))) if "Latest rating" in s2h else ""
        s8 = F.s8row(p)
        if s8 is not None:
            hh = M["s8_hdr"]
            talent = " · ".join(f"{h}: {J.text(s8[hh.index(h)])}" for h in ("Latest rating", "Potential", "9-box", "Readiness") if J.has(s8[hh.index(h)]))
        shows = data_shows(M, p, lv) + (" · " + exp_edu_shows(R, M, p, lv) if V.band(lv) != "X" else "")
        mw, _ = market_weights(R, rg) if rg else ({k_: "" for k_ in PKEYS}, None)
        f4 = R["f4"].get(lv, {})
        cross = R["cross"].get(lv, [])
        # benchmarking input: the six ratings on the market-average weights, without IRIS's tilt
        bench_lv, bench_txt = "", "Consultant · outside the ladder"
        if rg:
            try:
                rs = [float(p["tm"][H[h]]) for h in ("Complexity rating", "Autonomy rating", "Accountab. rating", "Leadership rating",
                                                      "Exp + Qual rating", "Bus. Impact rating")]
                w6 = [mw["C"], mw["Au"], mw["Ac"], mw["L"], mw["Exp"] + mw["Edu"], mw["BI"]]
                sc = sum(a * b for a, b in zip(rs, w6)) / sum(w6)
                bench_lv = V.LEVELS_STD[max(0, min(11, int(sc + 0.5) - 1))]
                rng = exp_range(R, bench_lv)
                yrs = J.text(p["tm"][H["Total exp. · yrs"]])
                bench_txt = (f"The six ratings on the market-average weights of {RG_NAME[rg]} score {sc:.2f} → {bench_lv} · "
                             f"Hay, Mercer IPE and public plans · experience {yrs} yrs vs the {bench_lv} market range "
                             f"{R['exp'].get(bench_lv, {}).get('Market range') or R['f4'].get(bench_lv, {}).get('Experience', '—')}")
            except (TypeError, ValueError):
                bench_lv = lv
                bench_txt = "Placed by a category rule — no full set of ratings · benchmark follows the HR proposed level"
        vals = {
            "Business": p["biz"], "Department": J.text(p["tm"][H["Department"]]), "Layer from the CEO": pr[i]["depth"],
            "Emp ID": p["emp"], "Employee name": p["name"], "Current designation": J.text(p["tm"][H["Current designation"]]),
            "Proposed designation": J.text(p["tm"][H["Final designation"]]), "Job family · v1.14": J.text(p["tm"][H["Job family"]]),
            "Sub-family": J.text(p["tm"][H["Sub-family · what they do"]]), "Track": J.text(p["tm"][H["Track"]]), "Role ID": links[0],
            "Reporting manager": J.text(p["tm"][H["Reporting manager"]]), "Manager Emp ID": links[1],
            "Total exp. · yrs": p["tm"][H["Total exp. · yrs"]], "Exp. at IRIS · yrs": p["tm"][H["Exp. at IRIS · yrs"]],
            "Qualification": J.text(p["tm"][H["Qualification"]]), "Current level · as is": J.text(p["tm"][H["Current level"]]),
            "Model level · v1.13": J.text(p["tm"][H["FINAL LEVEL"]]),
            "Survey · questions answered": f"{sv[0]} of 20", "What they shared · summary": short(survey),
            "KRAs received · summary": (k[5] + " · " + short(k[0], 320)) if J.text(k[5]).startswith("Received") else k[0],
            "Manager talent mapping · rating · potential · 9-box · readiness": talent or "Not in the manager mapping",
            "What the data shows · vs HR proposed band": shows, "All their data ▸": "Open ▸",
            "Market experience · yrs": R["exp"].get(lv, {}).get("Market range") or f4.get("Experience", ""),
            "Education norm": f4.get("Education", ""), "Span norm": f4.get("Span of control", ""),
            "Radford / Mercer · SFIA 9": f"{cross[1] if len(cross) > 1 else ''} · SFIA {cross[3] if len(cross) > 3 else ''}" if cross else "—",
            "References": f"{REFS['weights']} · {REFS['split']} · {REFS['exp']} · {REFS['levels']}",
            "HR proposed band · market benchmark and people inputs": lv,
            "HR proposed · band and title": f"{V.BAND_NAME.get(V.band(lv), '')} · {V.level_title(lv, p['tm'][H['Track']])}",
            "How HR set it": f"Model {J.text(p['tm'][H['FINAL LEVEL']])} from the six ratings, the market-derived weights and rules R1–R13 · then {pr[i]['note']}",
            "Input from benchmarking": bench_lv, "Benchmarking · how": bench_txt, "Business head comment": "",
            "Business head override · optional": "", "Ladder key": p["dl_key"],
        }
        for (pk, pn), wc in zip(PARAMS, wcols):
            vals[f"{pn} · weightage"] = f'=IFERROR(INDEX(DL_{pk},MATCH(${col["Ladder key"]}{r},DL_KEY,0)),"M")' if rg else ""
            vals[f"Market weight · {pn} %"] = (mw[pk] / 100.0) if rg else ""
        for (pk, pn), wc, ac, pc, sc_ in zip(PARAMS, wcols, acols, pcols, scols):
            vals[f"{pk} points"] = f'=IFERROR(INDEX(LMH_PTS,MATCH({wc}{r},LMH_KEY,0)),0)'
            vals[f"{pk} step"] = f'=IFERROR(INDEX(LMH_STEP,MATCH({ac}{r},LMH_KEY,0)),0)'
        HRc = col["HR proposed band · market benchmark and people inputs"]
        net = col["Weighted net · −1 to +1, live"]
        pts = ",".join(f"{c_}{r}" for c_ in pcols)
        prod = "+".join(f"{a}{r}*{b}{r}" for a, b in zip(pcols, scols))
        vals["Weighted net · −1 to +1, live"] = f'=IF(SUM({pts})=0,0,({prod})/SUM({pts}))'
        step = f'IF({net}{r}>=0.5,1,IF({net}{r}<=-0.5,-1,0))'
        vals["Final proposed band · business inputs, live"] = (
            f'=IF(LEFT({HRc}{r},1)="X",{HRc}{r},IF(OR({HRc}{r}="B+",{HRc}{r}="C+"),'
            f'INDEX(LAD_PC,MAX(1,MIN(14,MATCH({HRc}{r},LAD_PC,0)+{step}))),'
            f'INDEX(LAD_STD,MAX(1,MIN(12,MATCH({HRc}{r},LAD_STD,0)+{step})))))')
        live = col["Final proposed band · business inputs, live"]
        ov = col["Business head override · optional"]
        vals["FINAL PROPOSED BAND"] = f'=IF({ov}{r}<>"",{ov}{r},{live}{r})'
        vals["Band"] = f'=IFERROR(INDEX(B_RANK,MATCH({FIN}{r},B_LVL,0)),"")'
        vals["Band"] = f'=IFERROR(CHOOSE(INDEX(B_RANK,MATCH({FIN}{r},B_LVL,0)),"Band A","Band B","B+ · Player-Coach","Band C","C+ · Player-Coach","Band D","Band E"),"Band X · Consultant")'
        vals["Move vs HR proposed"] = (f'=IF({FIN}{r}={HRc}{r},"• Same",IF(INDEX(B_VAL,MATCH({FIN}{r},B_LVL,0))>'
                                       f'INDEX(B_VAL,MATCH({HRc}{r},B_LVL,0)),"▲ Up","▼ Down"))')
        mgr_col = col["Manager Emp ID"]
        vals["Manager's final proposed band"] = (f'=IF(OR(${mgr_col}{r}="",${mgr_col}{r}="—",${mgr_col}{r}="Not in this file"),"",'
                                                 f'IFERROR(INDEX(${FIN}$7:${FIN}${last},MATCH(${mgr_col}{r},${E}$7:${E}${last},0)),""))')
        mf = col["Manager's final proposed band"]
        vp = f"INDEX(B_VAL,MATCH({FIN}{r},B_LVL,0))"
        vm = f"INDEX(B_VAL,MATCH({mf}{r},B_LVL,0))"
        bp = f"INDEX(B_RANK,MATCH({FIN}{r},B_LVL,0))"
        bm = f"INDEX(B_RANK,MATCH({mf}{r},B_LVL,0))"
        vals["Manager rule check · live"] = (
            f'=IFERROR(IF({mf}{r}="","—",IF(OR(LEFT({FIN}{r},1)="X",LEFT({mf}{r},1)="X"),"—",'
            f'IF(AND(LEFT({FIN}{r},1)="E",LEFT({mf}{r},1)="E"),"Exception · Band E",'
            f'IF(AND({vm}>=7,{bp}>={bm}),"⚠ Same band as a head manager",'
            f'IF({vp}>={vm},"⚠ At or above the manager\'s level","OK"))))),"—")')
        prev = None
        for c_, h, dark, title in cols:
            v = vals.get(h, "")
            x = cell(ws, r, c_, v, TINT.get(dark, "FFFFFF"), first=title != prev)
            prev = title
            if title.startswith("WEIGHTAGE"):
                x.fill = Y.F_YEL
                x.alignment = Alignment(horizontal="center", vertical="top")
                x.font = Y.font(10, True)
            if title.startswith("YOUR ASSESSMENT"):
                if rg:
                    x.value = "M"
                x.fill = Y.F_YEL
                x.alignment = Alignment(horizontal="center", vertical="top")
                x.font = Y.font(10, True)
            if h.startswith("Market weight"):
                x.number_format = "0%"
            if h in ("Business head override · optional", "Business head comment"):
                x.fill = Y.F_YEL
            if h in ("HR proposed band · market benchmark and people inputs", "FINAL PROPOSED BAND", "Input from benchmarking",
                     "Final proposed band · business inputs, live"):
                x.font = Y.font(11, True)
                x.alignment = Alignment(horizontal="center", vertical="top")
            if h == "Weighted net · −1 to +1, live":
                x.number_format = "0.00"
            if h == "All their data ▸":
                x.hyperlink = f"#{Q(RAWN)}!A{raw_row[i]}"
                x.font = Y.font(9, True, "0563C1", underline="single")
        ws.row_dimensions[r].height = 110
    # drop-downs
    dv_t = DataValidation(type="list", formula1="LIST_LMH", allow_blank=True)
    dv_o = DataValidation(type="list", formula1="LIST_LEVELS", allow_blank=True)
    ws.add_data_validation(dv_t)
    ws.add_data_validation(dv_o)
    dv_t.add(f"{wcols[0]}7:{wcols[-1]}{last}")
    dv_t.add(f"{acols[0]}7:{acols[-1]}{last}")
    dv_o.add(f"{col['Business head override · optional']}7:{col['Business head override · optional']}{last}")
    for k in (pcols + scols + [col["Ladder key"]]):
        ws.column_dimensions[k].hidden = True
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:{L(len(cols))}{last}"
    ws.sheet_view.zoomScale = 85
    for wsx in {w for w, _ in dv_lmh_all}:
        dv = DataValidation(type="list", formula1="LIST_LMH", allow_blank=True)
        wsx.add_data_validation(dv)
        for w2, ref in dv_lmh_all:
            if w2 is wsx:
                dv.add(ref)

    # ---- Raw Data
    ws = ws_raw
    Y.banner(ws, "Raw Data — everything we hold on each person",
             "Sheet 3 · Same order as Talent Mapping, top down. HRIS, the survey count and per-person summary, the manager talent-mapping "
             "files, KRAs in full and the CARBON Mapping Master. People in the source files but not yet in the architecture are at the "
             "bottom. Personal fields are never carried. ◂ Back takes you to the person's row on Talent Mapping.", J.STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    s2_extra = [h for h, _ in P.KRA_COLS] + [h for h, _ in P.SURVEY_COLS]
    raw_groups = [("PERSON", "1F3864", [("◂ Back", 8), ("Status", 12)])]
    cur = None
    for j, h in enumerate(s3h):
        if s3grp[j]:
            cur = [J.text(s3grp[j]), ["1F3864", "806000", "548235", "7030A0"][min(3, len(raw_groups) - 1)], []]
            raw_groups.append(cur)
        cur[2].append((J.text(h), 18 if j > 21 else 14))
    raw_groups.append(["KRAs IN FULL & WHAT THEY SHARED IN THE SURVEY", "375623", [(h, 45) for h in s2_extra]])
    rcols = headers(ws, [(t, d, c) for t, d, c in raw_groups])
    for n, i in enumerate(order):
        r = raw_row[i]
        row3 = s3rows[i]
        row2 = s2rows[i]
        x = cell(ws, r, 1, "◂ Back", "EEF2F9", first=True)
        x.hyperlink = f"#{Q(TMN)}!A{tm_rows[i]}"
        x.font = Y.font(9, True, "0563C1", underline="single")
        cell(ws, r, 2, "Mapped", "EEF2F9")
        for j, v in enumerate(row3):
            cell(ws, r, 3 + j, J.debracket(v) if isinstance(v, str) else v, "FFFFFF" if j > 21 else "EEF2F9", first=bool(s3grp[j]))
        for k, h in enumerate(s2_extra):
            cell(ws, r, 3 + len(row3) + k, row2[s2h[h]], "EAF2E6", first=k == 0)
        ws.row_dimensions[r].height = 90
    r = 7 + len(order)
    seen = set()
    for u in unmapped:
        kk = J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"])
        if kk in seen:
            continue
        seen.add(kk)
        recs = [x for x in unmapped if (J.key(x["rec"]["id"]) or J.nname(x["rec"]["name"])) == kk]
        cell(ws, r, 1, "", "FFC7CE", first=True)
        removed = J.nname(u["rec"]["name"]) == "SHREYASH PAWAR"
        cell(ws, r, 2, "Removed by HR · 9 Oct 2026" if removed else "NOT YET MAPPED", "FFC7CE", True)
        vals = {}
        for x in recs:
            pre = J.PREFIX[x["source"]["kind"]]
            for h, v in x["rec"]["row"].items():
                if J.has(v):
                    for j, mh in enumerate(s3h):
                        if J.text(mh).startswith(pre) and J.hkey(mh) == J.hkey(h):
                            vals.setdefault(j, J.clean(v))
        vals[0] = {"MM": "CARBON", "CRAW": "CARBON"}.get(u["source"]["kind"], u["source"]["label"].split("· ")[-1])
        vals[1] = u["rec"]["id"] or "—"
        vals[2] = u["rec"]["name"]
        for j in range(len(s3h)):
            cell(ws, r, 3 + j, vals.get(j), "FDECEC")
        ws.row_dimensions[r].height = 60
        r += 1
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:{L(len(rcols))}{r - 1}"
    ws.sheet_view.zoomScale = 85

    # ---- Market Benchmark
    ws = ws_mb
    Y.banner(ws, "Market Benchmark — the numbers behind the weightage and the bands",
             "Sheet 4 · Published job-evaluation weights by role group (Korn Ferry Hay, Mercer IPE and five public point-factor plans), "
             "the Experience / Education split, how Low / Medium / High reproduce the market weights, and the level crosswalk with "
             "experience, education and span norms. Every number carries its source.", J.STAMP, width_cols=14)
    r = 5

    def sec(title, refs):
        nonlocal r
        Y.header(ws.cell(r, 1, title), Y.NAVY, 11)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        r += 1
        x = ws.cell(r, 1, "References: " + refs)
        x.font = Y.font(8, True, "375623", italic=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        r += 1

    def trow(vals, head=False, tint="FFFFFF", fmt=None):
        nonlocal r
        for j, v in enumerate(vals, 1):
            if head:
                Y.header(ws.cell(r, j, v), "595959")
            else:
                x = cell(ws, r, j, v, tint, j == 1)
                if fmt and j > 1 and isinstance(v, float):
                    x.number_format = fmt
        r += 1

    sec("1 · Market weights by role group — % of the score each parameter carries", REFS["weights"] + " · " + REFS["split"])
    trow(["Role group", "Source"] + [n for _, n in PARAMS] + ["L/M/H at IRIS default"], head=True)
    calib = []
    for rg in ("A", "B-IC", "B-M", "B+", "C-IC", "C-M", "C+", "DE"):
        mw, src = market_weights(R, rg)
        split = EXP_SPLIT["AB"] if rg in ("A", "B-IC", "B-M", "B+") else EXP_SPLIT["CE"]
        for sname, v in src.items():
            c, au, ac, l_, e, bi = v
            trow([RG_NAME[rg], sname, c / 100, au / 100, ac / 100, l_ / 100, bi / 100, e * split[0] / 100, e * split[1] / 100, ""], fmt="0%")
        lmh = {k: to_lmh(mw[k]) for k in PKEYS}
        pts = {k: {"L": 1, "M": 3, "H": 9}[lmh[k]] for k in PKEYS}
        tot = sum(pts.values())
        trow([RG_NAME[rg], "Market average"] + [mw[k] / 100 for k in PKEYS] + [" · ".join(f"{k} {lmh[k]}" for k in PKEYS)], tint="FFF2CC", fmt="0%")
        trow([RG_NAME[rg], "What L/M/H gives · L1 M3 H9"] + [pts[k] / tot for k in PKEYS] + [""], tint="EAF2E6", fmt="0%")
        calib.append((RG_NAME[rg], statistics.mean(abs(pts[k] / tot * 100 - mw[k]) for k in PKEYS)))
    r += 1
    sec("2 · How Low / Medium / High were set", REFS["weights"])
    for t in (f"Weightage: a market weight of {WEIGHT_H:g}% or more is High, {WEIGHT_M:g}% to under {WEIGHT_H:g}% is Medium, under {WEIGHT_M:g}% is Low. "
              "Points L 1 · M 3 · H 9 turn the choices back into shares of the score.",
              "Average gap between the L/M/H shares and the market average, per role group: " +
              " · ".join(f"{n} {g:.1f} pts" for n, g in calib),
              "Assessment: Low = −1, Medium = 0, High = +1 level against what the role & band matrix describes. The weighted net of the "
              "seven moves the person one level from the HR proposed level when it reaches ±0.5, so a business head can move a "
              "person by one level through the parameters, or set any band with the override.",
              "Experience and Education are split from the single Experience & Qualification weight: Bands A–B and B+ 60 / 40, "
              "Bands C–E 80 / 20 — education counts less as proven work grows. This split is an IRIS rule, supported by the "
              "skills-first evidence in S58, S60 and S61.",
              "Type of work shifts ±5 points in Bands A–C and B+, as in the HR Super Master: technology +Complexity −Accountability; "
              "product +Autonomy −Complexity; delivery +Accountability −Complexity; support +Business Impact −Complexity; sales "
              "+Business Impact +Autonomy −Complexity −Accountability; corporate +Education −Complexity."):
        x = ws.cell(r, 1, t)
        x.font = Y.font(9)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        ws.row_dimensions[r].height = Y.est_height([(t, 160)])
        r += 1
    r += 1
    sec("3 · Level crosswalk — market codes, experience, education and span by level", REFS["levels"] + " · " + REFS["exp"] + " · " + REFS["span"])
    trow(["Level", "Radford / Mercer", "Pave", "SFIA 9", "Google", "Infosys", "Accenture India", "Big Four", "IRIS experience",
          "Market experience", "Education norm", "Span norm"], head=True)
    ch = R["cross_hdr"]
    gi = lambda row, name: row[ch.index(name)] if name in ch and row else ""
    for lv in reversed(V.LEVELS_PC):
        cr = R["cross"].get(lv, [])
        ex = R["exp"].get(lv, {})
        f4 = R["f4"].get(lv, {})
        trow([lv, gi(cr, "Radford / Mercer code"), gi(cr, "Pave level name"), gi(cr, "SFIA 9 level"), gi(cr, "Google"), gi(cr, "Infosys"),
              gi(cr, "Accenture India"), gi(cr, "Big Four — Deloitte"), f4.get("Experience", ""), ex.get("Market range") or "—",
              f4.get("Education", ""), f4.get("Span of control", "")])
    r += 1
    sec("4 · Spans of control", REFS["span"])
    for row in R["span"]:
        trow([x for x in row if x != ""], head=row is R["span"][0])
    r += 1
    sec("5 · The public point-factor plans", REFS["weights"])
    for k, row in enumerate(R["plans"]):
        trow([x for x in row if x != ""], head=k == 0)
    for colx, w in zip("ABCDEFGHIJKL", (26, 22, 12, 12, 12, 12, 12, 12, 14, 14, 28, 26)):
        ws.column_dimensions[colx].width = w

    # ---- How to use
    ws = ws_use
    Y.banner(ws, "Business-Head Workbook — how to use it", "Read this first · five minutes.", J.STAMP, back=False, width_cols=8)
    moves = collections.Counter(("down" if V.VAL.get(pr[i]["level"], 0) < V.VAL.get(org.level[i], 0) else "up")
                                for i in pr if pr[i]["level"] != org.level[i] and V.band(pr[i]["level"]) != "X")
    steps = [
        ("What this is", "Every person at IRIS, top down from the CEO, with the band HR proposes and the evidence behind it. You review "
                         "it, weigh what matters for each role, assess each person, and the final proposed band updates itself."),
        ("1 · Read the row", "Core details, then what the employee and their manager told us. 'All their data ▸' opens everything we hold "
                             "on the person in 3. Raw Data; '◂ Back' returns."),
        ("2 · Weightage · L / M / H", "How much each of the seven parameters counts for the role: Complexity & Expertise, Autonomy & "
                                      "Decision, Accountability & Impact, Leadership & Strategy, Business Impact, Experience, Education. "
                                      "It comes from 5. Department Ladder, which starts from the family's RB sheet. Change it there for a "
                                      "whole role, or type L, M or H on a person's row for that person only."),
        ("3 · Your assessment · L / M / H", "Against what the RB sheet describes for the HR proposed band: Medium = meets it, Low = works "
                                            "at the level below on that parameter, High = already works at the level above. Everything "
                                            "starts at Medium, so the final band equals HR's until you change something. 'What the data "
                                            "shows' gives our reading of the evidence to start from."),
        ("4 · The band moves live", "Weighted net = Σ weightage × assessment ÷ Σ weightage, from −1 to +1. At +0.5 or more the person "
                                    "moves one level up; at −0.5 or less one level down. Use the override for any other band, and say why "
                                    "in your comment."),
        ("5 · The manager check", "Shown live on every row: under a head at C1 or above a reportee sits in a lower band; below C1 a "
                                  "reportee sits at least one level below the manager; CEO with CXOs is the agreed exception. A ⚠ means "
                                  "your change breaks it — adjust the person or their manager."),
        ("Market benchmark", "Each row shows the market weight of each parameter for the role group, the market experience range, the "
                             "education and span norms and the market codes — sheet 4 has every source."),
        ("HR proposed band", f"From the model · six ratings × market-derived weights and rules R1–R13 · then the top-down manager rule "
                             f"HR agreed on 12 October 2026: {moves['down']} people move down, pay and title protected for one cycle, "
                             f"{moves['up']} move up. Department and function heads sit in Executive Leadership · Business Head."),
        ("Input from benchmarking", "The same six ratings on the plain market-average weights, without IRIS's adjustments, rounded to a "
                                    "level, with the person's experience against the market range for it."),
        ("Colours", "Yellow = yours to edit · red = something missing · band colours on levels."),
        ("Privacy", "HR and business heads only. Survey content is shown against names by HR's decision. Pay is not in this file."),
    ]
    r = 5
    for a, b in steps:
        x = ws.cell(r, 1, a)
        x.font = Y.font(10, True, Y.NAVY)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        y = ws.cell(r, 2, b)
        y.font = Y.font(10)
        y.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = Y.est_height([(b, 120)], base=14)
        r += 1
    r += 1
    Y.header(ws.cell(r, 1, "Sheets"), Y.NAVY, 10)
    r += 1
    for nm in wb.sheetnames:
        x = ws.cell(r, 1, nm)
        x.hyperlink = f"#{Q(nm)}!A1"
        x.font = Y.font(9, False, "0563C1", underline="single")
        r += 1
    ws.column_dimensions["A"].width = 30
    for c_ in "BCDEFGH":
        ws.column_dimensions[c_].width = 18

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTFILE)
    work = os.path.join(os.environ.get("JA_WORK", "/tmp"), "ja_bh")
    os.makedirs(work, exist_ok=True)
    n_f, errors = X.cache_values(OUTFILE, work)
    print(f"Wrote {OUTFILE}: people {len(order)} · department ladder roles {len(dl_sorted)} · family sheets {len(families)} · "
          f"formulas {n_f} · formula errors {len(errors)}")
    if errors:
        print(errors[:20])


if __name__ == "__main__":
    main(sys.argv[1])
