"""IRIS Job Architecture Model v2.0 — rebuilt from scratch on HR's brief of 12 October 2026.

The story the workbook follows, sheet by sheet:
  1 How it works     — survey → ladder → placement → market → leaders, in plain words
  2 Ladder           — the levels, each with its own non-overlapping experience range, scope and span
  3 Role Model       — role-based: one block per role (sub-family × IC / Manager), its levels one below
                       the other, what Low / Medium / High mean at each level, and the role's weightage
                       IRIS → Market → Leader. A leader's weightage here flows to everyone in the role.
  4 Market Benchmark — Low / Medium / High for every parameter, with the published evidence behind it
  5 Talent Mapping   — every person, top down: who they are, what they told us, IRIS placement,
                       market, weightage IRIS → Market → Leader, ratings IRIS → Leader, and the
                       Leader proposed level and band, which stay blank until a leader gives an input
  6 Person Card      — one person on one page
  7 Raw Data         — survey, performance, KRAs and core HR data only, one row per person
  8 Backend Data     — the lookups

Placement: the v1.13 evidence level, then the fewest moves that keep every reportee in a lower band than
their manager · moving up costs twice moving down; Band E with Band E is the agreed exception; where a
reporting chain is deeper than the bands, the pair is flagged for a structure decision, never forced.

Usage: python3 build_model.py <folder with the five talent-mapping source files>
"""
import collections
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402
import ja_person as P  # noqa: E402
import ja_style as Y  # noqa: E402
import ja_v114 as V  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_bh import (PARAMS, PKEYS, LMH, REFS, RG_NAME, TINT, WEIGHT_H, WEIGHT_M, cell, default_weights, headers,  # noqa: E402
                      load_reference, name_range, qual_rank, rg_of, short, to_lmh)
from build_super import person_values  # noqa: E402

VERSION = "v2.0"
OUTDIR = os.path.join(J.ROOT, "IRIS_JA_Model_v2.0")
OUTFILE = os.path.join(OUTDIR, "IRIS_JA_Model_v2.0.xlsx")
STAMP = "IRIS · Job Architecture Model v2.0 · October 2026 · Internal — Confidential · HR and business leaders"
S_HOW, S_LAD, S_RM, S_MB, S_TM, S_PC, S_RAW, S_BE = ("1. How it works", "2. Ladder", "3. Role Model", "4. Market Benchmark",
                                                     "5. Talent Mapping", "6. Person Card", "7. Raw Data", "Backend Data")
Q = lambda n: f"'{n}'"
BANDS = ["A", "B", "B+", "C", "C+", "D", "E"]
# Distinct, non-overlapping experience ranges — years: from (inclusive) to (exclusive)
EXP_RANGE = {"A1": (0, 1, "Intern / trainee · first year"), "A2": (0, 2, "0–2 yrs"), "A3": (2, 4, "2–4 yrs"), "B1": (4, 6, "4–6 yrs"),
             "B2": (6, 8, "6–8 yrs"), "B3": (8, 10, "8–10 yrs"), "B+": (8, 12, "8–12 yrs · Player-Coach, set by the role, not by years"),
             "C1": (10, 13, "10–13 yrs"), "C2": (13, 16, "13–16 yrs"), "C+": (13, 17, "13–17 yrs · Player-Coach, set by the role, not by years"),
             "D1": (16, 19, "16–19 yrs"), "D2": (19, 22, "19–22 yrs"), "E1": (22, 99, "22+ yrs"), "E2": (22, 99, "22+ yrs · board appointment")}
FACT_COLS = {"C": ("Complexity rating", "Complexity & Expertise — justification"),
             "Au": ("Autonomy rating", "Autonomy & Decision — justification"),
             "Ac": ("Accountab. rating", "Accountability & Impact — justification"),
             "L": ("Leadership rating", "Leadership & Strategy — justification"),
             "BI": ("Bus. Impact rating", "Business Impact — justification")}


# ================================================================ placement: fewest moves, no shared band
def place(org, pen=6.0, up_cost=2.0, span=2):
    sys.setrecursionlimit(10000)

    def options(i):
        lv = org.level[i]
        b = V.band(lv)
        if b == "X" or not b:
            return {lv: 0.0}
        k = BANDS.index(b)
        out = {}
        for x in V.LEVELS_PC:
            nb = V.band(x)
            if nb in ("B+", "C+") and nb != b:
                continue
            if (nb == "E") != (b == "E"):
                continue
            if abs(BANDS.index(nb) - k) > span:
                continue
            d = V.VAL[x] - V.VAL[lv]
            out[x] = d * up_cost if d > 0 else -d
        return out

    def rel(cl, pl):
        cb, pb = V.band(cl), V.band(pl)
        if "X" in (cb, pb) or not cb or not pb or (cb == "E" and pb == "E"):
            return 0.0
        if V.VAL[cl] >= V.VAL[pl]:
            return None
        return pen if cb == pb else 0.0

    memo = {}

    def solve(i):
        if i in memo:
            return memo[i]
        res = {}
        for lv, c in options(i).items():
            tot, ok = c, True
            for k in org.kids[i]:
                best = None
                for kl, kc in solve(k).items():
                    r = rel(kl, lv)
                    if r is not None and (best is None or kc + r < best):
                        best = kc + r
                if best is None:
                    ok = False
                    break
                tot += best
            if ok:
                res[lv] = tot
        memo[i] = res
        return res

    new, flag = {}, {}

    def assign(i, lv):
        new[i] = lv
        for k in org.kids[i]:
            best = None
            for kl, kc in solve(k).items():
                r = rel(kl, lv)
                if r is not None and (best is None or kc + r < best[0]):
                    best = (kc + r, kl, r)
            if best[2] > 0:
                flag[k] = i
            assign(k, best[1])
    for r in org.roots:
        sr = solve(r)
        assign(r, min(sr, key=sr.get))
    return new, flag


# ================================================================ helpers
def track_group(track):
    t = J.text(track)
    if t == "Consultant":
        return "Consultant"
    return "Manager" if t in ("People Manager", "Player-Coach") else "IC"


def has_evidence(k, just):
    t = J.text(just)
    if k == "BI":
        return "kept at the start point" not in t and "No survey or manager-mapping signal" not in t
    return "Evidence reads like" in t


def evidence_line(k, just):
    for l in J.text(just).split("\n"):
        if l.strip().startswith("Evidence reads like"):
            return re.sub(r"^Evidence reads like ", "", l.strip())
    return ""


def split_title(des):
    parts = [x.strip() for x in J.text(des).split(" – ")]
    parts = [x for x in parts if x and x != "POD Lead"]
    return (parts[0] if parts else ""), " – ".join(parts[1:])


LEAD_HEAD = {"C1": "Associate Head", "C2": "Head", "C+": "Group Head", "D1": "Associate Vice President", "D2": "Senior Vice President"}


def std_title(rr, lv, des, track, model_lv=None):
    """(standard title for this role and level, what changes for the person). Moves down keep their title for one cycle.
    Leadership titles are function-specific: they change only when the level moves up, to the level's title."""
    des = J.text(des)
    up = model_lv is not None and V.VAL.get(lv, 0) > V.VAL.get(model_lv, 0)
    down = model_lv is not None and V.VAL.get(lv, 0) < V.VAL.get(model_lv, 0)
    head, qual = split_title(des)
    if rr["fam"] in (V.HEAD_FAMILY, V.ENT_FAMILY):
        sh = LEAD_HEAD.get(lv, head) if up else head
    else:
        sh = rr["std_head"].get(lv, head)
    if qual and (qual.lower() in sh.lower() or " – " in sh):
        qual = ""   # the level's built title already names the work
    std = sh + (f" – {qual}" if qual else "")
    if J.text(track) == "POD Lead":
        std += " – POD Lead"
    if std == des:
        return std, ""
    if down:
        return std, f"Keeps {des} for one cycle · then {std}"
    return std, f"{des} → {std}"


def rate(diff):
    return "H" if diff >= 0.75 else ("L" if diff <= -0.75 else "M")


def exp_lmh(years, level):
    lo, hi, _ = EXP_RANGE.get(level, (None, None, ""))
    if lo is None or years is None:
        return ""
    return "L" if years < lo else ("H" if years >= hi else "M")


def put_row(ws, r, cols, vals, tint_by_title=True):
    prev = None
    out = {}
    for c_, h, dark, title in cols:
        v = vals.get(h, "")
        x = cell(ws, r, c_, v, TINT.get(dark, "FFFFFF"), first=title != prev)
        prev = title
        out[h] = x
    return out


def descriptors(R, fam, lv, cats, cj, work):
    """Low / Medium / High at this level for each parameter: Medium = the level as defined; Low = the level below; High = the level above."""
    f4 = R["f4"].get(lv, {})
    g = R["grid"].get(lv, {})
    pos = V.LEVELS_PC.index(lv)
    if lv in ("B+", "C+"):
        below, above = V.LEVELS_PC[pos - 1], V.LEVELS_PC[pos + 1]
    else:
        below = next((x for x in reversed(V.LEVELS_PC[:pos]) if x not in ("B+", "C+")), None)
        above = next((x for x in V.LEVELS_PC[pos + 1:] if x not in ("B+", "C+")), None)
    theme = lambda col: next((short(J.text(c[cj[col]]), 140) for c in cats if J.has(c[cj[col]])), "")
    out = {}
    for pk, _ in PARAMS:
        if pk == "Exp":
            rng = EXP_RANGE[lv][2]
            out[pk] = (f"Less than {rng}", f"{rng} · market {R['exp'].get(lv, {}).get('Market range') or '—'} yrs", f"More than {rng}")
        elif pk == "Edu":
            md = f4.get("Education", "")
            if work == "Corporate functions":
                md += " · licence-to-practise roles: CA, CS, CMA or LLB as the role needs"
            out[pk] = ("Below: " + R["f4"].get(below, {}).get("Education", "") if below else "Still studying",
                       md, "Above: " + R["f4"].get(above, {}).get("Education", "") if above else "Credentials beyond the role")
        else:
            md = g.get(pk, "")
            if pk == "L" and V.band(lv) == "A":
                md = "Helps peers and newer people informally · " + f4.get("Span of control", "")
            if pk == "C" and theme("Skills & knowledge · counts"):
                md += f" · in this role: {theme('Skills & knowledge · counts')}"
            if pk == "Ac" and theme("Manager-set KRAs · common themes"):
                md += f" · KRA themes: {theme('Manager-set KRAs · common themes')}"
            lo = ("Works at " + below + ": " + R["grid"].get(below, {}).get(pk, "")) if below else "Not yet at the entry standard"
            hi = ("Works at " + above + ": " + R["grid"].get(above, {}).get(pk, "")) if above else "Beyond the role · a board mandate"
            if pk == "L" and below and V.band(below) == "A":
                lo = "Works at " + below + ": contributes alone; guides others only informally"
            out[pk] = (lo, md, hi)
    return out


# ================================================================ main
def main(srcdir):
    os.makedirs(OUTDIR, exist_ok=True)
    M, S, people, unmapped = J.build(srcdir)
    F_prev = P.Folds(M, people)
    org = V.Org(M, people, F_prev)
    heads = V.department_heads(M, people, org)
    changed = V.apply_heads(M, people, F_prev, heads)
    F = P.Folds(M, people)
    R = load_reference(J.MASTER)
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    new, flag = place(org)
    order = [i for i, d in org.top_down()]
    depth = dict(org.top_down())
    by_i = {p["i"]: p for p in people}
    cj = {h: j for j, h in enumerate(M["cat_hdr"])}

    # ---- roles: sub-family × IC / Manager
    roles = collections.OrderedDict()
    for i in order:
        p = by_i[i]
        fam, sub = J.text(p["tm"][H["Job family"]]), J.text(p["tm"][H["Sub-family · what they do"]])
        tg = track_group(p["tm"][H["Track"]])
        key = (fam, sub, tg)
        roles.setdefault(key, {"people": [], "fam": fam, "sub": sub, "tg": tg})
        roles[key]["people"].append(i)
        p["role"] = key
    role_keys = sorted(roles, key=lambda k: (k[0], k[1], k[2]))
    for n, k in enumerate(role_keys, 1):
        roles[k]["id"] = f"R{n:03d}"
        roles[k]["name"] = f"{k[1]} · {k[2]}"
    # standard title per role × level: the title before any dash is aligned to what most holders at that level use;
    # the person's own function after the dash, and a POD Lead marker, are kept
    for k in role_keys:
        rr = roles[k]
        at = collections.defaultdict(collections.Counter)
        for i in rr["people"]:
            head, _ = split_title(J.text(by_i[i]["tm"][H["Final designation"]]))
            if new[i] == org.level[i]:
                at[new[i]][head] += 1   # only people who stay at their level set the title for it
        mgr = k[2] == "Manager"
        rr["std_head"], used = {}, set()
        for lv in sorted({new[i] for i in rr["people"]}, key=lambda x: -V.VAL.get(x, 0)):
            pick_ = next((t for t, _ in at[lv].most_common() if t not in used), None)
            if pick_ is None:
                base = V.LEVEL_TITLE.get(lv, (lv, lv))[1 if mgr else 0]
                pick_ = f"{base} – {k[1]}" if base not in ("—", "") else k[1]
            rr["std_head"][lv] = pick_
            used.add(pick_)
        rr["std"] = dict(rr["std_head"])
        levels = [new[i] for i in rr["people"] if V.band(new[i]) != "X"]
        if levels:
            lo, hi = min(V.VAL[x] for x in levels), max(V.VAL[x] for x in levels)
            pc = any(x in ("B+", "C+") for x in levels)
            lad = V.LEVELS_PC if pc else V.LEVELS_STD
            rr["levels"] = [x for x in reversed(lad) if lo <= V.VAL[x] <= hi]
            main = collections.Counter(V.band(x) for x in levels).most_common(1)[0][0]
            main_lv = collections.Counter(x for x in levels if V.band(x) == main).most_common(1)[0][0]
            rr["rg"] = rg_of(main_lv, "People Manager" if k[2] == "Manager" else "IC")
        else:
            rr["levels"] = sorted({new[i] for i in rr["people"]})
            rr["rg"] = None
        works = collections.Counter(R["dept_type"].get((by_i[i]["biz"], J.text(by_i[i]["tm"][H["Department"]]))) for i in rr["people"])
        works.pop(None, None)
        rr["work"] = works.most_common(1)[0][0] if works else "General management"
        if rr["rg"]:
            dw = default_weights(R, rr["rg"], rr["work"] if rr["rg"] not in ("C+", "DE") else "General management")
            rr["iris_w"] = {pk: to_lmh(dw[pk]) for pk in PKEYS}
            src = {s: v for s, v in R["weights"][rr["rg"]].items() if "IRIS" not in s}
            # the market as published: experience and education are part of the knowledge factor, never weighted on their own
            avg = [sum(v[j] for v in src.values()) / len(src) for j in range(6)]
            c, au, ac, l_, e, bi = avg
            tot = c + e + au + ac + l_ + bi
            mk = {"C": (c + e) * 100 / tot, "Au": au * 100 / tot, "Ac": ac * 100 / tot, "L": l_ * 100 / tot, "BI": bi * 100 / tot}
            rr["mkt_pct"] = mk
            rr["mkt_w"] = {pk: to_lmh(mk[pk]) for pk in ("C", "Au", "Ac", "L", "BI")}
            rr["mkt_w"].update({"Exp": "L", "Edu": "L"})
        else:
            rr["iris_w"] = {pk: "" for pk in PKEYS}
            rr["mkt_w"] = {pk: "" for pk in PKEYS}
            rr["mkt_pct"] = {}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_how, ws_lad, ws_rm, ws_mb, ws_tm, ws_pc, ws_raw, ws_be = (wb.create_sheet(n) for n in (S_HOW, S_LAD, S_RM, S_MB, S_TM, S_PC, S_RAW, S_BE))

    # ================================================================ Backend Data
    ws = ws_be
    Y.banner(ws, "Backend Data — the lookups", "Last sheet · what L, M and H are worth, the ladders, the band order and the drop-down lists. "
             "Change the yellow cells only if HR agrees to change the model.", STAMP, back=False, width_cols=8)
    r = 5
    for j, h in enumerate(["Code", "Means", "Weightage points", "Rating step", "Explained"], 1):
        Y.header(ws.cell(r, j, h), Y.NAVY)
    r += 1
    top = r
    for code, nm, pts, step, expl in LMH:
        for j, v in enumerate((code, nm, pts, step, expl), 1):
            cell(ws, r, j, v, "FFF2CC" if j in (3, 4) else "FFFFFF", j == 1)
        r += 1
    name_range(wb, "LMH_KEY", S_BE, f"$A${top}:$A${r - 1}")
    name_range(wb, "LMH_PTS", S_BE, f"$C${top}:$C${r - 1}")
    name_range(wb, "LMH_STEP", S_BE, f"$D${top}:$D${r - 1}")
    name_range(wb, "LIST_LMH", S_BE, f"$A${top}:$A${r - 1}")
    r += 1
    for j, h in enumerate(["Level", "Value", "Band", "Band rank", "Standard ladder", "Player-Coach ladder"], 1):
        Y.header(ws.cell(r, j, h), Y.NAVY)
    r += 1
    top = r
    for k, lv in enumerate(V.LEVELS_PC + ["X1", "X2", "X3"]):
        b = V.band(lv)
        vals = (lv, V.VAL.get(lv, 0), b, BANDS.index(b) + 1 if b in BANDS else 0,
                V.LEVELS_STD[k] if k < len(V.LEVELS_STD) else None, V.LEVELS_PC[k] if k < len(V.LEVELS_PC) else None)
        for j, v in enumerate(vals, 1):
            cell(ws, r, j, v, "FFFFFF", j == 1)
        r += 1
    name_range(wb, "B_LVL", S_BE, f"$A${top}:$A${r - 1}")
    name_range(wb, "B_VAL", S_BE, f"$B${top}:$B${r - 1}")
    name_range(wb, "B_RANK", S_BE, f"$D${top}:$D${r - 1}")
    name_range(wb, "LAD_STD", S_BE, f"$E${top}:$E${top + 11}")
    name_range(wb, "LAD_PC", S_BE, f"$F${top}:$F${top + 13}")
    name_range(wb, "LIST_LEVELS", S_BE, f"$A${top}:$A${r - 1}")
    r += 1
    Y.header(ws.cell(r, 1, "People · for the Person Card"), Y.NAVY)
    r += 1
    people_top = r
    for i in order:
        p = by_i[i]
        cell(ws, r, 1, f"{p['name']} · {p['emp']}", "FFFFFF")
        cell(ws, r, 2, p["emp"], "FFFFFF")
        r += 1
    name_range(wb, "LIST_PEOPLE", S_BE, f"$A${people_top}:$A${r - 1}")
    name_range(wb, "PEOPLE_ID", S_BE, f"$B${people_top}:$B${r - 1}")
    ws.column_dimensions["A"].width = 40
    for c_ in "BCDEF":
        ws.column_dimensions[c_].width = 14
    ws.column_dimensions["E"].width = 60

    # ================================================================ Ladder
    ws = ws_lad
    Y.banner(ws, "Ladder — one level, one place",
             "Sheet 2 · Every level has its own experience range, scope and span — the ranges do not overlap, so each person fits one "
             "level by the evidence. Player-Coach steps B+ and C+ are set by the role, not by years. Titles follow HR's decisions.", STAMP, width_cols=12)
    cols = headers(ws, [("LEVEL", "1F3864", [("Level", 7), ("Band", 22), ("IC title", 24), ("Manager title", 26)]),
                        ("WHAT DEFINES IT — distinct by level", "375623", [("Experience · distinct range", 22), ("What the role owns", 34),
                                                                            ("Span of control", 26), ("Education", 30)]),
                        ("MARKET", "3A3838", [("Market experience · yrs", 14), ("Radford / Mercer", 14), ("SFIA 9", 12), ("References", 34)]),
                        ("PEOPLE", "595959", [("People · IRIS proposed", 10)])])
    r = 7
    cnt = collections.Counter(new.values())
    for lv in reversed(V.LEVELS_PC):
        f4 = R["f4"].get(lv, {})
        cr = R["cross"].get(lv, [])
        vals = {"Level": lv, "Band": V.BAND_NAME.get(V.band(lv), ""), "IC title": V.LEVEL_TITLE[lv][0], "Manager title": V.LEVEL_TITLE[lv][1],
                "Experience · distinct range": EXP_RANGE[lv][2], "What the role owns": f4.get("What the role owns", ""),
                "Span of control": f4.get("Span of control", ""), "Education": f4.get("Education", ""),
                "Market experience · yrs": R["exp"].get(lv, {}).get("Market range") or "—",
                "Radford / Mercer": cr[1] if len(cr) > 1 else "", "SFIA 9": cr[3] if len(cr) > 3 else "",
                "References": f"{REFS['levels']} · {REFS['exp']} · {REFS['span']}", "People · IRIS proposed": cnt.get(lv, 0)}
        xs = put_row(ws, r, cols, vals)
        b = Y.band_of(lv)
        if b:
            xs["Level"].fill = PatternFill("solid", fgColor=Y.BAND[b])
        xs["Level"].font = Y.font(11, True)
        ws.row_dimensions[r].height = 48
        r += 1
    ws.freeze_panes = "B7"

    # ================================================================ Role Model
    ws = ws_rm
    Y.banner(ws, "Role Model — every role, its levels one below the other",
             "Sheet 3 · Role-based: a role is a sub-family done as an individual contributor or as a manager. The ROLE row carries the "
             "role's weightage: IRIS · Market · Leader. Type L, M or H in the yellow Leader cells and everyone in the role is re-weighted "
             "on Talent Mapping at once. The LEVEL rows below say what Low, Medium and High mean at each level. Filter column A or D by role.",
             STAMP, width_cols=14)
    Y.colour_key(ws, 4)
    groups = [("ROLE", "1F3864", [("Role ID", 8), ("Role", 34), ("Job family", 22), ("Track", 10), ("Row", 7), ("Level", 7),
                                  ("Standard title at this level", 30), ("People", 8)]),
              ("WEIGHTAGE · IRIS — what we set", "806000", [(f"IRIS · {n}", 9) for _, n in PARAMS]),
              ("WEIGHTAGE · MARKET — published evidence, sheet 4", "3A3838", [(f"Market · {n}", 9) for _, n in PARAMS]),
              ("WEIGHTAGE · LEADER — your update, blank = agree with IRIS", "BF8F00", [(f"Leader · {n}", 9) for _, n in PARAMS]),
              ("IN USE · leader if given, else IRIS", "595959", [(f"In use · {n}", 8) for _, n in PARAMS] + [("Leader updates", 8)])]
    for pk, pn in PARAMS:
        groups.append((f"{pn.upper()} — what L / M / H mean at this level", "375623" if pk in ("L", "BI") else ("2E75B6" if pk in ("Exp", "Edu") else "843C0C"),
                       [(f"{pn} · Low", 26), (f"{pn} · Medium", 30), (f"{pn} · High", 26)]))
    cols = headers(ws, groups)
    col = {h: L(c_) for c_, h, _, _ in cols}
    r = 7
    rm_top = r
    cat_by = collections.defaultdict(list)
    for c in M["cat"]:
        cat_by[(J.text(c[1]), J.text(c[2]), J.text(c[3]))].append(c)
    dv_rm = DataValidation(type="list", formula1="LIST_LMH", allow_blank=True)
    ws.add_data_validation(dv_rm)
    for k in role_keys:
        rr = roles[k]
        rr["row"] = r
        vals = {"Role ID": rr["id"], "Role": rr["name"], "Job family": rr["fam"], "Track": rr["tg"], "Row": "ROLE",
                "Level": " · ".join(reversed(rr["levels"])) if rr["levels"] else "", "People": len(rr["people"]),
                "Standard title at this level": f"{RG_NAME.get(rr['rg'], 'Consultant · outside the ladder')} · {rr['work']}"}
        for pk, pn in PARAMS:
            vals[f"IRIS · {pn}"] = rr["iris_w"][pk]
            vals[f"Market · {pn}"] = rr["mkt_w"][pk]
            lc, ic = col[f"Leader · {pn}"], col[f"IRIS · {pn}"]
            vals[f"In use · {pn}"] = f'=IF({lc}{r}<>"",{lc}{r},{ic}{r})'
        vals["Leader updates"] = f'=COUNTA({col["Leader · " + PARAMS[0][1]]}{r}:{col["Leader · " + PARAMS[-1][1]]}{r})'
        xs = put_row(ws, r, cols, vals)
        for c_, h, dark, title in cols:
            x = ws.cell(r, c_)
            x.font = Y.font(9, True)
            if title == "ROLE":
                x.fill = PatternFill("solid", fgColor="D9E1F2")
            if title.startswith(("WEIGHTAGE", "IN USE")):
                x.alignment = Alignment(horizontal="center", vertical="top")
            if title.startswith("WEIGHTAGE · LEADER"):
                x.fill = Y.F_YEL
                dv_rm.add(f"{L(c_)}{r}")
        ws.row_dimensions[r].height = 30
        r += 1
        for lv in rr["levels"]:
            if V.band(lv) == "X":
                continue
            cats = cat_by.get((rr["fam"], rr["sub"], lv), [])
            d = descriptors(R, rr["fam"], lv, cats, cj, rr["work"])
            held = sum(1 for i in rr["people"] if new[i] == lv)
            vals = {"Role ID": rr["id"], "Role": rr["name"], "Job family": rr["fam"], "Track": rr["tg"], "Row": "level", "Level": lv,
                    "Standard title at this level": rr["std"].get(lv, "— no one at this level today"), "People": held}
            for pk, pn in PARAMS:
                vals[f"{pn} · Low"], vals[f"{pn} · Medium"], vals[f"{pn} · High"] = d[pk]
            xs = put_row(ws, r, cols, vals)
            b = Y.band_of(lv)
            if b:
                xs["Level"].fill = PatternFill("solid", fgColor=Y.BAND[b])
            ws.row_dimensions[r].height = 90
            r += 1
    rm_end = r - 1
    name_range(wb, "RM_KEY", S_RM, f"$A${rm_top}:$A${rm_end}")
    for pk, pn in PARAMS:
        for pre, tag in (("IRIS", "I"), ("Market", "M"), ("Leader", "L"), ("In use", "E")):
            name_range(wb, f"RM_{tag}_{pk}", S_RM, f"${col[f'{pre} · {pn}']}${rm_top}:${col[f'{pre} · {pn}']}${rm_end}")
    name_range(wb, "RM_LCNT", S_RM, f"${col['Leader updates']}${rm_top}:${col['Leader updates']}${rm_end}")
    ws.freeze_panes = "C7"
    ws.auto_filter.ref = f"A6:{L(len(cols))}{rm_end}"
    ws.sheet_view.zoomScale = 85

    # ================================================================ Market Benchmark
    ws = ws_mb
    Y.banner(ws, "Market Benchmark — Low / Medium / High, with the evidence behind each",
             "Sheet 4 · How much each parameter counts in published job-evaluation methods, by role group, turned into Low / Medium / High. "
             f"Rule: a share of {WEIGHT_H:g}% or more of the points is High; {WEIGHT_M:g}% to under {WEIGHT_H:g}% is Medium; under {WEIGHT_M:g}% "
             "is Low. Each row shows the published numbers it rests on. Experience and Education: every method folds them into knowledge, "
             "so the market weights them Low on their own.", STAMP, width_cols=12)
    r = 5
    Y.header(ws.cell(r, 1, "1 · Weightage by role group and parameter"), Y.NAVY, 11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    hdr = ["Role group", "Parameter", "Market · L/M/H", "Korn Ferry Hay", "Mercer IPE", "Public plans · average of 5", "Market average",
           "Why this rating", "References"]
    for j, h in enumerate(hdr, 1):
        Y.header(ws.cell(r, j, h), "3A3838")
    r += 1
    fac_idx = {"C": 0, "Au": 1, "Ac": 2, "L": 3, "BI": 5}
    for rg in ("A", "B-IC", "B-M", "B+", "C-IC", "C-M", "C+", "DE"):
        src = {s: v for s, v in R["weights"][rg].items() if "IRIS" not in s}
        def share(v, pk):
            tot = sum(v)
            if pk == "C":
                return (v[0] + v[4]) * 100 / tot
            return v[fac_idx[pk]] * 100 / tot
        for pk, pn in PARAMS:
            if pk in ("Exp", "Edu"):
                vals = [RG_NAME[rg], pn, "L", "in know-how", "in knowledge", "in knowledge / KTE", "not separate",
                        "No published method weights it on its own: Hay know-how includes all knowledge and experience however acquired; "
                        "Mercer IPE knowledge; NHS knowledge, training and experience. Skills-first evidence argues for a low weight.",
                        "S05, S49 Hay · S04, S52 Mercer IPE · S53 NHS · S58 McKinsey · S59 Deloitte · S60 WEF · S61 EY India"]
            else:
                per = {s: share(v, pk) for s, v in src.items()}
                avg = sum(per.values()) / len(per)
                lab = to_lmh(avg)
                why = (f"Average {avg:.0f}% of the points · {'at or above' if lab == 'H' else 'below'} {WEIGHT_H:g}% "
                       + ("→ High" if lab == "H" else (f"and at or above {WEIGHT_M:g}% → Medium" if lab == "M" else f"and below {WEIGHT_M:g}% → Low")))
                if pk == "C":
                    why += " · Complexity here includes knowledge and experience, as the methods publish it"
                vals = [RG_NAME[rg], pn, lab] + [f"{per.get(s, 0):.0f}%" for s in ("Korn Ferry Hay", "Mercer IPE", "Public plans")] + \
                       [f"{avg:.0f}%", why, REFS["weights"]]
            for j, v in enumerate(vals, 1):
                x = cell(ws, r, j, v, "FFF2CC" if j == 3 else "FFFFFF", j == 1)
                if j == 3:
                    x.font = Y.font(11, True)
                    x.alignment = Alignment(horizontal="center", vertical="top")
            ws.row_dimensions[r].height = 42
            r += 1
    r += 1
    Y.header(ws.cell(r, 1, "2 · The public point-factor plans — published factor shares"), Y.NAVY, 11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for k, row in enumerate(R["plans"]):
        vals = [x for x in row if x != ""]
        for j, v in enumerate(vals, 1):
            if k == 0:
                Y.header(ws.cell(r, j, v), "595959")
            else:
                cell(ws, r, j, v, "FFFFFF", j == 1)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    Y.header(ws.cell(r, 1, "3 · Level benchmarks — the market data the ratings are checked against"), Y.NAVY, 11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for j, h in enumerate(["Level", "IRIS experience · distinct", "Market experience", "Accenture India", "Infosys", "Microsoft / Amazon",
                           "Radford / Mercer", "SFIA 9", "Span norm · McKinsey", "References"], 1):
        Y.header(ws.cell(r, j, h), "595959")
    r += 1
    for lv in reversed(V.LEVELS_PC):
        ex = R["exp"].get(lv, {})
        cr = R["cross"].get(lv, [])
        vals = [lv, EXP_RANGE[lv][2], ex.get("Market range", "—"), ex.get("Accenture India", "—"), ex.get("Infosys", "—"),
                ex.get("Microsoft / Amazon", "—"), cr[1] if len(cr) > 1 else "", cr[3] if len(cr) > 3 else "",
                R["f4"].get(lv, {}).get("Span of control", ""), f"{REFS['exp']} · {REFS['levels']} · {REFS['span']}"]
        for j, v in enumerate(vals, 1):
            cell(ws, r, j, v, "FFFFFF", j == 1)
        ws.row_dimensions[r].height = 30
        r += 1
    for c_, w in zip("ABCDEFGHIJ", (26, 22, 12, 14, 14, 18, 14, 12, 50, 50)):
        ws.column_dimensions[c_].width = w

    # ================================================================ Talent Mapping
    ws = ws_tm
    Y.banner(ws, "Talent Mapping — every person, top down",
             "Sheet 5 · Read left to right: who they are → what they and their manager told us → where IRIS placed them → what the market "
             "says → how much each parameter counts · IRIS, Market, Leader → how they rate · IRIS, then yours → your proposed level and "
             "band. Blank = nothing received. Yellow = yours. Your columns stay blank until you give an input.", STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    groups = [
        ("WHO", "1F3864", [("Business", 10), ("Department", 18), ("Layer from the CEO", 7), ("Emp ID", 12), ("Employee name", 22),
                           ("Current designation", 24), ("Reporting manager", 20), ("Manager Emp ID", 12), ("Total exp. · yrs", 8),
                           ("Qualification", 16), ("Current level · as is", 9)]),
        ("WHAT THEY TOLD US — blank = not received", "7030A0", [("Survey · questions answered", 9), ("What they shared", 44), ("KRAs from the manager", 44),
                                                                  ("Performance · rating · potential · 9-box", 24), ("Evidence used for the ratings", 40),
                                                                  ("All their data ▸", 10), ("Person card ▸", 10)]),
        ("IRIS PLACEMENT", "C00000", [("Job family", 20), ("Sub-family", 22), ("Role ID", 8), ("Role", 30), ("IRIS proposed level", 9),
                                      ("IRIS proposed band", 20), ("Standard title for the role and level", 28), ("Title change", 20),
                                      ("How IRIS placed them", 40), ("Structure flag", 30)]),
        ("MARKET", "3A3838", [("Market level · published weights", 9), ("Experience vs the level · market", 30),
                              ("Span vs the norm", 22), ("Radford / Mercer · SFIA 9", 18)]),
        ("WEIGHTAGE · IRIS", "806000", [(f"Weightage IRIS · {n}", 9) for _, n in PARAMS]),
        ("WEIGHTAGE · MARKET", "3A3838", [(f"Weightage Market · {n}", 9) for _, n in PARAMS]),
        ("WEIGHTAGE · LEADER — set on 3. Role Model", "BF8F00", [(f"Weightage Leader · {n}", 9) for _, n in PARAMS]),
        ("RATING · IRIS — vs the level · blank = no input", "806000", [(f"Rating IRIS · {n}", 9) for _, n in PARAMS]),
        ("RATING · LEADER — yours · L / M / H vs the level", "BF8F00", [(f"Rating Leader · {n}", 9) for _, n in PARAMS]),
        ("LEADER PROPOSAL — blank until you give an input", "C00000", [("Leader override · any level", 10), ("Leader proposed level", 10),
                                                                       ("Leader proposed band", 20), ("Move vs IRIS", 12),
                                                                       ("Level in force · leader if given, else IRIS", 10),
                                                                       ("Manager check · live", 24), ("Leader comment", 30)]),
        ("ENGINE", "595959", [(f"pI {k}", 5) for k in PKEYS] + [(f"pE {k}", 5) for k in PKEYS] + [(f"dI {k}", 5) for k in PKEYS] +
         [(f"dE {k}", 5) for k in PKEYS] + [("Inputs", 5), ("Net IRIS", 6), ("Net leader", 6)]),
    ]
    cols = headers(ws, groups)
    col = {h: L(c_) for c_, h, _, _ in cols}
    last = 6 + len(order)
    E, INF = col["Emp ID"], col["Level in force · leader if given, else IRIS"]
    raw_row = {i: 7 + n for n, i in enumerate(order)}
    tm_row = {}
    for n, i in enumerate(order):
        r = 7 + n
        tm_row[i] = r
        p = by_i[i]
        tm = p["tm"]
        rr = roles[p["role"]]
        k, sv, links = person_values(M, F, p)
        s8 = F.s8row(p)
        perf = ""
        if s8 is not None:
            hh = M["s8_hdr"]
            perf = " · ".join(f"{h}: {J.text(s8[hh.index(h)])}" for h in ("Latest rating", "Potential", "9-box") if J.has(s8[hh.index(h)]))
        shared = sv[1] if J.text(sv[-1]).startswith("Per-person") else ""
        kras = "\n".join(f"{lab}: {v}" for lab, v in (("Annual · set by manager", k[0]), ("Quarterly", k[1]), ("Self-reported", k[2]))
                         if J.has(v) and not J.text(v).startswith(P.MISSING_KRA))
        ev = []
        iris_r = {}
        lv = new[i]
        for pk, (rc, jc) in FACT_COLS.items():
            if has_evidence(pk, tm[H[jc]]):
                line = evidence_line(pk, tm[H[jc]])
                if line:
                    ev.append(f"{dict(PARAMS)[pk]}: {line}")
                try:
                    iris_r[pk] = rate(float(tm[H[rc]]) - V.VAL[lv]) if V.band(lv) != "X" else ""
                except (TypeError, ValueError):
                    iris_r[pk] = ""
            else:
                iris_r[pk] = ""
        try:
            yrs = float(tm[H["Total exp. · yrs"]])
        except (TypeError, ValueError):
            yrs = None
        iris_r["Exp"] = exp_lmh(yrs, lv) if V.band(lv) != "X" else ""
        q = qual_rank(tm[H["Qualification"]])
        iris_r["Edu"] = ("" if q is None or V.band(lv) == "X" else ("L" if q < (1 if V.VAL[lv] < 6 else 2) else "H" if q > (1 if V.VAL[lv] < 6 else 2) else "M"))
        model = J.text(tm[H["FINAL LEVEL"]])
        if lv == model:
            how = f"Placed at {lv} by the evidence: six ratings × market-derived weights and rules R1–R13 · no move needed for the reporting line"
        else:
            how = (f"Evidence level {model} · moved {'down' if V.VAL.get(lv, 0) < V.VAL.get(model, 0) else 'up'} to {lv} so that no reportee shares "
                   f"a band with their manager, with the fewest moves across IRIS" + (" · pay and title protected for one cycle" if V.VAL.get(lv, 0) < V.VAL.get(model, 0) else ""))
        if p["i"] in changed:
            how += f" · {'Band E · Enterprise Leadership' if changed[p['i']][2] == V.ENT_FAMILY else 'department head · Executive Leadership · Business Head'}"
        m = org.mgr[i]
        sflag = (f"Deep chain: shares a band with manager {by_i[m]['name']} · {new[m]} · decide: report to the next manager up, or merge "
                 "this layer") if i in flag else ""
        des = J.text(tm[H["Final designation"]])
        std, title_change = std_title(rr, lv, des, tm[H["Track"]], J.text(tm[H["FINAL LEVEL"]]))
        # market
        mlev = ""
        if rr["rg"] and V.band(lv) != "X":
            try:
                rs = [float(tm[H[h]]) for h in ("Complexity rating", "Autonomy rating", "Accountab. rating", "Leadership rating",
                                                 "Exp + Qual rating", "Bus. Impact rating")]
                srcs = [v for s, v in R["weights"][rr["rg"]].items() if "IRIS" not in s]
                w6 = [sum(v[j] for v in srcs) / len(srcs) for j in range(6)]
                w6 = [w6[0], w6[1], w6[2], w6[3], w6[4], w6[5]]
                sc = sum(a * b for a, b in zip(rs, w6)) / sum(w6)
                mlev = V.LEVELS_STD[max(0, min(11, int(sc + 0.5) - 1))]
            except (TypeError, ValueError):
                mlev = ""
        mr = R["exp"].get(lv, {}).get("Market range", "")
        exp_txt = ""
        if yrs is not None and mr:
            nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", mr)]
            if nums:
                tag = "Low" if yrs < nums[0] else ("High" if len(nums) > 1 and yrs > nums[-1] else "Medium")
                exp_txt = f"{tag} · {yrs:g} yrs vs market {mr} yrs for {lv}"
        reps = len([k_ for k_ in org.kids[i] if V.band(new[k_]) != "X"])
        span_txt = ""
        if reps:
            span_txt = f"{reps} direct reports · norm {R['f4'].get(lv, {}).get('Span of control', '—')}"
        cr = R["cross"].get(lv, [])
        vals = {
            "Business": p["biz"], "Department": J.text(tm[H["Department"]]), "Layer from the CEO": depth[i], "Emp ID": p["emp"],
            "Employee name": p["name"], "Current designation": J.text(tm[H["Current designation"]]),
            "Reporting manager": J.text(tm[H["Reporting manager"]]), "Manager Emp ID": links[1] if links[1] not in ("—", "Not in this file") else "",
            "Total exp. · yrs": tm[H["Total exp. · yrs"]] if yrs is not None else "", "Qualification": J.text(tm[H["Qualification"]]) if q is not None else "",
            "Current level · as is": J.text(tm[H["Current level"]]) if "Not in current" not in J.text(tm[H["Current level"]]) else "",
            "Survey · questions answered": f"{sv[0]} of 20" if sv[0] else "", "What they shared": short(shared, 500) if shared else "",
            "KRAs from the manager": short(kras, 500) if kras else "", "Performance · rating · potential · 9-box": perf,
            "Evidence used for the ratings": "\n".join(ev), "All their data ▸": "Open ▸", "Person card ▸": "Card ▸",
            "Job family": rr["fam"], "Sub-family": rr["sub"], "Role ID": rr["id"], "Role": rr["name"], "IRIS proposed level": lv,
            "IRIS proposed band": V.BAND_NAME.get(V.band(lv), ""), "Standard title for the role and level": std, "Title change": title_change,
            "How IRIS placed them": how, "Structure flag": sflag, "Market level · published weights": mlev,
            "Experience vs the level · market": exp_txt, "Span vs the norm": span_txt,
            "Radford / Mercer · SFIA 9": f"{cr[1]} · SFIA {cr[3]}" if len(cr) > 3 else "",
        }
        rk = col["Role ID"]
        for pk, pn in PARAMS:
            vals[f"Weightage IRIS · {pn}"] = f'=IFERROR(IF(INDEX(RM_I_{pk},MATCH(${rk}{r},RM_KEY,0))="","",INDEX(RM_I_{pk},MATCH(${rk}{r},RM_KEY,0))),"")'
            vals[f"Weightage Market · {pn}"] = f'=IFERROR(IF(INDEX(RM_M_{pk},MATCH(${rk}{r},RM_KEY,0))="","",INDEX(RM_M_{pk},MATCH(${rk}{r},RM_KEY,0))),"")'
            vals[f"Weightage Leader · {pn}"] = f'=IFERROR(IF(INDEX(RM_L_{pk},MATCH(${rk}{r},RM_KEY,0))="","",INDEX(RM_L_{pk},MATCH(${rk}{r},RM_KEY,0))),"")'
            vals[f"Rating IRIS · {pn}"] = iris_r.get(pk, "")
            wi, wl = col[f"Weightage IRIS · {pn}"], col[f"Weightage Leader · {pn}"]
            ri, rl = col[f"Rating IRIS · {pn}"], col[f"Rating Leader · {pn}"]
            vals[f"pI {pk}"] = f'=IFERROR(INDEX(LMH_PTS,MATCH({wi}{r},LMH_KEY,0)),0)'
            vals[f"pE {pk}"] = f'=IFERROR(INDEX(LMH_PTS,MATCH(IF({wl}{r}<>"",{wl}{r},{wi}{r}),LMH_KEY,0)),0)'
            vals[f"dI {pk}"] = f'=IF({ri}{r}="",0,IFERROR(INDEX(LMH_STEP,MATCH({ri}{r},LMH_KEY,0)),0))'
            vals[f"dE {pk}"] = f'=IF({rl}{r}<>"",IFERROR(INDEX(LMH_STEP,MATCH({rl}{r},LMH_KEY,0)),0),{col["dI " + pk]}{r})'
        pi = [f"{col['pI ' + k_]}{r}" for k_ in PKEYS]
        pe = [f"{col['pE ' + k_]}{r}" for k_ in PKEYS]
        di = [f"{col['dI ' + k_]}{r}" for k_ in PKEYS]
        de = [f"{col['dE ' + k_]}{r}" for k_ in PKEYS]
        vals["Net IRIS"] = f'=IF(({"+".join(pi)})=0,0,({"+".join(a + "*" + b for a, b in zip(pi, di))})/({"+".join(pi)}))'
        vals["Net leader"] = f'=IF(({"+".join(pe)})=0,0,({"+".join(a + "*" + b for a, b in zip(pe, de))})/({"+".join(pe)}))'
        ov = col["Leader override · any level"]
        rl0, rl1 = col[f"Rating Leader · {PARAMS[0][1]}"], col[f"Rating Leader · {PARAMS[-1][1]}"]
        vals["Inputs"] = (f'=COUNTA({rl0}{r}:{rl1}{r})+IF({ov}{r}<>"",1,0)+IFERROR(INDEX(RM_LCNT,MATCH(${rk}{r},RM_KEY,0)),0)')
        iris_c = col["IRIS proposed level"]
        diff = f'({col["Net leader"]}{r}-{col["Net IRIS"]}{r})'
        step = f'IF({diff}>=0.5,1,IF({diff}<=-0.5,-1,0))'
        shift = (f'IF(LEFT({iris_c}{r},1)="X",{iris_c}{r},IF(OR({iris_c}{r}="B+",{iris_c}{r}="C+"),'
                 f'INDEX(LAD_PC,MAX(1,MIN(14,MATCH({iris_c}{r},LAD_PC,0)+{step}))),'
                 f'INDEX(LAD_STD,MAX(1,MIN(12,MATCH({iris_c}{r},LAD_STD,0)+{step})))))')
        lp = col["Leader proposed level"]
        vals["Leader proposed level"] = f'=IF({col["Inputs"]}{r}=0,"",IF({ov}{r}<>"",{ov}{r},{shift}))'
        vals["Leader proposed band"] = (f'=IF({lp}{r}="","",IFERROR(CHOOSE(INDEX(B_RANK,MATCH({lp}{r},B_LVL,0)),"Band A · Practitioner",'
                                        f'"Band B · Specialist / Manager","B+ · Player-Coach","Band C · Principal / Head","C+ · Player-Coach",'
                                        f'"Band D · Business leadership","Band E · Enterprise leadership"),"Band X · Consultant"))')
        vals["Move vs IRIS"] = (f'=IF({lp}{r}="","",IF({lp}{r}={iris_c}{r},"• Same",IF(INDEX(B_VAL,MATCH({lp}{r},B_LVL,0))>'
                                f'INDEX(B_VAL,MATCH({iris_c}{r},B_LVL,0)),"▲ Up","▼ Down")))')
        vals["Level in force · leader if given, else IRIS"] = f'=IF({lp}{r}<>"",{lp}{r},{iris_c}{r})'
        mg = col["Manager Emp ID"]
        mf = f'INDEX(${INF}$7:${INF}${last},MATCH(${mg}{r},${E}$7:${E}${last},0))'
        vals["Manager check · live"] = (
            f'=IFERROR(IF(${mg}{r}="","—",IF(OR(LEFT({INF}{r},1)="X",LEFT({mf},1)="X"),"—",'
            f'IF(AND(LEFT({INF}{r},1)="E",LEFT({mf},1)="E"),"Exception · Band E",'
            f'IF(INDEX(B_VAL,MATCH({INF}{r},B_LVL,0))>=INDEX(B_VAL,MATCH({mf},B_LVL,0)),"⚠ At or above the manager",'
            f'IF(INDEX(B_RANK,MATCH({INF}{r},B_LVL,0))=INDEX(B_RANK,MATCH({mf},B_LVL,0)),"⚠ Same band as the manager","OK"))))),"—")')
        xs = put_row(ws, r, cols, vals)
        for h, x in xs.items():
            if h.startswith(("Weightage", "Rating")):
                x.alignment = Alignment(horizontal="center", vertical="top")
                x.font = Y.font(10, True)
            if h.startswith(("Rating Leader", "Leader override", "Leader comment")):
                x.fill = Y.F_YEL
            if h in ("IRIS proposed level", "Leader proposed level", "Level in force · leader if given, else IRIS", "Market level · published weights"):
                x.font = Y.font(11, True)
                x.alignment = Alignment(horizontal="center", vertical="top")
                b = Y.band_of(vals.get(h, ""))
                if b and h == "IRIS proposed level":
                    x.fill = PatternFill("solid", fgColor=Y.BAND[b])
            if h == "Structure flag" and sflag:
                x.fill, x.font = Y.F_RED, Y.T_RED
            if h == "Title change" and title_change:
                x.fill = PatternFill("solid", fgColor="F8CBAD")
            if h == "All their data ▸":
                x.hyperlink = f"#{Q(S_RAW)}!A{raw_row[i]}"
                x.font = Y.font(9, True, "0563C1", underline="single")
            if h == "Person card ▸":
                x.hyperlink = f"#{Q(S_PC)}!A1"
                x.font = Y.font(9, True, "0563C1", underline="single")
        ws.row_dimensions[r].height = 70
    dv_r = DataValidation(type="list", formula1="LIST_LMH", allow_blank=True)
    dv_o = DataValidation(type="list", formula1="LIST_LEVELS", allow_blank=True)
    ws.add_data_validation(dv_r)
    ws.add_data_validation(dv_o)
    dv_r.add(f"{col['Rating Leader · ' + PARAMS[0][1]]}7:{col['Rating Leader · ' + PARAMS[-1][1]]}{last}")
    dv_o.add(f"{col['Leader override · any level']}7:{col['Leader override · any level']}{last}")
    for c_, h, dark, title in cols:
        if title == "ENGINE":
            ws.column_dimensions[L(c_)].hidden = True
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:{L(len(cols))}{last}"
    ws.sheet_view.zoomScale = 85
    tm_cols = cols

    # ================================================================ Raw Data
    ws = ws_raw
    Y.banner(ws, "Raw Data — the core of what we hold, one row per person",
             "Sheet 7 · Only what is useful: HR record, experience and education, everything shared in the survey, performance and talent "
             "data from the managers, KRAs and the work assigned. Blank = not received. ◂ Back returns to the person on Talent Mapping. "
             "People in the source files but not yet in the architecture are at the bottom.", STAMP, width_cols=16)
    Y.colour_key(ws, 4)
    sup = openpyxl.load_workbook(os.path.join(J.ROOT, "IRIS_JA_Final_v1.15", "IRIS_JA_2_HR_Super_Master_v1.15.xlsx"), read_only=True, data_only=True)
    s3 = list(sup["3. Raw Data & Market Standards"].iter_rows(min_row=6, values_only=True))
    sup.close()
    s3h = {J.text(h): j for j, h in enumerate(s3[0])}
    s3rows = [r_ for r_ in s3[1:] if r_[1] or r_[2]]
    TMP, CRP, MMP = J.PREFIX["TM"], J.PREFIX["CRAW"], J.PREFIX["MM"]
    RAW = [
        ("HR RECORD", "1F3864", [("Business", ["Business"]), ("Emp ID", ["Emp ID"]), ("Employee name", ["Employee name"]), ("Department", ["Department"]),
                                 ("Designation", ["Current designation"]), ("Employee category", ["Employee Category"]), ("Employee status", ["Employee Status"]),
                                 ("Date of joining", ["DoJ"]), ("Reporting manager", ["Reporting Manager"]),
                                 ("Location / region", [TMP + "Location", MMP + "Region / Operating Area"])]),
        ("EXPERIENCE & EDUCATION", "2E75B6", [("Experience before IRIS", ["EXP Outside IRIS"]), ("Experience at IRIS", ["EXP Current IRIS"]),
                                              ("Total experience", [TMP + "Total Experience", CRP + "Total Experience"]),
                                              ("Qualification", ["Highest Qualifications"]), ("Skill sets", ["Skill Sets"])]),
        ("SURVEY — what they shared", "7030A0", [("Survey questions answered · of 20", ["Survey questions answered · of 20"]),
                                                 ("Role mandate & core purpose", [MMP + "Role Mandate"]),
                                                 ("Deliverables & responsibilities", [MMP + "Reported Deliverables & Responsibilities"]),
                                                 ("Capabilities, tools & workload", [MMP + "Current Capabilities"]),
                                                 ("Development priorities & learning needs", [MMP + "Development Priorities & Learning Needs"]),
                                                 ("Challenges & support required", [MMP + "Challenges & Support Required"]),
                                                 ("Evidence used for the ratings", ["@evidence"])]),
        ("PERFORMANCE & TALENT — from the managers", "C55A11", [
            ("Q1 tentative rating · 1-5", [TMP + "Tentative Rating for Q1 · Scale 1-5"]),
            ("Quarterly performance rating", [TMP + "Quarterly Performance Rating", CRP + "Quarterly Performance Rating"]),
            ("Potential", [TMP + "Potential Rating · High/Med/Low", CRP + "Potential Rating · High/Med/Low"]),
            ("9-box", [TMP + "9-Box Placement", CRP + "9-Box Placement"]), ("Ready for promotion", [TMP + "Ready for Promotion", CRP + "Ready for Promotion"]),
            ("Fast tracker", [TMP + "Fast Tracker · Y/N", CRP + "Fast Tracker · Y/N"]), ("Critical role", [TMP + "Critical Role · Y/N", CRP + "Critical Role · Y/N"]),
            ("Successor named", [TMP + "Name of Successor Identified", CRP + "Name of Successor Identified"]),
            ("Last performance discussion", [TMP + "Last Performance Discussion date"]),
            ("Correct designation · manager's view", [TMP + "Correct Designation", CRP + "Correct Designation Based on Work"]),
            ("Structural change suggested", [TMP + "Is a structural Change needed for better future", CRP + "Is a structural Change needed for better future · If yes, What changes"]),
            ("Training needs · current role", [TMP + "Training Needs identified for current Role", CRP + "Training Needs identified for current Role"]),
            ("Training needs · if promoted", [TMP + "Training Needs identified if · If promotion Ready", CRP + "Training Needs identified if · If promotion Ready"]),
            ("Manager comments", [CRP + "Comments", MMP + "Comments"])]),
        ("KRAs & WORK", "375623", [("KRAs set by manager · annual", [TMP + "Annual KRAs · objective Measurable Minimum 5 - set by Manager", MMP + "KRAs Submitted"]),
                                    ("KRAs · quarterly", [TMP + "Quarterly KRAs"]), ("KRAs · self-reported", [TMP + "Annual KRAs · objective Measurable Minimum 5 - Self"]),
                                    ("Actual JD · from manager", [CRP + "Actual JD"]),
                                    ("Projects, products & clients", [TMP + "Projects and Products", TMP + "Projects/Delivery/client Assigned", MMP + "Acual Work"])]),
    ]
    raw_groups = [("PERSON", "1F3864", [("◂ Back", 8), ("Status", 12)])] + [(t, d, [(h, 16 if len(src) and "KRAs" not in h and "Survey" not in t else 36) for h, src in cs]) for t, d, cs in RAW]
    rcols = headers(ws, raw_groups)
    ev_by_i = {}
    for i in order:
        tm = by_i[i]["tm"]
        ev_by_i[i] = "\n".join(f"{dict(PARAMS)[pk]}: {evidence_line(pk, tm[H[jc]])}" for pk, (rc, jc) in FACT_COLS.items()
                               if has_evidence(pk, tm[H[jc]]) and evidence_line(pk, tm[H[jc]]))

    def pick(row, srcs, i=None):
        for s in srcs:
            if s == "@evidence":
                return ev_by_i.get(i, "")
            j = s3h.get(s)
            if j is not None and J.has(row[j]):
                v = row[j]
                return J.debracket(v) if isinstance(v, str) else v
        return ""
    for i in order:
        r = raw_row[i]
        row = s3rows[i]
        x = cell(ws, r, 1, "◂ Back", "EEF2F9", first=True)
        x.hyperlink = f"#{Q(S_TM)}!A{tm_row[i]}"
        x.font = Y.font(9, True, "0563C1", underline="single")
        cell(ws, r, 2, "Mapped", "EEF2F9")
        c_ = 3
        for t, d, cs in RAW:
            for k_, (h, srcs) in enumerate(cs):
                v = pick(row, srcs, i)
                if h == "Survey questions answered · of 20" and J.text(v) in ("0", ""):
                    v = ""
                cell(ws, r, c_, v, TINT.get(d, "FFFFFF"), first=k_ == 0)
                c_ += 1
        ws.row_dimensions[r].height = 80
    r = 7 + len(order)
    seen = set()
    for u in unmapped:
        kk = J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"])
        if kk in seen:
            continue
        seen.add(kk)
        recs = [x for x in unmapped if (J.key(x["rec"]["id"]) or J.nname(x["rec"]["name"])) == kk]
        cell(ws, r, 1, "", "FFC7CE", first=True)
        cell(ws, r, 2, "Removed by HR · 9 Oct" if J.nname(u["rec"]["name"]) == "SHREYASH PAWAR" else "NOT YET MAPPED", "FFC7CE", True)
        c_ = 3
        for t, d, cs in RAW:
            for h, srcs in cs:
                v = ""
                if h == "Emp ID":
                    v = u["rec"]["id"]
                elif h == "Employee name":
                    v = u["rec"]["name"]
                elif h == "Business":
                    v = {"MM": "CARBON", "CRAW": "CARBON"}.get(u["source"]["kind"], u["source"]["label"].split("· ")[-1])
                else:
                    for x in recs:
                        pre = J.PREFIX[x["source"]["kind"]]
                        for s in srcs:
                            base = s[len(pre):] if s.startswith(pre) else (s if not s.startswith("[") else None)
                            if base is None:
                                continue
                            for hh, vv in x["rec"]["row"].items():
                                if J.hkey(hh) == J.hkey(base) and J.has(vv):
                                    v = J.clean(vv)
                                    break
                            if v:
                                break
                        if v:
                            break
                cell(ws, r, c_, v, "FDECEC")
                c_ += 1
        ws.row_dimensions[r].height = 50
        r += 1
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:{L(len(rcols))}{r - 1}"
    ws.sheet_view.zoomScale = 85

    # ================================================================ Person Card
    ws = ws_pc
    Y.banner(ws, "Person Card — one person on one page",
             "Sheet 6 · Pick a name, or type an Emp ID · the Emp ID wins. Everything below follows from Talent Mapping and the Role Model, "
             "live — including the leader's proposal once given.", STAMP, width_cols=8)
    ws["A5"], ws["A6"] = "Name", "Emp ID — type or paste · wins over the name"
    for a in ("A5", "A6"):
        ws[a].font = Y.font(10, True)
    first = next((by_i[i] for i in order if J.has(person_values(M, F, by_i[i])[1][1]) and
                  person_values(M, F, by_i[i])[1][-1].startswith("Per-person") and person_values(M, F, by_i[i])[0][5].startswith("Received")), by_i[order[0]])
    ws["B5"] = f"{first['name']} · {first['emp']}"
    for c in ("B5", "B6"):
        ws[c].fill = Y.F_YEL
        ws[c].font = Y.font(11, True)
    ws.merge_cells("B5:F5")
    ws.merge_cells("B6:F6")
    dv = DataValidation(type="list", formula1="LIST_PEOPLE", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B5")
    ws["H5"] = "Row"
    ws["H6"] = (f'=IFERROR(IF(TRIM($B$6)<>"",MATCH(TRIM($B$6),{Q(S_TM)}!${E}$7:${E}${last},0),'
                f'MATCH(INDEX(PEOPLE_ID,MATCH($B$5,LIST_PEOPLE,0)),{Q(S_TM)}!${E}$7:${E}${last},0)),"")')
    ws["H5"].font = ws["H6"].font = Y.font(8, False, "808080")
    ws["A8"] = "Showing"
    ws["A8"].font = Y.font(10, True)
    ws["B8"] = f'=IF($H$6="","No match — check the Emp ID",INDEX({Q(S_TM)}!$E$7:$E${last},$H$6)&" · "&INDEX({Q(S_TM)}!${E}$7:${E}${last},$H$6))'
    ws["B8"].font = Y.font(12, True, Y.NAVY)
    ws.merge_cells("B8:F8")
    ws.column_dimensions["A"].width = 34
    for c in "BCDEF":
        ws.column_dimensions[c].width = 18
    r = 10
    tcol = {h: L(c_) for c_, h, _, _ in tm_cols}
    def look(h):
        cl = tcol[h]
        return f'=IFERROR(IF(INDEX({Q(S_TM)}!${cl}$7:${cl}${last},$H$6)="","",INDEX({Q(S_TM)}!${cl}$7:${cl}${last},$H$6)),"")'
    sections = [
        ("WHO", ["Business", "Department", "Emp ID", "Employee name", "Current designation", "Reporting manager", "Total exp. · yrs",
                 "Qualification", "Current level · as is"]),
        ("IRIS PLACEMENT", ["Job family", "Sub-family", "Role", "IRIS proposed level", "IRIS proposed band", "Standard title for the role and level",
                            "Title change", "How IRIS placed them", "Structure flag"]),
        ("MARKET", ["Market level · published weights", "Experience vs the level · market", "Span vs the norm", "Radford / Mercer · SFIA 9"]),
        ("LEADER PROPOSAL", ["Leader proposed level", "Leader proposed band", "Move vs IRIS", "Manager check · live", "Leader comment"]),
        ("WHAT THEY TOLD US", ["Survey · questions answered", "What they shared", "KRAs from the manager", "Performance · rating · potential · 9-box",
                               "Evidence used for the ratings"]),
    ]
    for title, fields in sections[:4]:
        Y.header(ws.cell(r, 1, title), Y.NAVY, 10)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
        for h in fields:
            a = ws.cell(r, 1, h)
            a.font = Y.font(9, True)
            a.fill = PatternFill("solid", fgColor="F2F2F2")
            b = ws.cell(r, 2, look(h))
            b.font = Y.font(10, h.startswith(("IRIS proposed level", "Leader proposed level")))
            b.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.row_dimensions[r].height = 45 if h.startswith(("How", "Structure", "Title")) else 18
            r += 1
        r += 1
    Y.header(ws.cell(r, 1, "WEIGHTAGE AND RATINGS — by parameter"), Y.NAVY, 10)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    for j, h in enumerate(["Parameter", "Weightage · IRIS", "Weightage · Market", "Weightage · Leader", "Rating · IRIS", "Rating · Leader"], 1):
        Y.header(ws.cell(r, j, h), "595959")
    r += 1
    for pk, pn in PARAMS:
        ws.cell(r, 1, pn).font = Y.font(9, True)
        for j, h in enumerate((f"Weightage IRIS · {pn}", f"Weightage Market · {pn}", f"Weightage Leader · {pn}", f"Rating IRIS · {pn}",
                               f"Rating Leader · {pn}"), 2):
            x = ws.cell(r, j, look(h))
            x.alignment = Alignment(horizontal="center")
            x.font = Y.font(10, True)
        r += 1
    r += 1
    title, fields = sections[4]
    Y.header(ws.cell(r, 1, title), Y.NAVY, 10)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    for h in fields:
        a = ws.cell(r, 1, h)
        a.font = Y.font(9, True)
        a.fill = PatternFill("solid", fgColor="F2F2F2")
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b = ws.cell(r, 2, look(h))
        b.font = Y.font(9)
        b.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 150 if h in ("What they shared", "KRAs from the manager") else (90 if h.startswith("Evidence") else 18)
        r += 1
    ws.freeze_panes = "A9"

    # ================================================================ How it works
    ws = ws_how
    Y.banner(ws, "IRIS Job Architecture Model — how it works", "Read this first · five steps, in the order they happened.", STAMP, back=False, width_cols=8)
    moved_dn = sum(1 for i in order if V.VAL.get(new[i], 0) < V.VAL.get(org.level[i], 0))
    moved_up = sum(1 for i in order if V.VAL.get(new[i], 0) > V.VAL.get(org.level[i], 0))
    steps = [
        ("1 · We asked", "The How We Grow survey: 20 questions to every employee about the work, decisions, accountability, skills and learning; "
                         "and the managers' talent mapping: KRAs, rating, potential, 9-box. Where nothing came back, the cell is blank."),
        ("2 · We built the ladder", "There was no level framework. From the answers and the market we built one ladder of 12 levels in five "
                                    "bands, plus two Player-Coach steps. Each level has its own experience range, scope and span, so the "
                                    "levels do not overlap · sheet 2."),
        ("3 · We placed everyone", "Each person got a job family, a sub-family, a role and a level from the evidence. Then we moved the fewest "
                                   f"people needed so that no one shares a band with their manager: {moved_dn} down, pay and title protected, "
                                   f"and {moved_up} up. Where a reporting chain is deeper than the bands, we flag it for a structure "
                                   f"decision instead of forcing a level: {len(flag)} cases."),
        ("4 · We checked the market", "Every parameter's weight is set against published job-evaluation methods — Korn Ferry Hay, Mercer "
                                      "IPE and five public plans — and shown as Low / Medium / High with the numbers behind it · sheet 4. "
                                      "Levels map to Radford / Mercer codes, SFIA and peer experience ranges."),
        ("5 · Leaders give their view", "On sheet 5 you see, for each person, what IRIS proposes and why. Rate any parameter Low, Medium or High "
                                        "against the level; re-weight a role on sheet 3; or set any level with the override. Your proposed "
                                        "level and band appear as soon as you give an input, and the manager check updates live."),
        ("The model in one line", "Each parameter is rated Low / Medium / High against the level; each role weights the parameters Low / "
                                  "Medium / High. If your weighted rating is half a level or more above IRIS's, the person moves one level "
                                  "up; half a level or more below, one level down."),
        ("The seven parameters", "Complexity & Expertise · Autonomy & Decision · Accountability & Impact · Leadership & Strategy · "
                                 "Business Impact · Experience · Education."),
        ("Low · Medium · High", "Rating: Low = works at the level below on this parameter; Medium = as the level is defined; High = already "
                                "at the level above. Weightage: Low counts 1, Medium 3, High 9 — the shares follow the published methods "
                                "within a few points."),
        ("Colours", "Yellow = yours to fill · red = a structure decision · orange = a title to align · band colours on levels."),
        ("Privacy", "HR and business leaders only. Personal fields are never written; pay is not in this file."),
    ]
    r = 5
    for a, b in steps:
        x = ws.cell(r, 1, a)
        x.font = Y.font(11, True, Y.NAVY)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        y = ws.cell(r, 2, b)
        y.font = Y.font(10)
        y.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = Y.est_height([(b, 120)], base=14)
        r += 1
    r += 1
    Y.header(ws.cell(r, 1, "The sheets"), Y.NAVY, 10)
    r += 1
    for nm in wb.sheetnames:
        x = ws.cell(r, 1, nm)
        x.hyperlink = f"#{Q(nm)}!A1"
        x.font = Y.font(10, False, "0563C1", underline="single")
        r += 1
    ws.column_dimensions["A"].width = 30
    for c in "BCDEFGH":
        ws.column_dimensions[c].width = 18

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTFILE)
    work = os.path.join(os.environ.get("JA_WORK", "/tmp"), "ja_model")
    os.makedirs(work, exist_ok=True)
    n_f, errors = X.cache_values(OUTFILE, work)
    print(f"Wrote {OUTFILE}: people {len(order)} · roles {len(roles)} · moved down {moved_dn} up {moved_up} · structure flags {len(flag)} · "
          f"title changes {sum(1 for i in order if std_title(roles[by_i[i]['role']], new[i], by_i[i]['tm'][H['Final designation']], by_i[i]['tm'][H['Track']], J.text(by_i[i]['tm'][H['FINAL LEVEL']]))[1])} · "
          f"formulas {n_f} · errors {len(errors)}")
    if errors:
        print(errors[:10])


if __name__ == "__main__":
    main(sys.argv[1])
