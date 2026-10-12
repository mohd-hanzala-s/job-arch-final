"""Checks for IRIS Job Architecture Model v2.1 · five-step weightage and ratings, live re-scoring. Exit 0 = all pass. Usage: python3 verify_model.py"""
import collections
import os
import re
import shutil
import sys
import tempfile
import zipfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402
import ja_raw as RW  # noqa: E402
import ja_scale as SC  # noqa: E402
import ja_v114 as V  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_model import EXP_RANGE, OUTFILE  # noqa: E402
from build_bh import PARAMS, PKEYS  # noqa: E402

RES = []


def check(name, ok, detail=""):
    RES.append(bool(ok))
    print(("PASS " if ok else "FAIL ") + name + (f" — {detail}" if detail else ""))


def main():
    wb = openpyxl.load_workbook(OUTFILE, read_only=True, data_only=True)
    want = ["1. How it works", "2. Ladder", "3. Role Ladder & Weightage", "4. Role Definitions", "5. Role Journey", "6. Market Benchmark",
            "7. Talent Mapping", "8. Person Card", "9. Raw Data", "Backend Data", "References & Benchmarks"]
    check("Sheets in the order of the story, References & Benchmarks last", wb.sheetnames == want, str(wb.sheetnames))
    t = list(wb["7. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(t[0])}
    d = [r for r in t[1:] if r[h["Emp ID"]] or r[h["Employee name"]]]
    check("Every person, top down from the CEO", len(d) == 586 and d[0][h["Employee name"]] == "Balachandran K", f"{len(d)}")
    boiler = sum(1 for ws in wb.worksheets if ws.title != "References & Benchmarks" for row in ws.iter_rows(values_only=True) for v in row
                 if isinstance(v, str) and re.search(r"no survey answers or KRAs on file|rated like \d+ colleagues|kept at the start point|KRAs not received|No survey response|not in the files received", v))
    check("No-input cells are blank — no placeholder text anywhere", boiler == 0, f"{boiler} cells")
    check("Leader columns blank until a leader gives an input", all(r[h["Leader proposed level"]] in (None, "") for r in d))
    mc = collections.Counter(r[h["Manager check · live"]] for r in d)
    flagged = sum(1 for r in d if r[h["Structure flag"]])
    check("No reportee at or above their manager; a shared band only where a deep chain is flagged",
          mc.get("⚠ At or above the manager", 0) == 0 and mc.get("⚠ Same band as the manager", 0) == flagged, f"{dict(mc)} · flagged {flagged}")
    regular = [lv for lv in V.LEVELS_STD if lv not in ("A1", "E2")]
    overl = [(a, b) for a, b in zip(regular, regular[1:]) if EXP_RANGE[a][1] > EXP_RANGE[b][0]]
    check("Ladder experience ranges do not overlap", not overl, str(overl))
    seen = collections.defaultdict(set)
    for r in d:
        if r[h["Job family"]] in (V.HEAD_FAMILY, V.ENT_FAMILY):
            continue
        seen[(r[h["Role ID"]], J.text(r[h["Standard title for the role and level"]]).replace(" – POD Lead", ""))].add(r[h["IRIS proposed level"]])
    dup = [k for k, v in seen.items() if len(v) > 1]
    check("One standard title per role and level", not dup, f"{len(dup)} {dup[:3]}")
    rm = list(wb["4. Role Definitions"].iter_rows(min_row=6, values_only=True))
    rh = {x: i for i, x in enumerate(rm[0])}
    role_rows = [r for r in rm[1:] if r[rh["Row"]] == "ROLE"]
    ids = {r[rh["Role ID"]] for r in role_rows}
    check("Every person's role is in the Role Definitions", all(r[h["Role ID"]] in ids for r in d), f"{len(ids)} roles")
    vac_def = [r for r in role_rows if r[rh["People"]] == "Vacant"]
    check("Role Definitions include roles nobody holds yet, with every level", len(vac_def) >= 20 and
          all(sum(1 for x in rm[1:] if x[rh["Role ID"]] == r[rh["Role ID"]] and x[rh["Row"]] == "level") >= 2 for r in vac_def), f"{len(vac_def)} vacant roles")
    rl = list(wb["3. Role Ladder & Weightage"].iter_rows(min_row=6, values_only=True))
    lh = {x: i for i, x in enumerate(rl[0])}
    lrows = [r for r in rl[1:] if r[lh["Row"]] == "level"]
    keys = {r[lh["Key"]] for r in rl[1:]}
    check("Every placed person reads their weightage from their own department × role × level row",
          all(r[h["Ladder key"]] in keys and r[h["Ladder key"]].split("|")[3] == r[h["IRIS proposed level"]] for r in d
              if V.band(r[h["IRIS proposed level"]]) != "X"), f"{len(lrows)} level rows")
    ic_a1 = [r for r in rl[1:] if r[lh["Row"]] == "ROLE" and r[lh["Job family"]] not in (V.HEAD_FAMILY, V.ENT_FAMILY)]
    lv_of = collections.defaultdict(set)
    for r in lrows:
        lv_of[r[lh["Key"]].rsplit("|", 1)[0]].add(r[lh["Level"]])
    need_ic = {x for x in V.LEVELS_PC if V.VAL[x] <= V.VAL["D1"]}
    short_ = [r[lh["Key"]] for r in ic_a1 if not need_ic <= lv_of[r[lh["Key"]].rsplit("|", 1)[0]]
              or (r[lh["Track"]] == "Manager" and set(V.LEVELS_PC) - lv_of[r[lh["Key"]].rsplit("|", 1)[0]])]
    check("Each department-role ladder runs from A1 Intern to the top of the role · D1 for IC, E2 for managers · vacant levels included",
          not short_, f"{len(ic_a1)} role ladders · {sum(1 for r in lrows if r[lh['People in IRIS']] == 'Vacant')} vacant level rows · {short_[:2]}")
    check("Roles nobody holds yet have a ladder too", sum(1 for r in ic_a1 if r[lh["Department"]] == "Not yet held") >= 20)
    wcols = [lh[x] for x in lh if str(x).startswith(("IRIS ·", "Market ·", "In use ·"))]
    codes = {c for c, *_ in SC.W5} | {0}
    check("Every level row has IRIS, Market and In-use weightage in five steps (0 = not weighed) and a written justification",
          all(r[c] in codes for r in lrows for c in wcols) and
          all("MARKET —" in J.text(r[lh["Justification · IRIS and Market"]]) and "IRIS —" in J.text(r[lh["Justification · IRIS and Market"]]) for r in lrows))
    scols = [lh[x] for x in lh if str(x).startswith("Share ·")]
    bad = [r[lh["Key"]] for r in lrows if abs(sum(float(r[c] or 0) for c in scols) - 1) > 1e-6]
    check("Each level row's shares of the decision add up to 100%", not bad, f"{len(bad)} rows off {bad[:2]}")
    check("Nothing changed yet: no level row reports a change", all((r[lh["Changes here"]] or 0) == 0 for r in rl[1:]))
    lt = collections.defaultdict(set)
    for r in lrows:
        lt[(r[lh["Role ID"]], r[lh["Standard title at this level"]])].add(r[lh["Level"]])
    check("One standard title per role and level on the ladder, vacant levels included", all(len(v) == 1 for v in lt.values()))
    sv = [r[h["Inputs received in the survey"]] for r in d]
    check("Inputs received in the survey: a detailed summary for everyone who answered, blank for the rest",
          sum(1 for x in sv if x) >= 500 and all((x or "").startswith(("Answered", "Shared")) for x in sv if x)
          and all(not x or r[h["Survey · questions answered"]] or "CARBON" in x for x, r in zip(sv, d)), f"{sum(1 for x in sv if x)} people")
    sub_track = [(r[rh["Role"]]) for r in role_rows]
    check("A role is a sub-family done as IC or Manager", all(re.search(r" · (IC|Manager|Consultant)$", x) for x in sub_track))
    placed = [r for r in d if V.band(r[h["IRIS proposed level"]]) != "X"]
    wi = [h[x] for x in h if str(x).startswith("Weightage IRIS")]
    wm = [h[x] for x in h if str(x).startswith("Weightage Market")]
    check("IRIS and Market weightage for every placed person, in five steps",
          all(r[c] in codes for r in placed for c in wi + wm), f"{len(placed)} people")
    check("Market weights Experience and Education VL · no method weights them on their own",
          all(r[h["Weightage Market · Experience"]] == "VL" and r[h["Weightage Market · Education"]] == "VL" for r in placed))
    ri = [h[x] for x in h if str(x).startswith("Rating IRIS")]
    check("IRIS ratings are 1 to 5 against the level, blank where nothing was received",
          all(r[c] in (None, "", 1, 2, 3, 4, 5) for r in d for c in ri) and sum(1 for r in d for c in ri if r[c]) > 2000)
    gap = [r[h["Emp ID"]] for r in d if abs(float(r[h["Net live"]] or 0) - float(r[h["Anchor · IRIS as built"]] or 0)) > 1e-5]
    check("Before any input, the live score equals IRIS's anchor for everyone — no one moves by accident", not gap, f"{len(gap)} {gap[:3]}")
    mb = list(wb["6. Market Benchmark"].iter_rows(values_only=True))
    lmh_rows = [r for r in mb if r and r[2] in codes and r[1] and r[0] != "Role group"]
    check("Market Benchmark gives five steps with the published numbers, the importance and references on every row",
          len(lmh_rows) >= 56 and all(r[3] and r[7] and r[9] and r[10] for r in lmh_rows), f"{len(lmh_rows)} rows")
    rj = list(wb["5. Role Journey"].iter_rows(min_row=6, values_only=True))
    jh = {x: i for i, x in enumerate(rj[0])}
    jr = [r for r in rj[1:] if r[jh["Role ID"]]]
    per = collections.Counter(r[jh["Role ID"]] for r in jr)
    kinds = collections.Counter(r[jh["Row"]] for r in jr)
    laddered = {r[rh["Role ID"]] for r in rm[1:] if r[rh["Row"]] == "level"}
    lv_cols = [k for k in jh if k in V.LEVELS_PC]
    pairs = [(a, b) for a, b in zip(jr[0::2], jr[1::2])]
    cells = [(lv, b[jh[lv]]) for a, b in pairs for lv in lv_cols if a[jh[lv]] not in (None, "", "—")]
    filled = sum(1 for lv, x in cells if x not in (None, "", "—"))
    check("Role Journey: two rows per role — IRIS and Market benchmark — for every role on a ladder, market filled at its levels",
          set(per) == laddered and all(v == 2 for v in per.values()) and kinds["IRIS"] == kinds["Market benchmark"]
          and all(a[jh["Row"]] == "IRIS" and b[jh["Row"]] == "Market benchmark" for a, b in pairs) and filled == len(cells),
          f"{len(per)} roles · market at {filled} of {len(cells)} role-levels")
    ref = wb["References & Benchmarks"]
    links = sum(1 for row in ref.iter_rows(values_only=True) for v in row if isinstance(v, str) and v.startswith("http"))
    srcs = {v.split()[0] for row in ref.iter_rows(values_only=True) for v in row[:1] if isinstance(v, str) and re.match(r"^S\d\d", v)}
    check("References & Benchmarks, last sheet: every source with its link, the crosswalk and the method sources",
          links >= 100 and {"S01", "S70", "S71"} <= srcs, f"{links} links · {len(srcs)} source IDs")
    raw = list(wb["9. Raw Data"].iter_rows(min_row=6, values_only=True))
    rwh = [x for x in raw[0] if x]
    personal = [x for x in rwh if RW.fkey(x.split(" · ")[0]) in RW.DROP]
    qcols = [x for x in rwh if re.match(r"^Q\d+ · ", x)]
    check("Raw Data holds every column we received, all 20 survey questions, no personal fields",
          len(rwh) >= 150 and len(qcols) == 20 and not personal, f"{len(rwh)} columns · {len(qcols)} questions · personal {personal}")
    grp = [J.text(x) for x in list(wb["9. Raw Data"].iter_rows(min_row=5, max_row=5, values_only=True))[0] if x]
    check("Raw Data also holds the analysis on file from every earlier version — placement, scoring, profile, talent, data quality",
          sum(1 for g in grp if g.startswith("ON FILE")) == 6, str([g[:30] for g in grp if g.startswith("ON FILE")]))
    status = collections.Counter(r[1] for r in raw[1:] if r[1])
    check("Raw Data holds everyone plus the people not yet mapped and those only in the source files", status.get("Mapped") == 586, dict(status))
    rh_ = {x: i for i, x in enumerate(raw[0])}
    by_id, by_nm = {}, collections.defaultdict(str)
    for r in raw[1:]:
        blob = " ".join(re.sub(r"\s+", " ", str(x).replace("\u200b", "")).lower() for x in r if x not in (None, ""))
        for x in str(r[rh_["Emp ID"]] or "").split("\n"):
            if J.key(x.split(": ")[-1]).startswith(("iris", "intn")):
                by_id.setdefault(J.key(x.split(": ")[-1]), blob)
        for x in str(r[rh_["Employee Name"]] or "").split("\n"):
            by_nm[J.nname(x.split(": ")[-1])] += " ¦ " + blob
    n_ans, lost = 0, []
    if RW.v1_files():
        for biz, hdr, r in RW.v1_rows():
            j = {x: i for i, x in enumerate(hdr)}
            b = by_id.get(J.key(r[j["Emp Id"]])) or by_nm.get(J.nname(r[j["Employee Name"]]), "")
            for q in RW.SURVEY_Q + [RW.ADDITIONAL]:
                v = RW.val(r[j[q]], q) if q in j else ""
                if v != "":
                    n_ans += 1
                    if re.sub(r"\s+", " ", str(v).replace("\u200b", "")).strip().lower() not in b:
                        lost.append((r[j["Employee Name"]], q))
    check("Every survey answer received, question by question, is on the person's Raw Data row",
          n_ans > 9000 and not lost, f"{n_ans} answers · {len(lost)} missing {lost[:3]}")
    with_q = sum(1 for x in sv if x and "Q1 · " in x)
    check("Inputs received in the survey on Talent Mapping carries the answers question by question", with_q >= 515, f"{with_q} people")
    pc = list(wb["8. Person Card"].iter_rows(min_row=8, max_row=8, values_only=True))[0]
    check("Person Card resolves", isinstance(pc[1], str) and "·" in pc[1], str(pc[1]))
    z = zipfile.ZipFile(OUTFILE)
    xml = "".join(z.read(n).decode("utf8") for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
    check("Excel-safe: no '=' in list validations, no formula errors", "<formula1>=" not in xml and 't="e"' not in xml)
    wb.close()
    live()
    print(f"\n{sum(RES)} of {len(RES)} checks pass")
    return 0 if all(RES) else 1


def ladder_of(lv):
    return V.LEVELS_PC if lv in ("B+", "C+") else V.LEVELS_STD


def predict(lv, weights, ratings, anchor):
    """The workbook's rule, recomputed here: shift = weighted rating − anchor; round half away from zero; at most two levels."""
    sh = SC.net(weights, ratings, PKEYS) - anchor
    st = int(abs(sh) + 0.5) * (1 if sh >= 0 else -1)
    st = max(-SC.MAX_SHIFT, min(SC.MAX_SHIFT, st))
    lad = ladder_of(lv)
    return lad[max(0, min(len(lad) - 1, lad.index(lv) + st))], sh


def live():
    work = tempfile.mkdtemp(prefix="jam_")
    base = openpyxl.load_workbook(OUTFILE, read_only=True, data_only=True)
    t0 = list(base["7. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    base.close()
    h0 = {x: i for i, x in enumerate(t0[0])}
    pn = dict(PARAMS)
    rows0 = {n + 6: r for n, r in enumerate(t0[1:], 1) if r[h0["Emp ID"]]}

    def w_of(r):
        return {k: r[h0[f"Weightage in use · {pn[k]}"]] for k in PKEYS}

    def r_of(r):
        return {k: r[h0[f"Rating IRIS · {pn[k]}"]] for k in PKEYS}

    def anc(r):
        return float(r[h0["Anchor · IRIS as built"]] or 0)

    wb = openpyxl.load_workbook(OUTFILE)
    ws = wb["7. Talent Mapping"]
    hd = {c.value: c.column for c in ws[6]}
    rl = wb["3. Role Ladder & Weightage"]
    lh = {c.value: c.column for c in rl[6]}
    lrow = {rl.cell(r, lh["Key"]).value: r for r in range(7, rl.max_row + 1)}
    by_key = collections.defaultdict(list)
    for r, v in rows0.items():
        if v[h0["Ladder key"]]:
            by_key[v[h0["Ladder key"]]].append(r)
    role_of = {r: v[h0["Role key"]] for r, v in rows0.items()}
    # A · a role (ROLE row) and one of its levels (level row) in one department
    roles = collections.Counter(role_of[r] for r in rows0 if role_of[r])
    rkey = next(k for k, n in roles.most_common() if n >= 6 and len({rows0[r][h0["IRIS proposed level"]] for r in rows0 if role_of[r] == k}) >= 2)
    lvl = collections.Counter(rows0[r][h0["IRIS proposed level"]] for r in rows0 if role_of[r] == rkey).most_common(1)[0][0]
    lkey = rkey.replace("|ROLE", f"|{lvl}")
    rl.cell(lrow[rkey], lh["Leader · Complexity & Expertise"]).value = "VL"
    rl.cell(lrow[lkey], lh["Leader · Complexity & Expertise"]).value = "VH"
    # B · HR changes the IRIS base for another role and level
    bkey = next(k for k, rs in sorted(by_key.items(), key=lambda x: -len(x[1])) if len(rs) >= 3 and not k.startswith(rkey[:-4]))
    rl.cell(lrow[bkey], lh["IRIS · Leadership & Strategy"]).value = "VH"
    rl.cell(lrow[bkey], lh["IRIS · Complexity & Expertise"]).value = 15
    used = set(by_key[bkey]) | {r for r in rows0 if role_of[r] == rkey}
    # C · a leader types weightage for one person on Talent Mapping
    cand = []
    for r, v in rows0.items():
        if r in used or not v[h0["Ladder key"]] or v[h0["IRIS proposed level"]] in ("E1", "E2", "D2"):
            continue
        rs = r_of(v)
        top = max((k for k in PKEYS if rs[k] not in (None, "")), key=lambda k: rs[k], default=None)
        if top and SC.R5_MOVE[int(rs[top])] - anc(v) >= 0.5:
            cand.append((r, top))
    p_row, p_top = cand[0]
    for k in PKEYS:
        ws.cell(p_row, hd[f"Weightage Leader · {pn[k]}"]).value = "VH" if k == p_top else 0
    used.add(p_row)
    # D · a leader rates every parameter 5 for a B1, and sets an override for an A3
    t1 = next(r for r, v in rows0.items() if v[h0["IRIS proposed level"]] == "B1" and r not in used and v[h0["Ladder key"]])
    for k in PKEYS:
        ws.cell(t1, hd[f"Rating Leader · {pn[k]}"]).value = 5
    used.add(t1)
    t2 = next(r for r, v in rows0.items() if v[h0["IRIS proposed level"]] == "A3" and r not in used)
    ws.cell(t2, hd["Leader override · any level"]).value = "E2"
    used.add(t2)
    # E · a vacant level
    vac = next(r for r in range(7, rl.max_row + 1) if rl.cell(r, lh["People in IRIS"]).value == "Vacant" and rl.cell(r, lh["Row"]).value == "level")
    rl.cell(vac, lh["Leader · Business Impact"]).value = "H"
    wb.save(os.path.join(work, "e.xlsx"))
    out = X.recalc_copy(os.path.join(work, "e.xlsx"), work)
    v = openpyxl.load_workbook(out, read_only=True, data_only=True)
    t = list(v["7. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(t[0])}
    got = {r: t[r - 6] for r in rows0}

    def expect(r, wts):
        return predict(rows0[r][h0["IRIS proposed level"]], wts, r_of(rows0[r]), anc(rows0[r]))[0]

    # D
    w1 = w_of(rows0[t1])
    e1 = predict("B1", w1, {k: 5 for k in PKEYS}, anc(rows0[t1]))[0]
    check("Live: a leader rating every parameter 5 moves a B1 up — to exactly the level the rule gives",
          got[t1][h["Leader proposed level"]] == e1 and e1 != "B1", f"{got[t1][h['Leader proposed level']]} · expected {e1}")
    check("Live: an override shows and the manager check flags it",
          got[t2][h["Leader proposed level"]] == "E2" and str(got[t2][h["Manager check · live"]]).startswith("⚠"))
    # C
    wC = {k: ("VH" if k == p_top else 0) for k in PKEYS}
    eC = expect(p_row, wC)
    check("Live: weightage typed for one person on Talent Mapping re-scores them and changes their level and band",
          got[p_row][h["Leader proposed level"]] == eC and eC != rows0[p_row][h0["IRIS proposed level"]]
          and got[p_row][h["Leader proposed band"]] not in (None, ""),
          f"{rows0[p_row][h0['IRIS proposed level']]} → {got[p_row][h['Leader proposed level']]} · expected {eC} · band {got[p_row][h['Leader proposed band']]}")
    # A
    in_role = [r for r in rows0 if role_of[r] == rkey]
    bad_a, moved_a = [], 0
    for r in in_role:
        wts = w_of(rows0[r])
        wts["C"] = "VH" if rows0[r][h0["Ladder key"]] == lkey else "VL"
        e = expect(r, wts)
        moved_a += e != rows0[r][h0["IRIS proposed level"]]
        if got[r][h["Leader proposed level"]] != e or got[r][h["Weightage in use · Complexity & Expertise"]] != wts["C"]:
            bad_a.append((rows0[r][h0["Employee name"]], got[r][h["Leader proposed level"]], e))
    check("Live: a ROLE-row weightage reaches everyone in the department's role, a level row wins for its level, and each "
          "person lands exactly where the rule puts them", not bad_a and in_role, f"{len(in_role)} people · {moved_a} change level · {bad_a[:2]}")
    # B
    bad_b, moved_b = [], 0
    for r in by_key[bkey]:
        wts = w_of(rows0[r])
        wts["L"], wts["C"] = "VH", 15
        e = expect(r, wts)
        moved_b += e != rows0[r][h0["IRIS proposed level"]]
        if got[r][h["Leader proposed level"]] != e:
            bad_b.append((rows0[r][h0["Employee name"]], got[r][h["Leader proposed level"]], e))
    check("Live: HR changing the IRIS base on the Role Ladder (a code and a number) re-scores everyone at that role and level",
          not bad_b, f"{len(by_key[bkey])} people · {moved_b} change level · {bad_b[:2]}")
    others = sum(1 for r in rows0 if got[r][h["Leader proposed level"]] not in (None, "") and r not in used and r not in by_key[bkey])
    check("Live: nobody else gets a leader proposal", others == 0, f"{others}")
    lr = list(v["3. Role Ladder & Weightage"].iter_rows(min_row=6, values_only=True))
    lx = {x: i for i, x in enumerate(lr[0])}
    vr = lr[vac - 6]
    check("Live: a vacant level can be weighted too · In use and the shares follow",
          vr[lx["In use · Business Impact"]] == "H" and vr[lx["Changes here"]] == 1
          and abs(sum(float(vr[lx[f"Share · {pn[k]}"]] or 0) for k in PKEYS) - 1) < 1e-6)
    br = lr[lrow[bkey] - 6]
    check("Live: the Role Ladder counts HR's base changes", br[lx["Changes here"]] == 2, str(br[lx["Changes here"]]))
    errs = sum(1 for ws_ in v.worksheets for row in ws_.iter_rows(values_only=True) for x in row if isinstance(x, str) and re.match(r"^#(REF|N/A|VALUE|NAME|DIV)", x))
    check("Live: no formula errors after the edits", errs == 0, str(errs))
    v.close()
    shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
