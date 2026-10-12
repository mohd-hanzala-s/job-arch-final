"""Checks for IRIS Job Architecture Model v2.1. Exit 0 = all pass. Usage: python3 verify_model.py"""
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
import ja_v114 as V  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_model import EXP_RANGE, OUTFILE  # noqa: E402

RES = []


def check(name, ok, detail=""):
    RES.append(bool(ok))
    print(("PASS " if ok else "FAIL ") + name + (f" — {detail}" if detail else ""))


def main():
    wb = openpyxl.load_workbook(OUTFILE, read_only=True, data_only=True)
    want = ["1. How it works", "2. Ladder", "3. Role Ladder & Weightage", "4. Role Definitions", "5. Market Benchmark", "6. Talent Mapping",
            "7. Person Card", "8. Raw Data", "Backend Data"]
    check("Sheets in the order of the story, Backend Data last", wb.sheetnames == want, str(wb.sheetnames))
    t = list(wb["6. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(t[0])}
    d = [r for r in t[1:] if r[h["Emp ID"]] or r[h["Employee name"]]]
    check("Every person, top down from the CEO", len(d) == 586 and d[0][h["Employee name"]] == "Balachandran K", f"{len(d)}")
    boiler = sum(1 for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for v in row
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
    check("Every level row has IRIS, Market and In-use weightage in L / M / H and a written justification",
          all(r[c] in ("L", "M", "H") for r in lrows for c in wcols) and
          all("MARKET —" in J.text(r[lh["Justification · IRIS and Market"]]) and "IRIS —" in J.text(r[lh["Justification · IRIS and Market"]]) for r in lrows))
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
    check("IRIS and Market weightage for every placed person, in L / M / H",
          all(r[c] in ("L", "M", "H") for r in placed for c in wi + wm), f"{len(placed)} people")
    check("Market weights Experience and Education Low · no method weights them on their own",
          all(r[h["Weightage Market · Experience"]] == "L" and r[h["Weightage Market · Education"]] == "L" for r in placed))
    mb = list(wb["5. Market Benchmark"].iter_rows(values_only=True))
    lmh_rows = [r for r in mb if r and r[2] in ("L", "M", "H")]
    check("Market Benchmark gives L / M / H with the published numbers and references on every row",
          len(lmh_rows) >= 56 and all(r[3] and r[7] and r[8] for r in lmh_rows), f"{len(lmh_rows)} rows")
    raw = list(wb["8. Raw Data"].iter_rows(min_row=6, values_only=True))
    rwh = [x for x in raw[0] if x]
    personal = [x for x in rwh if RW.fkey(x.split(" · ")[0]) in RW.DROP]
    qcols = [x for x in rwh if re.match(r"^Q\d+ · ", x)]
    check("Raw Data holds every column we received, all 20 survey questions, no personal fields",
          len(rwh) >= 150 and len(qcols) == 20 and not personal, f"{len(rwh)} columns · {len(qcols)} questions · personal {personal}")
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
    pc = list(wb["7. Person Card"].iter_rows(min_row=8, max_row=8, values_only=True))[0]
    check("Person Card resolves", isinstance(pc[1], str) and "·" in pc[1], str(pc[1]))
    z = zipfile.ZipFile(OUTFILE)
    xml = "".join(z.read(n).decode("utf8") for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
    check("Excel-safe: no '=' in list validations, no formula errors", "<formula1>=" not in xml and 't="e"' not in xml)
    wb.close()
    live()
    print(f"\n{sum(RES)} of {len(RES)} checks pass")
    return 0 if all(RES) else 1


def live():
    work = tempfile.mkdtemp(prefix="jam_")
    wb = openpyxl.load_workbook(OUTFILE)
    ws = wb["6. Talent Mapping"]
    hd = {c.value: c.column for c in ws[6]}
    rl = wb["3. Role Ladder & Weightage"]
    lh = {c.value: c.column for c in rl[6]}
    rkeys = collections.Counter(ws.cell(r, hd["Role key"]).value for r in range(7, ws.max_row + 1))
    # a department-role with several people at two or more levels
    lvls = collections.defaultdict(collections.Counter)
    for r in range(7, ws.max_row + 1):
        lvls[ws.cell(r, hd["Role key"]).value][ws.cell(r, hd["IRIS proposed level"]).value] += 1
    rkey = next(k for k, n in rkeys.most_common() if k and n >= 5 and len(lvls[k]) >= 2)
    lvl = lvls[rkey].most_common(1)[0][0]
    t1 = next(r for r in range(7, ws.max_row + 1) if ws.cell(r, hd["IRIS proposed level"]).value == "B1"
              and str(ws.cell(r, hd["Role"]).value).endswith("IC") and ws.cell(r, hd["Role key"]).value != rkey)
    for k, v in hd.items():
        if str(k).startswith("Rating Leader"):
            ws.cell(t1, v).value = "H"
    t2 = next(r for r in range(t1 + 1, ws.max_row + 1) if ws.cell(r, hd["IRIS proposed level"]).value == "A3" and ws.cell(r, hd["Role key"]).value != rkey)
    ws.cell(t2, hd["Leader override · any level"]).value = "E2"
    rows = {rl.cell(r, lh["Key"]).value: r for r in range(7, rl.max_row + 1)}
    rl.cell(rows[rkey], lh["Leader · Complexity & Expertise"]).value = "L"            # ROLE row: every level of the role here
    lkey = rkey.replace("|ROLE", f"|{lvl}")
    rl.cell(rows[lkey], lh["Leader · Complexity & Expertise"]).value = "H"            # one level row: wins over the ROLE row
    vac = next(r for r in range(7, rl.max_row + 1) if rl.cell(r, lh["People in IRIS"]).value == "Vacant" and rl.cell(r, lh["Row"]).value == "level")
    rl.cell(vac, lh["Leader · Business Impact"]).value = "H"
    vac_before = rl.cell(vac, lh["IRIS · Business Impact"]).value
    wb.save(os.path.join(work, "e.xlsx"))
    out = X.recalc_copy(os.path.join(work, "e.xlsx"), work)
    v = openpyxl.load_workbook(out, read_only=True, data_only=True)
    t = list(v["6. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(t[0])}
    check("Live: a leader rating every parameter High moves a B1 one level up", t[t1 - 6][h["Leader proposed level"]] == "B2", str(t[t1 - 6][h["Leader proposed level"]]))
    check("Live: an override shows and the manager check flags it", t[t2 - 6][h["Leader proposed level"]] == "E2" and str(t[t2 - 6][h["Manager check · live"]]).startswith("⚠"))
    hold = [x for x in t[1:] if x[h["Role key"]] == rkey]
    at_l = [x for x in hold if x[h["IRIS proposed level"]] == lvl]
    rest = [x for x in hold if x[h["IRIS proposed level"]] != lvl]
    check("Live: a ROLE-row weightage on the Role Ladder reaches everyone in that department's role",
          rest and all(x[h["Weightage Leader · Complexity & Expertise"]] == "L" and x[h["Leader proposed level"]] for x in rest), f"{len(rest)} people · {rkey}")
    check("Live: a level-row weightage wins over the ROLE row for that level",
          at_l and all(x[h["Weightage Leader · Complexity & Expertise"]] == "H" for x in at_l), f"{len(at_l)} people at {lvl}")
    others = sum(1 for x in t[1:] if x[h["Leader proposed level"]] not in (None, "") and x[h["Role key"]] != rkey)
    check("Live: nobody else gets a leader proposal", others == 2, f"{others}")
    lr = list(v["3. Role Ladder & Weightage"].iter_rows(min_row=6, values_only=True))
    lx = {x: i for i, x in enumerate(lr[0])}
    vr = lr[vac - 6]
    check("Live: a vacant level can be weighted too · In use follows the leader", vr[lx["In use · Business Impact"]] == "H" and vr[lx["Leader updates"]] == 1,
          f"IRIS {vac_before} → in use {vr[lx['In use · Business Impact']]}")
    lrr = lr[rows[lkey] - 6]
    check("Live: In use on the ladder shows level → ROLE row → IRIS", lrr[lx["In use · Complexity & Expertise"]] == "H" and
          all(x[lx["In use · Complexity & Expertise"]] == "L" for x in lr[1:] if x[lx["Key"]] and x[lx["Key"]].startswith(rkey[:-4]) and x[lx["Row"]] == "level"
              and x[lx["Key"]] != lkey))
    errs = sum(1 for ws_ in v.worksheets for row in ws_.iter_rows(values_only=True) for x in row if isinstance(x, str) and re.match(r"^#(REF|N/A|VALUE|NAME|DIV)", x))
    check("Live: no formula errors after the edits", errs == 0, str(errs))
    v.close()
    shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
