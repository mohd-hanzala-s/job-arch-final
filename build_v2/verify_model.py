"""Checks for IRIS Job Architecture Model v2.0. Exit 0 = all pass. Usage: python3 verify_model.py"""
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
import ja_v114 as V  # noqa: E402
import ja_xlsx as X  # noqa: E402
from build_model import EXP_RANGE, OUTFILE  # noqa: E402

RES = []


def check(name, ok, detail=""):
    RES.append(bool(ok))
    print(("PASS " if ok else "FAIL ") + name + (f" — {detail}" if detail else ""))


def main():
    wb = openpyxl.load_workbook(OUTFILE, read_only=True, data_only=True)
    want = ["1. How it works", "2. Ladder", "3. Role Model", "4. Market Benchmark", "5. Talent Mapping", "6. Person Card", "7. Raw Data", "Backend Data"]
    check("Sheets in the order of the story, Backend Data last", wb.sheetnames == want, str(wb.sheetnames))
    t = list(wb["5. Talent Mapping"].iter_rows(min_row=6, values_only=True))
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
    rm = list(wb["3. Role Model"].iter_rows(min_row=6, values_only=True))
    rh = {x: i for i, x in enumerate(rm[0])}
    role_rows = [r for r in rm[1:] if r[rh["Row"]] == "ROLE"]
    ids = {r[rh["Role ID"]] for r in role_rows}
    check("Every person's role is in the Role Model", all(r[h["Role ID"]] in ids for r in d), f"{len(ids)} roles")
    sub_track = [(r[rh["Role"]]) for r in role_rows]
    check("A role is a sub-family done as IC or Manager", all(re.search(r" · (IC|Manager|Consultant)$", x) for x in sub_track))
    placed = [r for r in d if V.band(r[h["IRIS proposed level"]]) != "X"]
    wi = [h[x] for x in h if str(x).startswith("Weightage IRIS")]
    wm = [h[x] for x in h if str(x).startswith("Weightage Market")]
    check("IRIS and Market weightage for every placed person, in L / M / H",
          all(r[c] in ("L", "M", "H") for r in placed for c in wi + wm), f"{len(placed)} people")
    check("Market weights Experience and Education Low · no method weights them on their own",
          all(r[h["Weightage Market · Experience"]] == "L" and r[h["Weightage Market · Education"]] == "L" for r in placed))
    mb = list(wb["4. Market Benchmark"].iter_rows(values_only=True))
    lmh_rows = [r for r in mb if r and r[2] in ("L", "M", "H")]
    check("Market Benchmark gives L / M / H with the published numbers and references on every row",
          len(lmh_rows) >= 56 and all(r[3] and r[7] and r[8] for r in lmh_rows), f"{len(lmh_rows)} rows")
    raw = list(wb["7. Raw Data"].iter_rows(min_row=6, values_only=True))
    rwh = [x for x in raw[0] if x]
    personal = [x for x in rwh if J.hkey(x) in J.PERSONAL]
    check("Raw Data keeps only core and useful columns, no personal fields", len(rwh) <= 50 and not personal, f"{len(rwh)} columns")
    status = collections.Counter(r[1] for r in raw[1:] if r[1])
    check("Raw Data holds everyone plus the people not yet mapped", status.get("Mapped") == 586, dict(status))
    pc = list(wb["6. Person Card"].iter_rows(min_row=8, max_row=8, values_only=True))[0]
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
    ws = wb["5. Talent Mapping"]
    hd = {c.value: c.column for c in ws[6]}
    t1 = next(r for r in range(7, ws.max_row + 1) if ws.cell(r, hd["IRIS proposed level"]).value == "B1" and str(ws.cell(r, hd["Role"]).value).endswith("IC"))
    for k, v in hd.items():
        if str(k).startswith("Rating Leader"):
            ws.cell(t1, v).value = "H"
    t2 = next(r for r in range(t1 + 1, ws.max_row + 1) if ws.cell(r, hd["IRIS proposed level"]).value == "A3")
    ws.cell(t2, hd["Leader override · any level"]).value = "E2"
    rm = wb["3. Role Model"]
    rh = {c.value: c.column for c in rm[6]}
    rr = next(r for r in range(7, rm.max_row + 1) if rm.cell(r, rh["Row"]).value == "ROLE" and (rm.cell(r, rh["People"]).value or 0) >= 5)
    rid = rm.cell(rr, rh["Role ID"]).value
    rm.cell(rr, rh["Leader · Complexity & Expertise"]).value = "L"
    wb.save(os.path.join(work, "e.xlsx"))
    out = X.recalc_copy(os.path.join(work, "e.xlsx"), work)
    v = openpyxl.load_workbook(out, read_only=True, data_only=True)
    t = list(v["5. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(t[0])}
    check("Live: a leader rating every parameter High moves a B1 one level up", t[t1 - 6][h["Leader proposed level"]] == "B2", str(t[t1 - 6][h["Leader proposed level"]]))
    check("Live: an override shows and the manager check flags it", t[t2 - 6][h["Leader proposed level"]] == "E2" and str(t[t2 - 6][h["Manager check · live"]]).startswith("⚠"))
    hold = [x for x in t[1:] if x[h["Role ID"]] == rid]
    check("Live: a leader weightage on the Role Model reaches everyone in the role",
          hold and all(x[h["Weightage Leader · Complexity & Expertise"]] == "L" and x[h["Leader proposed level"]] for x in hold), f"{len(hold)} people")
    others = sum(1 for x in t[1:] if x[h["Leader proposed level"]] not in (None, "") and x[h["Role ID"]] != rid)
    check("Live: nobody else gets a leader proposal", others == 2, f"{others}")
    errs = sum(1 for ws_ in v.worksheets for row in ws_.iter_rows(values_only=True) for x in row if isinstance(x, str) and re.match(r"^#(REF|N/A|VALUE|NAME|DIV)", x))
    check("Live: no formula errors after the edits", errs == 0, str(errs))
    v.close()
    shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
