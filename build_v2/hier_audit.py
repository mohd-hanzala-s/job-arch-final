"""IRIS Job Architecture v2.2 — Hierarchy Audit.

Checks every reporting line in the Master workbook and lists, for HR, where the source files disagree about who reports to
whom. Nothing here changes a reporting line: the model follows the Employee Master (HRIS) of 23 Sep 2026.

Usage: JA_V1=<v1 folder> python3 hier_audit.py <source folder>
"""
import collections
import glob
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402
import ja_raw as RW  # noqa: E402
import ja_style as Y  # noqa: E402
import ja_v114 as V  # noqa: E402
import build_model as BM  # noqa: E402
from build_bh import cell  # noqa: E402

OUT = os.path.join(BM.OUTDIR, "IRIS_JA_Hierarchy_Audit_v2.2.xlsx")
STAMP = "IRIS · Job Architecture v2.2 · Hierarchy Audit · October 2026 · Internal — Confidential · HR only"
HRIS = "Employee Master · HRIS 23 Sep 2026"
OTHER = ["Employee Master · older copy", "CARBON Mapping Master · current", "CARBON Talent Mapping · Raw", "Talent Mapping · iFile collated",
         "Talent Mapping · iDEAL", "Talent Mapping · Database Content", "Talent Mapping · HR", "v1 workbook"]


def same(a, b):
    a, b = J.nname(a), J.nname(b)
    if a == b:
        return True
    ta, tb = a.split(), b.split()
    return bool(ta and tb and ta[0] == tb[0] and (ta[-1] == tb[-1] or ta[-1][:4] == tb[-1][:4]))


def main(srcdir):
    wb_ = openpyxl.load_workbook(BM.MASTER_FILE, read_only=True, data_only=True)
    t = list(wb_["7. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    wb_.close()
    h = {x: i for i, x in enumerate(t[0])}
    d = [r for r in t[1:] if r[h["Emp ID"]]]
    by_row = {n: r for n, r in enumerate(d)}
    key_of = lambda r: (J.text(r[h["Emp ID"]]), J.nname(r[h["Employee name"]]))
    idx = {key_of(r): n for n, r in by_row.items()}
    by_emp = {}
    for n, r in by_row.items():
        e = J.text(r[h["Emp ID"]])
        if e.startswith(("IRIS", "INTN")):
            by_emp[e] = n
    by_nm = collections.defaultdict(list)
    for n, r in by_row.items():
        by_nm[J.nname(r[h["Employee name"]])].append(n)

    def who(emp, name):
        e = J.text(emp)
        if e in by_emp:
            return by_emp[e]
        c = by_nm.get(J.nname(name), [])
        return c[0] if len(c) == 1 else None

    # every source's view of each person's current manager
    srcs = collections.defaultdict(dict)
    proposed = {}

    def add(label, emp, name, m):
        m = J.text(m)
        n = who(emp, name)
        if n is None or not J.has(m) or RW.PLACEHOLDER.match(m):
            return
        srcs[n].setdefault(label, m)
    for f in glob.glob(os.path.join(RW.V1_DIR, "Employee Master_New_*.xlsx")):
        for sheet, hdr, r, emp, name in RW.em_rows(f):
            j = {J.text(x).strip(): i for i, x in enumerate(hdr)}
            mc = j.get("Reporting Manager", j.get("Line Manager"))
            if mc is not None:
                add(HRIS, emp, name, r[mc])
    for f in glob.glob(os.path.join(RW.V1_DIR, "Employee Master_OLD*.xlsx")):
        for sheet, hdr, r, emp, name in RW.em_rows(f):
            j = {J.text(x).strip(): i for i, x in enumerate(hdr)}
            mc = j.get("Reporting Manager", j.get("Line Manager"))
            if mc is not None:
                add(OTHER[0], emp, name, r[mc])
    M, S, people, un = J.build(srcdir)
    lab_of = {"MM": OTHER[1], "CRAW": OTHER[2]}
    for s in S:
        lab = lab_of.get(s["kind"]) or {"iFile": OTHER[3], "iDEAL": OTHER[4], "Database": OTHER[5], "HR": OTHER[6]}.get(
            next((k for k in ("iFile", "iDEAL", "Database", "HR") if k in s["label"]), ""), s["label"])
        for rec in s["recs"]:
            for hh, v in rec["row"].items():
                k = J.key(hh)
                if k in ("reportingmanager", "currentreportingmanager"):
                    add(lab, rec["id"], rec["name"], v)
                if k == "newreportingmanager":
                    n = who(rec["id"], rec["name"])
                    if n is not None and J.has(v):
                        proposed[n] = J.text(v)
    for biz, hdr, r in RW.v1_rows():
        j = {x: i for i, x in enumerate(hdr)}
        add(OTHER[-1], r[j["Emp Id"]], r[j["Employee Name"]], r[j["Reporting Manager"]])

    lv = {n: J.text(r[h["IRIS proposed level"]]) for n, r in by_row.items()}
    mgr_n = {n: who(r[h["Manager Emp ID"]], "") for n, r in by_row.items()}
    on = lambda x: V.band(x) not in ("X", "")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_sum, ws_all, ws_diff, ws_prop = (wb.create_sheet(n) for n in ("1. Summary", "2. Every reporting line", "3. Source differences",
                                                                      "4. CARBON proposed lines"))

    # 2 · every reporting line
    ws = ws_all
    Y.banner(ws, "Every reporting line — checked against the level rule",
             "A reportee sits at a lower level than their manager (B1 under B2 is fine; B2 under B2 is not). Reporting lines are the "
             "Employee Master's (HRIS, 23 Sep 2026), unchanged. Gap = steps between the two levels on the 14-level ladder "
             "A1 A2 A3 · B1 B2 B3 B+ · C1 C2 C+ · D1 D2 · E1 E2.", STAMP, back=False, width_cols=14)
    cols = BM.headers(ws, [("PERSON", "1F3864", [("Business", 10), ("Department", 18), ("Layer from the CEO", 7), ("Emp ID", 12),
                                                  ("Employee name", 22), ("Level", 7), ("Band", 20)]),
                           ("MANAGER — as in the Employee Master", "375623", [("Manager", 22), ("Manager Emp ID", 12),
                                                                               ("Manager level", 8), ("Gap · steps", 7)]),
                           ("CHECK", "C00000", [("Level rule", 22), ("Other files name a different manager", 40),
                                                ("How IRIS placed them", 50)])])
    stat = collections.Counter()
    gaps = collections.Counter()
    r_ = 7
    for n, r in by_row.items():
        m = mgr_n[n]
        mname = J.text(r[h["Reporting manager"]])
        if mname.startswith("—"):
            status = "Top of the organisation"
        elif not mname:
            status = "No manager on record" if on(lv[n]) else "Consultant · no manager on record"
        elif m is None:
            status = "⚠ Manager not found"
        elif not on(lv[n]) or not on(lv[m]):
            status = "Consultant · outside the ladder"
        elif V.VAL[lv[n]] < V.VAL[lv[m]]:
            status = "OK"
        else:
            status = "⚠ At or above the manager"
        stat[status] += 1
        gap = ""
        if status == "OK":
            gap = V.LEVELS_PC.index(lv[m]) - V.LEVELS_PC.index(lv[n])
            gaps[gap] += 1
        others = "; ".join(f"{lab}: {v}" for lab, v in srcs.get(n, {}).items() if lab != HRIS and mname and not same(mname, v))
        vals = [r[h["Business"]], r[h["Department"]], r[h["Layer from the CEO"]], r[h["Emp ID"]], r[h["Employee name"]], lv[n],
                r[h["IRIS proposed band"]], "" if mname.startswith("—") else mname, r[h["Manager Emp ID"]], lv.get(m, "") if m is not None else "",
                gap, status, others, r[h["How IRIS placed them"]]]
        for j, v in enumerate(vals, 1):
            x = cell(ws, r_, j, v, "FFFFFF", j == 1)
            if j == 12 and str(v).startswith("⚠"):
                x.fill, x.font = Y.F_RED, Y.T_RED
            if j == 13 and v:
                x.fill = PatternFill("solid", fgColor="FFF2CC")
        b = Y.band_of(lv[n])
        if b:
            ws.cell(r_, 6).fill = PatternFill("solid", fgColor=Y.BAND[b])
        ws.row_dimensions[r_].height = 30 if not others else 45
        r_ += 1
    ws.freeze_panes = "F7"
    ws.auto_filter.ref = f"A6:N{r_ - 1}"

    # 3 · where the files disagree
    ws = ws_diff
    Y.banner(ws, "Source differences — who reports to whom, file by file",
             "People whose manager in another file differs from the Employee Master (HRIS, 23 Sep 2026), which the model follows. "
             "These are differences in the source data for HR to confirm or correct; no reporting line has been changed.",
             STAMP, back=False, width_cols=12)
    hdr = ["Business", "Emp ID", "Employee name", "Level", HRIS + " · used by the model"] + OTHER
    for j, x in enumerate(hdr, 1):
        Y.header(ws.cell(6, j, x), "1F3864" if j <= 5 else "595959")
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 22 if j > 3 else (12 if j != 3 else 24)
    ws.row_dimensions[6].height = 45
    r_ = 7
    n_diff = 0
    for n, r in by_row.items():
        mname = J.text(r[h["Reporting manager"]])
        base = "" if mname.startswith("—") else mname
        alt = {lab: v for lab, v in srcs.get(n, {}).items() if lab != HRIS and base and not same(base, v)}
        if not alt:
            continue
        n_diff += 1
        vals = [r[h["Business"]], r[h["Emp ID"]], r[h["Employee name"]], lv[n], base] + [srcs.get(n, {}).get(lab, "") for lab in OTHER]
        for j, v in enumerate(vals, 1):
            x = cell(ws, r_, j, v, "FFFFFF", j == 1)
            if j > 5 and v and not same(base, v):
                x.fill = PatternFill("solid", fgColor="FFF2CC")
        r_ += 1
    ws.freeze_panes = "D7"
    ws.auto_filter.ref = f"A6:{openpyxl.utils.get_column_letter(len(hdr))}{max(r_ - 1, 7)}"

    # 4 · CARBON's proposed lines, not applied
    ws = ws_prop
    Y.banner(ws, "CARBON Mapping Master — proposed new reporting lines, not applied",
             "CARBON's own mapping proposes a new manager for these people. The model keeps the Employee Master's line; this sheet shows "
             "what each proposal would mean for the level rule if HR decides to adopt it.", STAMP, back=False, width_cols=8)
    hdr = ["Emp ID", "Employee name", "Level", "Current manager · model", "Current manager's level", "Proposed manager",
           "Proposed manager's level", "Level rule if adopted"]
    for j, x in enumerate(hdr, 1):
        Y.header(ws.cell(6, j, x), "375623")
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = (12, 24, 8, 24, 10, 34, 10, 30)[j - 1]
    ws.row_dimensions[6].height = 40
    r_ = 7
    n_prop = 0
    for n, pm in proposed.items():
        r = by_row[n]
        cur = J.text(r[h["Reporting manager"]])
        if same(cur, pm):
            continue
        cand = [k for k, rr in by_row.items() if same(J.text(rr[h["Employee name"]]), pm.split(" — ")[0].replace("K. Balachandran (CEO)", "Balachandran K"))]
        pl = lv[cand[0]] if len(cand) == 1 else ""
        rule = ("—" if not pl or not on(pl) or not on(lv[n]) else
                ("Holds" if V.VAL[lv[n]] < V.VAL[pl] else f"Would break: {lv[n]} under {pl}"))
        cm = mgr_n[n]
        vals = [r[h["Emp ID"]], r[h["Employee name"]], lv[n], cur, lv.get(cm, "") if cm is not None else "", pm, pl, rule]
        for j, v in enumerate(vals, 1):
            x = cell(ws, r_, j, v, "FFFFFF", j == 1)
            if j == 8 and str(v).startswith("Would"):
                x.fill, x.font = Y.F_RED, Y.T_RED
        n_prop += 1
        r_ += 1
    ws.freeze_panes = "C7"

    # 1 · summary
    ws = ws_sum
    Y.banner(ws, "Hierarchy Audit — every reporting line, checked",
             "Read this first. The model keeps every reporting line exactly as the Employee Master (HRIS, 23 Sep 2026) has it, and "
             "places levels so that every reportee sits at a lower level than their manager.", STAMP, back=False, width_cols=6)
    checked = stat["OK"] + stat["⚠ At or above the manager"]
    loops = 0
    for n in by_row:
        x, steps = n, 0
        while mgr_n.get(x) is not None and steps <= len(by_row):
            x, steps = mgr_n[x], steps + 1
        loops += steps > len(by_row) or mgr_n.get(n) == n
    rows = [("People in the architecture", len(by_row)),
            ("Reporting lines checked · both on the ladder", checked),
            ("Reportee at or above their manager", stat["⚠ At or above the manager"]),
            ("Manager named but not found in the roster", stat["⚠ Manager not found"]),
            ("Top of the organisation · no manager above", stat["Top of the organisation"]),
            ("Employee with no manager on record in the HRIS", stat["No manager on record"]),
            ("Consultant with no manager on record · no Emp ID, outside the ladder", stat["Consultant · no manager on record"]),
            ("Consultant lines · outside the ladder, not levelled", stat["Consultant · outside the ladder"]),
            ("People reporting to themselves or in a loop", loops),
            ("Deepest reporting line · people", max(int(J.text(r[h["Layer from the CEO"]]) or 0) for r in d) + 1),
            ("Level gap manager − reportee · steps: count", " · ".join(f"{g}: {c}" for g, c in sorted(gaps.items()))),
            ("Reporting-line differences between the source files", f"{n_diff} people · sheet 3 · for HR to confirm; nothing changed"),
            ("CARBON's proposed new lines", f"{n_prop} · sheet 4 · not applied")]
    for j, x in enumerate(("Check", "Result"), 1):
        Y.header(ws.cell(5, j, x), Y.NAVY)
    for k, (a, b) in enumerate(rows, 6):
        x = cell(ws, k, 1, a, "F2F2F2", True)
        x.font = Y.font(10, True)
        y = cell(ws, k, 2, b, "FFFFFF")
        if a.startswith(("Reportee at or above", "Manager named", "Employee with no manager")) and b:
            y.fill, y.font = Y.F_RED, Y.T_RED
        ws.merge_cells(start_row=k, start_column=2, end_row=k, end_column=6)
        ws.row_dimensions[k].height = 22
    ws.column_dimensions["A"].width = 52
    for c in "BCDEF":
        ws.column_dimensions[c].width = 20
    for w in wb.worksheets:
        w.sheet_view.zoomScale = 90
    wb.save(OUT)
    print(f"Wrote {OUT}: {dict(stat)} · differences {n_diff} people · CARBON proposals {n_prop}")
    return stat, n_diff, n_prop


if __name__ == "__main__":
    main(sys.argv[1])
