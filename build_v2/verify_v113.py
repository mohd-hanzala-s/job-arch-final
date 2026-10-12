"""Checks for the two v1.13 workbooks. Exit code 0 = every check passes.

Usage: python3 verify_v113.py
"""
import collections
import os
import re
import sys
import zipfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_common as J  # noqa: E402

SUPER = os.path.join(J.OUT, "IRIS_JA_2_HR_Super_Master_v1.13.xlsx")
SIMPL = os.path.join(J.OUT, "IRIS_JA_2S_HR_Master_Simplified_v1.13.xlsx")
RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append(ok)
    print(("PASS " if ok else "FAIL ") + name + (f" — {detail}" if detail else ""))


def num_eq(a, b):
    return isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-6


def main():
    orig = openpyxl.load_workbook(J.MASTER, read_only=True, data_only=True)
    sup = openpyxl.load_workbook(SUPER, read_only=True, data_only=True)
    sim = openpyxl.load_workbook(SIMPL, read_only=True, data_only=True)

    # 1 every original sheet kept, every original value kept (sheet 3 restorations and stamps aside)
    missing = [n for n in orig.sheetnames if n not in sup.sheetnames]
    check("Super Master keeps every v1.12.2 sheet", not missing, f"{len(orig.sheetnames)} sheets · missing {missing}")
    changed = collections.Counter()
    for n in orig.sheetnames:
        a = list(orig[n].iter_rows(values_only=True))
        b = list(sup[n].iter_rows(values_only=True))
        for i, ra in enumerate(a):
            rb = b[i] if i < len(b) else ()
            for j, v in enumerate(ra):
                w = rb[j] if j < len(rb) else None
                if v != w and not num_eq(v, w):
                    if i < 3 and isinstance(v, str) and "Final v1.12" in v:
                        continue
                    changed[n] += 1
    allowed = {"3. Raw Data & Market Standards"}
    bad = {k: v for k, v in changed.items() if k not in allowed}
    check("Super Master: every v1.12.2 value unchanged, except values restored on sheet 3", not bad,
          f"restored cells on sheet 3: {changed.get('3. Raw Data & Market Standards', 0)} · other changes {bad}")

    # 2 new columns present and filled
    tm = list(sup["2. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    h = {x: i for i, x in enumerate(tm[0])}
    need = ["KRAs set by manager · annual", "Survey questions answered · of 20", "Role mandate & core purpose · as shared",
            "Survey status", "Role ID · catalogue", "Data quality", "Weight sensitivity"]
    check("Super Master sheet 2 carries the KRA, survey and link columns", all(x in h for x in need), f"{len(tm[0])} columns")
    rows = [r for r in tm[1:] if r[1] or r[2]]
    kra = sum(1 for r in rows if not str(r[h["KRAs set by manager · annual"]] or "").startswith("KRAs not received"))
    srv = sum(1 for r in rows if str(r[h["Survey status"]] or "").startswith("Per-person"))
    check("KRAs received shown per person; missing ones marked red with who supplies them", kra > 400, f"{kra} of {len(rows)} on file")
    check("Survey summary per person where the source holds one", srv == 205, f"{srv} CARBON people")
    check("Every person has a Role ID", all(r[h["Role ID · catalogue"]] for r in rows), f"{len(rows)} people")

    # 3 privacy and pay
    for label, wb in (("Super Master", sup), ("Simplified", sim)):
        bad_cols = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=8, values_only=True):
                for v in row:
                    if isinstance(v, str) and J.hkey(v) in J.PERSONAL and len(v) < 30:
                        bad_cols.append((ws.title, v))
        check(f"{label}: no personal-field columns · gender, DoB, age, e-mail, resignation, relieving", not bad_cols, str(bad_cols[:5]))
    t = list(sim["2. Talent Mapping"].iter_rows(min_row=6, values_only=True))
    hs = {x: i for i, x in enumerate(t[0])}
    pay = [x for x in t[0] if x and ("pay" in str(x).lower() and "CTC" in str(x) or str(x).startswith("Proposed range") or x == "Compa-ratio")]
    filled = sum(1 for r in t[1:] for x in pay if r[hs[x]] not in (None, ""))
    check("Simplified: pay columns blank", filled == 0, f"{len(pay)} columns, {filled} filled")

    # 4 simplified = super master, person by person
    srows = [r for r in t[1:] if r[hs["Emp ID"]] or r[hs["Employee name"]]]
    check("Simplified sheet 2 has every person, in the same order", [r[hs["Emp ID"]] for r in srows] == [r[1] for r in rows], f"{len(srows)} people")
    mism = 0
    scored = 0
    for a, b in zip(rows, srows):
        sa, sb = a[h["WEIGHTED SCORE = Σ points"]], b[hs["WEIGHTED SCORE = Σ points"]]
        if isinstance(sa, (int, float)):
            scored += 1
            if not (isinstance(sb, (int, float)) and abs(sa - sb) < 0.005):
                mism += 1
        for k in ("Level the score indicates", "FINAL LEVEL", "Final designation"):
            if a[h[k]] != b[hs[k]]:
                mism += 1
    check("Simplified: live scores, indicated level, final level and designation equal the Super Master", mism == 0, f"{scored} scored · {mism} differ")
    lost = sum(1 for k, i in h.items() if k in hs for a, b in zip(rows, srows) if a[i] not in (None, "") and b[hs[k]] in (None, ""))
    absent = [k for k in h if k not in hs]
    check("Simplified: every Super Master sheet-2 column and value is on sheet 2", not absent and lost == 0, f"missing columns {absent} · lost values {lost}")
    A8 = list(sup["8. Talent, Learning & Mobility"].iter_rows(min_row=6, values_only=True))
    lost8 = 0
    idx = {(str(r[hs["Emp ID"]]), str(r[hs["Employee name"]])): r for r in srows}
    for r in A8[1:]:
        if r[1] or r[2]:
            q = idx[(str(r[1]), str(r[2]))]
            lost8 += sum(1 for j, k in enumerate(A8[0]) if k in hs and r[j] not in (None, "") and q[hs[k]] in (None, ""))
    check("Simplified: every sheet-8 talent, competency, learning and mobility value is on sheet 2", lost8 == 0, f"lost {lost8}")
    RA = list(sup["3. Raw Data & Market Standards"].iter_rows(min_row=6, values_only=True))
    RB = list(sim["3. Raw Data"].iter_rows(min_row=6, values_only=True))
    lostr = sum(1 for a, b in zip(RA[1:], RB[1:]) for j, v in enumerate(a) if v not in (None, "") and (j + 1 >= len(b) or b[j + 1] in (None, "")))
    status = collections.Counter(r[0] for r in RB[1:] if r[0])
    check("Simplified: every raw value is on sheet 3, plus the people not yet mapped", lostr == 0, f"lost {lostr} · {dict(status)}")
    rp = [r for r in sim["5. Role Profiles"].iter_rows(min_row=7, values_only=True) if r[0]]
    hp = {x: i for i, x in enumerate(next(sim["5. Role Profiles"].iter_rows(min_row=6, max_row=6, values_only=True)))}
    holders = sum(len(str(r[hp["Holders · name · Emp ID · business"]] or "").split("\n")) for r in rp if r[hp["Holders · name · Emp ID · business"]])
    check("Role Profiles: every catalogue role, every holder", len(rp) == 175 and holders == len(rows), f"{len(rp)} roles · {holders} holders")
    pc = list(sim["4. Person Card"].iter_rows(min_row=8, max_row=8, values_only=True))[0]
    check("Person Card resolves the default person", isinstance(pc[1], str) and "IRIS" in pc[1], str(pc[1]))

    # 5 workbook hygiene
    for label, path in (("Super Master", SUPER), ("Simplified", SIMPL)):
        z = zipfile.ZipFile(path)
        dv_eq = 0
        f_err = 0
        brackets = 0
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                x = z.read(n).decode("utf8")
                dv_eq += len(re.findall(r"<formula1>=", x))
                f_err += len(re.findall(r't="e"', x))
        check(f"{label}: no '=' at the start of a list validation · Excel repair rule", dv_eq == 0, f"{dv_eq}")
        check(f"{label}: no formula returns an error", f_err == 0, f"{f_err}")
        charts = [n for n in z.namelist() if n.startswith("xl/charts/chart")]
        check(f"{label}: charts present", len(charts) >= (8 if label == "Super Master" else 1), f"{len(charts)}")
    new_text = 0
    for ws in sim.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and not v.startswith("=") and re.search(r"[()]", v):
                    new_text += 1
    check("Simplified: no text in round brackets", new_text == 0, f"{new_text} cells")
    print(f"\n{sum(RESULTS)} of {len(RESULTS)} checks pass")
    return 0 if all(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
