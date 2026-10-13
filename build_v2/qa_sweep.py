"""IRIS Job Architecture v2.2: QA sweep.

Reads the finished Master or Simplified with cached values and checks:
- every sheet agrees with every other sheet: counts, levels, bands, roles, titles, managers and layers
- the baseline state before any leader input
- the Person Card
- privacy
- text glitches

It prints every finding and the figures worth highlighting.

Usage: python3 qa_sweep.py <workbook.xlsx>
"""
import collections
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ja_v114 as V  # noqa: E402

BAND_LABEL = {"A": "Band A", "B": "Band B", "B+": "B+", "C": "Band C", "C+": "C+", "D": "Band D", "E": "Band E", "X": "Band X"}
issues = collections.defaultdict(list)
facts = collections.OrderedDict()


def bad(kind, *what):
    issues[kind].append(what)


def table(ws, hdr_row=6):
    rows = list(ws.iter_rows(min_row=hdr_row, values_only=True))
    h = {}
    for i, x in enumerate(rows[0]):
        if x is not None and x not in h:
            h[x] = i
    return h, [r for r in rows[1:] if any(v not in (None, "") for v in r)]


def t(v):
    return "" if v is None else str(v).strip()


def main(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sn = {re.sub(r"^\d+\. ", "", n): n for n in wb.sheetnames}
    simple = "Role Definitions" not in sn

    # ---------------------------------------------------------------- Talent Mapping
    h, tm = table(wb[sn["Talent Mapping"]])
    people = [r for r in tm if t(r[h["Employee name"]])]
    facts["People on Talent Mapping"] = len(people)
    by_emp = collections.defaultdict(list)
    for r in people:
        by_emp[t(r[h["Emp ID"]])].append(r)
    real_ids = {e: rs for e, rs in by_emp.items() if e.startswith("IRIS")}
    dup = {e: [t(x[h["Employee name"]]) for x in rs] for e, rs in real_ids.items() if len(rs) > 1}
    if dup:
        bad("Duplicate Emp IDs", dup)
    names = collections.Counter(t(r[h["Employee name"]]).lower() for r in people)
    facts["Same name held by two people"] = [n for n, c in names.items() if c > 1]
    emp_row = {t(r[h["Emp ID"]]): r for r in people if t(r[h["Emp ID"]]).startswith("IRIS")}
    lvl_count = collections.Counter()
    key_count = collections.Counter()
    for r in people:
        nm, emp = t(r[h["Employee name"]]), t(r[h["Emp ID"]])
        lv, band = t(r[h["IRIS proposed level"]]), t(r[h["IRIS proposed band"]])
        lvl_count[lv] += 1
        b = V.band(lv)
        if not band.startswith(BAND_LABEL.get(b, "?")):
            bad("Band label does not match the level", nm, lv, band)
        key = t(r[h["Ladder key"]])
        if b != "X":
            key_count[key] += 1
            if key.split("|")[1:2] != [lv] or key.split("|")[0] != t(r[h["Role ID"]]):
                bad("Ladder key does not match role and level", nm, key, lv)
            if t(r[h["Lane"]]) not in ("IC", "Manager", "Shared"):
                bad("Lane missing", nm, t(r[h["Lane"]]))
        # baseline: no leader input → nothing moves
        for c in ("Leader proposed level", "Leader proposed band", "Move vs IRIS", "Score shift · levels", "Level rule · live"):
            if t(r[h[c]]):
                bad("Baseline: a leader column is not blank", nm, c, t(r[h[c]]))
        if t(r[h["Level in force · after the level rule"]]) != lv:
            bad("Baseline: level in force differs from IRIS level", nm, lv, t(r[h["Level in force · after the level rule"]]))
        mc = t(r[h["Manager check · live"]])
        if mc not in ("OK", "—"):
            bad("Baseline: manager check is not OK", nm, mc)
        for p in ("Complexity & Expertise", "Autonomy & Decision", "Accountability & Impact", "Leadership & Strategy", "Business Impact",
                  "Experience", "Education"):
            if t(r[h[f"Weightage in use · {p}"]]) != t(r[h[f"Weightage IRIS · {p}"]]):
                bad("Baseline: weightage in use differs from IRIS base", nm, p)
        # manager
        mid, mname = t(r[h["Manager Emp ID"]]), t(r[h["Reporting manager"]])
        if mid.startswith("IRIS"):
            m = emp_row.get(mid)
            if m is None:
                bad("Manager Emp ID not in the roster", nm, mid)
            else:
                if t(m[h["Employee name"]]).lower() != mname.lower():
                    bad("Manager name does not match the manager's Emp ID", nm, mname, t(m[h["Employee name"]]))
                try:
                    if int(r[h["Layer from the CEO"]]) != int(m[h["Layer from the CEO"]]) + 1:
                        bad("Layer from the CEO is not the manager's + 1", nm, r[h["Layer from the CEO"]], m[h["Layer from the CEO"]])
                except (TypeError, ValueError):
                    bad("Layer from the CEO missing", nm)
                ml = t(m[h["IRIS proposed level"]])
                if b != "X" and V.band(ml) != "X" and V.VAL[lv] >= V.VAL[ml]:
                    bad("Reportee at or above the manager", nm, lv, mname, ml)
        # titles
        cur, std, chg = t(r[h["Current designation"]]), t(r[h["Standard title for the role and level"]]), t(r[h["Title change"]])
        if b != "X" and not std:
            bad("Standard title missing", nm, lv)
        if chg and cur == std:
            bad("Title change shown but the title is unchanged", nm, cur)
        if not chg and std and cur and cur != std and b != "X":
            bad("Title differs from the standard but no change is shown", nm, cur, std)
        if b == "X" and not t(r[h["How IRIS placed them"]]).startswith("Consultant"):
            bad("Consultant not described as a consultant", nm)
        for c in ("Job family", "Role ID", "Role · sub-family"):
            if not t(r[h[c]]):
                bad(f"{c} missing", nm)
    facts["Levels"] = dict(sorted(lvl_count.items(), key=lambda x: -V.VAL.get(x[0], 0)))

    # ---------------------------------------------------------------- Ladder agrees with Talent Mapping
    hl, lad = table(wb[sn["Ladder"]])
    for r in lad:
        lv = t(r[hl["Level"]])
        if lv in V.VAL:
            n = int(r[hl["People · IRIS proposed"]] or 0)
            if n != lvl_count.get(lv, 0):
                bad("Ladder head-count differs from Talent Mapping", lv, n, lvl_count.get(lv, 0))

    # ---------------------------------------------------------------- Role Ladder agrees with Talent Mapping
    hr, rl = table(wb[sn["Role Ladder & Weightage"]])
    rl_keys = {}
    for r in rl:
        k = t(r[hr["Key"]])
        rl_keys[k] = r
        if t(r[hr["Row"]]) == "level":
            n = t(r[hr["People in IRIS"]])
            n = 0 if n in ("", "Vacant") else int(float(n))
            if n != key_count.get(k, 0):
                bad("Role Ladder head-count differs from Talent Mapping", k, n, key_count.get(k, 0))
            if t(r[hr["Level"]]) == "E2" and "Executive" not in t(r[hr["Job family"]]) and "Enterprise" not in t(r[hr["Job family"]]):
                bad("A function's ladder runs to E2", k)
    missing = [k for k in key_count if k not in rl_keys]
    if missing:
        bad("Ladder key with no Role Ladder row", missing[:5])
    for r in people:
        k = t(r[h["Ladder key"]])
        if k in rl_keys and V.band(t(r[h["IRIS proposed level"]])) != "X":
            # the person's role title is the row's title, or the row's title for their own function / POD Lead
            row_title = t(rl_keys[k][hr["Standard title at this level"]])
            if t(r[h["Role · sub-family"]]) != t(rl_keys[k][hr["Role · sub-family"]]):
                bad("Role name differs between Talent Mapping and Role Ladder", t(r[h["Employee name"]]), k)
    for r in rl:
        if t(r[hr["Row"]]) == "ROLE" and "not recorded" in t(r[hr["Role · sub-family"]]).lower():
            bad("A role on the ladder is named 'not recorded'", t(r[hr["Role ID"]]), t(r[hr["Role · sub-family"]]))
    roles_tm = {t(r[h["Role ID"]]) for r in people if t(r[h["Role ID"]])}
    roles_rl = {t(r[hr["Role ID"]]) for r in rl if t(r[hr["Row"]]) == "ROLE"}
    facts["Roles"] = f"{len(roles_rl)} on Role Ladder · {len(roles_tm)} held · {len(roles_rl - roles_tm)} not yet held"
    facts["Role Ladder level rows"] = sum(1 for r in rl if t(r[hr["Row"]]) == "level")
    facts["Vacant level rows"] = sum(1 for r in rl if t(r[hr["Row"]]) == "level" and t(r[hr["People in IRIS"]]) in ("", "Vacant", "0"))
    # weightage shares add up to 100% on every level row
    for r in rl:
        if t(r[hr["Row"]]) != "level":
            continue
        sh = [r[hr[f"Share · {p}"]] for p in ("Complexity & Expertise", "Autonomy & Decision", "Accountability & Impact",
                                              "Leadership & Strategy", "Business Impact", "Experience", "Education")]
        try:
            s = sum(float(x or 0) for x in sh)
            if abs(s - 1) > 0.011 and abs(s - 100) > 1.1:
                bad("Shares do not add up to 100%", t(r[hr["Key"]]), round(s, 3))
        except (TypeError, ValueError):
            bad("Share not a number", t(r[hr["Key"]]), sh)

    # ---------------------------------------------------------------- Role Journey
    hj, rj = table(wb[sn["Role Journey"]])
    ids = collections.Counter(t(r[hj["Role ID"]]) for r in rj if t(r[hj["Role ID"]]))
    if any(c != 2 for c in ids.values()):
        bad("Role Journey: a role without exactly two rows", [k for k, c in ids.items() if c != 2][:5])
    if set(ids) != roles_rl:
        bad("Role Journey and Role Ladder list different roles", sorted(set(ids) ^ roles_rl)[:5])
    for r in rj:
        for lv in V.VAL:
            if lv in hj and "Chief Executive Officer – " in t(r[hj[lv]]):
                bad("Role Journey: a function titled CEO", t(r[hj["Role ID"]]))

    # ---------------------------------------------------------------- Backend people list
    hb = wb[sn["Backend Data"]]
    plist = [r[0] for r in hb.iter_rows(min_row=39, max_row=hb.max_row, values_only=True) if r[0]]
    if len(plist) != len(people):
        bad("Person Card list differs from Talent Mapping", len(plist), len(people))
    if len(set(plist)) != len(plist):
        bad("Person Card list has duplicate entries", [x for x, c in collections.Counter(plist).items() if c > 1][:5])

    # ---------------------------------------------------------------- Raw Data covers everyone
    hw, raw = table(wb[sn["Raw Data"]])
    raw_ids = collections.Counter(t(r[hw["Emp ID"]]).split("\n")[0].strip() for r in raw)
    lost = [e for e in emp_row if raw_ids.get(e, 0) == 0]
    if lost:
        bad("People on Talent Mapping missing from Raw Data", lost[:5])
    facts["Raw Data rows"] = len(raw)
    facts["Raw Data columns"] = len([x for x in hw if x])

    # ---------------------------------------------------------------- privacy and glitches across every sheet
    pii = {"e-mail": re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "phone": re.compile(r"(?<!\d)(\+91[\s-]?)?[6-9]\d{9}(?!\d)"),
           "PAN": re.compile(r"\b[A-Z]{5}\d{4}[A-Z]\b"), "Aadhaar": re.compile(r"\b\d{4}\s\d{4}\s\d{4}\b")}
    bad_hdr = re.compile(r"\b(gender|date of birth|dob|birth|age|e-?mail|mobile|phone|salary|ctc|pay ?scale|compensation|"
                         r"resignation|relieving|aadhaar|pan|address)\b", re.I)
    glitch = {"error value": re.compile(r"^(#(REF|N/A|VALUE|NAME|DIV|NUM|NULL)|Err:\d+)"),
              "stray None/nan": re.compile(r"(^|[\s·:(])(None|nan|NaN|null)([\s·:)]|$)"),
              "empty separator": re.compile(r"·\s*·|\(\s*\)|^\s*·|·\s*$"),
              "placeholder": re.compile(r"^(Not provided|Not rated|Not placed|TBD|TBC|xx+)\b", re.I),
              "unfilled template": re.compile(r"\{[a-z_]+\}|%s|<[A-Z_]+>")}
    for ws in wb.worksheets:
        src_cols = set()
        is_raw = ws.title == sn["Raw Data"]
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            for j, v in enumerate(row):
                if not isinstance(v, str):
                    continue
                if i == 6 and bad_hdr.search(v) and not re.search(r"Comment|address(es)?\b.*(gap|need)", v, re.I):
                    bad("Header that may be a personal field", ws.title, v)
                for k, p in pii.items():
                    m = p.search(v)
                    if m and not (k == "PAN" and ws.title.startswith("References")):
                        bad(f"Possible personal data · {k}", ws.title, f"row {i}", m.group(0))
                for k, p in glitch.items():
                    if p.search(v):
                        # answers people typed are kept as given in Raw Data and the survey column
                        if k in ("stray None/nan", "empty separator") and (is_raw or ws.title == sn["Talent Mapping"] and j in (
                                h.get("Inputs received in the survey"), h.get("KRAs from the manager"),
                                h.get("Evidence used for the ratings"), h.get("Performance · rating · potential · 9-box"))):
                            continue
                        bad(f"Text glitch · {k}", ws.title, f"row {i} col {j + 1}", v[:80])
    return simple


if __name__ == "__main__":
    simple = main(sys.argv[1])
    print(f"QA sweep · {os.path.basename(sys.argv[1])}")
    for k, v in facts.items():
        print(f"  {k}: {v}")
    if not issues:
        print("  No findings.")
    for k, v in issues.items():
        print(f"  ✗ {k}: {len(v)}")
        for x in v[:6]:
            print("      ", x)
    sys.exit(1 if issues else 0)
