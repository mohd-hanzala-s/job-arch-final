"""Every bit of data held on each person, from every source, joined for the Raw Data sheet.

Sources, newest first where the same field appears in more than one:
  · the five manager talent-mapping uploads (CARBON Mapping Master, CARBON Raw, iFile, iDEAL, Database Content, HR)
  · the v1.15 HR Super Master, sheet 3 (the uploads merged with the v1 raw data)
  · the six v1 workbooks IRIS_<Business>_Talent_Mapping.xlsx, sheet 2. Raw Data — the How We Grow survey, all 20
    questions and the earlier survey round, with the HRIS record and the bracketed source columns
  · the Employee Master (HRIS) of 23 Sep 2026, and the older Employee Master — its differing values are kept apart
  · the CARBON upload's per-person extras: proposed reporting change, FY27 leadership scorecard, open hiring items

Personal fields (gender, date of birth, age, e-mail, mobile, resignation and relieving dates) are never read into the
record, nor are processing notes (serial numbers, source-file names, how a survey response was matched).
Values are kept as received. Where sources disagree, the newest is shown first and the others follow, so nothing is lost.
"""
import datetime
import glob
import os
import re

import openpyxl

import ja_common as J

V1_DIR = os.environ.get("JA_V1", "/home/user/job-arc")
SURVEY_Q = ["What region do you serve?", "What is this role there to do (core purpose)?", "What do you produce or deliver?",
            "What are you personally accountable for?", "What can you decide and act on without asking anyone?",
            "How heavy is the load in a typical week (1-5)?", "Top 5 responsibilities", "What's different about this role?",
            "Is the way this is measured (KPI) what you expected?", "Is this something only you do?",
            "What happens when something is missed or done wrong?", "What skills or technical knowledge does this role need?",
            "What tools, systems, or platforms are used?", "A typical judgment call / way of working",
            "What consistently slows you down?", "If you could change one thing about how this role is set up...",
            "What matters most in the first 30/60/90 days?", "What training, mentoring, or experience made a difference?",
            "What would you need to learn to move into the next role?", "Is anything about this role changing?"]
SURVEY_META = ["Total Question Answered", "Status", "Status Detail"]
ADDITIONAL = "Additional Input (Earlier/Alternate Survey Round)"
DROP = J.PERSONAL | {"srno", "sourcefile", "mobilenumber", "dob", "newemployeeid", "employeeid"}
ALIAS = {"correctdesignationbasedonwork": "correctdesignation", "correctdesignationinlinewithcurrentrole": "correctdesignation",
         "nameofsuccessor": "nameofsuccessoridentified", "quarterlyperformancerating15": "quarterlyperformancerating",
         "linemanager": "reportingmanager", "ncategory": "nlevels"}
GROUPS = [("HRIS", "HR RECORD — Employee Master · HRIS, 23 Sep 2026", "1F3864"),
          ("SURVEY", "HOW WE GROW SURVEY — all 20 questions, as answered", "7030A0"),
          ("MM", "CARBON MAPPING MASTER — as received", "843C0C"),
          ("MGR", "MANAGER TALENT MAPPING — KRAs, ratings, potential, succession, training, as received", "C55A11"),
          ("CARBON", "CARBON — reporting change, FY27 leadership scorecard, open items", "375623"),
          ("OLD", "EMPLOYEE MASTER · OLDER COPY — only where it differs", "595959")]
GROUP_NAME = {g: n for g, n, _ in GROUPS}


def fkey(h):
    k = J.key(re.sub(r"^\[[^\]]*\]\s*", "", J.text(h)))
    return ALIAS.get(k, k)


def pretty(h):
    return re.sub(r"\s+", " ", re.sub(r"^\[[^\]]*\]\s*", "", J.text(h))).strip()


def source_group(h):
    s = J.text(h)
    if not s.startswith("["):
        return None
    if "Mapping Master" in s:
        return "MM"
    return "MGR"


def val(v, h=""):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%d %b %Y") if (v.hour, v.minute) == (0, 0) else v.strftime("%d %b %Y %H:%M")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if not isinstance(v, str):
        return v
    s = v.strip()
    if h == ADDITIONAL:
        s = re.sub(r"^\[[^\]]*\]\s*", "", s)          # drop the note naming the file and the row it was matched from
    if re.fullmatch(r"\d{4}-\d{2}-\d{2} 00:00:00", s):
        s = datetime.datetime.strptime(s[:10], "%Y-%m-%d").strftime("%d %b %Y")
    return s if J.has(s) else ""


def norm(v):
    return re.sub(r"\s+", " ", J.text(v).replace("\u200b", "")).strip().lower()


class Record:
    """One person's fields: (group, key) → [values in source priority], plus the order fields were first seen."""
    def __init__(self):
        self.f = {}

    def add(self, group, h, v, prio, label=None):
        k = fkey(h)
        if not k or k in DROP:
            return
        v = val(v, J.text(h).strip())
        if v == "":
            return
        self.f.setdefault((group, k), []).append((prio, v, label))

    def get(self, group, k):
        vs = sorted(self.f.get((group, k), []), key=lambda x: x[0])
        out, seen = [], []
        for _, v, lab in vs:
            n = norm(v)
            if any(n == s or (len(n) > 12 and n in s) for s in seen):
                continue
            seen.append(n)
            out.append((v, lab))
        if not out:
            return ""
        if len(out) == 1:
            return out[0][0]
        first = J.text(out[0][0])
        return first + "".join(f"\n‖ also on file{(' · ' + lab) if lab else ''}: {J.text(v)}" for v, lab in out[1:])


def _rows(ws, hrow):
    if isinstance(hrow, str):   # find the header row by a column it must have
        hrow = next((n for n, r in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1)
                     if any(J.text(x) == hrow for x in r)), 1)
    hdr = None
    blank = 0
    for n, r in enumerate(ws.iter_rows(min_row=hrow, values_only=True)):
        if n == 0:
            hdr = [J.text(x) for x in r]
            continue
        if not any(x not in (None, "") for x in r):
            blank += 1
            if blank > 50:
                break
            continue
        blank = 0
        yield hdr, r


class Joiner:
    def __init__(self, people, unmapped):
        self.people = people
        self.by_id = {}
        for p in people:
            for x in [p["emp"]] + sorted(p["src_ids"]):
                if J.key(x):
                    self.by_id.setdefault(J.key(x), ("p", p["i"]))
        self.un = {}
        for u in unmapped:
            kk = J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"])
            self.un.setdefault(kk, ("u", kk))
            if u["rec"]["id"]:
                self.by_id.setdefault(J.key(u["rec"]["id"]), ("u", kk))
        self.un_names = {J.nname(u["rec"]["name"]): ("u", J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"])) for u in unmapped}

        self.extra = {}

    def find(self, emp, name, add=True, biz=None):
        """The person a source row belongs to; a row for someone outside the architecture gets its own ('x', …) record.
        Two people with the same name are told apart by business, then by whether either has an Emp ID."""
        if J.key(emp) and J.key(emp) in self.by_id:
            return self.by_id[J.key(emp)]
        if J.has(name):
            p = J._name_match(name, self.people)
            if p is None:
                same = [q for q in self.people if J.nname(q["name"]) == J.nname(name)]
                if len(same) > 1 and biz:
                    same = [q for q in same if J.text(q["biz"]).lower() == J.text(biz).lower()] or same
                if len(same) > 1 and not J.key(emp):
                    same = [q for q in same if not J.text(q["emp"]).startswith(("IRIS", "INTN"))] or same
                p = same[0] if len(same) == 1 else None
            if p is not None:
                return ("p", p["i"])
            if J.nname(name) in self.un_names:
                return self.un_names[J.nname(name)]
        k = J.key(emp) if J.has(emp) else ""
        n = J.nname(name) if J.has(name) else ""
        who = self.extra.get(("id", k)) if k else None
        who = who or (self.extra.get(("nm", n)) if n else None)
        if who is None and add and (k or n):
            who = ("x", k or n)
        if who is not None:
            if k:
                self.extra.setdefault(("id", k), who)
            if n:
                self.extra.setdefault(("nm", n), who)
        return who


def v1_files(v1dir=V1_DIR):
    return sorted(glob.glob(os.path.join(v1dir, "IRIS_*_Talent_Mapping.xlsx")))


def v1_rows(v1dir=V1_DIR):
    """(business, header, row) for every person row in the six v1 '2. Raw Data' sheets."""
    for f in v1_files(v1dir):
        biz = os.path.basename(f).split("_")[1]
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for hdr, r in _rows(wb["2. Raw Data"], 4):
            j = {h: i for i, h in enumerate(hdr)}
            if J.has(r[j["Emp Id"]]) or J.has(r[j["Employee Name"]]):
                yield biz, hdr, r
        wb.close()


def em_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        for hdr, r in _rows(ws, 1):
            j = {h.strip(): i for i, h in enumerate(hdr)}
            idc = next((j[c] for c in ("Emp Id", "New Employee ID", "Employee ID") if c in j), None)
            emp = J.text(r[idc]) if idc is not None else ""
            name = J.text(r[j["Employee Name"]])
            if J.has(emp) or J.has(name):
                yield ws.title, hdr, r, emp, name
    wb.close()


def em_biz(hdr, r):
    j = {J.text(h).strip(): i for i, h in enumerate(hdr)}
    return J.text(r[j["Product/Others"]]) if "Product/Others" in j else None


def collect(srcdir, people, unmapped, s3h, s3rows, v1dir=V1_DIR):
    """{('p', i) | ('u', key): Record} and the ordered column list [(group, key, header)]."""
    jn = Joiner(people, unmapped)
    recs = {}
    order = {g: [] for g, _, _ in GROUPS}
    names = {}

    def rec(who):
        return recs.setdefault(who, Record())

    def note(group, h):
        k = fkey(h)
        if k and k not in DROP and (group, k) not in names:
            names[(group, k)] = pretty(h)
            order[group].append(k)

    def put(who, group, h, v, prio, label=None):
        if who is None:
            return
        note(group, h)
        rec(who).add(group, h, v, prio, label)

    found = {"v1": 0, "em_new": 0, "em_old": 0}
    # 1 · Employee Master 23 Sep 2026 — the HR record
    em_new = glob.glob(os.path.join(v1dir, "Employee Master_New_*.xlsx"))
    for path in em_new:
        for sheet, hdr, r, emp, name in em_rows(path):
            who = jn.find(emp, name, biz=em_biz(hdr, r))
            if who:
                found["em_new"] += 1
            for h, v in zip(hdr, r):
                if h:
                    put(who, "HRIS", "Emp ID" if h.strip() in ("Emp Id", "New Employee ID", "Employee ID") else h, v, 1)
            put(who, "HRIS", "HRIS sheet", sheet, 1)
    # 2 · uploads (newest), v1.15 sheet 3, v1 raw — the HRIS fields fill in behind the Employee Master
    for p in people:
        who = ("p", p["i"])
        put(who, "HRIS", "Business", p["biz"], 0)
        put(who, "HRIS", "Emp ID", p["emp"], 0)
        put(who, "HRIS", "Employee Name", p["name"], 0)
        for kind, items in p["src"].items():
            for s, r in items:
                for h, v in r["row"].items():
                    g = "MM" if kind == "MM" else "MGR"
                    if fkey(h) in ("empid", "employeename"):
                        put(who, "HRIS", "Emp ID" if fkey(h) == "empid" else "Employee Name", v, 5, s["label"])
                        continue
                    put(who, g, h, v, 2, s["file"].rsplit(".", 1)[0].replace("_", " "))
        row = s3rows[p["i"]]
        on = False
        for h, j in sorted(s3h.items(), key=lambda x: x[1]):
            if h == "Sr No":
                on = True
            if not on or not h:
                continue
            g = source_group(h)
            if g is None:
                g = "SURVEY" if (h in SURVEY_META or h in SURVEY_Q) else "HRIS"
                if h in ("Emp Id · 2", "Employee Name · 2", "Department · 2"):
                    h = h.replace(" · 2", "")
            elif fkey(h) in ("empid", "employeename"):
                continue
            put(who, g, h, row[j], 3)
    for u in unmapped:
        who = jn.un.get(J.key(u["rec"]["id"]) or J.nname(u["rec"]["name"]))
        g = "MM" if u["source"]["kind"] == "MM" else "MGR"
        for h, v in u["rec"]["row"].items():
            if fkey(h) not in ("empid", "employeename"):
                put(who, g, h, v, 2, u["source"]["file"].rsplit(".", 1)[0].replace("_", " "))
            else:
                put(who, "HRIS", "Emp ID" if fkey(h) == "empid" else "Employee Name", v, 5, u["source"]["label"])
        put(who, "HRIS", "Emp ID", u["rec"]["id"], 1)
        put(who, "HRIS", "Employee Name", u["rec"]["name"], 1)
    for biz, hdr, r in v1_rows(v1dir):
        j = {h: i for i, h in enumerate(hdr)}
        who = jn.find(r[j["Emp Id"]], r[j["Employee Name"]], biz=biz)
        if who:
            found["v1"] += 1
        for h, v in zip(hdr, r):
            if not h or h == "None":
                continue
            g = source_group(h)
            if g is None:
                g = "SURVEY" if (h in SURVEY_META or h in SURVEY_Q or h == ADDITIONAL) else "HRIS"
            elif fkey(h) in ("empid", "employeename"):
                put(who, "HRIS", "Emp ID" if fkey(h) == "empid" else "Employee Name", v, 5, f"v1 {biz} workbook")
                continue
            put(who, g, h, v, 4, f"v1 {biz} workbook")
    # 3 · the older Employee Master — kept apart, shown only where it differs
    for path in glob.glob(os.path.join(v1dir, "Employee Master_OLD*.xlsx")):
        for sheet, hdr, r, emp, name in em_rows(path):
            who = jn.find(emp, name, biz=em_biz(hdr, r))
            if who:
                found["em_old"] += 1
            if who and who[0] == "x":
                put(who, "HRIS", "Emp ID", emp, 6, "older Employee Master")
                put(who, "HRIS", "Employee Name", name, 6, "older Employee Master")
            for h, v in zip(hdr, r):
                if h and h.strip() not in ("Emp Id", "New Employee ID", "Employee ID", "Employee Name"):
                    put(who, "OLD", h, v, 1)
    # 4 · CARBON upload extras
    carbon = next(iter(glob.glob(os.path.join(srcdir, "*CARBON_Talent_Mapping*.xlsx"))), None)
    if carbon:
        wb = openpyxl.load_workbook(carbon, read_only=True, data_only=True)
        if "Sheet1" in wb.sheetnames:
            for hdr, r in _rows(wb["Sheet1"], 1):
                j = {h: i for i, h in enumerate(hdr)}
                if "Change Reporting" in j:
                    put(jn.find(r[j["Emp Number"]], r[j["Employee Name"]]), "CARBON", "Proposed new reporting manager · CARBON file",
                        r[j["Change Reporting"]], 1)
        first = {}
        for p in people:
            if p["biz"] == "CARBON":
                first.setdefault(J.nname(p["name"]).split()[0], []).append(p["i"])
        if "Leadership Scorecard" in wb.sheetnames:
            for hdr, r in _rows(wb["Leadership Scorecard"], "Leader"):
                j = {h: i for i, h in enumerate(hdr)}
                lead = J.nname(r[j["Leader"]]).split()[0] if J.has(r[j["Leader"]]) else ""
                if len(first.get(lead, [])) == 1 and J.has(r[j["Function & mandate"]]):
                    txt = " · ".join(f"{h}: {val(r[i])}" for h, i in j.items() if h and h != "Leader" and val(r[i]) != "")
                    put(("p", first[lead][0]), "CARBON", "FY27 leadership scorecard · CARBON", txt, 1)
        if "Hiring & Risk Tracker" in wb.sheetnames:
            items = {}
            for hdr, r in _rows(wb["Hiring & Risk Tracker"], "Owner"):
                j = {h: i for i, h in enumerate(hdr)}
                lead = J.nname(r[j["Owner"]]).split()[0] if J.has(r[j["Owner"]]) else ""
                if len(first.get(lead, [])) == 1:
                    items.setdefault(first[lead][0], []).append(" · ".join(f"{h}: {val(r[i])}" for h, i in j.items() if h and h != "Owner" and val(r[i]) != ""))
            for i, xs in items.items():
                put(("p", i), "CARBON", "Open hiring and risk items owned · CARBON", "\n".join(f"• {x}" for x in xs), 1)
        wb.close()
    # columns: survey questions in the survey's order; the older master only where it differs from the current record
    order["SURVEY"] = [fkey(h) for h in SURVEY_META] + [fkey(q) for q in SURVEY_Q] + [fkey(ADDITIONAL)] + \
                      [k for k in order["SURVEY"] if k not in {fkey(h) for h in SURVEY_META + SURVEY_Q + [ADDITIONAL]}]
    for n, q in enumerate(SURVEY_Q, 1):
        names[("SURVEY", fkey(q))] = f"Q{n} · {q}"
    names[("SURVEY", fkey("Total Question Answered"))] = "Survey questions answered · of 20"
    names[("SURVEY", fkey("Status"))] = "Survey status"
    names[("SURVEY", fkey("Status Detail"))] = "Survey status detail"
    names[("SURVEY", fkey(ADDITIONAL))] = "Earlier / alternate survey round · their answers"
    old_keep = []
    for k in order["OLD"]:
        if any(r.get("OLD", k) and norm(r.get("OLD", k)) != norm(r.get("HRIS", k)) for r in recs.values()):
            old_keep.append(k)
    order["OLD"] = old_keep
    for k in old_keep:
        names[("OLD", k)] = names[("OLD", k)] + " · older master"
    hris_first = ["business", "empid", "employeename", "employeecategory", "employeestatus", "doj", "designation", "highestqualifications",
                  "reportingmanager", "nlevels", "managerorindividualcontributor", "businessline", "productothers", "department",
                  "subfunctionalcategorization", "coredevelopment", "functionalcategorization", "skillsets", "technicalnontechnical",
                  "expoutsideiris", "expcurrentiris", "totalexperience", "avgexpiniris"]
    order["HRIS"] = [k for k in hris_first if k in order["HRIS"]] + [k for k in order["HRIS"] if k not in hris_first]
    names[("HRIS", "empid")] = "Emp ID"
    cols = [(g, k, names[(g, k)]) for g, _, _ in GROUPS for k in order[g]
            if any(r.get(g, k) for r in recs.values())]
    seen, tag = set(), {"MM": "mapping master", "MGR": "manager file", "CARBON": "CARBON file", "SURVEY": "survey"}
    for n, (g, k, h) in enumerate(cols):   # one name per column, so filters and lookups are never ambiguous
        if h in seen:
            cols[n] = (g, k, f"{h} · {tag.get(g, g.lower())}")
        seen.add(cols[n][2])
    return recs, cols, found


def survey_of(record):
    """(answered, [(question, answer)], earlier-round text) for the Talent Mapping summary."""
    n = J.text(record.get("SURVEY", fkey("Total Question Answered")))
    try:
        answered = int(float(n.split("\n")[0])) if n else 0
    except ValueError:
        answered = 0
    qa = [(q, record.get("SURVEY", fkey(q))) for q in SURVEY_Q]
    qa = [(q, a) for q, a in qa if J.has(a)]
    return answered, qa, J.text(record.get("SURVEY", fkey(ADDITIONAL)))
