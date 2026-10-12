"""v1.14 decisions (HR, 12 October 2026) and the numbers behind them.

1. Titles, from HR's edits to the Simplified ladder table:
   D2 = Senior Vice President / Business Head · D1 = Associate Vice President / Distinguished Principal
   (the D1 specialist title is decided: Distinguished Principal) · B1 = Team Lead / Specialist / Asst Manager.
2. Department and function heads — the top person of each department who leads a team at C1 or above —
   move to job family Executive Leadership, sub-family Business Head.
3. Top-down manager rule · hybrid, HR decision of 12 October 2026: under a head at C1 or above a
   reportee sits in a lower band; below C1 a reportee sits at least one level below the manager.
   Applied from the CEO down. A Band A lead with peer-level reportees becomes B1 Team Lead. Band E
   with Band E is the agreed exception; B+ and C+ count as their own step; consultants sit outside
   the ladder. Moves down keep pay and title for one cycle. The strict band rule was tested first: it
   cascades through chains up to 9 layers deep and left 172 people with no band below their manager.
"""
import collections

import ja_common as J

VAL = {"A1": 1, "A2": 2, "A3": 3, "B1": 4, "B2": 5, "B3": 6, "B+": 6.5, "C1": 7, "C2": 8, "C+": 8.5,
       "D1": 9, "D2": 10, "E1": 11, "E2": 12}
LEVELS_STD = ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "D1", "D2", "E1", "E2"]
LEVELS_PC = ["A1", "A2", "A3", "B1", "B2", "B3", "B+", "C1", "C2", "C+", "D1", "D2", "E1", "E2"]
BAND_ORDER = ["A", "B", "B+", "C", "C+", "D", "E"]
BAND_NAME = {"A": "Band A · Practitioner", "B": "Band B · Specialist / Manager", "B+": "B+ · Player-Coach",
             "C": "Band C · Principal / Head", "C+": "C+ · Player-Coach", "D": "Band D · Business leadership",
             "E": "Band E · Enterprise leadership", "X": "Band X · Consultant"}

# Whole-cell replacements first, then text inside longer cells.
TITLE_CELL = {
    "Associate Vice President / Business Head · Chief Specialist / Distinguished Principal":
        "Associate Vice President / Distinguished Principal",
    "Team Lead / Specialist": "Team Lead / Specialist / Asst Manager",
}
TITLE_REPLACE = [
    ("Chief Specialist / Distinguished Principal", "Distinguished Principal"),
    ("Associate Vice President / Business Head", "Associate Vice President"),
]
# Whole-cell values replaced only inside IRIS title columns.
TITLE_EXACT = {"Vice President": "Senior Vice President / Business Head", "Team Lead": "Team Lead / Asst Manager"}
LEVEL_TITLE = {  # level → (IC title, manager title) after the v1.14 decisions
    "A1": ("Intern / Trainee", "Intern / Trainee"), "A2": ("Associate", "Associate"), "A3": ("Senior Associate", "Senior Associate"),
    "B1": ("Specialist", "Team Lead / Asst Manager"), "B2": ("Senior Specialist", "Manager"), "B3": ("Lead Specialist", "Senior Manager"),
    "B+": ("Practice Lead · Player-Coach", "Practice Lead · Player-Coach"), "C1": ("Principal", "Associate Head"),
    "C2": ("Senior Principal / Architect", "Head of Function"), "C+": ("Group Head · Player-Coach", "Group Head · Player-Coach"),
    "D1": ("Distinguished Principal", "Associate Vice President"), "D2": ("—", "Senior Vice President / Business Head"),
    "E1": ("—", "CXO / President"), "E2": ("—", "Chief Executive Officer"),
}
HEAD_FAMILY, HEAD_SUB = "Executive Leadership", "Business Head"


def band(level):
    s = str(level or "")
    if s in ("B+", "C+"):
        return s
    return s[:1] if s[:1] in "ABCDEX" and s else ""


def is_manager_track(track):
    return J.text(track) in ("People Manager", "Player-Coach", "POD Lead")


class Org:
    """Reporting tree over the 586 people, with the v1.13 model level of each."""

    def __init__(self, M, people, F):
        self.H = H = {h: i for i, h in enumerate(M["tm_hdr"])}
        self.people = people
        self.level = {p["i"]: J.text(p["tm"][H["FINAL LEVEL"]]) for p in people}
        self.mgr = {}
        for p in people:
            m = self._mgr(F, p)
            self.mgr[p["i"]] = m["i"] if m is not None and m is not p else None
        self.kids = collections.defaultdict(list)
        for p in people:
            if self.mgr[p["i"]] is not None:
                self.kids[self.mgr[p["i"]]].append(p["i"])
        self.by_i = {p["i"]: p for p in people}
        self.roots = [p["i"] for p in people if self.mgr[p["i"]] is None]

    def _mgr(self, F, p):
        mn = J.nname(p["tm"][self.H["Reporting manager"]])
        m = F.by_name.get(mn)
        if m is None and mn:
            t = mn.split()
            c = [q for n, q in F.by_name.items() if n.split()[:1] == t[:1] and n.split()[-1:] == t[-1:]]
            m = c[0] if len(c) == 1 else None
        return m

    def top_down(self):
        """Depth-first order from the top: each manager before their reportees, seniors first."""
        out, seen = [], set()

        def visit(i, depth):
            if i in seen:
                return
            seen.add(i)
            out.append((i, depth))
            for k in sorted(self.kids[i], key=lambda k: (-VAL.get(self.level[k], 0), J.text(self.by_i[k]["name"]))):
                visit(k, depth + 1)
        for r in sorted(self.roots, key=lambda k: (-VAL.get(self.level[k], 0), J.text(self.by_i[k]["name"]))):
            visit(r, 0)
        for p in self.people:
            if p["i"] not in seen:
                visit(p["i"], 0)
        return out


def department_heads(M, people, org):
    """Top person of each department who leads a team and sits at C1 or above in v1.13."""
    H = org.H
    deps = collections.defaultdict(list)
    for p in people:
        deps[(p["biz"], J.text(p["tm"][H["Department"]]))].append(p)
    heads = []
    for (b, d), ps in deps.items():
        ids = {p["i"] for p in ps}
        for p in ps:
            if J.text(p["tm"][H["Track"]]) == "Consultant":
                continue
            top = org.mgr[p["i"]] is None or org.mgr[p["i"]] not in ids
            if top and org.kids[p["i"]] and VAL.get(org.level[p["i"]], 0) >= 7:
                heads.append(p)
    return heads


def cap_level(manager_level, person_level):
    """Highest level allowed under a manager: the top regular level of the band below the manager's band."""
    mb = band(manager_level)
    if mb not in BAND_ORDER:
        return None
    k = BAND_ORDER.index(mb)
    if k == 0:
        return None
    pc = person_level in ("B+", "C+")
    allowed = [lv for lv in LEVELS_PC if BAND_ORDER.index(band(lv)) < k and (pc or lv not in ("B+", "C+"))]
    return allowed[-1] if allowed else None


def propose(org):
    """HR proposed level for everyone, top down — the hybrid rule HR chose on 12 October 2026.

    Heads and above · manager at C1 or higher: the reportee sits in a lower band than the manager.
    Below C1: the reportee sits at least one level below the manager.
    A Band A lead whose reportees sit at or above their own level becomes B1 Team Lead, where the
    lead's own manager sits above B1. Band E with Band E is the agreed exception.
    Returns {i: dict(level, note, flag, depth, mgr_level, check)}.
    """
    out, new = {}, {}
    for i, depth in org.top_down():
        lv = org.level[i]
        m = org.mgr[i]
        notes = []
        target = lv
        ml = new.get(m, org.level[m]) if m is not None else None
        if band(lv) != "X" and ml and band(ml) not in ("X", ""):
            mb, pb = band(ml), band(lv)
            lad = LEVELS_PC if lv in ("B+", "C+") else LEVELS_STD
            if mb == "E" and pb == "E":
                notes.append(f"Band E with Band E · agreed exception · manager {ml}")
            elif VAL[ml] >= 7:
                if BAND_ORDER.index(pb) >= BAND_ORDER.index(mb):
                    target = cap_level(ml, lv) or lv
                    notes.append(f"Band rule: manager {ml} is a head in {BAND_NAME[mb]}; a reportee sits in a lower band · {lv} → {target}")
            elif VAL[lv] >= VAL[ml]:
                below = [x for x in lad if VAL[x] < VAL[ml]]
                if below:
                    target = below[-1]
                    notes.append(f"Level rule: manager {ml}; a reportee sits at least one level below · {lv} → {target}")
        if band(target) == "A":
            reps = [k for k in org.kids[i] if band(org.level[k]) not in ("X", "")]
            if reps and any(VAL[org.level[k]] >= VAL[target] for k in reps) and (ml is None or VAL.get(ml, 0) > 4):
                notes.append(f"Leads people at or above their own level · {target} → B1 Team Lead")
                target = "B1"
        if target != lv and VAL.get(target, 0) < VAL.get(lv, 0):
            notes.append("move down · pay and title protected for one cycle")
        new[i] = target
        if band(lv) == "X":
            notes = ["Consultant · outside the ladder"]
        out[i] = {"level": target, "note": " · ".join(notes) or ("No change · already below the manager" if m is not None else "Top of the reporting line"),
                  "flag": "", "depth": depth}
    for i, d in out.items():
        m = org.mgr[i]
        d["mgr_level"] = out[m]["level"] if m is not None else ""
        a, b = d["level"], d["mgr_level"]
        if m is None or "X" in (band(a), band(b)) or not band(a) or not band(b):
            d["check"] = "—"
        elif band(a) == "E" and band(b) == "E":
            d["check"] = "Exception · Band E"
        elif VAL[b] >= 7 and BAND_ORDER.index(band(a)) >= BAND_ORDER.index(band(b)):
            d["check"] = "Same band as a head manager"
        elif VAL[a] >= VAL[b]:
            d["check"] = "At or above the manager's level"
        else:
            d["check"] = "OK"
    return out


def level_title(level, track):
    t = LEVEL_TITLE.get(level)
    if not t:
        return ""
    return t[1] if is_manager_track(track) else t[0]


# ---------------------------------------------------------------- heads → Executive Leadership · Business Head
HEAD_PURPOSE = ("Heads a department or function: sets its direction, owns its results, budget, people and risk, "
                "and speaks for it in the business leadership team.")


def apply_heads(M, people, F_prev, heads):
    """Move heads to Executive Leadership · Business Head in the person records and add the catalogue roles they now hold.
    Returns {i: (old family, old sub-family)}."""
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    ch = M["cat_hdr"]
    cj = {h: j for j, h in enumerate(ch)}
    fj, sj = H["Job family"], H["Sub-family · what they do"]
    changed = {}
    by_level = collections.defaultdict(list)
    for p in heads:
        old = (J.text(p["tm"][fj]), J.text(p["tm"][sj]))
        changed[p["i"]] = old
        by_level[J.text(p["tm"][H["FINAL LEVEL"]])].append((p, F_prev.cat_of.get(p["i"])))
        p["tm"] = list(p["tm"])
        p["tm"][fj], p["tm"][sj] = HEAD_FAMILY, HEAD_SUB
    ladder = {J.text(r[1]): r for r in M.get("f4", [])}
    for lv, items in sorted(by_level.items(), key=lambda x: VAL.get(x[0], 0)):
        row = [None] * len(ch)
        def put(h, v):
            if h in cj:
                row[cj[h]] = v
        put(ch[0], f"RC-EL-BH-{lv.replace('+', 'P')}")
        put("Job family", HEAD_FAMILY)
        put("Sub-family", HEAD_SUB)
        put("Level", lv)
        put("Level title", LEVEL_TITLE.get(lv, ("", ""))[1])
        put("Track", "Manager")
        ps = [p for p, _ in items]
        put("Designations used at IRIS", "\n".join(sorted({J.text(p["tm"][H["Final designation"]]) for p in ps})))
        put("Holders", len(ps))
        put("Who", ", ".join(J.text(p["name"]) for p in ps))
        put("Business", ", ".join(sorted({p["biz"] for p in ps})))
        put("Purpose", HEAD_PURPOSE)
        # carry every descriptive field of the roles they held before, so nothing is lost
        for h in ch[13:]:
            if h in ("Purpose",):
                continue
            vals = []
            for _, c in items:
                if c is not None and J.has(c[cj[h]]) and J.text(c[cj[h]]) not in vals:
                    vals.append(J.text(c[cj[h]]))
            if vals:
                row[cj[h]] = "\n".join(vals) if len(vals) == 1 else "\n".join("• " + v for v in vals)
        M["cat"].append(row)
    return changed


TITLE_SKIP = {"1a. Placements", "8. Talent, Learning & Mobility", "Card Lists", "Chart data"}


def retitle(wb):
    """Apply the v1.14 title decisions to every IRIS title in a workbook. Raw data and people's designations are left as received."""
    n = collections.Counter()
    for ws in wb.worksheets:
        if ws.title in TITLE_SKIP:
            continue
        rows = list(ws.iter_rows())
        only_col = None
        if ws.title == "2. Talent Mapping":
            hdr = {J.text(c.value): c.column for c in rows[5]}
            only_col = hdr.get("Level title")
        col_level = {}
        for row in rows:
            vals = [J.text(c.value) for c in row]
            first = vals[0] if vals else ""
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                if ws.title == "3. Raw Data & Market Standards" and (c.column > 22 or c.row < 7):
                    continue   # raw data stays as received; only the market-standard columns carry IRIS titles
                if only_col and c.row > 6 and c.column != only_col:
                    if not any(o in v for o, _ in TITLE_REPLACE):
                        continue
                new = TITLE_CELL.get(v.strip())
                if new is None:
                    new = v
                    for o, r in TITLE_REPLACE:
                        if o in new and (only_col is None or c.column == only_col or c.row <= 6 or True):
                            new = new.replace(o, r)
                for exact, lv in (("Vice President", "D2"), ("Team Lead", "B1")):
                    if v.strip() == exact and (only_col is None or c.column == only_col):
                        if lv in vals or (first.startswith("IRIS — ") and col_level.get(c.column) == lv):
                            new = TITLE_EXACT[exact]
                if new != v:
                    c.value = new
                    n[ws.title] += 1
            if first.startswith("Ladder / company"):
                col_level = {c.column: J.text(c.value) for c in row}
    return n
