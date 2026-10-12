"""Per-person fields added in v1.13: KRAs received from managers, what the employee
shared in the survey, data quality, links, and the per-person lines folded in from
sheets 1a, 8, 12–17, T2 and T3 of the HR Super Master."""
import collections
import re

from ja_common import clean, has, key, nname, text, alnum, debracket, PREFIX, hkey

MISSING_SURVEY = "Full survey answers not in the files received · How We Grow export needed · HR"
MISSING_KRA = "KRAs not received · manager to supply"

# New columns, in order: (group, header, width, kind) — kind 'red' marks a cell red when it holds a 'needed' text.
KRA_COLS = [
    ("KRAs set by manager · annual", 60),
    ("KRAs · quarterly", 55),
    ("KRAs · self-reported annual", 55),
    ("Q1 tentative rating · scale 1-5", 12),
    ("Actual JD · from manager", 45),
    ("KRA status", 24),
    ("KRA source", 30),
]
SURVEY_COLS = [
    ("Survey questions answered · of 20", 12),
    ("Role mandate & core purpose · as shared", 55),
    ("Deliverables, responsibilities & accountability · as shared", 60),
    ("Capabilities, tools, workload & way of working · as shared", 55),
    ("Development priorities & learning needs · as shared", 50),
    ("Challenges & support required · as shared", 45),
    ("Survey & KRA evidence used for each rating", 60),
    ("Survey themes for this role · role-holders", 50),
    ("Learning topics named in the survey", 35),
    ("Survey status", 34),
]


def _distinct(vals):
    out, seen = [], set()
    for v in vals:
        a = alnum(v)
        if a and a not in seen:
            seen.add(a)
            out.append(v)
    return out


def all_src(p, kind, header):
    """Every distinct non-empty value for a header across the person's records of one kind."""
    hk = hkey(header)
    vals = []
    for s, rec in p["src"].get(kind, []):
        for h, v in rec["row"].items():
            if hkey(h) == hk and has(v):
                vals.append((clean(v), s["label"]))
    d = _distinct([v for v, _ in vals])
    labels = sorted({l for v, l in vals})
    return "\n\n".join(d), labels


def raw(M, p, header):
    j = M["raw_hdr"].index(header)
    return p["raw"][j]


TMP = PREFIX["TM"]
MMP = PREFIX["MM"]
CRP = PREFIX["CRAW"]


def kras(M, p):
    srcs = set()
    def pick(kind, src_header, master_header):
        v, labels = all_src(p, kind, src_header)
        if v:
            srcs.update(labels)
            return v
        m = raw(M, p, master_header)
        if has(m):
            srcs.add("HR Super Master v1.12.2 sheet 3")
            return clean(m)
        return ""
    mgr = pick("TM", "Annual KRAs (objective Measurable) Minimum 5 - set by Manager",
               TMP + "Annual KRAs · objective Measurable Minimum 5 - set by Manager")
    carbon = pick("MM", "KRAs Submitted", MMP + "KRAs Submitted")
    if carbon:
        mgr = (mgr + "\n\n" if mgr else "") + ("CARBON Mapping Master · KRAs submitted:\n" if mgr else "") + carbon
    qtr = pick("TM", "Quarterly KRAs", TMP + "Quarterly KRAs")
    slf = pick("TM", "Annual KRAs (objective Measurable) Minimum 5 - Self",
               TMP + "Annual KRAs · objective Measurable Minimum 5 - Self")
    q1 = pick("TM", "Tentative Rating for Q1 (Scale 1-5)", TMP + "Tentative Rating for Q1 · Scale 1-5")
    jd = pick("CRAW", "Actual JD", CRP + "Actual JD")
    got = [n for n, v in (("manager annual", mgr), ("quarterly", qtr), ("self", slf)) if v]
    status = ("Received · " + ", ".join(got)) if got else MISSING_KRA
    return [mgr or MISSING_KRA, qtr, slf, q1, jd, status, " · ".join(sorted(srcs))]


def _lines(s, pred):
    return [l.strip() for l in text(s).split("\n") if l.strip() and pred(l.strip())]


FACTORS = ["Complexity", "Autonomy", "Accountability", "Leadership", "Experience", "Business Impact"]


def survey(M, p, s8):
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    tm = p["tm"]
    n = text(raw(M, p, "Survey questions answered · of 20"))
    mand, _ = all_src(p, "MM", "Role Mandate")
    deliv, _ = all_src(p, "MM", "Reported Deliverables & Responsibilities")
    capab, _ = all_src(p, "MM", "Current Capabilities")
    devel, _ = all_src(p, "MM", "Development Priorities & Learning Needs")
    chall, _ = all_src(p, "MM", "Challenges & Support Required")
    # The rater's per-factor reading of this person's survey and KRA evidence (sheet 2, columns AX–BC).
    ev = []
    for f, col in zip(FACTORS, range(H["Complexity & Expertise — justification"], H["Business Impact — justification"] + 1)):
        for l in _lines(tm[col], lambda l: l.startswith("Evidence reads like") or "survey" in l.lower()):
            ev.append(f"{f}: {l}")
    themes = []
    for l in _lines(tm[H["Role profile — what they do, responsibilities & competencies required"]],
                    lambda l: l.lower().startswith("common themes")):
        themes.append(l)
    for l in _lines(tm[H["Business impact — from survey and data"]], lambda l: "themes" in l.lower()):
        themes.append(l)
    topics = []
    if s8 is not None:
        for lab, h in (("Self-identified", "Self-identified — topics · survey; not quoted"),
                       ("For the next role", "To move to the next role — topics · survey; not quoted")):
            v = text(s8[M["s8_hdr"].index(h)])
            if has(v):
                topics.append(f"{lab}: {v}")
    try:
        answered = int(float(n or 0))
    except ValueError:
        answered = 0
    if mand or deliv:
        status = "Per-person summary · CARBON Mapping Master"
    elif answered:
        status = MISSING_SURVEY
    else:
        status = "No survey response on file"
    first = mand or (MISSING_SURVEY if answered else "No survey response on file")
    rest = lambda v: v or "—"
    return [answered, first, rest(deliv), rest(capab), rest(devel), rest(chall),
            "\n".join(_distinct(ev)), "\n".join(_distinct(themes)), "\n".join(topics), status]


# ---------------------------------------------------------------- folds from other sheets
def index_by(rows, idcol, namecol):
    d = collections.defaultdict(list)
    for r in rows:
        if r[idcol]:
            d[("id", key(r[idcol]))].append(r)
        if r[namecol]:
            d[("nm", nname(r[namecol]))].append(r)
    return d


def look(idx, p):
    return idx.get(("id", key(p["emp"]))) or idx.get(("nm", nname(p["name"]))) or []


def block_rows(rows, header_first):
    """Rows of the block whose header row starts with header_first (12, 13, 14 have several blocks)."""
    out, on = [], False
    for r in rows:
        f = text(r[0]) if r else ""
        if f == header_first:
            on = True
            continue
        if on:
            if not any(v is not None for v in r):
                break
            out.append(r)
    return out


class Folds:
    def __init__(self, M, people):
        self.M = M
        self.s8 = index_by(M["s8"], 1, 2)
        self.pl = index_by(M["pl"], 0, 1)
        self.t2 = index_by(M["t2"], 3, 4)
        self.t3 = index_by(M["t3"], 3, 4)
        self.s17 = index_by(M["s17"], 0, 1)
        self.s15a = index_by(M["s15a"], 1, 0)
        self.s15 = {nname(r[0]): r for r in M["s15"] if r[0]}
        self.s12 = {nname(r[0]): r for r in block_rows(M["s12"], "Manager") if r[0]}
        self.s13 = {nname(r[0]): r for r in block_rows(M["s13"], "Person") if r[0]}
        self.s14 = {nname(r[0]): r for r in block_rows(M["s14"], "Person") if r[0]}
        # Catalogue ID per person: family × sub-family × level × track (Manager for people managers).
        cidx = collections.defaultdict(list)
        for c in M["cat"]:
            cidx[(text(c[1]), text(c[2]), text(c[3]), text(c[5]))].append(c)
        H = {h: i for i, h in enumerate(M["tm_hdr"])}
        self.H = H
        self.cat_of = {}
        for p in people:
            r = p["tm"]
            fam, sub, lv, tr = (text(r[H["Job family"]]), text(r[H["Sub-family · what they do"]]),
                                text(r[H["FINAL LEVEL"]]), text(r[H["Track"]]))
            t = "Manager" if tr == "People Manager" else "IC"
            hit = cidx.get((fam, sub, lv, t)) or cidx.get((fam, sub, lv, "Manager" if t == "IC" else "IC")) or []
            self.cat_of[p["i"]] = hit[0] if hit else None
        # Reporting manager lookup by name.
        self.by_name = {}
        for p in people:
            self.by_name.setdefault(nname(p["name"]), p)
        self.reports = collections.Counter(nname(p["tm"][H["Reporting manager"]]) for p in people)

    def s8row(self, p):
        r = look(self.s8, p)
        return r[0] if r else None

    def manager(self, p):
        H = self.H
        mn = nname(p["tm"][H["Reporting manager"]])
        m = self.by_name.get(mn)
        if m is None and mn:
            t = mn.split()
            c = [q for n, q in self.by_name.items() if n.split()[:1] == t[:1] and n.split()[-1:] == t[-1:]]
            m = c[0] if len(c) == 1 else None
        if m is None:
            return ["Not in this file" if mn else "—", "—"]
        return [m["emp"], text(m["tm"][H["FINAL LEVEL"]])]

    def placement(self, p):
        r = look(self.pl, p)
        if not r:
            return ["", "", ""]
        h = self.M["pl_hdr"]
        return [text(r[0][h.index("What set the level")]), text(r[0][h.index("Flags")]), text(r[0][h.index("Confidence")])]

    def sensitivity(self, p):
        r = self.s13.get(nname(p["name"]))
        if r:
            return f"Moves in {text(r[5])} of 30 weight tests · calibrate first"
        return "Stable · moves in fewer than 5 of 30 weight tests"

    def org(self, p):
        r = self.s12.get(nname(p["name"]))
        a3 = self.s14.get(nname(p["name"]))
        out = []
        if r:
            out.append(f"Span: {text(r[4])} direct reports · layer {text(r[5])} · {text(r[6])}")
        if a3:
            out.append(f"A3 deep-dive: score {text(a3[5])} · gap to B1 {text(a3[6])} · {text(a3[7])}")
        return "\n".join(out)

    def succession(self, p):
        r = self.s15.get(nname(p["name"]))
        out = []
        if r:
            out.append(f"Succession: {text(r[4])} · ready now {text(r[5])} · in 6–12 months {text(r[6])} · high potential {text(r[7])}")
        q = look(self.s15a, p)
        risk, cover, action = "", "", ""
        if q:
            q = q[0]
            risk = f"{text(q[9])} · critical role {text(q[6])} · only this person {text(q[7])} · successor named {text(q[8])}"
            cover, action = text(q[10]), text(q[11])
        return ["\n".join(out), risk, cover, action]

    def decisions(self, p):
        rs = look(self.t2, p)
        return "\n".join(f"{text(r[0])} · {text(r[7])} · {text(r[8])}: {text(r[9])} · owner {text(r[10])} · {text(r[11])}" for r in rs)

    def requests(self, p):
        rs = look(self.t3, p)
        return "\n".join(f"{text(r[0])} · {text(r[6])} · {text(r[7])} · {text(r[9])}" for r in rs)

    def gaps(self, p):
        rs = look(self.s17, p)
        return "\n".join(f"{text(r[5])} · {text(r[6])}" for r in rs)

    def quality(self, p, kra_status, survey_status):
        H = self.H
        out = []
        out.append("Survey: " + ("summary on file" if survey_status.startswith("Per-person") else
                                  "answered · full text needed" if survey_status == MISSING_SURVEY else "no response"))
        out.append("KRAs: " + ("received" if kra_status.startswith("Received") else "not received"))
        tot = text(p["tm"][H["Total exp. · yrs"]])
        out.append("Experience: " + ("missing" if not has(tot) or "input" in tot.lower() else "on file"))
        if "R2" in text(p["tm"][H["Gates, limits & Player-Coach transition steps applied — with the numbers"]]) or \
                "provisional" in text(p["tm"][H["Points to confirm"]]).lower():
            out.append("Placement provisional · R2")
        if p["src_ids"]:
            out.append("Emp ID differs in source files: " + ", ".join(sorted(p["src_ids"])))
        return " · ".join(out)


# ---------------------------------------------------------------- raw data: restore and extend
import difflib as _dl

EXTRA_RAW = [
    ("MM", "Comments", "[CARBON Mapping Master] Comments"),
    ("MM", "Role Mandate", "[CARBON Mapping Master] Role Mandate"),
    ("MM", "Reported Deliverables & Responsibilities", "[CARBON Mapping Master] Reported Deliverables & Responsibilities"),
    ("MM", "Current Capabilities", "[CARBON Mapping Master] Current Capabilities"),
    ("MM", "Development Priorities & Learning Needs", "[CARBON Mapping Master] Development Priorities & Learning Needs"),
    ("MM", "Challenges & Support Required", "[CARBON Mapping Master] Challenges & Support Required"),
]
RAW_LOG = "Values restored from the source files · field names"
RAW_IDS = "Emp ID in the source files · where different"


def _lost_content(master, src):
    """True when the source holds the master's text plus content the master lost (bracket clean-up, blanks)."""
    a, b = alnum(master), alnum(src)
    if not a:
        return True
    if len(b) < len(a) + 4:
        return False
    m = sum(x.size for x in _dl.SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks())
    return m >= 0.97 * len(a)


def restore_raw(M, p):
    """The person's sheet-3 row with values restored from the source files, plus the extra columns."""
    row = [clean(v) if isinstance(v, str) else v for v in p["raw"]]
    restored = []
    kinds = {k: v for k, v in PREFIX.items()}
    for j, h in enumerate(M["raw_hdr"]):
        kind = next((k for k, pre in kinds.items() if h.startswith(pre)), None)
        if not kind:
            continue
        sv, _ = all_src(p, kind, h[len(kinds[kind]):])
        if not sv:
            continue
        mv = p["raw"][j]
        if not has(mv):
            row[j] = sv
            restored.append(h[len(kinds[kind]):])
        elif isinstance(mv, str) and _lost_content(mv, sv):
            row[j] = sv
            restored.append(h[len(kinds[kind]):])
    extra = []
    for kind, src_h, _ in EXTRA_RAW:
        v, labels = all_src(p, kind, src_h)
        extra.append(v)
    extra.append("; ".join(restored))
    extra.append(", ".join(sorted(p["src_ids"])))
    return row, extra, restored
