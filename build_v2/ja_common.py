"""Shared loader for the v1.13 build: reads the v1.12.2 HR Super Master and the
manager talent-mapping source files, and joins them into one record per person.

Nothing here writes a workbook. build_super.py and build_simplified.py use the
records; verify.py checks the result.

Ground rules kept (Build Guide section 3): personal fields (gender, DoB, age,
e-mail, mobile, resignation / relieving dates) are never carried; pay stays blank;
no round brackets in text (converted losslessly to ' · ').
"""
import collections
import difflib
import datetime as dt
import os
import re

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORE = os.path.join(ROOT, "IRIS_JA_Core_Files_v1.12.2")
MASTER = os.path.join(CORE, "IRIS_JA_2_HR_Super_Master.xlsx")
OUT = os.path.join(ROOT, "IRIS_JA_Final_v1.15")
VERSION = "Final v1.15"
VER = "v1.15"
STAMP = "IRIS · Job Architecture & Job Design · Final v1.15 · October 2026 · Internal — Confidential · HR only"

PERSONAL = {"gender", "dob", "age", "emailid", "today", "dateofresignation", "dateofrelieving", "mobile"}

# The five manager talent-mapping files and how to read them.
SOURCES = [
    # key in file name, sheet, emp-id col, name col, kind, label
    ("CARBON_Talent_Mapping", "Mapping Master", 0, 1, "MM", "CARBON Mapping Master"),
    ("CARBON_Talent_Mapping", "Raw", 1, 4, "CRAW", "CARBON Talent Mapping · Raw"),
    ("iFile_Collated", "Collated", 1, 4, "TM", "Talent Mapping · iFile collated"),
    ("IDeal_Talent_Mapping", "Talent Mapping", 0, 3, "TM", "Talent Mapping · iDEAL"),
    ("Database_content", "Talent Mapping", 0, 3, "TM", "Talent Mapping · Database Content"),
    ("Dhobale_HR", "Definition", 0, 3, "TM", "Talent Mapping · Datatech"),
]
# Master sheet-3 column prefix for each kind of source.
PREFIX = {
    "MM": "[CARBON Mapping Master] ",
    "CRAW": "[CARBON Talent Mapping · Raw] ",
    "TM": "[Talent Mapping · iFile / iDEAL / DataTech / Database Content] ",
}
# Source headers that the master spells differently.
ALIAS = {
    "correctdesignationinlinewithcurrentrole": "correctdesignation",
    "nameofsuccessor": "nameofsuccessoridentified",
    "quarterlyperformancerating15": "quarterlyperformancerating",
}

ZW = re.compile(r"[​‌‍﻿]")


def text(v):
    """Display text of a cell value: zero-width characters removed, dates as dd Mon yyyy."""
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%d %b %Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = ZW.sub("", str(v)).replace("\r\n", "\n").replace("\r", "\n")
    return s.strip()


def debracket(s):
    """Round brackets → ' · ' without losing what was inside (house style, Build Guide §3)."""
    if not isinstance(s, str) or "(" not in s and ")" not in s:
        return s
    # list numbering such as '1)' or 'a)' keeps its meaning as '1.' / 'a.'
    s = re.sub(r"(^|\s)(\d{1,2}|[a-zA-Z])\)(?=[A-Za-z])", lambda m: f"{m.group(1)}{m.group(2)}. ", s)
    s = re.sub(r"(^|\s)(\d{1,2}|[a-zA-Z])\)(?=\s)", lambda m: f"{m.group(1)}{m.group(2)}.", s)
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\s*\(([^()]*)\)", lambda m: (" · " + m.group(1).strip()) if m.group(1).strip() else "", s)
    s = s.replace("(", " · ").replace(")", "")
    s = re.sub(r"[ \t]+·", " ·", s)
    s = re.sub(r"·\s*·", "·", s)
    s = re.sub(r"^\s*·\s*", "", s, flags=re.M)
    return s.strip()


def clean(v):
    return debracket(text(v))


def has(v):
    s = text(v)
    return s not in ("", "-", "—", "NA", "N/A", "na", "n/a", "None", "#REF!", "#N/A", "#VALUE!")


def key(s):
    return re.sub(r"[^a-z0-9]", "", text(s).lower())


def hkey(h):
    """Header key: drop '[source] ' prefix and punctuation; apply aliases."""
    h = re.sub(r"^\[[^\]]*\]\s*", "", text(h))
    k = key(h)
    return ALIAS.get(k, k)


def nname(s):
    return re.sub(r"\s+", " ", text(s)).strip().upper()


def alnum(s):
    return re.sub(r"\W", "", text(s).lower())


# ---------------------------------------------------------------- master
def read_block(ws, header_row, first_data, idcol=1, namecol=2):
    rows = list(ws.iter_rows(min_row=header_row, values_only=True))
    hdr = [text(h) for h in rows[0]]
    skip = first_data - header_row
    data = [list(r) for r in rows[skip:] if r and (r[idcol] or r[namecol])]
    return hdr, data


def load_master(path=MASTER):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    M = {"wb_sheetnames": wb.sheetnames}
    M["tm_hdr"], M["tm"] = read_block(wb["2. Talent Mapping"], 6, 7)
    M["tm_grp"] = [text(v) for v in next(wb["2. Talent Mapping"].iter_rows(min_row=5, max_row=5, values_only=True))]
    M["raw_hdr"], M["raw"] = read_block(wb["3. Raw Data & Market Standards"], 6, 7)
    M["raw_grp"] = [text(v) for v in next(wb["3. Raw Data & Market Standards"].iter_rows(min_row=5, max_row=5, values_only=True))]
    M["s8_hdr"], M["s8"] = read_block(wb["8. Talent, Learning & Mobility"], 6, 7)
    M["pl_hdr"], M["pl"] = read_block(wb["1a. Placements"], 6, 7, 0, 1)
    M["cat_hdr"], cat = read_block(wb["6. Role Catalogue"], 6, 7, 0, 1)
    M["cat"] = [r for r in cat if r[0] and r[1]]
    M["req_hdr"], M["req"] = read_block(wb["7. Role Requirements"], 6, 7, 0, 4)
    M["comp_hdr"], M["comp"] = read_block(wb["4. Role Competency Framework"], 6, 7, 0, 3)
    M["t2_hdr"], M["t2"] = read_block(wb["T2. Decisions"], 6, 7, 0, 4)
    M["t3_hdr"], M["t3"] = read_block(wb["T3. Data Requests"], 6, 7, 0, 4)
    M["s17_hdr"], M["s17"] = read_block(wb["17. Data Gaps"], 5, 6, 0, 1)
    M["s15a_hdr"], M["s15a"] = read_block(wb["15a. Critical Roles at Risk"], 5, 6, 0, 1)
    M["s15_hdr"], M["s15"] = read_block(wb["15. Succession"], 5, 6, 0, 1)
    # 12, 13, 14 hold more than one block; read them whole.
    M["s12"] = [list(r) for r in wb["12. Org Health"].iter_rows(values_only=True)]
    M["s13"] = [list(r) for r in wb["13. Weight Sensitivity"].iter_rows(values_only=True)]
    M["s14"] = [list(r) for r in wb["14. A3 Deep-Dive"].iter_rows(values_only=True)]
    wb.close()
    return M


# ---------------------------------------------------------------- sources
def find_file(srcdir, key_):
    for f in sorted(os.listdir(srcdir)):
        if key_.lower() in f.lower() and f.lower().endswith(".xlsx"):
            return os.path.join(srcdir, f)
    raise FileNotFoundError(f"No source file containing '{key_}' in {srcdir}")


def load_sources(srcdir):
    out = []
    for fkey, sheet, ic, nc, kind, label in SOURCES:
        path = find_file(srcdir, fkey)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = list(wb[sheet].iter_rows(values_only=True))
        wb.close()
        hdr = [text(h) for h in rows[0]]
        recs = []
        for r in rows[1:]:
            r = list(r) + [None] * (len(hdr) - len(r))
            if not (has(r[ic]) or has(r[nc])) or key(r[ic]) == "empid":
                continue
            d = {h: r[j] for j, h in enumerate(hdr) if h and hkey(h) not in PERSONAL}
            recs.append({"id": text(r[ic]), "name": text(r[nc]), "row": d})
        fname = re.sub(r"^[0-9a-f]{8}-", "", os.path.basename(path))
        out.append({"kind": kind, "label": label, "file": fname, "sheet": sheet, "hdr": hdr, "recs": recs})
    return out


# ---------------------------------------------------------------- join
def _name_match(nm, people):
    """Exact name, else same first name and close surname (e.g. 'Utkarsh Bafna' ~ 'Utkarsh Sagar Bafana')."""
    n = nname(nm)
    exact = [p for p in people if nname(p["name"]) == n]
    if len(exact) == 1:
        return exact[0]
    t = n.split()
    if len(t) < 2:
        return None
    cands = []
    for p in people:
        u = nname(p["name"]).split()
        if len(u) >= 2 and u[0] == t[0] and difflib.SequenceMatcher(None, u[-1], t[-1]).ratio() >= 0.8:
            cands.append(p)
    return cands[0] if len(cands) == 1 else None


def build(srcdir):
    M = load_master()
    S = load_sources(srcdir)
    H = {h: i for i, h in enumerate(M["tm_hdr"])}
    people = []
    for i, r in enumerate(M["tm"]):
        people.append({"i": i, "emp": text(r[1]), "name": text(r[2]), "biz": text(r[0]), "tm": r,
                       "raw": M["raw"][i], "src": collections.defaultdict(list), "src_ids": set()})
    assert len(people) == len(M["raw"]) == 586
    for p, raw in zip(people, M["raw"]):
        assert text(raw[1]) == p["emp"] and text(raw[2]) == p["name"], "sheet 2 and 3 out of step"
    by_id = {key(p["emp"]): p for p in people if p["emp"].startswith(("IRIS", "INTN"))}

    unmapped = []
    for s in S:
        for rec in s["recs"]:
            p = by_id.get(key(rec["id"])) if rec["id"] else None
            how = "Emp ID"
            if p is None:
                p = _name_match(rec["name"], people)
                how = "name"
            if p is None:
                unmapped.append({"source": s, "rec": rec})
                continue
            p["src"][s["kind"]].append((s, rec))
            if rec["id"] and key(rec["id"]) != key(p["emp"]):
                p["src_ids"].add(rec["id"])
    return M, S, people, unmapped


# ---------------------------------------------------------------- per-person derived fields
def src_val(p, kind, *names):
    """First non-empty value of the given header names in the person's records of that kind."""
    keys = [hkey(n) for n in names]
    for s, rec in p["src"].get(kind, []):
        for h, v in rec["row"].items():
            if hkey(h) in keys and has(v):
                return v, s
    return None, None


def raw_val(M, p, header):
    j = M["raw_hdr"].index(header)
    return p["raw"][j]
