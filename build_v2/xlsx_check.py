"""Excel-validity scan of a written .xlsx at the XML level: the things that make Excel say "We found a problem with some
content" even when openpyxl and LibreOffice open the file happily."""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{%s}id" % NS["r"]

# Functions newer than Excel 2007 must be written with the _xlfn. prefix, or Excel shows #NAME? / repairs.
NEW_FUNCS = ("XLOOKUP", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "TEXTJOIN", "CONCAT", "IFS", "SWITCH",
             "MAXIFS", "MINIFS", "IFNA", "LET", "LAMBDA", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT", "VSTACK", "HSTACK",
             "STDEV.S", "STDEV.P", "RANK.EQ", "PERCENTILE.INC", "FLOOR.MATH", "CEILING.MATH", "AGGREGATE", "DAYS", "ISFORMULA")


def col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def name_problem(nm):
    """Why Excel would reject a defined name, or ''."""
    if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", nm):
        return "bad characters"
    if len(nm) > 255:
        return "too long"
    m = re.fullmatch(r"([A-Za-z]{1,3})(\d+)", nm)
    if m and col_num(m.group(1).upper()) <= 16384 and 1 <= int(m.group(2)) <= 1048576:
        return "reads as an A1 cell reference"
    m = re.match(r"([A-Za-z]{1,3})(\d+)[_.\\]", nm)
    if m and col_num(m.group(1).upper()) <= 16384 and 1 <= int(m.group(2)) <= 1048576:
        return "starts with a cell reference"
    if re.match(r"[Rr]\d*[Cc]?\d*[_.\\]", nm) and re.match(r"([Rr]\d+|[Cc]\d+)", nm):
        return "starts with an R1C1 reference"
    if re.fullmatch(r"[RrCc]", nm) or re.fullmatch(r"([Rr]\d*)?([Cc]\d*)?", nm):
        return "reads as an R1C1 reference"
    if re.fullmatch(r"[Rr]\d+.*|[Cc]\d+.*", nm) and re.match(r"([Rr]\d+([Cc]\d*)?|[Cc]\d+)$", nm):
        return "reads as an R1C1 reference"
    if nm.upper() in ("TRUE", "FALSE"):
        return "reserved word"
    return ""


def ranges_overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def rng(ref):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref)
    if not m:
        return None
    c1, r1 = col_num(m.group(1)), int(m.group(2))
    c2, r2 = (col_num(m.group(3)), int(m.group(4))) if m.group(3) else (c1, r1)
    return (c1, r1, c2, r2)


def scan(path):
    out = []
    z = zipfile.ZipFile(path)
    parts = set(z.namelist())
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    sheets = []
    seen_sheet = set()
    for s in wbx.find("m:sheets", NS):
        nm = s.get("name")
        if len(nm) > 31 or re.search(r"[\[\]:*?/\\]", nm) or nm.lower() in seen_sheet:
            out.append(f"sheet name invalid: {nm!r}")
        seen_sheet.add(nm.lower())
        t = rels.get(s.get(RID), "")
        t = t.lstrip("/") if t.startswith("/") else "xl/" + t
        if t not in parts:
            out.append(f"sheet part missing: {nm} → {t}")
        sheets.append((nm, t))
    seen = set()
    dns = wbx.find("m:definedNames", NS)
    for d in (dns if dns is not None else []):
        nm, scope = d.get("name"), d.get("localSheetId")
        why = "" if nm.startswith("_xlnm.") else name_problem(nm)
        if why:
            out.append(f"defined name {nm}: {why}")
        key = (nm.lower(), scope)
        if key in seen:
            out.append(f"defined name {nm}: duplicate")
        seen.add(key)
        if "#REF!" in (d.text or "") or not (d.text or "").strip():
            out.append(f"defined name {nm}: broken reference {d.text!r}")
    styles = ET.fromstring(z.read("xl/styles.xml"))
    xfs = styles.find("m:cellXfs", NS)
    if xfs is not None and len(xfs) > 64000:
        out.append(f"too many cell styles: {len(xfs)}")
    n_xf = len(xfs) if xfs is not None else 1
    dx = styles.find("m:dxfs", NS)
    n_dxf = len(dx) if dx is not None else 0
    if "xl/sharedStrings.xml" in parts:
        for i, si in enumerate(ET.fromstring(z.read("xl/sharedStrings.xml"))):
            if len("".join(si.itertext())) > 32767:
                out.append(f"shared string {i} longer than 32,767 characters")
    for nm, part in sheets:
        if part not in parts:
            continue
        x = ET.fromstring(z.read(part))
        srels = {}
        rp = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rp in parts:
            srels = {r.get("Id"): (r.get("Target"), r.get("TargetMode")) for r in ET.fromstring(z.read(rp))}
        prev_row = 0
        seen_refs = set()
        for row in x.iter("{%s}row" % NS["m"]):
            rn = int(row.get("r", prev_row + 1))
            if rn <= prev_row:
                out.append(f"{nm}: rows out of order at {rn}")
            prev_row = rn
            ht = row.get("ht")
            if ht and float(ht) > 409:
                out.append(f"{nm}: row {rn} taller than 409")
            prev_c = 0
            for c in row.findall("m:c", NS):
                ref = c.get("r")
                cc = rng(ref) if ref else None
                if cc:
                    if cc[0] <= prev_c:
                        out.append(f"{nm}: cells out of order at {ref}")
                    prev_c = cc[0]
                f = c.find("m:f", NS)
                if f is not None and f.text:
                    if len(f.text) > 8192:
                        out.append(f"{nm}!{ref}: formula longer than 8,192")
                    up = f.text.upper()
                    for fn in NEW_FUNCS:
                        if re.search(r"(?<![\w.])" + re.escape(fn) + r"\(", up):
                            out.append(f"{nm}!{ref}: {fn} needs the _xlfn. prefix")
                            break
                t = c.get("t")
                txt = "".join((c.find("m:is", NS) if c.find("m:is", NS) is not None else ET.Element("x")).itertext()) if t == "inlineStr" else \
                      (c.findtext("m:v", "", NS) if t in ("str", "s") else "")
                if t != "s" and len(txt) > 32767:
                    out.append(f"{nm}!{ref}: text longer than 32,767 characters")
                v = c.findtext("m:v", None, NS)
                if t in (None, "n") and v not in (None, ""):
                    try:
                        fv = float(v)
                        if fv != fv or fv in (float("inf"), float("-inf")):
                            out.append(f"{nm}!{ref}: number is not finite ({v})")
                    except ValueError:
                        out.append(f"{nm}!{ref}: number cell holds text {v[:30]!r}")
                if t == "b" and v not in ("0", "1"):
                    out.append(f"{nm}!{ref}: bad boolean {v!r}")
                si = int(c.get("s", 0))
                if si >= n_xf:
                    out.append(f"{nm}!{ref}: style {si} beyond {n_xf} styles")
                if ref in seen_refs:
                    out.append(f"{nm}!{ref}: duplicate cell")
                seen_refs.add(ref)
                if t == "e" and c.findtext("m:v", "", NS) not in ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"):
                    out.append(f"{nm}!{ref}: unknown error value")
        cols = x.find("m:cols", NS)
        last_max = 0
        for col in (cols if cols is not None else []):
            if int(col.get("min")) <= last_max or int(col.get("max")) < int(col.get("min")):
                out.append(f"{nm}: column ranges overlap or out of order at {col.get('min')}")
            last_max = int(col.get("max"))
            if col.get("width") and float(col.get("width")) > 255:
                out.append(f"{nm}: column width over 255")
        mc = x.find("m:mergeCells", NS)
        merges = [rng(m.get("ref")) for m in (mc if mc is not None else [])]
        if None in merges:
            out.append(f"{nm}: invalid merge reference")
        merges = sorted(m for m in merges if m)
        for i in range(len(merges)):
            for j in range(i + 1, len(merges)):
                if merges[j][0] > merges[i][2] and False:
                    break
                if ranges_overlap(merges[i], merges[j]):
                    out.append(f"{nm}: merged ranges overlap {merges[i]} {merges[j]}")
                    break
        hl = x.find("m:hyperlinks", NS)
        for h in (hl if hl is not None else []):
            loc = h.get("location") or ""
            tgt = srels.get(h.get(RID), ("", None))[0] if h.get(RID) else ""
            if h.get(RID) and h.get(RID) not in srels:
                out.append(f"{nm}!{h.get('ref')}: hyperlink relationship missing")
            if len(loc) > 255 or len(tgt) > 2079:
                out.append(f"{nm}!{h.get('ref')}: hyperlink too long")
            if tgt and (tgt.startswith("#") or re.search(r"[\s<>\"{}|^`]", tgt)):
                out.append(f"{nm}!{h.get('ref')}: hyperlink target is not a valid external URL: {tgt[:60]}")
        for cf in x.findall("m:conditionalFormatting", NS):
            for rule in cf.findall("m:cfRule", NS):
                if rule.get("dxfId") is not None and int(rule.get("dxfId")) >= n_dxf:
                    out.append(f"{nm}: conditional format points to a missing style")
                for f in rule.findall("m:formula", NS):
                    if len(f.text or "") > 8192:
                        out.append(f"{nm}: conditional format formula too long")
        dv = x.find("m:dataValidations", NS)
        for d in (dv if dv is not None else []):
            f1 = d.findtext("m:formula1", "", NS)
            if len(f1) > 255:
                out.append(f"{nm}: validation formula longer than 255")
            if len(d.get("error") or "") > 225 or len(d.get("prompt") or "") > 255 or len(d.get("errorTitle") or "") > 32 \
                    or len(d.get("promptTitle") or "") > 32:
                out.append(f"{nm}: validation message longer than Excel allows")
            if not d.get("sqref"):
                out.append(f"{nm}: validation with no range")
            if d.get("type") == "list" and f1.startswith('"') and len(f1) > 257:
                out.append(f"{nm}: list validation longer than 255")
        if dv is not None and int(dv.get("count", len(dv))) != len(dv):
            out.append(f"{nm}: validation count mismatch")
    return out


if __name__ == "__main__":
    bad = 0
    for p in sys.argv[1:]:
        res = scan(p)
        print(p, "→", "clean" if not res else f"{len(res)} problems")
        for r in res[:40]:
            print("  ", r)
        bad += len(res)
    sys.exit(1 if bad else 0)
