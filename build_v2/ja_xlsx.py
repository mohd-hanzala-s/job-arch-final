"""Package-level helpers for workbooks written by openpyxl.

reinject_drawings(): openpyxl drops charts when it re-saves an existing workbook;
    this copies the original drawings and charts back, sheet by sheet (by name).
cache_values(): stores each formula's result next to the formula (computed by
    LibreOffice), so phone / browser previews show values. Formulas stay live.
"""
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
import html

import openpyxl

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_TYPE = NS_R + "/drawing"


def _sheet_paths(z):
    """sheet name → zip path of its XML."""
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = {m.group(2): m.group(1) for m in re.finditer(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)}
    rid.update({m.group(1): m.group(2) for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)})
    out = {}
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        t = rid[m.group(2)].lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[html.unescape(m.group(1))] = t
    return out


def _rels_path(sheet_path):
    d, f = os.path.split(sheet_path)
    return f"{d}/_rels/{f}.rels"


def reinject_drawings(orig, new):
    zo = zipfile.ZipFile(orig)
    zn = zipfile.ZipFile(new)
    po, pn = _sheet_paths(zo), _sheet_paths(zn)
    files = {n: zn.read(n) for n in zn.namelist()}
    names_new = set(files)
    ct = files["[Content_Types].xml"].decode("utf8")
    ct_orig = zo.read("[Content_Types].xml").decode("utf8")
    moved = []
    for sheet, path in po.items():
        rp = _rels_path(path)
        if rp not in zo.namelist():
            continue
        rels = zo.read(rp).decode("utf8")
        for m in re.finditer(r'<Relationship[^>]*/>', rels):
            rel = m.group(0)
            if 'relationships/drawing"' not in rel:
                continue
            target = re.search(r'Target="([^"]+)"', rel).group(1).lstrip("/")
            moved.append((sheet, target))
    used = 0
    for sheet, dpath in moved:
        npath = pn[sheet]
        # copy the drawing, its rels and every chart it points to, under fresh names
        used += 1
        dname = f"xl/drawings/drawingC{used}.xml"
        files[dname] = zo.read(dpath)
        drels_o = _rels_path(dpath)
        drels = zo.read(drels_o).decode("utf8")
        def chart_copy(t):
            t = t[1:] if t.startswith("/") else ("xl/" + t[3:] if t.startswith("../") else t)
            nt = f"xl/charts/c{used}_{os.path.basename(t)}"
            files[nt] = zo.read(t)
            if _rels_path(t) in zo.namelist():
                files[_rels_path(nt)] = zo.read(_rels_path(t))
            return f'Target="/{nt}"'
        drels = re.sub(r'Target="([^"]+)"', lambda m: chart_copy(m.group(1)), drels)
        files[_rels_path(dname)] = drels.encode("utf8")
        # sheet rels
        srp = _rels_path(npath)
        srels = files.get(srp, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode("utf8")
        rid = f"rIdDrw{used}"
        srels = srels.replace("</Relationships>",
                              f'<Relationship Id="{rid}" Type="{DRAWING_TYPE}" Target="/{dname}"/></Relationships>')
        files[srp] = srels.encode("utf8")
        sx = files[npath].decode("utf8")
        if "<drawing " in sx:
            raise RuntimeError(f"{sheet} already has a drawing")
        tag = f'<drawing xmlns:r="{NS_R}" r:id="{rid}"/>'
        for anchor in ("<legacyDrawing ", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"):
            k = sx.find(anchor)
            if k >= 0:
                sx = sx[:k] + tag + sx[k:]
                break
        files[npath] = sx.encode("utf8")
    # content types
    adds = []
    for n in files:
        if n not in names_new and n.endswith(".xml"):
            if n.startswith("xl/charts/"):
                adds.append(f'<Override PartName="/{n}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
            elif n.startswith("xl/drawings/") and "/_rels/" not in n:
                adds.append(f'<Override PartName="/{n}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    ct = ct.replace("</Types>", "".join(adds) + "</Types>")
    files["[Content_Types].xml"] = ct.encode("utf8")
    zn.close()
    zo.close()
    _write(new, files)
    return len(moved)


def _write(path, files):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        order = ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]
        for n in order:
            z.writestr(n, files[n])
    os.replace(tmp, path)


# ---------------------------------------------------------------- cached values
def recalc_copy(path, workdir):
    """LibreOffice recalculates a copy and returns the path of its xlsx output."""
    src = os.path.join(workdir, "in_" + os.path.basename(path))
    shutil.copy(path, src)
    outdir = os.path.join(workdir, "out")
    os.makedirs(outdir, exist_ok=True)
    profile = "file://" + os.path.join(workdir, "lo_profile")
    subprocess.run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--calc",
                    "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, src],
                   check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=3000)
    return os.path.join(outdir, os.path.basename(src))


def _xml_esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(<f[^>]*>.*?</f>|<f[^>]*/>)(<v>.*?</v>|<v/>)?</c>', re.S)


def cache_values(path, workdir):
    calc = recalc_copy(path, workdir)
    vals = openpyxl.load_workbook(calc, read_only=True, data_only=True)
    z = zipfile.ZipFile(path)
    paths = _sheet_paths(z)
    files = {n: z.read(n) for n in z.namelist()}
    z.close()
    errors = []
    n_f = 0
    for sheet, sp in paths.items():
        sx = files[sp].decode("utf8")
        if "<f" not in sx:
            continue
        ws = vals[sheet]
        grid = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and hasattr(c, "coordinate"):
                    grid[c.coordinate] = c.value

        def fix(m):
            nonlocal n_f
            n_f += 1
            ref, attrs, f = m.group(1), m.group(2), m.group(3)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            v = grid.get(ref)
            if v is None:
                return f'<c r="{ref}"{attrs} t="str">{f}<v></v></c>'
            if isinstance(v, bool):
                return f'<c r="{ref}"{attrs} t="b">{f}<v>{int(v)}</v></c>'
            if isinstance(v, (int, float)):
                return f'<c r="{ref}"{attrs}>{f}<v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
            s = str(v)
            if s.startswith("#") and s.rstrip("!?0/").upper() in ("#REF", "#N/A", "#VALUE", "#NAME", "#DIV", "#NUM", "#NULL") or s in ("#DIV/0!", "#NAME?", "#N/A", "#REF!", "#VALUE!", "#NUM!", "#NULL!"):
                errors.append((sheet, ref, s))
                return f'<c r="{ref}"{attrs} t="e">{f}<v>{_xml_esc(s)}</v></c>'
            return f'<c r="{ref}"{attrs} t="str">{f}<v>{_xml_esc(s)}</v></c>'
        sx = CELL.sub(fix, sx)
        files[sp] = sx.encode("utf8")
    vals.close()
    _write(path, files)
    return n_f, errors
