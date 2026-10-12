"""Shared look for the v1.13 workbooks — same palette and conventions as v1.12.2
(Build Guide Appendix E.6): banner rows 1–3, colour key row 4, group headers row 5,
column headers row 6, data from row 7."""
import copy
import math

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

NAVY = "1F3864"
THIN = Side(style="thin", color="BFBFBF")
MED = Side(style="medium", color="1F3864")
F_RED = PatternFill("solid", fgColor="FFC7CE")
T_RED = Font(name="Calibri", size=9, color="9C0006")
F_YEL = PatternFill("solid", fgColor="FFF2CC")
F_GREY = PatternFill("solid", fgColor="EDEDED")
F_ORANGE = PatternFill("solid", fgColor="F8CBAD")
T_ORANGE = Font(name="Calibri", size=9, color="833C0B", bold=True)
BAND = {"A": "DDEBF7", "B": "E2EFDA", "B+": "C6E0B4", "C": "D0E8E5", "C+": "A9D08E", "D": "E4DFEC", "E": "EDE7DD", "X": "EDEDED"}
CHANGE = {"▲": ("C6EFCE", "006100"), "▼": ("FFEB9C", "9C5700"), "●": ("DDEBF7", "1F4E79"), "•": ("F2F2F2", "404040")}
KEY = [("Section colours: each column group has its own header colour and a matching light tint", "D9E1F2", 4),
       ("Changed vs current", "F8CBAD", 2), ("Level up", "C6EFCE", 1), ("Level down · pay & title protected", "FFEB9C", 2),
       ("Newly mapped", "DDEBF7", 1), ("Player-Coach · B+ · C+", "C6E0B4", 2),
       ("RED = critical input needed · says what and who", "FFC7CE", 3), ("YELLOW = to complete or edit", "FFF2CC", 2),
       ("GREY = pay · set later", "EDEDED", 2)]


def font(size=9, bold=False, color="000000", italic=False, underline=None):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic, underline=underline)


def header(c, fill, size=9):
    c.font = font(size, True, "FFFFFF")
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def body(c, tint, first=False, wrap=True):
    c.font = font()
    if tint:
        c.fill = PatternFill("solid", fgColor=tint)
    c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal="left")
    c.border = Border(left=MED if first else THIN, right=THIN, top=THIN, bottom=THIN)


def banner(ws, title, purpose, stamp, back=True, width_cols=12):
    ws["A1"] = title
    ws["A1"].font = font(16, True, NAVY)
    ws["A2"] = purpose
    ws["A2"].font = font(10, False, "404040")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width_cols)
    ws.row_dimensions[2].height = max(30, 15 * math.ceil(len(purpose) / (width_cols * 14)))
    ws["A3"] = stamp
    ws["A3"].font = font(8, False, "808080", italic=True)
    ws.row_dimensions[1].height = 24
    if back:
        c = ws.cell(1, width_cols + 1, "◂ Start")
        c.hyperlink = "#'1. Framework & Model'!A1"
        c.font = font(9, False, "0563C1", underline="single")


def colour_key(ws, row=4):
    col = 1
    for text_, fill, span in KEY:
        c = ws.cell(row, col, text_)
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = font(8, True, "404040")
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if span > 1:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
        col += span
    ws.row_dimensions[row].height = 30


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format


def band_of(level):
    s = str(level or "")
    if s in ("B+", "C+"):
        return s
    return s[:1] if s[:1] in BAND else None


def est_height(texts_widths, base=12.5, minimum=15, maximum=409):
    """Row height for wrapped text: texts_widths = [(text, column width in characters)]."""
    lines = 1
    for t, w in texts_widths:
        if t is None:
            continue
        t = str(t)
        n = 0
        for part in t.split("\n"):
            n += max(1, math.ceil(len(part) * 1.05 / max(4, w - 1)))
        lines = max(lines, n)
    return max(minimum, min(maximum, lines * base + 3))
