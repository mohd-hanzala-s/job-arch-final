"""The v2.1 measurement scales — five steps for weightage and five steps for ratings.

WEIGHTAGE · how much a parameter matters for a role at a level, as an importance score from 0 to 100 in five 20-point
bands. A cell takes a step code or an exact number:
    VL Very low 0–20 · L Low 20–40 · M Medium 40–60 · H High 60–80 · VH Very high 80–100
A step counts at its band's midpoint — 10, 30, 50, 70, 90 — so the steps stand in the ratio 1 : 3 : 5 : 7 : 9, the odd
steps of Saaty's 1–9 fundamental scale for relative importance. Each parameter's share of the decision is its importance
divided by the sum of the seven importances (direct-rating weights, as in SMART · Edwards). A number typed instead of a
code is used as it is, so a leader can fine-tune within a band.

How the published weights become steps: a method publishes each factor's share of the points. The importance of a factor
is its share as a percentage of the largest factor's share, so the most important factor scores 100 and the rest follow
in proportion; the step is the 20-point band that importance falls in.

RATING · how the person works against the level's definition on that parameter, in half-level steps:
    1 Well below — works at the level below · 2 Below — part of the way there · 3 Meets — as the level is defined ·
    4 Above — part of the way to the level above · 5 Well above — already works at the level above
Step s moves the person (s − 3) ÷ 2 of a level on that parameter: −1, −½, 0, +½, +1.

PLACEMENT · the weighted rating is Σ share × move. IRIS's placement is the anchor — the weighted rating with IRIS's
ratings and IRIS's weightage as built. When any weightage or rating changes, the new weighted rating minus the anchor is
the shift in levels; the level moves by the shift rounded to the nearest whole level (half a level or more moves one),
at most two levels either way, along the person's own ladder.
"""

W5 = [("VL", "Very low", 0, 20, 10), ("L", "Low", 20, 40, 30), ("M", "Medium", 40, 60, 50),
      ("H", "High", 60, 80, 70), ("VH", "Very high", 80, 100, 90)]
W5_MID = {c: m for c, _, _, _, m in W5}
W5_NAME = {c: n for c, n, _, _, _ in W5}
R5 = [(1, "Well below", "works at the level below on this parameter", -1.0),
      (2, "Below", "part of the way to the level's standard", -0.5),
      (3, "Meets", "as the level is defined", 0.0),
      (4, "Above", "part of the way to the level above", 0.5),
      (5, "Well above", "already works at the level above on this parameter", 1.0)]
R5_MOVE = {s: m for s, _, _, m in R5}
MAX_SHIFT = 2


def w5(importance):
    """Importance 0–100 → step code."""
    x = max(0.0, min(100.0, float(importance)))
    for c, _, lo, hi, _ in W5:
        if x < hi:
            return c
    return "VH"


def importance(shares):
    """{param: share %} → {param: importance 0–100}, the largest share = 100."""
    top = max(shares.values()) or 1
    return {k: v * 100.0 / top for k, v in shares.items()}


def steps(shares):
    return {k: w5(v) for k, v in importance(shares).items()}


def rebuilt(codes):
    """Shares the step midpoints give back, to compare with the published shares."""
    tot = sum(W5_MID[c] for c in codes.values()) or 1
    return {k: W5_MID[c] * 100.0 / tot for k, c in codes.items()}


def imp_of(x):
    """A weightage cell → importance: a number is used as typed (0–100), a code at its midpoint."""
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return max(0.0, min(100.0, float(x)))
    return W5_MID.get(str(x).strip().upper())


def rate5(diff):
    """A rating against the level, in levels (+0.7 = 0.7 of a level above), → step 1–5 in half-level steps."""
    h = max(-1.0, min(1.0, round(diff * 2) / 2))
    return int(3 + 2 * h)


def exp5(years, lo, hi):
    if years is None or lo is None:
        return ""
    span = max(1.0, hi - lo) if hi < 90 else 4.0
    if years < lo - span:
        return 1
    if years < lo:
        return 2
    if years < hi:
        return 3
    if years < hi + span:
        return 4
    return 5


def edu5(q, norm):
    if q is None:
        return ""
    d = q - norm
    return 1 if d <= -2 else 2 if d == -1 else 3 if d == 0 else 4 if d == 1 else 5


def net(weights, ratings, keys):
    """Weighted rating in levels: Σ importance × move ÷ Σ importance. Blank rating = meets (no move)."""
    num = den = 0.0
    for k in keys:
        w = imp_of(weights.get(k))
        if w is None:
            continue
        r = ratings.get(k)
        m = R5_MOVE.get(int(r), 0.0) if r not in (None, "") else 0.0
        num += w * m
        den += w
    return num / den if den else 0.0
