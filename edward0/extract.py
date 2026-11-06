"""
extract.py - workbook -> one entry per value cell, with the labels around it.

Taken from bronze.py (same rules), without SQLite and without labels.toml.
Layout is used here, and only here: to find the tables on a sheet and to collect the
labels that apply to each value cell. The labels are returned as an unordered set; what
they mean (context, dimension, noise) is decided in clean.py from the sets alone.

Per value cell, the labels are:
  * header  - its column-header path and row-header path in its table
  * title   - the table's title rows, and captions just above the table
  * layout  - text to its left in the same row, and header rows above the table
"""
from __future__ import annotations

import csv
import os
import re
from array import array
from decimal import Decimal, InvalidOperation
from bisect import bisect_left, bisect_right
from collections import defaultdict
from datetime import date, datetime, time as dtime

from openpyxl import load_workbook
from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet._reader import WorkSheetParser

# --------------------------------------------------------------------------- #
# Values and labels
# --------------------------------------------------------------------------- #
EXCEL_ERRORS = {"#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!",
                "#SPILL!", "#CALC!", "#GETTING_DATA", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!"}
_NUMERIC_TEXT = re.compile(r"^\s*[-+(]?\s*[\d,]*\.?\d+\s*\)?\s*%?\s*$")
_XL_DATE = [("yyyy", "%Y"), ("mmmm", "%B"), ("mmm", "%b"), ("dddd", "%A"), ("ddd", "%a"),
            ("yy", "%y"), ("mm", "%m"), ("dd", "%d"), ("m", "%-m"), ("d", "%-d")]


def is_empty(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def is_number(v) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    return isinstance(v, str) and bool(_NUMERIC_TEXT.match(v))


def is_label(v) -> bool:
    """Something that can name a row or column: text (not a number written as text) or a date."""
    return not is_empty(v) and not is_number(v)


def iso(v) -> str:
    if isinstance(v, datetime):
        return v.date().isoformat() if v.time() == dtime(0) else v.isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    return ""


def excel_date_text(v, number_format: str | None) -> str:
    """A date the way Excel displays it ('mmm-yy' -> 'Sep-26'); ISO if the format is unknown."""
    fmt = (number_format or "").split(";")[0]
    fmt = re.sub(r"\[[^\]]*\]|\\|\"", "", fmt).lower()
    if not fmt or "general" in fmt or not re.search(r"[dmy]", fmt) or re.search(r"[hs]", fmt):
        return iso(v)
    out, i = "", 0
    while i < len(fmt):
        for code, py in _XL_DATE:
            if fmt.startswith(code, i):
                out += py
                i += len(code)
                break
        else:
            out += fmt[i]
            i += 1
    try:
        return v.strftime(out)
    except (ValueError, TypeError):
        return iso(v)


class Label:
    """A label as displayed, plus its ISO date if it is a date, and the cell it was read from
    (top-left of a merged range; used only to name dimensions, never to decide them)."""
    __slots__ = ("text", "date_value", "origin")

    def __init__(self, text: str, date_value: str = "", origin=None):
        self.text, self.date_value, self.origin = text, date_value, origin

    def key(self):
        return (self.text, self.date_value)


def to_label(v, number_format=None) -> Label | None:
    if is_empty(v):
        return None
    if isinstance(v, (datetime, date)):
        return Label(excel_date_text(v, number_format), iso(v))
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return Label(str(v).strip())


def cell_value(v, has_formula: bool):
    """-> (value_type, value) or None for an empty cell.
    value_type: number | text | date | bool | error | none (formula without a cached value)"""
    if v is None:
        return ("none", None) if has_formula else None
    if isinstance(v, bool):
        return "bool", v
    if isinstance(v, (int, float)):
        return "number", v
    if isinstance(v, (datetime, date)):
        return "date", iso(v)
    if isinstance(v, dtime):
        return "text", v.isoformat()
    s = str(v)
    if not s.strip():
        return ("none", None) if has_formula else None
    if s.strip() in EXCEL_ERRORS:
        return "error", s.strip()
    return "text", s


# --------------------------------------------------------------------------- #
# One worksheet as a grid
# --------------------------------------------------------------------------- #
class Grid:
    """Cached values, formulas, number formats and merged ranges of one worksheet.
    Built by grid_xlsx() (a sheet read in read-only mode) or grid_csv()."""

    def __init__(self):
        self.value: dict[tuple[int, int], object] = {}
        self.formula: dict[tuple[int, int], str] = {}
        self.fmt: dict[tuple[int, int], str] = {}
        self.anchor: dict[tuple[int, int], tuple[int, int]] = {}
        self.unmap = None          # after hide(): grid position -> position on the sheet

    def put(self, rc, v, is_date: bool = False, number_format: str | None = None):
        """One cell of the sheet with formulas (as stored, not calculated)."""
        if v is None:
            return
        if isinstance(v, str) and v.startswith("="):
            self.formula[rc] = v
        elif hasattr(v, "text") and not isinstance(v, str):              # ArrayFormula / DataTableFormula
            self.formula[rc] = getattr(v, "text", None) or str(v)
        else:
            self.value[rc] = v
        if is_date:
            self.fmt[rc] = number_format

    def merge(self, refs):
        """Merged ranges ('A1:C2'). As openpyxl does on a normal load, only the top-left cell of a
        range keeps its content; the cells it covers are emptied."""
        for ref in refs:
            c0, r0, c1, r1 = range_boundaries(ref)
            top = (r0, c0)
            for r in range(r0, r1 + 1):
                for c in range(c0, c1 + 1):
                    if (r, c) != top:
                        self.value.pop((r, c), None)
                        self.formula.pop((r, c), None)
                        self.fmt.pop((r, c), None)
                    self.anchor[(r, c)] = top

    def hide(self, rows: set, cols: set):
        """Leave out hidden rows and columns, as on screen: their cells are dropped and the visible
        cells close up (a hidden column between Jan and Mar does not split the table). Tables and
        headers are then found on what is visible; self.unmap gives each cell's place on the sheet.
        A merged range keeps its content when part of it is visible: the content moves to its
        first visible cell, as Excel shows it."""
        if not rows and not cols:
            return
        hidden = lambda rc: rc[0] in rows or rc[1] in cols
        ranges = defaultdict(list)
        for rc, top in self.anchor.items():
            ranges[top].append(rc)
        anchor = {}
        for top, members in ranges.items():
            shown = sorted(rc for rc in members if not hidden(rc))
            if not shown:
                continue
            new = shown[0]
            if new != top:                                               # the content moves with the range
                for d in (self.value, self.formula, self.fmt):
                    if top in d:
                        d[new] = d.pop(top)
            for rc in shown:
                anchor[rc] = new
        self.anchor = anchor
        for d in (self.value, self.formula, self.fmt):
            for rc in [rc for rc in d if hidden(rc)]:
                del d[rc]

        # close up: row r becomes r minus the hidden rows above it (columns likewise)
        hr, hc = sorted(rows), sorted(cols)
        new = lambda rc: (rc[0] - bisect_left(hr, rc[0]), rc[1] - bisect_left(hc, rc[1]))
        self.value = {new(rc): v for rc, v in self.value.items()}
        self.formula = {new(rc): v for rc, v in self.formula.items()}
        self.fmt = {new(rc): v for rc, v in self.fmt.items()}
        self.anchor = {new(rc): new(top) for rc, top in self.anchor.items()}

        def back(x, h):                    # the visible row (column) that closed up to x
            o = x
            while (o2 := x + bisect_right(h, o)) != o:
                o = o2
            return o
        self.unmap = lambda r, c: (back(r, hr), back(c, hc))

    def raw(self, r, c):
        return self.value.get((r, c))

    def lab(self, r, c):
        """Value with merged cells resolved (for headers)."""
        return self.value.get(self.anchor.get((r, c), (r, c)))

    def label_at(self, r, c) -> Label | None:
        rc = self.anchor.get((r, c), (r, c))
        lb = to_label(self.value.get(rc), self.fmt.get(rc))
        if lb is not None:
            lb.origin = rc
        return lb

    def has_content(self, r, c) -> bool:
        rc = self.anchor.get((r, c), (r, c))
        return rc in self.formula or not is_empty(self.value.get(rc))

    def numeric_like(self, r, c) -> bool:
        v = self.raw(r, c)
        return is_number(v) or ((r, c) in self.formula and not is_label(v))

    def occupied(self) -> set[tuple[int, int]]:
        cells = {rc for rc in set(self.value) | set(self.formula) if self.has_content(*rc)}
        cells |= {rc for rc, a in self.anchor.items() if self.has_content(*a)}
        return cells


# --------------------------------------------------------------------------- #
# Tables on a sheet
# --------------------------------------------------------------------------- #
def find_regions(occupied: set) -> list[tuple[int, int, int, int]]:
    """Connected blocks of filled cells (8-neighbourhood) -> bounding boxes (r0, c0, r1, c1)."""
    seen, boxes = set(), []
    for start in sorted(occupied):
        if start in seen:
            continue
        stack, rows, cols = [start], [], []
        seen.add(start)
        while stack:
            r, c = stack.pop()
            rows.append(r); cols.append(c)
            for dr in (-1, 0, 1):
                for dc in (-1, 0, 1):
                    nb = (r + dr, c + dc)
                    if nb in occupied and nb not in seen:
                        seen.add(nb)
                        stack.append(nb)
        boxes.append((min(rows), min(cols), max(rows), max(cols)))
    changed = True                                   # merge overlapping boxes
    while changed:
        changed, out = False, []
        for a in boxes:
            for i, b in enumerate(out):
                if not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1]):
                    out[i] = (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))
                    changed = True
                    break
            else:
                out.append(a)
        boxes = out
    return sorted(boxes)


def _anchors(g: Grid, r, c0, c1):
    out = []
    for c in range(c0, c1 + 1):
        a = g.anchor.get((r, c), (r, c))
        if g.has_content(*a) and a not in out:
            out.append(a)
    return out


def split_stacked(g: Grid, box):
    """Blocks stacked without a blank row are one region; split it where a new block starts:
    after a row with numbers, one or more rows holding a single label (the block's title),
    then a header row (>= 2 cells, >= 80 % labels). Yields the boxes."""
    r0, c0, r1, c1 = box

    def single_label(r):
        a = _anchors(g, r, c0, c1)
        return len(a) == 1 and is_label(g.lab(*a[0]))

    def header(r):
        vals = [g.lab(r, c) for c in range(c0, c1 + 1) if g.has_content(r, c)]
        return len(vals) >= 2 and sum(is_label(v) for v in vals) / len(vals) >= 0.8

    def has_number(r):
        return any(g.numeric_like(r, c) for c in range(c0, c1 + 1))

    start = r0
    r = r0 + 1
    while r <= r1:
        if single_label(r) and has_number(r - 1):
            rr = r
            while rr <= r1 and single_label(rr):
                rr += 1
            if rr <= r1 and header(rr):
                yield (start, c0, r - 1, c1)
                start = r
                r = rr
        r += 1
    yield (start, c0, r1, c1)


class Table:
    __slots__ = ("box", "kind", "titles", "title_labels", "header_rows", "col_header_rows", "row_header_cols", "region",
                 "keyed")

    @property
    def ref(self):
        r0, c0, r1, c1 = self.box
        return f"{get_column_letter(c0)}{r0}:{get_column_letter(c1)}{r1}"


def analyse(g: Grid, box) -> Table | None:
    """Title rows, header rows, row-header columns and 1D/2D kind of one block."""
    r0, c0, r1, c1 = box
    titles, r = [], r0
    while r < r1 and c1 > c0 and len(_anchors(g, r, c0, c1)) <= 1:      # single-label rows on top
        titles.append(r)
        r += 1
    h0 = r
    if h0 >= r1:
        return None
    body_rows = range(h0 + 1, r1 + 1)

    def row_label(v):                                                    # an Excel error (#REF!) is not a label
        return is_label(v) and str(v).strip() not in EXCEL_ERRORS

    def body_vals(c):
        """Values of column c below the first header row, merged cells resolved; the covered part
        of a merged header cell (e.g. 'Year' merged over A1:A3) is header, not body."""
        return [g.lab(rr, c) for rr in body_rows
                if g.has_content(rr, c) and g.anchor.get((rr, c), (rr, c))[0] > h0]

    def label_col(c):
        vals = body_vals(c)
        return bool(vals) and sum(row_label(v) for v in vals) / len(vals) >= 0.8

    def number_col(c):
        vals = body_vals(c)
        return bool(vals) and sum(is_number(v) for v in vals) / len(vals) >= 0.8

    # 1. a key column first: the leftmost column whose values on the data rows (rows with a number in
    #    another column) are all filled and all different, and are labels or a row counter (1, 2, 3 ...).
    #    The index is the key plus the label columns directly left of it (its parent levels); columns of
    #    numbers left of it (e.g. '# OF Trades' left of 'Index') stay values.
    def key_col(c):
        rows_ = [rr for rr in body_rows if any(g.numeric_like(rr, x) for x in range(c0, c1 + 1) if x != c)]
        vals = [g.lab(rr, c) for rr in rows_]
        if len(vals) < 2 or any(is_empty(v) for v in vals):
            return False
        if len({" ".join(str(v).split()).casefold() for v in vals}) != len(vals):
            return False
        if all(row_label(v) for v in vals):
            return True
        ints = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer()]
        return len(ints) == len(vals) and all(b - a == 1 for a, b in zip(ints, ints[1:]))

    key = next((c for c in range(c0, c1) if key_col(c)), None)
    if key is not None:
        rh = [key]
        while rh[0] > c0 and label_col(rh[0] - 1):
            rh.insert(0, rh[0] - 1)
        t = Table()
        t.box, t.region, t.kind, t.keyed = box, None, "2D", True
        t.row_header_cols = rh
        val_cols = [c for c in range(c0, c1 + 1) if c not in rh]
        return _headers(g, t, h0, r1, c0, c1, val_cols, titles)

    # 2. no key: numbers left of the index are part of the index (e.g. years 1991, 1992 left of
    #    Winter/Spring/Fall): leading columns of numbers count as row headers when a label column follows
    lead = 0
    while c0 + lead < c1 - 1 and number_col(c0 + lead):
        lead += 1
    if not label_col(c0 + lead):
        lead = 0

    k = lead                                                             # leftmost text columns
    for c in range(c0 + lead, c1):
        vals = body_vals(c)
        if not vals:                                                     # an empty column at the left edge (under a title)
            if k == lead and label_col(c + 1):
                k += 1                                                   # is part of the row headers when the next
                continue                                                 # column is a row-label column
            break
        if sum(row_label(v) for v in vals) / len(vals) < 0.8:
            break
        k += 1
    val_cols = list(range(c0 + k, c1 + 1))
    body = [(rr, c) for rr in body_rows for c in val_cols if g.has_content(rr, c)]
    num_share = sum(g.numeric_like(*p) for p in body) / len(body) if body else 0
    t = Table()
    t.box = box
    t.region = None
    t.keyed = False
    t.kind = "2D" if k > 0 and num_share >= 0.5 else "1D"
    t.row_header_cols = list(range(c0, c0 + k)) if t.kind == "2D" else []
    if t.kind == "1D":
        val_cols = list(range(c0, c1 + 1))
    return _headers(g, t, h0, r1, c0, c1, val_cols, titles)


def _headers(g: Grid, t: Table, h0, r1, c0, c1, val_cols, titles) -> Table:
    """Column-header rows of each value column, and the table's title labels."""

    def run(row, c):
        """Columns of the contiguous run of filled cells in `row` around column c (an empty c takes the
        run on its left). Side-by-side tables that touch share a region, so header rows are judged per run."""
        cc = c
        while cc >= c0 and not g.has_content(row, cc):
            cc -= 1
        if cc < c0:
            return range(0)
        a = b = cc
        while a - 1 >= c0 and g.has_content(row, a - 1):
            a -= 1
        while b + 1 <= c1 and g.has_content(row, b + 1):
            b += 1
        return range(a, b + 1)

    def header_like(row, c):                    # >= 80 % of the run are labels
        vals = [g.lab(row, x) for x in run(row, c)]
        return bool(vals) and sum(is_label(v) for v in vals) / len(vals) >= 0.8

    col_hdr = {}                                                         # multi-level column headers, per column:
    for c in val_cols:                                                   # header-like rows under the first header row
        hdr, rr = [h0], h0 + 1
        while rr < r1 and len(hdr) < 4 and header_like(rr, c):
            hdr.append(rr)
            rr += 1
        if t.kind == "1D" and len(hdr) > 1 and not any(g.numeric_like(rr, x) for x in run(rr, c)):
            hdr = [h0]                                 # a 1D table of text records: those rows are data, not headers
        col_hdr[c] = hdr
    t.col_header_rows = col_hdr
    t.titles = titles
    t.header_rows = sorted({r for rows in col_hdr.values() for r in rows})
    t.title_labels = [lb for tr in titles for a in _anchors(g, tr, c0, c1)[:1] if (lb := g.label_at(*a))]
    return t


def table_cells(g: Grid, t: Table):
    """Yield (row, col, [header labels]) for every data cell of a table."""
    r0, c0, r1, c1 = t.box
    rh = t.row_header_cols
    val_cols = [c for c in range(c0, c1 + 1) if c not in rh]

    ch = t.col_header_rows                                               # header rows of each value column
    col_path: dict[int, list[Label]] = {c: [] for c in val_cols}         # column headers, upper levels filled from the left
    for h in t.header_rows:
        carry = None
        for c in val_cols:
            if h not in ch[c]:
                carry = None                                              # not a header row for this column
                continue
            v = g.lab(h, c)
            if is_empty(v):
                lb = None if h == ch[c][-1] else carry                    # blank upper-level header: take the one on its left
            else:
                lb = g.label_at(h, c) if is_label(v) else None            # numbers are never headers
                carry = lb
            if lb is not None:
                col_path[c].append(lb)

    # corner labels: header-row labels above the row headers, or above a column with no values,
    # name the whole table
    body_rows = range(t.header_rows[-1] + 1, r1 + 1)
    corner = []
    for h in t.header_rows:
        for c in range(c0, c1 + 1):
            if c in ch and any(g.has_content(rr, c) for rr in body_rows):
                break                                                     # first column with values: corner ends
            if is_label(g.lab(h, c)) and g.anchor.get((h, c), (h, c)) == (h, c):
                corner.append(g.label_at(h, c))

    def left_label(r, c):
        for cc in range(c - 1, c0 - 1, -1):
            if is_label(g.lab(r, cc)):
                return [g.label_at(r, cc)]
        return []

    def in_header(r, c):
        return r in ch[c] if c in ch else r == t.header_rows[0]

    for r in t.titles + t.header_rows:                                   # numbers in title/header rows are values
        for c in range(c0, c1 + 1):
            if (r in t.titles or in_header(r, c)) and g.numeric_like(r, c) \
                    and g.anchor.get((r, c), (r, c)) == (r, c):
                yield r, c, left_label(r, c)

    carry = [None] * len(rh)
    for r in range(t.header_rows[0] + 1, r1 + 1):
        if not any(g.has_content(r, c) for c in range(c0, c1 + 1)):
            continue
        row_path: list[Label] = []
        for i, c in enumerate(rh):                                        # row headers: a blank is filled down (ffill),
            v = g.lab(r, c)                                               # on every level; a new label on a level
            lb = g.label_at(r, c)                                         # clears the levels below it
            if is_empty(v):
                lb = carry[i]
            else:
                carry[i] = lb
                for j in range(i + 1, len(rh)):
                    carry[j] = None
            if lb is not None:
                row_path.append(lb)
        for c in val_cols:
            if (r, c) in g.anchor and g.anchor[(r, c)] != (r, c):
                continue                                                  # covered part of a merged cell
            if g.has_content(r, c) and r > ch[c][-1]:
                yield r, c, col_path[c] + row_path + corner


# --------------------------------------------------------------------------- #
# Labels around a cell
# --------------------------------------------------------------------------- #
class Layout:
    """Labels left of a cell in its row, and from header rows above it.
    A header row is judged within the table's columns: >= 2 filled cells, >= 80 % labels."""

    def __init__(self, g: Grid, max_above: int = 4, look_up: int = 80):
        self.g, self.max_above, self.look_up = g, max_above, look_up
        self.rows: dict[int, dict[int, object]] = defaultdict(dict)
        for (r, c), v in g.value.items():
            if not is_empty(v):
                self.rows[r][c] = v
        for (r, c), a in g.anchor.items():        # a merged label covers all its cells ('Grade' over H2:H3)
            if (r, c) != a and not is_empty(g.value.get(a)):
                self.rows[r][c] = g.value[a]
        self.sorted_rows = sorted(self.rows)
        self.boxes = []
        self._spans: dict = {}

    def headers(self, c0, c1):
        if (c0, c1) not in self._spans:
            hdr, filled = [], {}
            for r in self.sorted_rows:
                cs = {c: v for c, v in self.rows[r].items() if c0 <= c <= c1}
                if len(cs) >= 2 and sum(is_label(v) for v in cs.values()) / len(cs) >= 0.8:
                    hdr.append(r)
                    filled[r] = sorted(c for c, v in cs.items() if is_label(v))
            self._spans[(c0, c1)] = (hdr, filled)
        return self._spans[(c0, c1)]

    def labels(self, r, c, c0, c1, r0, floor, header_end, left: bool = True) -> list[Label]:
        """Text left of the cell in its row, plus header rows: the table's own (any distance) and
        those above the table within `look_up` rows and below `floor` (the bottom of the nearest
        table above that has its own header row). Rows in the table body are never headers."""
        out = [self.g.label_at(r, cc) for cc, v in sorted(self.rows[r].items())
               if left and c0 <= cc < c and is_label(v)]
        hdr, filled = self.headers(c0, c1)
        i = bisect_right(hdr, min(r - 1, header_end))
        for hr in reversed(hdr[max(0, i - self.max_above):i]):
            if hr < r0 and (r - hr > self.look_up or hr <= floor):
                break
            cols = filled[hr]
            j = bisect_right(cols, c) - 1
            if j >= 0 and self._reaches(hr, cols[j], c):
                out.append(self.g.label_at(hr, cols[j]))
        return [lb for lb in out if lb is not None]

    def _reaches(self, hr, hc, c) -> bool:
        """A header label inside a table only labels columns of that table."""
        for r0, c0, r1, c1 in self.boxes:
            if r0 <= hr <= r1 and c0 <= hc <= c1:
                return c0 <= c <= c1
        return True


# --------------------------------------------------------------------------- #
# Value cells in columns
# --------------------------------------------------------------------------- #
VALUE_TYPES = ("number", "text", "date", "bool", "error", "none")
_VTYPE = {t: i for i, t in enumerate(VALUE_TYPES)}


class Labels:
    """The distinct labels of one sheet. A label is one text and date, read from one cell, by one
    source (header | title | layout); cells refer to labels by id."""
    __slots__ = ("ids", "text", "date_value", "origin", "via")

    def __init__(self):
        self.ids: dict = {}
        self.text, self.date_value, self.origin, self.via = [], [], [], []

    def id(self, lb: Label, via: str) -> int:
        k = (lb.text, lb.date_value, lb.origin, via)
        i = self.ids.get(k)
        if i is None:
            i = self.ids[k] = len(self.text)
            self.text.append(lb.text)
            self.date_value.append(lb.date_value)
            self.origin.append(lb.origin)
            self.via.append(via)
        return i


class TableCells:
    """The value cells of one table, in columns: cell i is at row[i], col[i] with value_type
    VALUE_TYPES[vtype[i]] and value[i]; formula: {i: formula}. Its labels are
    lab[off[i]:off[i + 1]] (ids into the sheet's Labels), in order header, title, layout."""
    __slots__ = ("ref", "row", "col", "vtype", "value", "formula", "off", "lab")

    def __init__(self, ref: str):
        self.ref = ref
        self.row, self.col, self.vtype = array("i"), array("i"), array("b")
        self.value: list = []
        self.formula: dict[int, str] = {}
        self.off, self.lab = array("q", [0]), array("i")

    def __len__(self):
        return len(self.value)

    def add(self, r, c, value_type, value, formula, label_ids):
        if formula is not None:
            self.formula[len(self.value)] = formula
        self.row.append(r)
        self.col.append(c)
        self.vtype.append(_VTYPE[value_type])
        self.value.append(value)
        self.lab.extend(label_ids)
        self.off.append(len(self.lab))


class SheetCells:
    """The value cells of one sheet: its distinct labels, and its tables' cells in columns."""
    __slots__ = ("file", "sheet", "labels", "tables")

    def __init__(self, file: str, sheet: str):
        self.file, self.sheet, self.labels, self.tables = file, sheet, Labels(), []

    def __len__(self):
        return sum(len(t) for t in self.tables)

    def dicts(self) -> list[dict]:
        """The cells as dicts (extract()'s form): file, sheet, table, address, row, col,
        value_type, value, formula, via, labels."""
        L, out = self.labels, []
        for t in self.tables:
            for i in range(len(t)):
                ids = t.lab[t.off[i]:t.off[i + 1]]
                r, c = t.row[i], t.col[i]
                out.append({"file": self.file, "sheet": self.sheet, "table": t.ref,
                            "address": f"{get_column_letter(c)}{r}", "row": r, "col": c,
                            "value_type": VALUE_TYPES[t.vtype[i]], "value": t.value[i],
                            "formula": t.formula.get(i),
                            "via": [L.via[j] for j in ids],
                            "labels": [Label(L.text[j], L.date_value[j], L.origin[j]) for j in ids]})
        return out

    @staticmethod
    def from_dicts(cells: list[dict]) -> list["SheetCells"]:
        """Cell dicts (extract()'s form) -> one SheetCells per sheet, in the order first seen."""
        sheets: dict = {}
        tables: dict = {}
        for d in cells:
            sc = sheets.get((d["file"], d["sheet"]))
            if sc is None:
                sc = sheets[(d["file"], d["sheet"])] = SheetCells(d["file"], d["sheet"])
            tc = tables.get((d["file"], d["sheet"], d["table"]))
            if tc is None:
                tc = tables[(d["file"], d["sheet"], d["table"])] = TableCells(d["table"])
                sc.tables.append(tc)
            tc.add(d["row"], d["col"], d["value_type"], d["value"], d["formula"],
                   [sc.labels.id(lb, via) for lb, via in zip(d["labels"], d["via"])])
        return list(sheets.values())


# --------------------------------------------------------------------------- #
# Files
# --------------------------------------------------------------------------- #
_CSV_NUM = re.compile(r"^\(?-?[\d,]*\.?\d+%?\)?%?$")


def csv_value(s: str, text: bool = False):
    """A CSV field as Excel would show it: '105,299' -> 105299, '(491)' -> -491, '12%' -> 0.12,
    '(5%)' -> -0.05.
    Whole numbers are read exactly, never through a float ('9007199254740993' stays that integer);
    text=True keeps the field as written (an ID '00123' keeps its zeros)."""
    t = s.strip()
    if not t:
        return None
    if text or not _CSV_NUM.match(t) or t.count("%") > 1 or t.startswith("(") != t.rstrip("%").endswith(")"):
        return t                                         # not a number (or brackets that do not pair up)
    neg = t.startswith("(")                              # accounting negative: (491), (5%), (5)%
    pct = "%" in t
    try:
        x = Decimal(t.strip("()%").replace(",", ""))
    except InvalidOperation:
        return t
    x = -x if neg else x
    if pct:
        return float(x / 100)
    return int(x) if x == x.to_integral_value() else float(x)


def grid_csv(path: str, text_columns=None) -> Grid:
    """A CSV file as a Grid, cell by cell as Excel would hold it (a text starting with '=' is a
    formula, and is also its own cached value).
    text_columns: columns kept as written, named by their header (first row) or letter ('A')."""
    g = Grid()
    want, as_text = set(text_columns or ()), set()
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        for r, row in enumerate(csv.reader(f), 1):
            if r == 1 and want:
                as_text = {c for c, field in enumerate(row, 1)
                           if field.strip() in want or get_column_letter(c) in want}
            for c, field in enumerate(row, 1):
                v = csv_value(field, c in as_text)
                if isinstance(v, str):
                    v = v[:32767]                                        # Excel's cell limit
                g.put((r, c), v)
                if (r, c) in g.formula:
                    g.value[(r, c)] = v
    return g


def _hidden(attrs: dict, size: str) -> bool:
    """A row or column hidden in Excel: hidden="1", or a height / width of 0."""
    if attrs.get("hidden") in ("1", "true"):
        return True
    try:
        return float(attrs[size]) == 0
    except (KeyError, ValueError):
        return False


def _sheet_rows(ws, merges: list | None = None, hidden: dict | None = None):
    """The cells of a read-only worksheet as openpyxl's sheet parser reads them (the same parser a
    normal load uses): dicts with row, column, value, data_type, style_id. When the rows are done,
    the sheet's merged ranges (stored after them) are added to `merges`, and its hidden rows and
    columns to hidden["rows"] / hidden["cols"]."""
    wb = ws.parent
    with ws._get_source() as src:
        parser = WorkSheetParser(src, ws._shared_strings, data_only=wb.data_only, epoch=wb.epoch,
                                 date_formats=wb._date_formats, timedelta_formats=wb._timedelta_formats)
        for _, row in parser.parse():
            yield from row
        if merges is not None and parser.merged_cells:
            merges += [cr.ref for cr in parser.merged_cells.mergeCell]
        if hidden is not None:
            hidden["rows"] = {int(r) for r, a in parser.row_dimensions.items() if _hidden(a, "ht")}
            hidden["cols"] = {c for a in parser.column_dimensions.values() if _hidden(a, "width")
                              for c in range(int(a["min"]), int(a.get("max", a["min"])) + 1)}


def grid_xlsx(ws_f, ws_v, visible_only: bool = False) -> Grid:
    """One sheet, read in read-only mode: formulas from ws_f, their cached results from ws_v
    (the same sheet of the workbook opened with data_only=True).
    visible_only: hidden rows and columns are left out, as on screen (see Grid.hide)."""
    styles = {}

    def style(ws, sid):                                  # number format and is-date, per style
        if sid not in styles:
            nf = ReadOnlyCell(ws, 1, 1, None, "n", sid).number_format
            styles[sid] = (nf, is_date_format(nf))
        return styles[sid]

    def is_date(ws, cell):                               # as openpyxl's Cell.is_date
        t = cell["data_type"]
        return t == "d" or (t == "n" and style(ws, cell["style_id"])[1])

    g, merges, hidden = Grid(), [], {} if visible_only else None
    for cell in _sheet_rows(ws_f, merges, hidden):
        d = is_date(ws_f, cell)
        g.put((cell["row"], cell["column"]), cell["value"], d, style(ws_f, cell["style_id"])[0] if d else None)
    g.merge(merges)
    if g.formula:                                        # cached results
        for rc in g.formula:
            g.value[rc] = None
        for cell in _sheet_rows(ws_v):
            rc = (cell["row"], cell["column"])
            if rc in g.formula:
                g.value[rc] = cell["value"]
                if is_date(ws_v, cell):
                    g.fmt[rc] = style(ws_v, cell["style_id"])[0]
    if visible_only:
        g.hide(hidden["rows"], hidden["cols"])
    return g


def extract(path: str, sheets=None, text_columns=None, visible_only: bool = False) -> list[dict]:
    """Every value cell of a workbook (.xlsx / .xlsm / .csv) with its labels.
    sheets: a sheet name or a list of them to read only those (default: every sheet).
    text_columns: CSV columns kept as text, by header or letter (see grid_csv).
    visible_only: hidden rows and columns are left out, as on screen (Excel files).

    Returns dicts: file, sheet, table, address, row, col, value_type, value, formula,
    labels (list of Label, unordered in meaning), via (where each label came from:
    header | title | layout; same order as labels)."""
    return [c for sheet in extract_sheets(path, sheets, text_columns=text_columns, visible_only=visible_only)
            for c in sheet.dicts()]


def extract_sheets(path: str, sheets=None, names: list | None = None, text_columns=None,
                   visible_only: bool = False):
    """Like extract(), one sheet at a time: yields the value cells of each sheet as SheetCells.
    Only one sheet's grid and cells are held at a time; sheets not asked for are not read.
    names: a list to receive the names of all sheets in the file."""
    want = {sheets} if isinstance(sheets, str) else set(sheets) if sheets is not None else None
    name = os.path.basename(path)
    if path.lower().endswith(".csv"):
        title = re.sub(r"[\\/*?:\[\]]", "_", os.path.splitext(name)[0])[:31]
        if names is not None:
            names.append(title)
        if want is None or title in want:
            yield _sheet_cells(name, title, grid_csv(path, text_columns))
        return
    wb_f = load_workbook(path, read_only=True, data_only=False)     # read-only: one sheet at a time
    wb_v = load_workbook(path, read_only=True, data_only=True)
    try:
        if names is not None:
            names += [ws.title for ws in wb_f.worksheets]
        for ws in wb_f.worksheets:
            if want is None or ws.title in want:
                yield _sheet_cells(name, ws.title, grid_xlsx(ws, wb_v[ws.title], visible_only))
    finally:
        wb_v.close()
        wb_f.close()


def _sheet_cells(name: str, sheet: str, g: Grid) -> SheetCells:
    """The value cells of one worksheet with their labels."""
    out = SheetCells(name, sheet)
    layout = Layout(g)
    um = g.unmap                                  # visible_only: grid positions -> sheet positions
    if um is None:
        label_id = out.labels.id
    else:
        def label_id(lb, via):
            o = um(*lb.origin) if lb.origin else None
            return out.labels.id(Label(lb.text, lb.date_value, o), via)

    # tables, and label-only blocks just above them (captions)
    parts = {}
    for region in find_regions(g.occupied()):
        for b in split_stacked(g, region):
            parts[b] = region
    tables, captions = [], []
    for box in parts:
        t = analyse(g, box)
        if t is not None:
            first = tables[-1] if tables and tables[-1].region == parts[box] else None
            t.region = parts[box]
            if first is not None and len(first.titles) > len(t.titles):
                # stacked blocks: the first block's extra top title rows head them all
                shared = first.titles[:len(first.titles) - len(t.titles)]
                t.title_labels = [lb for tr in shared for a in _anchors(g, tr, box[1], box[3])[:1]
                                  if (lb := g.label_at(*a))] + t.title_labels
            tables.append(t)
        elif all(is_label(g.lab(r, c)) for r in range(box[0], box[2] + 1)
                 for c in range(box[1], box[3] + 1) if g.has_content(r, c)):
            captions.append(box)
    layout.boxes = [t.box for t in tables]
    for t in tables:
        r0, c0, r1, c1 = t.box
        for cb in captions:
            if 0 < r0 - cb[2] <= 3 and cb[1] <= c1 and cb[3] >= c0:
                t.title_labels = [lb for r in range(cb[0], cb[2] + 1) for a in _anchors(g, r, cb[1], cb[3])
                                  if (lb := g.label_at(*a))] + t.title_labels

    def own_header(t):
        """The table's first row is a header row, or puts a label above a column of numbers."""
        r0, c0, r1, c1 = t.box
        h = t.header_rows[0]
        if h in layout.headers(c0, c1)[0]:
            return True
        for c in range(c0, c1 + 1):
            if is_label(g.lab(h, c)) and g.anchor.get((h, c), (h, c)) == (h, c):
                body = [g.numeric_like(r, c) for r in range(h + 1, r1 + 1) if g.has_content(r, c)]
                if body and sum(body) / len(body) >= 0.8:
                    return True
        return False

    for t in tables:
        box = t.box
        # a table with its own header never takes header rows from above the nearest table
        # above it; a block without its own header shares that table's header, nothing higher
        blockers = [o.box for o in tables if o is not t and o.box[2] < box[0]
                    and o.box[1] <= box[3] and o.box[3] >= box[1]
                    and o.header_rows[0] in layout.headers(o.box[1], o.box[3])[0]]
        own = own_header(t)
        floors: dict[int, int] = {}

        def floor(c):
            if c not in floors:
                near = max((b for b in blockers if b[1] <= c <= b[3]), key=lambda b: b[2], default=None)
                floors[c] = 0 if near is None else near[2] if own else near[0] - 1
            return floors[c]

        c0, c1 = box[1], box[3]
        if um is None:
            tc = TableCells(t.ref)
        else:
            (a, b), (y, z) = um(box[0], box[1]), um(box[2], box[3])        # the table's span on the sheet
            tc = TableCells(f"{get_column_letter(b)}{a}:{get_column_letter(z)}{y}")
        out.tables.append(tc)
        for r, c, header in table_cells(g, t):
            v = cell_value(g.raw(r, c), (r, c) in g.formula)
            if v is None:
                continue
            ids, seen = [], set()
            head_end = min(r - 1, t.col_header_rows.get(c, t.header_rows)[-1])
            for src, lbs in (("header", header), ("title", t.title_labels),
                             ("layout", layout.labels(r, c, c0, c1, box[0], floor(c), head_end,
                                                      left=not t.keyed))):
                for lb in lbs:
                    if lb.key() not in seen:              # first source wins: header > title > layout
                        seen.add(lb.key())
                        ids.append(label_id(lb, src))
            tc.add(*(um(r, c) if um else (r, c)), v[0], v[1], g.formula.get((r, c)), ids)
    out.tables = [tc for tc in out.tables if len(tc)]
    return out
