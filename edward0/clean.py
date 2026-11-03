"""
clean.py - value cells with label sets -> relational records.

Taken from silver.py (same rules), without SQLite. Uses only the label sets, never positions:

1. Clean labels   whitespace collapsed, case folded, trailing ':' dropped; dates compared by
                  their ISO date. Error labels (#REF!, #N/A ...) and notes (> 60 characters)
                  are dropped; header labels are never notes.
2. Raw data only  typed values, and formulas that are only arithmetic on numbers
                  (=1312+350000/4342). Cells calculated from other cells are left out.
3. Axes           inside each table, labels that never sit on the same cell and appear with
                  the same other labels form one unnamed axis, e.g. {May, Jun, July}.
                  Labels on every cell of the table are its context.
4. Datasets       tables whose axes match (Jaccard >= 0.5 per axis) are one dataset, e.g. the
                  same template on 30 dated sheets.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict

MAX_AXES = 6           # axis columns in the output; smaller axes go to `other`
SAME_AXIS = 0.5        # Jaccard of partner labels for a label to join an axis
SAME_DATASET = 0.5     # Jaccard per axis for two tables to be one dataset
NOTE_LEN = 60          # longer labels are notes, not identifiers

ERROR_LABELS = {"#ref!", "#value!", "#div/0!", "#n/a", "#name?", "#num!", "#null!", "#spill!", "#calc!"}
_CONSTANT = re.compile(r"=[\d\s.+\-*/^()%,eE]*")


def is_raw(formula) -> bool:
    """A typed value, or a formula that is only arithmetic on numbers (no cells, names or functions)."""
    return formula is None or bool(_CONSTANT.fullmatch(formula.strip()))


def clean_key(text: str, date_value: str = "", header: bool = False) -> tuple[str, str | None]:
    """-> (comparison key, drop reason or None).
    A long text is a note only when it is not a row/column header: a long ID in the index is kept."""
    t = re.sub(r"\s+", " ", text or "").strip()
    if not t:
        return "", "empty"
    if t.casefold() in ERROR_LABELS:
        return t.casefold(), "error"
    if len(t) > NOTE_LEN and not header:
        return t.casefold(), "note"
    if date_value:
        return "date:" + date_value, None
    return t.rstrip(":").strip().casefold(), None


def jaccard(a: set, b: set) -> float:
    return len(a & b) / len(a | b) if a or b else 1.0


# --------------------------------------------------------------------------- #
# Axes inside one table
# --------------------------------------------------------------------------- #
def table_axes(cells):
    """cells: label-key sets of the table's cells.
    Returns (context labels, axes) where axes is a list of label sets, largest coverage first.

    Linear in the cells: no per-label cell bitmasks (they were O(cells^2) memory on large tables).
    Two labels share a cell iff one is a partner of the other, so 'sits on the same cell as a
    member' is partners[k] ∩ members != ∅; and partners(axis) − members is kept up to date."""
    n = len(cells)
    count = Counter(k for s in cells for k in s)
    constants = {k for k, v in count.items() if v == n} if n > 1 else set(count)
    partners = defaultdict(set)
    for s in cells:
        var = s - constants if constants else s
        for k in var:
            partners[k] |= var
    for k, p in partners.items():
        p.discard(k)

    axes = []                                   # [members, partners of members − members]
    for k in sorted(partners, key=lambda x: (-count[x], x)):
        p = partners[k]
        best, best_score = None, SAME_AXIS
        for ax in axes:
            if not p.isdisjoint(ax[0]):
                continue                        # sits on the same cell as a member
            inter = len(p & ax[1])
            union = len(p) + len(ax[1]) - inter
            score = inter / union if union else 1.0
            if score >= best_score:
                best, best_score = ax, score
        if best is None:
            axes.append([{k}, set(p)])
        else:
            best[0].add(k)
            best[1] |= p                        # p has no member and not k (checked above)
    axis_of = {k: i for i, ax in enumerate(axes) for k in ax[0]}
    cover = Counter()
    for s in cells:
        cover.update({axis_of[k] for k in s if k in axis_of})
    order = sorted(range(len(axes)), key=lambda i: (-cover[i], min(axes[i][0])))
    return constants, [axes[i][0] for i in order]


# --------------------------------------------------------------------------- #
# Datasets across tables
# --------------------------------------------------------------------------- #
class Datasets:
    def __init__(self):
        self.rep: list[list[set]] = []        # representative axes (first table) per dataset
        self.flat: dict[frozenset, int] = {}  # tables without axes: grouped by their context
        self.index = defaultdict(set)         # label key -> datasets with it in an axis

    def match(self, axes: list[set], constants: set) -> int:
        if not axes:
            key = frozenset(constants)
            if key not in self.flat:
                self.flat[key] = self._new([])
            return self.flat[key]
        cand = set()
        for ax in axes:
            for k in ax:
                cand |= self.index.get(k, set())
        for d in sorted(cand):
            rep = self.rep[d]
            if len(rep) == len(axes) and self.mapping(axes, rep) is not None:
                return d
        d = self._new(axes)
        for ax in axes:
            for k in ax:
                self.index[k].add(d)
        return d

    def _new(self, axes):
        self.rep.append([set(a) for a in axes])
        return len(self.rep) - 1

    @staticmethod
    def mapping(axes, rep):
        """axis i of the table -> axis j of the dataset, every pair Jaccard >= SAME_DATASET."""
        pairs = sorted(((jaccard(a, r), i, j) for i, a in enumerate(axes) for j, r in enumerate(rep)), reverse=True)
        out, used = {}, set()
        for s, i, j in pairs:
            if s < SAME_DATASET:
                break
            if i not in out and j not in used:
                out[i] = j
                used.add(j)
        return out if len(out) == len(axes) else None


# --------------------------------------------------------------------------- #
# Records
# --------------------------------------------------------------------------- #
COLUMNS = ["file", "sheet", "table", "address", "row", "col", "value", "dataset", "context",
           *(f"axis{j + 1}" for j in range(MAX_AXES)), "other", "dropped", "is_formula", "dup"]


class Records:
    """Records in columns: one list per column, record i is position i in every list.

    cols   COLUMNS, plus `_index` (frozenset of each record's labels, when kept) and `_tid`
           (the record's table: a position in meta)
    meta   one entry per table, for relational.py to name the columns:
             key      (file, sheet, table)
             context  {label key: (text, source cell)} for the table's context labels
             origins  {axis j: source cells of the labels on that axis (dict keys, first seen first)}
             axes     axes j with a label on at least one record"""

    def __init__(self, cols: dict, meta: list):
        self.cols, self.meta = cols, meta

    def __len__(self):
        return len(self.cols["value"])


def _show(keys, own):
    """A record shows its cell's own spelling of each label."""
    return " · ".join(sorted({x for k in keys for x in own[k]}, key=str.casefold)) or None


class _Table:
    """One table after steps 1-3, kept compact: its axes, its table metadata and one entry per
    value cell in column lists (the cell dicts and their Label objects are not kept)."""
    __slots__ = ("key", "n", "constants", "main", "cols", "ax", "context", "origins")


def _sheet_tables(cells: list[dict], raw_only: bool, pool: dict, keep_index: bool) -> list[_Table]:
    """Steps 1-3 on one sheet's cells: clean labels, raw data only, axes per table.
    pool: shared strings, so a label's text is stored once however many records show it."""
    share = lambda x: pool.setdefault(x, x) if x is not None else None

    # 1. clean labels, 2. raw data only; cells without any kept label cannot be placed and are left out
    by_table = defaultdict(list)
    for c in cells:
        if c["value_type"] in ("error", "none") or (raw_only and not is_raw(c["formula"])):
            continue
        keys, own, dropped, origin = set(), defaultdict(set), [], {}
        for lb, via in zip(c["labels"], c["via"]):
            text = share(" ".join(lb.text.split()))
            key, reason = clean_key(lb.text, lb.date_value, via == "header")
            if reason is None:
                keys.add(key)
                own[key].add(text)
                origin.setdefault(key, lb.origin)               # where the label was read (first source)
            else:
                dropped.append(text)
        if keys:
            by_table[(c["file"], c["sheet"], c["table"])].append((c, frozenset(keys), own, dropped, origin))

    # 3. axes per table
    out = []
    for t, items in by_table.items():
        constants, axes = table_axes([x[1] for x in items])
        main, rest = axes[:MAX_AXES], axes[MAX_AXES:]
        rest_keys = set().union(*rest) if rest else set()
        items.sort(key=lambda x: (x[0]["row"], x[0]["col"]))    # sheet order: the metadata below
        cols = {k: [] for k in ("address", "row", "col", "value", "is_formula", "context", "other", "dropped")}
        if keep_index:
            cols["_index"] = []
        ax = [[] for _ in main]                                  # label on each axis, by the table's own axis order
        context, origins = {}, [{} for _ in main]               # origins: source cells, first seen first
        first = {}                                               # axis -> (first record with it, axis)
        for q, (c, keys, own, dropped, origin) in enumerate(items):
            ctx = constants & keys
            cols["address"].append(c["address"])
            cols["row"].append(c["row"])
            cols["col"].append(c["col"])
            cols["value"].append(c["value"])
            cols["is_formula"].append(c["formula"] is not None)
            cols["context"].append(share(_show(ctx, own)))
            for i, a in enumerate(main):
                h = keys & a
                ax[i].append(share(_show(h, own)) if h else None)
                if h and i not in first:
                    first[i] = (q, i)
                for k in h:
                    origins[i].setdefault(origin.get(k))
            cols["other"].append(share(_show(keys & rest_keys, own)))
            cols["dropped"].append(share(" · ".join(sorted(dropped, key=str.casefold)) or None))
            if keep_index:
                cols["_index"].append(frozenset(x for v in own.values() for x in v))
            for k in ctx:                                        # the table's context labels, with the
                context[k] = (_show({k}, own), origin.get(k))    # spelling of the last cell in sheet order
        tb = _Table()
        tb.key, tb.n, tb.constants, tb.main = t, len(items), constants, main
        tb.cols, tb.ax, tb.context = cols, ax, context
        tb.origins = {i: origins[i] for i in sorted(first, key=first.get)}   # axes in the order records show them
        out.append(tb)
    return out


def records(cells, raw_only: bool = True, keep_index: bool = True) -> Records:
    """cells from extract -> one record per value, as Records (columns + per-table metadata).

    cells:      an iterable of sheets, each a list of cell dicts (extract.extract_sheets), or one flat
                list of cell dicts (extract.extract). A table must not be split across sheets. Each
                sheet is reduced to its tables' axes and compact columns before the next one is read;
                datasets are matched once all sheets are in.
    keep_index: keep `_index`, the set of each record's labels (for raw() and read()'s `index`).

    Every label of a cell is in exactly one of context / axis1..6 / other / dropped.
    Records come in sheet order, row by row."""
    import numpy as np
    import pandas as pd

    if isinstance(cells, list) and (not cells or isinstance(cells[0], dict)):
        cells = [cells]
    pool, tables = {}, []
    for sheet in cells:
        tables += _sheet_tables(sheet, raw_only, pool, keep_index)
        del sheet
    pool.clear()

    # 4. datasets; largest tables first so a dataset's first table is its fullest
    names = COLUMNS[:-1] + ["_tid"] + (["_index"] if keep_index else [])
    out = {k: [] for k in names}
    meta, ds, size = [], Datasets(), Counter()
    tables.sort(key=lambda tb: (-tb.n, tb.key))
    for tid, tb in enumerate(tables):
        d = ds.match(tb.main, tb.constants)
        m = Datasets.mapping(tb.main, ds.rep[d]) if tb.main else {}
        n = tb.n
        f, sh, tref = tb.key
        out["file"] += [f] * n
        out["sheet"] += [sh] * n
        out["table"] += [tref] * n
        out["dataset"] += [d] * n
        out["_tid"] += [tid] * n
        for k, v in tb.cols.items():
            out[k] += v
        local = {j: i for i, j in m.items()}
        for j in range(MAX_AXES):
            out[f"axis{j + 1}"] += tb.ax[local[j]] if j in local else [None] * n
        meta.append({"key": tb.key, "context": tb.context,
                     "origins": {m[i]: o for i, o in tb.origins.items()},
                     "axes": {m[i] for i, a in enumerate(tb.ax) if any(x is not None for x in a)}})
        size[d] += n
        tables[tid] = None                                       # free the table's columns
    if not out["value"]:
        return Records({k: [] for k in COLUMNS + names[len(COLUMNS) - 1:]}, meta)

    # records of a dataset with the same full key: > 1 means a dimension is missing
    code = None
    for k in ["dataset", "file", "sheet", "context", *(f"axis{j + 1}" for j in range(MAX_AXES)), "other"]:
        c, u = pd.factorize(np.array(out[k], dtype=object))
        c = c.astype(np.int64) + 1                               # None -> 0
        code = c if code is None else pd.factorize(code * (len(u) + 1) + c)[0]
    out["dup"] = np.bincount(code)[code].tolist()
    del code

    # number datasets 1, 2, ... by size
    new_id = {d: i for i, (d, _) in enumerate(size.most_common(), 1)}
    out["dataset"] = [new_id[d] for d in out["dataset"]]

    # sheet order, row by row
    fc = pd.factorize(np.array(out["file"], dtype=object), sort=True)[0]
    sc = pd.factorize(np.array(out["sheet"], dtype=object), sort=True)[0]
    perm = np.lexsort((np.array(out["col"]), np.array(out["row"]), sc, fc)).tolist()
    del fc, sc
    for k in out:
        v = out[k]
        out[k] = [v[i] for i in perm]
    return Records(out, meta)
