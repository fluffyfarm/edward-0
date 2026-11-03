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


def table_axes_csr(indptr, kids, names: list[str]):
    """table_axes() on the table's cells x keys matrix in CSR form: cell i has the keys
    kids[indptr[i]:indptr[i + 1]] (numpy arrays, no key twice on a cell); names[k] is key k.

    Same rules and result as table_axes(), computed from the matrix: count = column sums,
    partners = keys on the same cell (A^T A), coverage = cells with a key of the axis.
    Returns (constant keys, axes), keys as ids; each axis lists its keys in the order they joined,
    axes largest coverage first."""
    import numpy as np
    n, K = len(indptr) - 1, len(names)
    count = np.bincount(kids, minlength=K)
    const = (count == n) if n > 1 else (count > 0)
    constants = set(np.flatnonzero(const).tolist())
    cell = np.repeat(np.arange(n, dtype=np.int64), np.diff(indptr))
    var = ~const[kids]
    vc, vk = cell[var], kids[var].astype(np.int64)

    # partners: pairs of keys on the same cell (a cell's keys are contiguous in vc / vk)
    codes = []
    width = int(np.bincount(vc).max()) if len(vc) else 0
    for d in range(1, width):
        same = vc[d:] == vc[:-d]
        a, b = vk[:-d][same], vk[d:][same]
        codes += [a * K + b, b * K + a]
    pc = np.unique(np.concatenate(codes)) if codes else np.empty(0, np.int64)
    start = np.searchsorted(pc // K, np.arange(K + 1)).tolist()
    pb = pc % K
    del codes, pc

    cnt = count.tolist()
    axes = []                                   # [members, partners of members - members, members in order]
    for k in sorted(np.flatnonzero((count > 0) & ~const).tolist(), key=lambda x: (-cnt[x], names[x])):
        p = set(pb[start[k]:start[k + 1]].tolist())
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
            axes.append([{k}, p, [k]])
        else:
            best[0].add(k)
            best[1] |= p                        # p has no member and not k (checked above)
            best[2].append(k)

    A = len(axes)
    if not A:
        return constants, []
    axis_of = np.full(K, -1, dtype=np.int64)
    for i, ax in enumerate(axes):
        axis_of[ax[2]] = i
    cover = np.bincount(np.unique(vc * A + axis_of[vk]) % A, minlength=A).tolist()
    order = sorted(range(A), key=lambda i: (-cover[i], min(names[k] for k in axes[i][2])))
    return constants, [axes[i][2] for i in order]


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


def _order(text: str):
    """Labels in a record: case-insensitive, then exact text (so 'TUFA' before 'Tufa' every run)."""
    return text.casefold(), text


def _show(keys, own):
    """A record shows its cell's own spelling of each label."""
    return " · ".join(sorted({x for k in keys for x in own[k]}, key=_order)) or None


class _Table:
    """One table after steps 1-3, kept compact: its axes, its table metadata and one entry per
    value cell in column lists."""
    __slots__ = ("key", "n", "constants", "main", "cols", "ax", "context", "origins")


def _sheet_tables(sheet, raw_only: bool, pool: dict, keep_index: bool) -> list[_Table]:
    """Steps 1-3 on one sheet (extract.SheetCells): clean labels, raw data only, axes per table.
    Labels are cleaned once per distinct label; a table is a cells x keys matrix.
    pool: shared strings, so a label's text is stored once however many records show it."""
    import numpy as np
    from openpyxl.utils import get_column_letter

    share = lambda x: pool.setdefault(x, x) if x is not None else None
    L = sheet.labels

    # 1. clean labels: once per distinct label; kid_of[label] = key id, or -1 for a dropped label
    kid_of, disp, names, kid_by_key = [], [], [], {}
    for text, date_value, via in zip(L.text, L.date_value, L.via):
        key, reason = clean_key(text, date_value, via == "header")
        disp.append(share(" ".join(text.split())))
        if reason is None:
            k = kid_by_key.get(key)
            if k is None:
                k = kid_by_key[key] = len(names)
                names.append(key)
            kid_of.append(k)
        else:
            kid_of.append(-1)
    kid_arr = np.array(kid_of, dtype=np.int64)
    K, origin = len(names), L.origin
    letters = {}

    def gtext(lids):                                             # _show() for the labels of one part
        if len(lids) == 1:
            return disp[lids[0]]
        return share(" · ".join(sorted({disp[x] for x in lids}, key=_order)))

    out = []
    for tc in sheet.tables:
        n0 = len(tc)
        if not n0 or not K:
            continue
        # 2. raw data only; cells without any kept label cannot be placed and are left out
        ok = np.frombuffer(tc.vtype, dtype=np.int8) < 4                   # not error / none
        if raw_only:
            for i, f in tc.formula.items():
                if not is_raw(f):
                    ok[i] = False
        off = np.frombuffer(tc.off, dtype=np.int64)
        lab = np.frombuffer(tc.lab, dtype=np.int32)
        cell = np.repeat(np.arange(n0, dtype=np.int64), np.diff(off))
        ek = kid_arr[lab]
        has = np.bincount(cell[ek >= 0], minlength=n0) > 0
        use = ok & has
        keep = np.flatnonzero(use)
        if not len(keep):
            continue

        # 3. axes: the cells x keys matrix (CSR, no key twice on a cell)
        sel = (ek >= 0) & use[cell]
        pairs = np.unique(cell[sel] * K + ek[sel])
        pos = np.full(n0, -1, dtype=np.int64)
        pos[keep] = np.arange(len(keep))
        indptr = np.searchsorted(pos[pairs // K], np.arange(len(keep) + 1))
        const_k, axes_k = table_axes_csr(indptr, pairs % K, names)
        del pairs, sel, cell, ek

        constants = {names[k] for k in const_k}
        main_k, rest_k = axes_k[:MAX_AXES], axes_k[MAX_AXES:]
        main = [{names[k] for k in a} for a in main_k]          # same insertion order as table_axes()
        G = len(main_k) + 1                                      # part of a key: 0 context, 1.. axis, G other
        grp = [-1] * K
        for k in const_k:
            grp[k] = 0
        for i, a in enumerate(main_k):
            for k in a:
                grp[k] = 1 + i
        for a in rest_k:
            for k in a:
                grp[k] = G

        rows = np.frombuffer(tc.row, dtype=np.int32)
        cols_ = np.frombuffer(tc.col, dtype=np.int32)
        order = keep[np.lexsort((cols_[keep], rows[keep]))].tolist()      # sheet order: the metadata below
        cols = {k: [] for k in ("address", "row", "col", "value", "is_formula", "context", "other", "dropped")}
        if keep_index:
            cols["_index"] = []
        c_addr, c_row, c_col, c_val, c_isf = (cols[k] for k in ("address", "row", "col", "value", "is_formula"))
        c_ctx, c_oth, c_drop = cols["context"], cols["other"], cols["dropped"]
        c_idx = cols.get("_index")
        ax = [[] for _ in main_k]
        origins, first = [{} for _ in main_k], {}
        lab_l, off_l, row_l, col_l, formula = tc.lab, tc.off, tc.row, tc.col, tc.formula
        nax = len(main_k)
        last = None
        for q, i in enumerate(order):
            parts, dropped, kept = {}, [], []
            for j in range(off_l[i], off_l[i + 1]):
                lid = lab_l[j]
                k = kid_of[lid]
                if k < 0:
                    dropped.append(disp[lid])
                else:
                    kept.append(lid)
                    g = grp[k]
                    if g in parts:
                        parts[g].append(lid)
                    else:
                        parts[g] = [lid]
            r, c = row_l[i], col_l[i]
            letter = letters.get(c)
            if letter is None:
                letter = letters[c] = get_column_letter(c)
            c_addr.append(f"{letter}{r}")
            c_row.append(r)
            c_col.append(c)
            c_val.append(tc.value[i])
            c_isf.append(i in formula)
            p = parts.get(0)
            c_ctx.append(gtext(p) if p else None)
            for a in range(nax):
                p = parts.get(1 + a)
                if p is None:
                    ax[a].append(None)
                    continue
                ax[a].append(gtext(p))
                if a not in first:
                    first[a] = (q, a)
                ks = {kid_of[x] for x in p}
                if len(ks) == 1:
                    origins[a].setdefault(origin[p[0]])
                else:                                            # several keys on one axis: in the order
                    fo, keys = {}, set()                         # table_axes()' sets give them
                    for x in kept:
                        keys.add(names[kid_of[x]])
                        fo.setdefault(names[kid_of[x]], origin[x])
                    for key in frozenset(keys) & main[a]:
                        origins[a].setdefault(fo[key])
            p = parts.get(G)
            c_oth.append(gtext(p) if p else None)
            c_drop.append(share(" · ".join(sorted(dropped, key=str.casefold))) if dropped else None)
            if keep_index:
                c_idx.append(frozenset(disp[x] for x in kept))
            last = kept

        # the table's context labels, with the spelling of the last cell in sheet order
        context, own, fo = {}, {}, {}
        for x in last:
            key = names[kid_of[x]]
            own.setdefault(key, set()).add(disp[x])
            fo.setdefault(key, origin[x])
        for key in constants & set(own):
            context[key] = (_show({key}, own), fo[key])

        tb = _Table()
        tb.key, tb.n, tb.constants, tb.main = (sheet.file, sheet.sheet, tc.ref), len(order), constants, main
        tb.cols, tb.ax, tb.context = cols, ax, context
        tb.origins = {i: origins[i] for i in sorted(first, key=first.get)}   # axes in the order records show them
        out.append(tb)
    return out


def records(cells, raw_only: bool = True, keep_index: bool = True) -> Records:
    """cells from extract -> one record per value, as Records (columns + per-table metadata).

    cells:      an iterable of sheets (extract.extract_sheets: SheetCells, or lists of cell dicts), or
                one flat list of cell dicts (extract.extract). A table must not be split across sheets. Each
                sheet is reduced to its tables' axes and compact columns before the next one is read;
                datasets are matched once all sheets are in.
    keep_index: keep `_index`, the set of each record's labels (for raw() and read()'s `index`).

    Every label of a cell is in exactly one of context / axis1..6 / other / dropped.
    Records come in sheet order, row by row."""
    import numpy as np
    import pandas as pd

    from .extract import SheetCells

    if isinstance(cells, list) and (not cells or isinstance(cells[0], dict)):
        cells = SheetCells.from_dicts(cells)
    pool, tables = {}, []
    for sheet in cells:
        for sc in SheetCells.from_dicts(sheet) if isinstance(sheet, list) else [sheet]:
            tables += _sheet_tables(sc, raw_only, pool, keep_index)
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
