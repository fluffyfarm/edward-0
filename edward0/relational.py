"""
relational.py - records -> one long relational table per dataset, with named columns.

    Year  Term    Mark         Mark.2   value
    1991  Winter  Assignments  Ass1     85
    1991  Winter  Grade        Grade    75

Each axis of a dataset becomes a column. The structure (which axes, which values) is already
decided in clean.py from label sets; here the sheet is used only to NAME and ORDER the columns:

  * an axis read down one column (row headers) is named by the label above it in that column
    ('Year' above 1991, 1992);
  * an axis read across one row (column headers) is named by the label above it whose column
    lies within the axis' columns ('Mark' above Assignments ... Grade); a second axis under the
    same name gets '.2' ('Mark.2' for Ass1 ... Final);
  * no such label: 'axis1', 'axis2' ...
  * order: row-header axes left to right, then column-header axes top to bottom.

A name can only come from a context label (a label on every cell of the table); labels used as
names leave the context. A ragged level is filled from its neighbour: when every label of a
lower axis has one fixed label on the axis above (Ass1 -> Assignments), a record with only one
of the two gets the same label on both ('Grade', 'Grade').
"""
from __future__ import annotations

from collections import Counter, defaultdict

from .clean import MAX_AXES, exact_values

RESERVED = ["context", "other", "value"]           # columns of every table
SOURCE_COLUMNS = ["file", "sheet", "address"]      # columns added by source=True


def _shape(origins):
    """('row', column) for labels read down one column, ('col', row) for labels read across one row."""
    pts = {o for o in origins if o}
    if not pts:
        return None
    rows = Counter(r for r, _ in pts)
    cols = Counter(c for _, c in pts)
    if len(cols) <= len(rows):
        return "row", cols.most_common(1)[0][0]
    return "col", rows.most_common(1)[0][0]


def _name_table(ctx, origins):
    """Names of the axes of one table: {axis: name}, {axis: (kind, position)}, context keys used.
    ctx: {label key: (text, source cell)} of the table's context; origins: {axis: source cells}."""
    shape = {j: s for j, o in origins.items() if (s := _shape(o))}

    names, used = {}, set()                          # used: context keys taken as row-axis names
    for j in sorted((j for j in shape if shape[j][0] == "row"), key=lambda j: shape[j][1]):
        c = shape[j][1]
        top = min(o[0] for o in origins[j] if o)
        cand = [(o[0], k, t) for k, (t, o) in ctx.items() if o and o[1] == c and o[0] < top and k not in used]
        if cand:
            _, k, t = max(cand)
            names[j] = t
            used.add(k)
    col_used = set()
    for j in sorted((j for j in shape if shape[j][0] == "col"), key=lambda j: shape[j][1]):
        h = shape[j][1]
        cs = [o[1] for o in origins[j] if o]
        lo, hi = min(cs), max(cs)
        cand = [(o[0], o[1], k, t) for k, (t, o) in ctx.items()
                if o and o[0] < h and lo <= o[1] <= hi and k not in used]
        if cand:                                     # nearest row above, then rightmost; may repeat
            *_, k, t = max(cand)                     # for a lower level (Mark, then Mark.2)
            names[j] = t
            col_used.add(k)
    return names, shape, used | col_used


def tables(recs, source: bool = False) -> dict:
    """recs from clean.records() -> {dataset: pandas DataFrame} in long form, one row per value.
    Columns are named from each table's metadata (recs.meta), never from per-record data."""
    import numpy as np
    import pandas as pd
    C, meta = recs.cols, recs.meta
    if not len(recs):
        return {}
    ds_arr, tid_arr = np.asarray(C["dataset"]), np.asarray(C["_tid"])

    out = {}
    for d in sorted(set(C["dataset"])):
        idx = np.flatnonzero(ds_arr == d)                                  # sheet order, row by row
        tids = tid_arr[idx]
        uniq, first, counts = np.unique(tids, return_index=True, return_counts=True)
        tbls = [(int(uniq[q]), int(counts[q])) for q in np.argsort(first, kind="stable")]   # first seen first
        votes, shapes, used_by = defaultdict(Counter), {}, {}
        for t, n in sorted(tbls, key=lambda x: -x[1]):                     # largest table first
            names, shape, used = _name_table(meta[t]["context"], meta[t]["origins"])
            for j, nm in names.items():
                votes[j][nm] += n
            for j, s in shape.items():
                shapes.setdefault(j, s)
            used_by[t] = used
        axes = [j for j in range(MAX_AXES) if any(j in meta[t]["axes"] for t, _ in tbls)]
        axes.sort(key=lambda j: ({"row": 0, "col": 1}[shapes[j][0]], shapes[j][1]) if j in shapes else (2, j))

        # column names, unique: a name already taken gets .2, .3 ... (the first free one). Taken are
        # the output's own columns ('value' ...) and the names given before: an axis headed 'value'
        # becomes 'value.2', and a header that literally reads 'value.2' then becomes 'value.2.2'
        cols, used = {}, set(RESERVED + (SOURCE_COLUMNS if source else []))
        for j in axes:
            base = votes[j].most_common(1)[0][0] if votes[j] else f"axis{j + 1}"
            name, k = base, 1
            while name in used:
                k += 1
                name = f"{base}.{k}"
            used.add(name)
            cols[j] = name

        il = idx.tolist()
        pick = lambda k: [C[k][i] for i in il]
        data = {}
        for j in axes:
            data[cols[j]] = pick(f"axis{j + 1}")
        ctx_text = {}
        for t, _ in tbls:
            ctx = [txt for k, (txt, _) in meta[t]["context"].items() if k not in used_by[t]]
            ctx_text[t] = " · ".join(sorted(ctx, key=str.casefold)) or None
        data["context"] = [ctx_text[t] for t in tids.tolist()]
        data["other"] = pick("other")
        data["value"] = exact_values(pick("value"))
        if source:
            data["file"], data["sheet"], data["address"] = pick("file"), pick("sheet"), pick("address")

        # ragged levels: neighbouring axes of the same kind where the lower one fixes the upper one
        for a, b in zip(axes, axes[1:]):
            if a not in shapes or b not in shapes or shapes[a][0] != shapes[b][0]:
                continue
            A, B = data[cols[a]], data[cols[b]]
            up = defaultdict(set)
            for x, y in zip(A, B):
                if x is not None and y is not None:
                    up[y].add(x)
            n_a = len({x for x in A if x is not None})
            if not up or any(len(v) > 1 for v in up.values()) or n_a >= len(up):
                continue
            for q in range(len(A)):
                if A[q] is None and B[q] is not None:
                    A[q] = B[q]
                elif B[q] is None and A[q] is not None:
                    B[q] = A[q]

        df = pd.DataFrame(data)
        if df["context"].nunique(dropna=False) <= 1:      # the same on every row: not a column
            df = df.drop(columns="context")
        if df["other"].isna().all():
            df = df.drop(columns="other")
        out[d] = df.reset_index(drop=True)
    return out
