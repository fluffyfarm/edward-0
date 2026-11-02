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

from .clean import MAX_AXES


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


def _name_table(recs):
    """Names of the axes of one table: {axis: name}, {axis: (kind, position)}, context keys used."""
    ctx = {}
    for r in recs:
        ctx.update(r["_context"])
    origins = defaultdict(list)
    for r in recs:
        for j, labs in r["_axes"].items():
            origins[j] += [o for _, o in labs.values()]
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
    """recs from clean.records() -> {dataset: pandas DataFrame} in long form, one row per value."""
    import pandas as pd
    by_ds = defaultdict(lambda: defaultdict(list))
    for r in recs:
        by_ds[r["dataset"]][(r["file"], r["sheet"], r["table"])].append(r)

    out = {}
    for d in sorted(by_ds):
        tbls = by_ds[d]
        votes, shapes, used_by = defaultdict(Counter), {}, {}
        for t, rs in sorted(tbls.items(), key=lambda kv: -len(kv[1])):     # largest table first
            names, shape, used = _name_table(rs)
            for j, n in names.items():
                votes[j][n] += len(rs)
            for j, s in shape.items():
                shapes.setdefault(j, s)
            used_by[t] = used
        axes = [j for j in range(MAX_AXES) if any(r[f"axis{j + 1}"] is not None for rs in tbls.values() for r in rs)]
        axes.sort(key=lambda j: ({"row": 0, "col": 1}[shapes[j][0]], shapes[j][1]) if j in shapes else (2, j))

        # column names, unique: a repeated name gets .2, .3 ...
        cols, seen = {}, Counter()
        for j in axes:
            n = votes[j].most_common(1)[0][0] if votes[j] else f"axis{j + 1}"
            seen[n] += 1
            cols[j] = n if seen[n] == 1 else f"{n}.{seen[n]}"

        rows = []
        for t, rs in tbls.items():
            for r in rs:
                row = {cols[j]: r[f"axis{j + 1}"] for j in axes}
                ctx = [txt for k, (txt, _) in r["_context"].items() if k not in used_by[t]]
                row["context"] = " · ".join(sorted(ctx, key=str.casefold)) or None
                row["other"] = r["other"]
                row["value"] = r["value"]
                if source:
                    row.update(file=r["file"], sheet=r["sheet"], address=r["address"])
                row["_pos"] = (r["file"], r["sheet"], r["row"], r["col"])
                rows.append(row)

        # ragged levels: neighbouring axes of the same kind where the lower one fixes the upper one
        for a, b in zip(axes, axes[1:]):
            if a not in shapes or b not in shapes or shapes[a][0] != shapes[b][0]:
                continue
            A, B = cols[a], cols[b]
            up = defaultdict(set)
            for row in rows:
                if row[A] is not None and row[B] is not None:
                    up[row[B]].add(row[A])
            n_a = len({row[A] for row in rows if row[A] is not None})
            if not up or any(len(v) > 1 for v in up.values()) or n_a >= len(up):
                continue
            for row in rows:
                if row[A] is None and row[B] is not None:
                    row[A] = row[B]
                elif row[B] is None and row[A] is not None:
                    row[B] = row[A]

        rows.sort(key=lambda row: row.pop("_pos"))
        df = pd.DataFrame(rows)
        if df["context"].nunique(dropna=False) <= 1:      # the same on every row: not a column
            df = df.drop(columns="context")
        if df["other"].isna().all():
            df = df.drop(columns="other")
        out[d] = df.reset_index(drop=True)
    return out
