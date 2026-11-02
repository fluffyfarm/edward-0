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
def table_axes(cells: list[frozenset]):
    """cells: label-key sets of the table's cells.
    Returns (context labels, axes) where axes is a list of label sets, largest coverage first."""
    n = len(cells)
    count = Counter(k for s in cells for k in s)
    constants = {k for k, v in count.items() if v == n} if n > 1 else set(count)
    bits, partners = defaultdict(int), defaultdict(set)
    for i, s in enumerate(cells):
        var = s - constants
        for k in var:
            bits[k] |= 1 << i
            partners[k] |= var
    for k in partners:
        partners[k].discard(k)

    axes = []                      # [members, cell bits, partner labels]
    for k in sorted(bits, key=lambda x: (-count[x], x)):
        best, best_score = None, SAME_AXIS
        for ax in axes:
            if bits[k] & ax[1]:
                continue                                    # sits on the same cell: not the same axis
            score = jaccard(partners[k], ax[2] - ax[0])
            if score >= best_score:
                best, best_score = ax, score
        if best is None:
            axes.append([{k}, bits[k], set(partners[k])])
        else:
            best[0].add(k)
            best[1] |= bits[k]
            best[2] |= partners[k]
    axes.sort(key=lambda ax: (-bin(ax[1]).count("1"), min(ax[0])))
    return constants, [ax[0] for ax in axes]


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
def records(cells: list[dict], raw_only: bool = True) -> list[dict]:
    """cells from extract.extract() -> one record per value.

    Record keys: file, sheet, table, address, row, col, value, dataset, context, axis1..axis6, other,
    dropped, labels, is_formula, dup. Every label of a cell is in exactly one of context /
    axis1..6 / other / dropped.
    Internal keys (start with '_'): _index (set of the cell's labels), _context and _axes
    ({label key: (text, source cell)}, used by relational.py to name the columns)."""
    # 1. clean labels; display each key with its most used spelling
    spell = defaultdict(Counter)
    for c in cells:
        for lb, via in zip(c["labels"], c["via"]):
            key, reason = clean_key(lb.text, lb.date_value, via == "header")
            if reason is None:
                spell[key][" ".join(lb.text.split())] += 1
    display = {k: s.most_common(1)[0][0] for k, s in spell.items()}

    # 2. raw data only; cells without any kept label cannot be placed and are left out
    by_table = defaultdict(list)
    for c in cells:
        if c["value_type"] in ("error", "none") or (raw_only and not is_raw(c["formula"])):
            continue
        keys, own, dropped, origin = set(), defaultdict(set), [], {}
        for lb, via in zip(c["labels"], c["via"]):
            text = " ".join(lb.text.split())
            key, reason = clean_key(lb.text, lb.date_value, via == "header")
            if reason is None:
                keys.add(key)
                own[key].add(text)
                origin.setdefault(key, lb.origin)               # where the label was read (first source)
            else:
                dropped.append(text)
        if keys:
            by_table[(c["file"], c["sheet"], c["table"])].append((c, frozenset(keys), own, dropped, origin))

    # 3. axes per table, 4. datasets; largest tables first so a dataset's first table is its fullest
    ds, info = Datasets(), {}
    for t in sorted(by_table, key=lambda t: (-len(by_table[t]), t)):
        constants, axes = table_axes([x[1] for x in by_table[t]])
        main, rest = axes[:MAX_AXES], axes[MAX_AXES:]
        d = ds.match(main, constants)
        m = Datasets.mapping(main, ds.rep[d]) if main else {}
        info[t] = (d, constants, main, rest, m)

    def show(keys, own):
        """A record shows its cell's own spelling of each label."""
        return " · ".join(sorted({x for k in keys for x in (own.get(k) or {display.get(k, k)})}, key=str.casefold)) or None

    out = []
    for t, (d, constants, main, rest, m) in info.items():
        rest_keys = set().union(*rest) if rest else set()
        for c, keys, own, dropped, origin in by_table[t]:
            rec = {"file": c["file"], "sheet": c["sheet"], "table": c["table"], "address": c["address"],
                   "row": c["row"], "col": c["col"], "value": c["value"], "dataset": d, "context": show(constants & keys, own)}
            for j in range(MAX_AXES):
                rec[f"axis{j + 1}"] = None
            for i, a in enumerate(main):
                hit = keys & a
                if hit:
                    rec[f"axis{m[i] + 1}"] = show(hit, own)
            rec["other"] = show(keys & rest_keys, own)
            rec["dropped"] = " · ".join(sorted(dropped, key=str.casefold)) or None
            rec["labels"] = " · ".join(sorted({x for v in own.values() for x in v}, key=str.casefold))
            rec["_index"] = frozenset(x for v in own.values() for x in v)
            rec["is_formula"] = c["formula"] is not None
            # for tables(): each label's text and source cell, by part of the record
            lab = lambda ks: {k: (show({k}, own), origin.get(k)) for k in ks}
            rec["_context"] = lab(constants & keys)
            rec["_axes"] = {m[i]: lab(keys & a) for i, a in enumerate(main) if keys & a}
            out.append(rec)

    # records of a dataset with the same full key: > 1 means a dimension is missing
    axis_cols = ["context"] + [f"axis{j + 1}" for j in range(MAX_AXES)] + ["other"]
    full = lambda r: (r["dataset"], r["file"], r["sheet"], *(r[k] for k in axis_cols))
    n = Counter(full(r) for r in out)
    for r in out:
        r["dup"] = n[full(r)]

    # number datasets 1, 2, ... by size
    size = Counter(r["dataset"] for r in out)
    new_id = {d: i for i, (d, _) in enumerate(size.most_common(), 1)}
    for r in out:
        r["dataset"] = new_id[r["dataset"]]
    out.sort(key=lambda r: (r["file"], r["sheet"], r["row"], r["col"]))      # sheet order, row by row
    return out
