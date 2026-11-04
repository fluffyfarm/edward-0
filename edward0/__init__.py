"""
edward0 - spreadsheets to relational records.

    import edward0 as cleaner

    df = cleaner.read("book.xlsx")          # one row per value, dimensions as axis1..axis6
    df = cleaner.read("data/2026-10-09")    # every .xlsx / .xlsm / .csv in a folder
    tables = cleaner.tables("book.xlsx", sheets="Sales")   # only some sheets (a name or a list)
    df = cleaner.read("ids.csv", text_columns=["Customer ID"])   # keep a CSV column as text ('00123')
    tables = cleaner.tables("book.xlsx")    # {dataset: long table with named columns}
    raw = cleaner.raw("book.xlsx")          # index (set of labels), value, dataset
    cells = cleaner.cells("book.xlsx")      # step 1 only: each value cell with its labels

Layout is used only to collect the labels around each value cell (extract.py). Which labels
are context, which form a dimension, and which tables belong together is decided from the
label sets alone (clean.py).
"""
from __future__ import annotations

import os

from .clean import COLUMNS, _order, exact_values, records
from .extract import extract, extract_sheets
from .relational import tables as _tables

__all__ = ["read", "raw", "tables", "cells"]
__version__ = "0.1.0"

_EXT = (".xlsx", ".xlsm", ".csv")


def _paths(src) -> list[str]:
    items = [src] if isinstance(src, (str, os.PathLike)) else list(src)
    out = []
    for p in map(os.fspath, items):
        if os.path.isdir(p):
            out += sorted(os.path.join(p, f) for f in os.listdir(p)
                          if f.lower().endswith(_EXT) and not f.startswith("~$"))
        elif os.path.isfile(p):
            out.append(p)
        else:
            raise FileNotFoundError(p)
    return out


def _sheets(src, sheets=None, text_columns=None):
    """Every sheet of every file, one at a time (a list of value cells per sheet).
    sheets: a sheet name or a list of them (default: every sheet); a name found in none of the
    files is an error."""
    want = None if sheets is None else [sheets] if isinstance(sheets, str) else list(sheets)
    names = []
    for p in _paths(src):
        yield from extract_sheets(p, want, names, text_columns)
    missing = [s for s in want or [] if s not in names]
    if missing:
        raise ValueError(f"sheet not found: {', '.join(map(repr, missing))}; "
                         f"sheets: {', '.join(map(repr, dict.fromkeys(names)))}")


def _extract_all(src, sheets=None, text_columns=None) -> list[dict]:
    return [c for sheet in _sheets(src, sheets, text_columns) for c in sheet.dicts()]


def cells(src, sheets=None, text_columns=None):
    """Every value cell with the labels around it (no cleaning).

    src:    a file, a folder, or a list of them.
    sheets: a sheet name or a list of them (default: every sheet).
    text_columns: CSV columns kept as written (e.g. IDs with leading zeros: "00123"), by header
              or letter: text_columns=["Customer ID"] or ["A"]. Excel files keep their own types.
    Columns: file, sheet, table, address, row, col, value_type, value, formula,
    labels (each with where it came from: header, title or layout)."""
    import pandas as pd
    rows = _extract_all(src, sheets, text_columns)
    for r in rows:
        r["labels"] = " · ".join(f"{lb.text} ({via})" for lb, via in zip(r["labels"], r.pop("via")))
    df = pd.DataFrame(rows, columns=["file", "sheet", "table", "address", "row", "col",
                                     "value_type", "value", "formula", "labels"])
    if rows:
        df["value"] = exact_values([r["value"] for r in rows])
    return df


def read(src, raw_only: bool = True, sheets=None, text_columns=None):
    """Relational records: one row per value.

    src:       a file, a folder, or a list of them (.xlsx / .xlsm / .csv).
    raw_only:  leave out cells calculated from other cells (default True).
    sheets:    a sheet name or a list of them (default: every sheet). Datasets are then
               matched among those sheets only.
    text_columns: CSV columns kept as written (e.g. IDs with leading zeros: "00123"), by header
              or letter: text_columns=["Customer ID"] or ["A"]. Excel files keep their own types.

    Columns:
      file, sheet, table, address  where the value is (row, col: its row and column numbers;
                                   records come in sheet order, row by row)
      value                        the value (as Excel last calculated it)
      dataset                      tables with matching axes share a number (1 = largest)
      context                      labels on every cell of the cell's table (title, unit ...)
      axis1 .. axis6               the cell's label on each dimension of its dataset
      other                        labels of smaller dimensions
      dropped                      labels left out: errors (#REF! ...) and notes
      index                        all kept labels of the cell (text, ' · ' between)
      is_formula                   the cell holds a formula
      dup                          records in the dataset with the same key (1 = unique)
    """
    import pandas as pd
    c = records(_sheets(src, sheets, text_columns), raw_only=raw_only).cols
    data = {k: c[k] for k in COLUMNS[:COLUMNS.index("is_formula")]}
    if c["value"]:
        data["value"] = exact_values(c["value"])
    data["index"] = [" · ".join(sorted(ix, key=_order)) for ix in c["_index"]]
    data["is_formula"], data["dup"] = c["is_formula"], c["dup"]
    return pd.DataFrame(data)


def raw(src, raw_only: bool = True, sheets=None, text_columns=None):
    """The simplest form: every value with the set of labels that identify it.

    Columns:
      index    frozenset of the cell's labels; query with set operations, e.g.
               raw[raw["index"].map({"May", "C2 SC"}.issubset)]
      value    the value
      dataset  tables with matching axes share a number (1 = largest)
    sheets: a sheet name or a list of them (default: every sheet).
    text_columns: CSV columns kept as written (e.g. IDs with leading zeros: "00123"), by header
              or letter: text_columns=["Customer ID"] or ["A"]. Excel files keep their own types.
    """
    import pandas as pd
    c = records(_sheets(src, sheets, text_columns), raw_only=raw_only).cols
    value = exact_values(c["value"]) if c["value"] else c["value"]
    return pd.DataFrame({"index": c["_index"], "value": value, "dataset": c["dataset"]})


def tables(src, raw_only: bool = True, source: bool = False, sheets=None, text_columns=None) -> dict:
    """One relational table per dataset, long form: one row per value.

    Each dimension is a column, named from the sheet ('Year', 'Term', 'Mark'; a lower level of
    the same header gets '.2'; no name found: 'axis1' ...). Row-header columns come first, then
    column-header columns, then `context` (only if it differs between rows), `other` (only if
    used) and `value`. source=True adds file, sheet, address.
    sheets: a sheet name or a list of them (default: every sheet), e.g. sheets="Sales".
    text_columns: CSV columns kept as written (e.g. IDs with leading zeros: "00123"), by header
              or letter: text_columns=["Customer ID"] or ["A"]. Excel files keep their own types.

    Returns {dataset number: DataFrame}."""
    return _tables(records(_sheets(src, sheets, text_columns), raw_only=raw_only, keep_index=False), source=source)
