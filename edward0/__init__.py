"""
edward0 - spreadsheets to relational records.

    import edward0 as cleaner

    df = cleaner.read("book.xlsx")          # one row per value, dimensions as axis1..axis6
    df = cleaner.read("data/2026-10-09")    # every .xlsx / .xlsm / .csv in a folder
    tables = cleaner.tables("book.xlsx")    # {dataset: long table with named columns}
    raw = cleaner.raw("book.xlsx")          # index (set of labels), value, dataset
    cells = cleaner.cells("book.xlsx")      # step 1 only: each value cell with its labels

Layout is used only to collect the labels around each value cell (extract.py). Which labels
are context, which form a dimension, and which tables belong together is decided from the
label sets alone (clean.py).
"""
from __future__ import annotations

import os

from .clean import records
from .extract import extract
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


def _extract_all(src) -> list[dict]:
    out = []
    for p in _paths(src):
        out += extract(p)
    return out


def cells(src):
    """Every value cell with the labels around it (no cleaning).

    src: a file, a folder, or a list of them.
    Columns: file, sheet, table, address, row, col, value_type, value, formula,
    labels (each with where it came from: header, title or layout)."""
    import pandas as pd
    rows = _extract_all(src)
    for r in rows:
        r["labels"] = " · ".join(f"{lb.text} ({via})" for lb, via in zip(r["labels"], r.pop("via")))
    return pd.DataFrame(rows, columns=["file", "sheet", "table", "address", "row", "col",
                                       "value_type", "value", "formula", "labels"])


def read(src, raw_only: bool = True):
    """Relational records: one row per value.

    src:       a file, a folder, or a list of them (.xlsx / .xlsm / .csv).
    raw_only:  leave out cells calculated from other cells (default True).

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
    recs = records(_extract_all(src), raw_only=raw_only)
    cols = ["file", "sheet", "table", "address", "row", "col", "value", "dataset", "context",
            "axis1", "axis2", "axis3", "axis4", "axis5", "axis6",
            "other", "dropped", "labels", "is_formula", "dup"]
    return pd.DataFrame(recs, columns=cols).rename(columns={"labels": "index"})


def raw(src, raw_only: bool = True):
    """The simplest form: every value with the set of labels that identify it.

    Columns:
      index    frozenset of the cell's labels; query with set operations, e.g.
               raw[raw["index"].map({"May", "C2 SC"}.issubset)]
      value    the value
      dataset  tables with matching axes share a number (1 = largest)"""
    import pandas as pd
    recs = records(_extract_all(src), raw_only=raw_only)
    return pd.DataFrame({"index": [r["_index"] for r in recs],
                         "value": [r["value"] for r in recs],
                         "dataset": [r["dataset"] for r in recs]})


def tables(src, raw_only: bool = True, source: bool = False) -> dict:
    """One relational table per dataset, long form: one row per value.

    Each dimension is a column, named from the sheet ('Year', 'Term', 'Mark'; a lower level of
    the same header gets '.2'; no name found: 'axis1' ...). Row-header columns come first, then
    column-header columns, then `context` (only if it differs between rows), `other` (only if
    used) and `value`. source=True adds file, sheet, address.

    Returns {dataset number: DataFrame}."""
    return _tables(records(_extract_all(src), raw_only=raw_only), source=source)
