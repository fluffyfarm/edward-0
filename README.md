# edward-0

Spreadsheets → relational records. Every number in a workbook becomes one row with the
labels that identify it, grouped into unnamed dimensions.

```bash
pip install -e .
```

```python
import edward0 as cleaner

df = cleaner.read("book.xlsx")        # a file, a folder, or a list of them
df[["sheet", "address", "value", "context", "axis1", "axis2"]]

tables = cleaner.tables("book.xlsx")  # {dataset: long relational table, named columns}
tables[1]
#    Year  Term    Mark         Mark.2  value
#    1991  Winter  Assignments  Ass1       85
#    1991  Winter  Grade        Grade      75

raw = cleaner.raw("book.xlsx")        # index (set of labels), value, dataset
raw[raw["index"].map({"May", "C2 SC"}.issubset)]     # any value, by any of its labels

cells = cleaner.cells("book.xlsx")    # first step only: value cells with their labels,
                                      # each tagged with where it came from (header, title, layout)
```

| column | meaning |
|---|---|
| `file`, `sheet`, `table`, `address` | where the value is |
| `row`, `col` | its row and column numbers; records come in sheet order, row by row |
| `value` | the value, as Excel last calculated it |
| `dataset` | tables with matching dimensions share a number (1 = largest) |
| `context` | labels on every cell of the table (title, unit …) |
| `axis1` … `axis6` | the cell's label on each dimension |
| `other`, `dropped` | labels of smaller dimensions; errors and notes left out |
| `index` | all kept labels of the cell, to check the row by eye |
| `dup` | records with the same key in the dataset (1 = unique; > 1 = a dimension is missing) |

The column is called `index`; read it with `df["index"]` (`df.index` is the pandas row index).

`tables()` columns: one per dimension, named from the sheet (a lower level of the same header
gets `.2`; no name found: `axis1` …), then `context` (only when it differs between rows),
`other` (only when used) and `value`; `source=True` adds `file`, `sheet`, `address`.

By default cells calculated from other cells are left out (`read(..., raw_only=False)`
keeps them). See `ALGORITHM.md` for how it works.
