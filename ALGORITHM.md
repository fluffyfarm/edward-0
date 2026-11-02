# edward-0: spreadsheets to relational records from label sets

One pass, workbook → relational records. Every number in the workbook becomes one row:
`file | sheet | address | value | context | axis1 … axisK`. No training, no LLM,
no hand-written dimension names. Python + openpyxl (+ pandas for the output).

```python
import edward0 as cleaner
df = cleaner.read("book.xlsx")
```

**Core idea.** The work is split in two parts that use different evidence:

1. **Collect labels** (`extract.py`): layout rules find the tables on a sheet and the
   labels around each value cell — its column and row headers, the table title, text to its
   left. The result is *a value with an unordered set of labels*. Order and position are
   not passed on.
2. **Structure from label sets** (`clean.py`): which labels are context, which form a
   dimension, and which tables belong together is decided from statistics on those sets
   alone, never from positions:
   - two labels that **never sit on the same cell** and **appear with the same other labels**
     are values of the same dimension (e.g. `May`, `Jun`, `Jul`);
   - a label on **every** cell of a table is context (e.g. the table title);
   - tables whose dimensions hold **mostly the same labels** are one dataset (e.g. the same
     template on 30 dated sheets).

So layout tells us *which labels belong to a value*; it never tells us *what the labels
mean*. Header detection is only asked to be generous enough to collect the right labels:
an extra label on every cell becomes context, and a stray label on a few cells ends up in a
small dimension in `other`, without changing the main dimensions.

## Output

`read()` returns a pandas DataFrame, one row per value:

| column | meaning |
|---|---|
| `file`, `sheet`, `table`, `address` | where the value is |
| `row`, `col` | its row and column numbers (records come in sheet order, row by row) |
| `value` | as Excel last calculated it |
| `dataset` | tables with matching dimensions share a number (1 = largest) |
| `context` | labels on every cell of the table |
| `axis1` … `axis6` | the cell's label on each dimension (unnamed) |
| `other` | labels of smaller dimensions |
| `dropped` | labels left out: Excel errors and notes |
| `index` | all kept labels of the cell |
| `dup` | records in the dataset with the same key (1 = unique; > 1 = a dimension is missing) |

Every label of a cell is in exactly one of `context`, `axis1..6`, `other`, `dropped`, so each
row can be checked by eye against the original cell.

## Part 1 — collect labels (layout)

### 1. Find tables
- Connected blocks of filled cells (8-neighbourhood) → boxes; overlapping boxes merged.
- A box is split where a new block starts without a blank row: a row with numbers, then
  single-label rows (block title), then a header row.
- Per box: title rows (≤ 1 filled cell on top) and header rows (≥ 80 % labels, up to 4 levels
  per column). Then the index (row-header columns), key first:
  1. **Key**: the leftmost column whose values on the data rows are all filled and all
     different, and are labels or a row counter (1, 2, 3 …), e.g. `Index` or `Protein IDs`.
     The index is the key plus the label columns directly left of it (its parent levels).
     Columns of numbers left of the key stay values (`# OF Trades` left of `Index`).
  2. **No key**: leftmost columns ≥ 80 % labels; columns of numbers left of them are part of
     the index too (years 1991, 1992 left of Winter/Spring/Fall). 2D if it has an index and
     ≥ 50 % numeric body, else 1D.

A *label* is any non-empty text that is not a number written as text, or a date (kept as
Excel shows it, with its ISO date alongside).

### 2. Labels of each value cell
The cell's label set is the union of:
- **header**: its column-header path (blank upper levels filled from the left, merged cells
  resolved) + row-header path (a blank row header is filled down from the row above, on every
  level; a new label on a level clears the levels below it) + corner labels;
- **title**: the table's title rows and captions ≤ 3 rows above it;
- **layout**: text to its left in the same row (not in a keyed table: the key already names the
  row), and header rows above the table (≤ 80 rows
  up, never past another table's header).

## Part 2 — structure from label sets (no positions)

### 3. Clean labels
Whitespace collapsed, case folded, trailing `:` removed; dates compared by ISO date.
Dropped (kept in `dropped`): Excel errors (`#REF!`, `#N/A` …) and notes: text over 60
characters that came from a title or the layout, not from a row or column header (a long ID
in the index, e.g. a list of protein IDs, is kept).

### 4. Keep raw data only
Keep typed values and formulas that are arithmetic on constants (`=1312+350000/4342`).
Cells whose formula uses other cells are calculated, not data, and are left out
(`read(..., raw_only=False)` keeps them).

### 5. Find the dimensions of each table
```
n        = number of cells in the table
context  = labels on all n cells
for each other label k:
    cells(k)    = set of cells carrying k
    partners(k) = all other non-context labels on those cells
axes = []
for k in labels, most frequent first:
    best = argmax over axes A with cells(k) ∩ cells(A) = ∅          # never on the same cell
           of jaccard(partners(k), partners(A) − members(A))
    if best exists and score ≥ 0.5: add k to best (union its cells and partners)
    else: start a new axis {k}
sort axes by number of cells covered
```
Axes are unnamed; the first `K = 6` become `axis1..6`, smaller ones go to `other`.

### 6. Group tables into datasets
Tables are processed largest first. A table joins an existing dataset when it has the same
number of axes and a one-to-one pairing of its axes with the dataset's axes exists where every
pair has Jaccard ≥ 0.5 (greedy, highest pairs first). Otherwise it starts a new dataset.
Tables with no axes are grouped by their context. Candidate datasets are found through an
inverted index label → datasets.

### 7. Emit records
One row per raw cell: dataset, context, each axis label in the dataset's axis column,
leftovers in `other`, and `dup`. Records are sorted by file, sheet, row, column, so they
read in the same order as the sheet (sorting by `address` text would put `B10` before `B5`).

## Part 3 — relational tables (`tables()`)

One long table per dataset, one row per value, one column per axis. The structure is already
fixed by part 2; the sheet is used here only to **name and order** the columns:

- **Row-header axis** (labels read down one column): named by the context label above it in
  that column (`Year` above 1991, 1992).
- **Column-header axis** (labels read across one row): named by the nearest context label above
  it whose column lies within the axis' columns (`Mark` over Assignments … Grade). A lower level
  under the same label gets `.2` (`Mark.2` for Ass1 … Final).
- No such label: `axis1`, `axis2` …. Labels used as names leave the context.
- Order: row-header axes left to right, then column-header axes top to bottom.
- **Ragged levels**: for neighbouring axes of the same kind where each lower label always has
  the same upper label (Ass1 → Assignments), a record with only one of the two gets the same
  label on both (`Grade`, `Grade`), so grouping by the upper level still adds up.
- `context` is kept as a column only when it differs between rows (e.g. `Ref date: 28 Apr` vs
  `Ref date: 7 May` on two snapshot tables of one dataset).

`raw()` is the simplest form: `index` (the set of the cell's labels), `value`, `dataset`; any
value can be found by any subset of its labels.

## Parameters

| name | value | used in | meaning |
|---|---|---|---|
| label share | 80 % | part 1 | a row/column is a header when ≥ 80 % of its filled cells are labels |
| header levels | 4 | part 1 | column-header rows per column |
| look-up | 80 rows | part 1 | how far above a table a header row may be |
| `SAME_AXIS` | 0.5 | part 2 | Jaccard of partner labels for a label to join an axis |
| `SAME_DATASET` | 0.5 | part 2 | Jaccard per paired axis for a table to join a dataset |
| `NOTE_LEN` | 60 | part 2 | longer title/layout labels are notes, not identifiers (header labels are never notes) |
| `MAX_AXES` | 6 | part 2 | axis columns in the output |

All are hand-set, not tuned.

## Worked example (Embley et al., *Factoring Web Tables*, Fig. 1)

36 marks; each cell's label set is e.g. `{year, term, mark, 1991, Winter, Assignments, Ass1}`.
Output of step 5:

| | labels |
|---|---|
| context | `year`, `term`, `mark` |
| axis | `1991`, `1992` |
| axis | `Ass1`, `Ass2`, `Ass3`, `Midterm`, `Final`, `Grade` |
| axis | `Winter`, `Spring`, `Fall` |
| axis | `Assignments`, `Examinations` (empty on `Grade` cells) |

Embley's factoring gives the same values but keeps `Grade` one level up, as a sibling of
`Assignments`/`Examinations`. Flat axes lose that hierarchy: see limitations.
Data: `data/paper_examples/embley_factoring_web_tables/`.

## Position vs. prior work

| work | how structure is found |
|---|---|
| Embley, Krishnamoorthy, Nagy, Seth — *Factoring Web Tables* | header positions → header paths → algebraic factoring into Wang categories |
| Nagy et al. 2015 — *Clustering header categories* | Jaccard on category headers to cluster tables across sources |
| Chen & Cafarella, KDD 2014 (Senbazuru) | CRF over 32 layout/format features (+ sibling co-occurrence) → header hierarchy |
| Vitagliano et al., VLDB 2022 (Mondrian) | rendered layout similarity → templates across files |
| Li et al., VLDB 2023 (Auto-Tables) | learned model predicts grid reshaping steps |
| Dong et al., ACL 2025 (RelationalCoder) | LLM writes SQL from the encoded sheet |
| **edward-0** | layout only to collect each cell's labels; dimensions, context and datasets from **label-set statistics**; deterministic, no training |

Closest overlap: Embley et al. (same output model: mutually exclusive label groups) and
Nagy et al. 2015 (Jaccard on headers across tables). The difference is the evidence used
for structure: unordered label sets instead of header positions.

## Limitations
- **Label collection uses layout rules** (part 1): a label that the rules never attach to a
  cell cannot be recovered later. Part 2 tolerates extra labels, not missing ones.
- **Numbers as labels**: years or codes typed as numbers are labels only when they sit left of
  the row-header columns; in header rows or elsewhere they are read as values.
- **Flat axes**: hierarchies are flattened; a ragged level (`Grade`) can land on another axis.
- **Order-dependent and greedy**: label order (by frequency) and table order (by size) can
  change the result; a dataset is matched against its first table only.
- **Sparse tables**: labels of one dimension that rarely share partners may split into two
  axes; labels of different dimensions that never share a cell by chance may merge.
- **Untuned thresholds**, and **no measured accuracy yet**: evaluate on the Auto-Tables
  benchmark, DECO, Troy_200 and the paper examples in `data/`.
