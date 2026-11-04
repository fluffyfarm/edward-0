# Edward recommended additions

Preserve source values and identities first, then expand table interpretation and make
conversion quality measurable. Reading one sheet at a time and storing cells and labels
as sparse arrays are already implemented.

Status below reflects commit `fefee1e` on 2026-11-04. Priority numbers retain the order
of the original review; the reserved column names fix has since been implemented.

| Priority | Area | Status | Recommended addition | Example |
|---|---|---|---|---|
| 1 — Urgent | Output column names | Reserved names fixed in `fefee1e`; regression coverage pending | Add regression tests for reserved names, repeated names, and names that already contain suffixes. Ensure every generated column name is unique. | A dimension headed `value` becomes `value.2`, retaining `North` and `South` alongside the numeric `value` column. |
| 2 — Urgent | CSV data types and precision | Open | Parse integer fields without converting through `float`. Allow explicit text or ID column types so leading zeros can be preserved. | Preserve the ID `00123` when configured as text, and read the integer `9007199254740993` exactly. |
| 3 — High | Source identity | Open | Keep a unique source identifier, such as a normalized path, alongside the display filename. Use it when tracking records and detecting duplicate keys. | Distinguish `north/book.csv` from `south/book.csv` even when their sheet names and cell addresses match. |
| 4 — High | Table structure interpretation | Open | Support numeric headers, composite keys, and hierarchical labels. Use label origins and positions alongside label statistics to validate the inferred relationships. | Treat `2025` and `2026` as column headers when the layout supports that role; associate a score with its `Test ID` and `Rank`. |
| 5 — High | Automated tests and benchmarks | Open | Commit representative fixtures and expected results, run them in CI, and benchmark time and peak memory across different table shapes. | Detect a lost year header or changed integer even when a performance optimization makes conversion faster. |
| 6 — Next | Conversion diagnostics | Open | Report retained and excluded candidate values, exclusion reasons, duplicate keys, and dimensions placed in `other`. Distinguish missing cached formula values from intentional formula filtering. | Explain that a formula cell was excluded because its cached result is unavailable, rather than silently omitting it. |

## Acceptance checks

1. **Output column names:** preserve every dimension and value with `source=False` and
   `source=True`. Cover `value`, `context`, `other`, `file`, `sheet`, `address`, repeated
   names, and literal suffixed names such as `value.2`.
2. **CSV data types:** preserve exact large integers, configured text IDs, valid zeros,
   negative numbers, percentages, and blanks. Test values on both sides of the exact
   integer range of a floating-point number.
3. **Source identity:** convert two files with the same basename in different folders.
   Trace each output row to its source without treating distinct sources as one file.
4. **Table structure:** test ordinary record tables, cross tables, numeric headers,
   merged and multilevel headers, repeated keys, composite keys, and multiple tables
   on one sheet. Expected results must verify relationships as well as values.
5. **Automation:** exercise `cells()`, `read()`, `raw()`, and `tables()`, including sheet
   selection, empty input, missing sheets, formulas, and Excel errors. Compare the sparse
   axis implementation with its reference implementation, and use separate expected
   outputs to measure interpretation accuracy. Include wide tables and cells with many
   labels in memory benchmarks.
6. **Diagnostics:** reconcile candidate value counts with retained and excluded counts.
   Give each excluded cell an address and reason, and keep ordinary filtering separate
   from unexpected conversion failures.

## Relevant implementation

- [Table detection and label extraction](edward0/extract.py)
- [Dimension inference and duplicate detection](edward0/clean.py)
- [Relational output and column naming](edward0/relational.py)
- [Public API](edward0/__init__.py)
- [Algorithm and current limitations](ALGORITHM.md)
