"""visible_only=True: cells in hidden rows and columns are treated as empty."""
import shutil
import zipfile

from openpyxl import Workbook

import edward0

CROSS = [["Region", "Jan", "Feb", "Mar"],
         ["North", 1, 2, 3],
         ["South", 4, 5, 6],
         ["East", 7, 8, 9]]
REGIONS = {"North", "South", "East"}


def save(tmp_path, rows, setup=None, name="book.xlsx", merges=()):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    for m in merges:
        ws.merge_cells(m)
    if setup:
        setup(ws)
    path = tmp_path / name
    wb.save(path)
    return str(path)


def patch_sheet(path, old, new):
    """Edit the sheet XML directly (for what openpyxl does not write, e.g. a row height of 0)."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                assert old.encode() in data
                data = data.replace(old.encode(), new.encode())
            zout.writestr(item, data)
    shutil.move(tmp, path)


def values(path, **kw):
    """{(region, month): value} from read()."""
    out = {}
    for r in edward0.read(path, **kw).itertuples():
        a, b = r.axis1, r.axis2
        out[(a, b) if a in REGIONS else (b, a)] = r.value
    return out


def labels(path, **kw):
    raw = edward0.raw(path, **kw)
    return set().union(*raw["index"]) if len(raw) else set()


def hide_east_and_feb(ws):
    ws.row_dimensions[4].hidden = True
    ws.column_dimensions["C"].hidden = True


def test_hidden_rows_and_columns_are_left_out(tmp_path):
    path = save(tmp_path, CROSS, hide_east_and_feb)
    assert len(values(path)) == 9                                       # default: everything
    assert values(path, visible_only=True) == {("North", "Jan"): 1, ("North", "Mar"): 3,
                                               ("South", "Jan"): 4, ("South", "Mar"): 6}
    assert not {"East", "Feb"} & labels(path, visible_only=True)


def test_every_function_takes_visible_only(tmp_path):
    path = save(tmp_path, CROSS, hide_east_and_feb)
    assert len(edward0.cells(path, visible_only=True)) == 4
    assert len(edward0.raw(path, visible_only=True)) == 4
    assert len(edward0.tables(path, visible_only=True)[1]) == 4


def test_default_is_unchanged(tmp_path):
    path = save(tmp_path, CROSS, hide_east_and_feb)
    assert edward0.read(path).equals(edward0.read(path, visible_only=False))
    plain = save(tmp_path, CROSS, name="plain.xlsx")
    a, b = edward0.read(plain), edward0.read(path)
    assert a.drop(columns="file").equals(b.drop(columns="file"))      # hiding alone changes nothing


def test_grouped_columns_and_zero_height_rows(tmp_path):
    path = save(tmp_path, CROSS, lambda ws: ws.column_dimensions.group("C", "D", hidden=True))   # <col min=3 max=4>
    patch_sheet(path, '<row r="3">', '<row r="3" ht="0" customHeight="1">')                       # South: height 0
    df = edward0.read(path, visible_only=True)
    assert sorted(df["address"]) == ["B2", "B4"]
    assert sorted(df["value"]) == [1, 7]
    assert not {"South", "Feb", "Mar"} & labels(path, visible_only=True)


def test_hidden_header_row_gives_no_labels(tmp_path):
    path = save(tmp_path, CROSS, lambda ws: setattr(ws.row_dimensions[1], "hidden", True))
    assert {"Jan", "Feb", "Mar", "Region"} <= labels(path)
    assert not {"Jan", "Feb", "Mar", "Region"} & labels(path, visible_only=True)


def test_hidden_column_does_not_split_a_table(tmp_path):
    """Visible cells close up, as on screen: Jan and Mar stay one table, at their own addresses."""
    path = save(tmp_path, CROSS, lambda ws: setattr(ws.column_dimensions["C"], "hidden", True))
    df = edward0.read(path, visible_only=True)
    assert df["table"].unique().tolist() == ["A1:D4"]
    assert sorted(df["address"]) == ["B2", "B3", "B4", "D2", "D3", "D4"]
    t = edward0.tables(path, visible_only=True)
    assert list(t) == [1] and len(t[1]) == 6


def test_hidden_helper_column_leaves_no_trace(tmp_path):
    rows = [["Item", "Q1", "helper", "Q2"],
            ["a", 1, "x", 5],
            ["b", 2, "x", 6]]
    path = save(tmp_path, rows, lambda ws: setattr(ws.column_dimensions["C"], "hidden", True))
    assert {"helper"} <= labels(path) or "x" in set(edward0.read(path)["value"])
    df = edward0.read(path, visible_only=True)
    assert sorted(df["value"]) == [1, 2, 5, 6]
    assert not {"helper", "x"} & labels(path, visible_only=True)


def test_merged_header_starting_in_a_hidden_column_still_labels(tmp_path):
    rows = [["Region", "Sales", None],
            [None, "Q1", "Q2"],
            ["North", 1, 2],
            ["South", 3, 4]]
    path = save(tmp_path, rows, lambda ws: setattr(ws.column_dimensions["B"], "hidden", True),
                merges=["B1:C1"])
    raw = edward0.raw(path, visible_only=True)
    assert sorted(raw["value"]) == [2, 4]
    assert all({"Sales", "Q2"} <= ix for ix in raw["index"])
