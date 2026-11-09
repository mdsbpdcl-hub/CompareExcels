"""The browser engine, run in Node on real files; outputs verified with openpyxl (independently)."""
from __future__ import annotations

import csv
import datetime as dt
import random
import zipfile

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill

from conftest import all_pass, filled, report_rows, soffice_convert, write_csv, write_xlsx

HEAD = ["ID", "Name", "Amount", "Date"]
BASE = [["001", "Ram", "500", "2024-01-05"], ["002", "Sita", "600", "2024-02-01"],
        ["003", "Mohan", "700", "2024-03-09"], ["004", "Geeta", "800", "2024-04-10"]]
K = [["ID", "ID"]]


def test_identical(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [HEAD] + BASE})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [HEAD] + BASE})
    r = runner(a, b, K)
    assert r["totals"]["matched"] == 4 and r["totals"]["changed_cells"] == 0
    assert filled(r["highlighted"])["S"] == set()
    all_pass(r)


def test_changed_cell_yellow_exact_address_reported_once(tmp_path, runner):
    rb = [x[:] for x in BASE]
    rb[2][2] = "750"
    a = write_xlsx(tmp_path / "a.xlsx", {"ConsumerData": [HEAD] + BASE})
    b = write_xlsx(tmp_path / "b.xlsx", {"UpdatedConsumerData": [HEAD] + rb})
    r = runner(a, b, K)
    assert filled(r["highlighted"])["UpdatedConsumerData"] == {"C4"}
    rep = report_rows(r["report"])
    hits = [x for x in rep if x["B: Cell"] == "C4"]
    assert len(hits) == 1 and len(rep) == 1
    h = hits[0]
    assert (h["A: Sheet"], h["A: Cell"], h["Before Value"], h["After Value"], h["Match Key"]) == \
           ("ConsumerData", "C4", "700", "750", "003")
    assert openpyxl.load_workbook(r["highlighted"])["UpdatedConsumerData"]["C4"].value == "750"
    all_pass(r)


def test_sheet_names_title_rows_order_headings(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"ConsumerData": [["Consumer ID", "Name", "Amount"], ["C1", "Ram", "500"],
                                                          ["C2", "Sita", "600"]]},
                   title_rows={"ConsumerData": [["Consumer report"], ["Sept"], []]})
    b = write_xlsx(tmp_path / "b.xlsx", {"Updated": [["Amount", "Name", "CA Number"], ["600", "Sita", "C2"],
                                                     ["550", "Ram", "C1"]]}, title_rows={"Updated": [["Updated list"]]})
    r = runner(a, b, [["Consumer ID", "CA Number"]], a_header=4, b_header=2, columns=[["Name", "Name"], ["Amount", "Amount"]])
    rep = report_rows(r["report"])
    assert [(x["A: Cell"], x["B: Cell"], x["B: Col Letter"]) for x in rep] == [("C5", "A4", "A")]
    assert filled(r["highlighted"])["Updated"] == {"A4"}
    all_pass(r)


def test_composite_key(tmp_path, runner):
    h = ["Div", "Acc", "Bill"]
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [h, ["P1", "1", "10"], ["P2", "1", "20"], ["P1", "2", "30"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [h, ["P2", "1", "25"], ["P1", "1", "10"], ["P1", "2", "30"]]})
    r = runner(a, b, [["Div", "Div"], ["Acc", "Acc"]], columns=[["Bill", "Bill"]])
    assert [(x["Match Key"], x["B: Cell"]) for x in report_rows(r["report"])] == [("P2 | 1", "C2")]
    all_pass(r)


def test_leading_zeros(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "Code"], ["007", "0042"], ["008", "9"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "Code"], ["7", "42"], ["008", "09"]]})
    r = runner(a, b, K)
    assert (r["totals"]["matched"], r["totals"]["only_a"], r["totals"]["changed_cells"]) == (1, 1, 1)
    r = runner(a, b, K, key_rules={"leading_zeros": True}, rules={"leading_zeros": True})
    assert (r["totals"]["matched"], r["totals"]["changed_cells"]) == (2, 0)
    ws = openpyxl.load_workbook(r["highlighted"])["S"]
    assert ws["A2"].value == "7" and ws["A3"].value == "008"
    all_pass(r)


def test_spaces_case(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "Name"], ["1", "Ram Kumar"], ["2", "sita"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "Name"], ["1", "  Ram   Kumar "], ["2", "SITA"]]})
    assert runner(a, b, K)["totals"]["changed_cells"] == 2
    assert runner(a, b, K, rules={"trim": True, "collapse_spaces": True})["totals"]["changed_cells"] == 1
    r = runner(a, b, K, rules={"trim": True, "collapse_spaces": True, "ignore_case": True})
    assert r["totals"]["changed_cells"] == 0 and filled(r["highlighted"])["S"] == set()
    all_pass(r)


def test_numbers_only_selected_columns(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "Amount", "Ref"], ["1", "1000", "1000"], ["2", "5.5", "7"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "Amount", "Ref"], ["1", "1,000.00", "1,000"], ["2", "5.50", "7"]]})
    r = runner(a, b, K, rules={"numeric": True, "numeric_columns": ["Amount"]})
    assert [(x["A: Column"], x["B: Cell"]) for x in report_rows(r["report"])] == [("Ref", "C2")]
    all_pass(r)


def test_number_cells_and_float_noise(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], ["1", 1000], ["2", 0.1 + 0.2], ["3", 1 / 3]]})
    b = write_csv(tmp_path / "b.csv", [["ID", "V"], ["1", "1000"], ["2", "0.3"], ["3", "0.333333333333333"]])
    assert runner(a, b, K)["totals"]["changed_cells"] == 0


def test_dates(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "D"], ["1", dt.datetime(2024, 1, 5)], ["2", "2024-06-05"],
                                               ["3", "2024-01-20"], ["4", "2024-03-03"], ["5", "31/02/2024"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "D"], ["1", "05-Jan-2024"], ["2", "05/06/2024"],
                                               ["3", "20/01/2024"], ["4", "03/03/2024"], ["5", "31/02/2024"]]})
    r = runner(a, b, K, rules={"dates": True, "date_columns": ["D"]})
    assert {x["Match Key"] for x in report_rows(r["report"])} == {"2"}
    assert runner(a, b, K, rules={"dates": True, "date_columns": ["D"], "date_order": "dmy"})["totals"]["changed_cells"] == 0
    all_pass(r)


def test_excel_date_cells_read_as_dates(tmp_path, runner):
    wa = openpyxl.Workbook(); ws = wa.active; ws.title = "S"
    ws.append(["ID", "D", "T"]); ws.append(["1", dt.datetime(2024, 1, 5), dt.time(12, 30)])
    ws["B2"].number_format = "dd-mmm-yyyy"; ws["C2"].number_format = "h:mm"
    wa.save(tmp_path / "a.xlsx")
    b = write_csv(tmp_path / "b.csv", [["ID", "D", "T"], ["1", "2024-01-05", "12:30:00"]])
    r = runner(tmp_path / "a.xlsx", b, K)
    assert r["totals"]["changed_cells"] == 0, report_rows(r["report"])


def test_symbols_breaks_blank(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "Amt", "Addr", "Note"], ["1", "₹500", "Line 1\nLine 2", None]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "Amt", "Addr", "Note"], ["1", "500", "Line 1 Line 2", "  "]]})
    assert runner(a, b, K)["totals"]["changed_cells"] == 3
    r = runner(a, b, K, rules={"symbols": True, "symbol_chars": "₹", "line_breaks": True, "blank_equals_empty": True})
    assert r["totals"]["changed_cells"] == 0
    all_pass(r)


def test_blank_keys(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [None, "x"], ["1", "a"], ["  ", "y"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [None, "z"], ["1", "a"]]})
    r = runner(a, b, K)
    assert (r["totals"]["matched"], r["totals"]["blank_key_a"], r["totals"]["blank_key_b"]) == (1, 2, 1)
    oa = report_rows(r["report"], "Only in A")
    assert sorted(x["Original Row"] for x in oa) == [2, 4]
    all_pass(r)


@pytest.mark.parametrize("policy,matched,excluded,changed", [
    ("keep_first", 2, 1, 0), ("keep_last", 2, 1, 1), ("pair", 3, 0, 1), ("exclude", 1, 2, 0)])
def test_duplicates(tmp_path, runner, policy, matched, excluded, changed):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], ["1", "a"], ["1", "b"], ["2", "c"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], ["1", "a"], ["2", "c"], ["1", "a"]]})
    r = runner(a, b, K, dup_policy=policy)
    t = r["totals"]
    assert (t["matched"], t["dup_excluded_a"], t["changed_cells"]) == (matched, excluded, changed)
    assert {(x["File"], x["Row"]) for x in report_rows(r["report"], "Duplicate Keys")} == \
           {("a.xlsx", 2), ("a.xlsx", 3), ("b.xlsx", 2), ("b.xlsx", 4)}
    all_pass(r)


def test_blank_value_and_missing_cell(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "P", "Q", "R"], ["1", None, "x", "keep"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "P", "Q", "R"], ["1", "new", None, "keep"]]})
    r = runner(a, b, K)
    assert {x["B: Cell"]: x["Difference Type"] for x in report_rows(r["report"])} == \
           {"B2": "Blank-to-Value", "C2": "Value-to-Blank"}
    assert filled(r["highlighted"])["S"] == {"B2", "C2"}
    all_pass(r)


def test_only_in_and_orange(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [HEAD] + BASE})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [HEAD] + BASE[:3] + [["005", "New", "1", "2024-05-05"], [None, "Total"]]})
    r = runner(a, b, K, highlight_only_b=True)
    assert [(x["Match Key"], x["Original Row"]) for x in report_rows(r["report"], "Only in A")] == [("004", 5)]
    assert filled(r["highlighted"], "FCE4D6")["S"] == {"A5", "B5", "C5", "D5"}
    all_pass(r)


def test_formula_difference(tmp_path, runner):
    wa = openpyxl.Workbook(); ws = wa.active; ws.title = "S"; ws.append(["ID", "X", "Y"]); ws.append(["1", 5, 10])
    wa.save(tmp_path / "a.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S"; ws.append(["ID", "X", "Y"]); ws.append(["1", 5, "=B2*2"])
    wb.save(tmp_path / "b_raw.xlsx")
    with zipfile.ZipFile(tmp_path / "b_raw.xlsx") as zin, zipfile.ZipFile(tmp_path / "b.xlsx", "w") as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<f>B2*2</f><v></v>", b"<f>B2*2</f><v>10</v>")
            zout.writestr(it, data)
    r = runner(tmp_path / "a.xlsx", tmp_path / "b.xlsx", K, compare_formulas=True)
    rep = report_rows(r["report"])
    assert [(x["B: Cell"], x["Difference Type"]) for x in rep] == [("C2", "Formula Difference")]
    assert "B2*2" in rep[0]["Remarks"]
    assert runner(tmp_path / "a.xlsx", tmp_path / "b.xlsx", K)["totals"]["changed_cells"] == 0
    all_pass(r)


CA = [["ID", "Name", "Amount"], ["001", "Ram, Jr.", "500"], ["002", 'Sita "S"', "600"], ["003", "Line1\nLine2", "700"]]
CB = [["ID", "Name", "Amount"], ["001", "Ram, Jr.", "550"], ["002", 'Sita "S"', "600"], ["003", "Line1\nLine2", "700"]]


@pytest.mark.parametrize("delim", [",", "\t", ";", "|"])
@pytest.mark.parametrize("enc", ["utf-8", "utf-8-sig", "cp1252"])
def test_csv_delims_encodings(tmp_path, runner, delim, enc):
    extra = [["004", "Café €" if enc == "cp1252" else "Café ₹", "1"]]
    a = write_csv(tmp_path / "a.csv", CA + extra, delim, enc)
    b = write_csv(tmp_path / "b.csv", CB + extra, delim, enc)
    r = runner(a, b, K)
    assert [(x["B: Cell"], x["Before Value"], x["After Value"]) for x in report_rows(r["report"])] == [("C2", "500", "550")]
    assert r["totals"]["matched"] == 4
    assert filled(r["highlighted"])["b"] == {"C2"}
    all_pass(r)


def test_csv_blank_lines_multiline(tmp_path, runner):
    (tmp_path / "a.csv").write_text('ID,V\n1,"x\ny"\n\n2,a\n', encoding="utf-8")
    (tmp_path / "b.csv").write_text('ID,V\n\n1,"x\ny"\n2,b\n', encoding="utf-8")
    r = runner(tmp_path / "a.csv", tmp_path / "b.csv", K)
    assert [(x["A: Cell"], x["B: Cell"]) for x in report_rows(r["report"])] == [("B4", "B4")]
    all_pass(r)


@pytest.mark.parametrize("combo", ["xlsx-csv", "csv-xlsx", "xlsx-xlsx", "csv-csv"])
def test_combinations(tmp_path, runner, combo):
    fa, fb = combo.split("-")
    rb = [x[:] for x in BASE]
    rb[1][1] = "Seeta"
    mk = {"xlsx": lambda p, rows: write_xlsx(p, {"Data": [HEAD] + rows}), "csv": lambda p, rows: write_csv(p, [HEAD] + rows)}
    r = runner(mk[fa](tmp_path / f"a.{fa}", BASE), mk[fb](tmp_path / f"b.{fb}", rb), K)
    assert [(x["B: Cell"], x["After Value"]) for x in report_rows(r["report"])] == [("B3", "Seeta")]
    assert filled(r["highlighted"])["Data" if fb == "xlsx" else "b"] == {"B3"}
    all_pass(r)


def test_xls(tmp_path, runner):
    rb = [x[:] for x in BASE]
    rb[0][1] = "Raam"
    a = write_xlsx(tmp_path / "a.xlsx", {"Data": [HEAD] + BASE, "Other": [["keep me"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"Data": [HEAD] + rb, "Other": [["keep me"]]})
    a2 = soffice_convert(a, "xls:MS Excel 97", tmp_path / "ca")
    b2 = soffice_convert(b, "xls:MS Excel 97", tmp_path / "cb")
    r = runner(a2, b2, K)
    assert [(x["B: Cell"], x["After Value"]) for x in report_rows(r["report"])] == [("B2", "Raam")]
    assert filled(r["highlighted"])["Data"] == {"B2"}
    wb = openpyxl.load_workbook(r["highlighted"])
    assert wb.sheetnames == ["Data", "Other"] and wb["Other"]["A1"].value == "keep me"
    all_pass(r)


def test_same_workbook_two_sheets(tmp_path, runner):
    rb = [x[:] for x in BASE]
    rb[3][2] = "880"
    p = write_xlsx(tmp_path / "book.xlsx", {"Before": [HEAD] + BASE, "After": [HEAD] + rb})
    r = runner(p, p, K, a_sheet="Before", b_sheet="After")
    f = filled(r["highlighted"])
    assert f["After"] == {"C5"} and f["Before"] == set()
    all_pass(r)


def test_multiple_pairs(tmp_path, runner):
    rb = [x[:] for x in BASE]
    rb[0][2] = "999"
    a = write_xlsx(tmp_path / "a.xlsx", {"Jan": [HEAD] + BASE, "Feb": [HEAD] + BASE})
    b = write_xlsx(tmp_path / "b.xlsx", {"Feb2": [HEAD] + rb, "Jan2": [HEAD] + BASE[:3]})
    cols = [{"a": c, "b": c, "compare": True} for c in HEAD[1:]]
    key = [{"a": "ID", "b": "ID"}]
    pairs = [{"a_sheet": "Jan", "b_sheet": "Jan2", "a_header": 1, "b_header": 1, "keys": key, "columns": cols},
             {"a_sheet": "Feb", "b_sheet": "Feb2", "a_header": 1, "b_header": 1, "keys": key, "columns": cols}]
    r = runner(a, b, None, pairs=pairs)
    f = filled(r["highlighted"])
    assert f["Feb2"] == {"C2"} and f["Jan2"] == set()
    wb = openpyxl.load_workbook(r["report"], read_only=True)
    assert "Only in A (P1)" in wb.sheetnames and "Only in B (P2)" in wb.sheetnames
    all_pass(r)


def test_layout_kept(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [HEAD] + BASE})
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S"
    ws.append(HEAD)
    for x in BASE:
        ws.append(x)
    ws["C3"] = "650"; ws["C3"].font = Font(bold=True, color="FF0000")
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws.column_dimensions["B"].width = 33
    ws.merge_cells("F1:G1"); ws["F1"] = "merged"
    ws["H2"] = "=C2*2"
    wb.create_sheet("Notes")["A1"] = "untouched"
    wb.save(tmp_path / "b.xlsx")
    r = runner(a, tmp_path / "b.xlsx", K, columns=[["Amount", "Amount"]])
    s = openpyxl.load_workbook(r["highlighted"])
    w = s["S"]
    assert w["C3"].value == "650" and w["C3"].font.b and w["C3"].fill.fgColor.rgb == "FFFFF2CC"
    assert w.column_dimensions["B"].width == 33 and "F1:G1" in {str(x) for x in w.merged_cells.ranges}
    assert w["H2"].value == "=C2*2" and s["Notes"]["A1"].value == "untouched"
    all_pass(r)


def test_injection_and_illegal_chars(tmp_path, runner):
    a = write_csv(tmp_path / "a.csv", [["ID", "V"], ["1", "safe"], ["=2", "x"]])
    b = write_csv(tmp_path / "b.csv", [["ID", "V"], ["1", '=HYPERLINK("http://x","click")'], ["=2", "@SUM(1)\x01￿"]])
    r = runner(a, b, K)
    for p in (r["report"], r["highlighted"]):
        with zipfile.ZipFile(p) as z:
            assert all("<f>" not in z.read(n).decode() for n in z.namelist() if n.startswith("xl/worksheets/"))
        openpyxl.load_workbook(p)
    vals = {x["After Value"] for x in report_rows(r["report"])}
    assert '=HYPERLINK("http://x","click")' in vals


def test_errors(tmp_path, runner):
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "X", "Y"], ["1", "a", "b"]]})
    b = write_xlsx(tmp_path / "b.xlsx", {"S": [["ID", "Z"], ["1", "a"]]})
    with pytest.raises(RuntimeError, match="more than one"):
        runner(a, b, K, columns=[["X", "Z"], ["Y", "Z"]])
    with pytest.raises(RuntimeError, match="key column"):
        runner(a, b, K, columns=[["X", "ID"]])
    with pytest.raises(RuntimeError, match="not found"):
        runner(a, b, K, columns=[["X", "Nope"]])


@pytest.mark.parametrize("seed", [1, 2])
@pytest.mark.parametrize("bk", ["xlsx", "csv"])
def test_randomized(tmp_path, runner, seed, bk):
    rnd = random.Random(seed)
    head = ["Key", "C1", "C2", "C3", "C4"]
    n = 400
    ar = [[f"K{i:05d}"] + [str(rnd.randint(0, 9)) for _ in range(4)] for i in range(n)]
    br = [x[:] for x in ar]
    rnd.shuffle(br)
    for _ in range(60):
        i, c = rnd.randrange(n), rnd.randint(1, 4)
        br[i][c] = rnd.choice(["", "10", "x", br[i][c] + "!"])
    am = {x[0]: x for x in ar}
    exp = {f"{'ABCDE'[c]}{k + 2}" for k, x in enumerate(br) for c in range(1, 5) if (x[c] or None) != (am[x[0]][c] or None)}
    a = write_xlsx(tmp_path / "a.xlsx", {"Data": [head] + ar})
    if bk == "xlsx":
        b = write_xlsx(tmp_path / "b.xlsx", {"Data": [head] + [[v or None for v in x] for x in br]}); sh = "Data"
    else:
        b = write_csv(tmp_path / "b.csv", [head] + br); sh = "b"
    r = runner(a, b, [["Key", "Key"]])
    assert filled(r["highlighted"])[sh] == exp
    cells = [x["B: Cell"] for x in report_rows(r["report"])]
    assert sorted(cells) == sorted(exp) and len(cells) == len(set(cells))
    all_pass(r)


def test_one_lakh_rows(tmp_path, runner):
    N = 100_000
    head = ["Consumer ID", "Name", "Division", "Load", "Arrear", "Mobile"]
    rows = [[f"{i:011d}", f"Name {i % 997}", f"Div {i % 31}", str(i % 10), f"{(i * 37) % 25000}.50", f"9{i:09d}"]
            for i in range(1, N + 1)]
    changed = sorted(random.Random(5).sample(range(N), 200))
    with open(tmp_path / "a.csv", "w", newline="") as fh:
        w = csv.writer(fh); w.writerow(head); w.writerows(rows)
    wb = openpyxl.Workbook(write_only=True); ws = wb.create_sheet("B"); ws.append(head)
    exp = set()
    for k, x in enumerate(rows):
        x = list(x)
        if k in changed:
            x[4] = "1.00"; exp.add(f"E{k + 2}")
        ws.append(x)
    wb.save(tmp_path / "b.xlsx")
    r = runner(tmp_path / "a.csv", tmp_path / "b.xlsx", [["Consumer ID", "Consumer ID"]])
    assert r["totals"]["matched"] == N and r["totals"]["changed_cells"] == 200
    assert {x["B: Cell"] for x in report_rows(r["report"])} == exp
    all_pass(r)
    print(f"\n1 lakh rows in {r['seconds']:.1f}s (Node)")


def test_xlsb_and_error_cells(tmp_path, runner):
    import os
    fx = os.path.join(os.path.dirname(__file__), "fixtures")
    r = runner(os.path.join(fx, "sheetjs_a.xlsb"), os.path.join(fx, "sheetjs_b.xlsb"), K, a_header=3, b_header=3)
    assert [(x["B: Cell"], x["Before Value"], x["After Value"]) for x in report_rows(r["report"])] == [("C5", "600", "650")]
    wb = openpyxl.load_workbook(r["highlighted"])
    assert wb.sheetnames == ["Data", "Other"]
    assert wb["Data"]["E4"].value == "#N/A" and wb["Data"]["E4"].data_type == "e"   # error cell kept as an error
    assert filled(r["highlighted"])["Data"] == {"C5"}
    all_pass(r)


def test_unpatchable_workbook_falls_back_to_values_copy(tmp_path, runner):
    import re
    a = write_xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], ["1", "a"], ["2", "b"]]})
    b0 = write_xlsx(tmp_path / "b0.xlsx", {"S": [["ID", "V"], ["1", "a"], ["2", "c"]]})
    with zipfile.ZipFile(b0) as zi, zipfile.ZipFile(tmp_path / "b.xlsx", "w") as zo:
        for it in zi.infolist():
            d = zi.read(it.filename)
            if it.filename == "xl/styles.xml":
                d = re.sub(rb"<fills[^>]*>.*?</fills>", b"", d, flags=re.S)
            zo.writestr(it, d)
    r = runner(a, tmp_path / "b.xlsx", K)
    assert r["method"] == "converted" and "values-only copy" in r["notes"][0]
    assert filled(r["highlighted"])["S"] == {"B3"}
    all_pass(r)
