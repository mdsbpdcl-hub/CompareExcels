"""Helpers: build test files with openpyxl/csv, run the browser engine in Node, read outputs independently."""
from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess

import openpyxl
import pytest

WEB = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNNER = os.path.join(WEB, "tests", "run_case.mjs")


def write_xlsx(path, sheets, title_rows=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in (title_rows or {}).get(name, []):
            ws.append(r)
        for r in rows:
            ws.append(list(r))
    wb.save(path)
    return str(path)


def write_csv(path, rows, delimiter=",", encoding="utf-8"):
    with open(path, "w", newline="", encoding=encoding) as fh:
        w = csv.writer(fh, delimiter=delimiter)
        for r in rows:
            w.writerow(["" if v is None else v for v in r])
    return str(path)


class Runner:
    def __init__(self, tmp):
        self.tmp = tmp
        self.n = 0

    def __call__(self, a, b, keys=None, **kw):
        self.n += 1
        cfg = {"a": str(a), "b": str(b), "keys": keys or []}
        cfg.update(kw)
        cp = self.tmp / f"cfg{self.n}.json"
        out = self.tmp / f"out{self.n}"
        cp.write_text(json.dumps(cfg))
        p = subprocess.run(["node", RUNNER, str(cp), str(out)], capture_output=True, text=True, timeout=600)
        rp = out / "result.json"
        assert rp.exists(), p.stderr[-2000:]
        res = json.loads(rp.read_text())
        if not res["ok"]:
            raise RuntimeError(res["error"])
        return res


@pytest.fixture
def runner(tmp_path):
    return Runner(tmp_path)


def filled(path, rgb="FFF2CC"):
    """{sheet: {cells}} whose solid fill is rgb — read with openpyxl (independent of the app)."""
    wb = openpyxl.load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        s = set()
        for row in ws.iter_rows():
            for c in row:
                f = c.fill
                if f is not None and f.fill_type == "solid" and str(f.fgColor.rgb or "").upper().endswith(rgb):
                    s.add(c.coordinate)
        out[ws.title] = s
    return out


def report_rows(path, sheet="Detailed Differences"):
    wb = openpyxl.load_workbook(path, read_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    wb.close()
    return [dict(zip(rows[0], r)) for r in rows[1:]]


def all_pass(res):
    failed = [f"{c['name']}: {c['detail']}" for c in res["checks"] if c["critical"] and not c["passed"]]
    assert not failed, failed


def soffice_convert(src, filt, outdir):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        pytest.skip("LibreOffice is needed to create .xls test files")
    subprocess.run([exe, "--headless", "--convert-to", filt, "--outdir", str(outdir), str(src)], capture_output=True,
                   timeout=120)
    out = os.path.join(str(outdir), os.path.splitext(os.path.basename(src))[0] + "." + filt.split(":")[0])
    if not os.path.exists(out):
        pytest.skip(f"LibreOffice cannot write {filt.split(':')[0]} files")
    return out
