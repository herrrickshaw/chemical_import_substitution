"""End-to-end test of build() and the CLI on small synthetic TradeStat-shaped workbooks."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

import hsn8_signal_analysis as m

SCRIPT = Path(m.__file__)


def _xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for _ in range(3):  # TradeStat header block
        ws.append(["header"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


@pytest.fixture
def tradestat(tmp_path):
    # value layout: [sno, hs, name, fy25, share, fy26, share, growth]
    imp = _xlsx(
        tmp_path / "imp.xlsx",
        [
            [1, "28010000", "CHLORINE", 100, 0, 85, 0, -15],
            [2, "29010000", "ETHYLENE", 100, 0, 130, 0, 30],
            [3, "3901", "chapter subtotal", 1, 0, 1, 0, 0],
            [4, "72010000", "PIG IRON", 100, 0, 200, 0, 100],
            [5, None, None, None, None, None, None, None],
        ],
    )
    exp = _xlsx(tmp_path / "exp.xlsx", [[1, "28010000", "CHLORINE", 1, 0, 2, 0, 100]])
    # qty layout: [sno, hs, name, unit, qty25, qty26, growth]
    qty = _xlsx(
        tmp_path / "qty.xlsx",
        [[1, "28010000", "CHLORINE", "KGS", 100, 85, -15], [2, "29010000", "ETHYLENE", "KGS", 100, 125, 25]],
    )
    return imp, exp, qty


def test_build_classifies_and_groups(tradestat):
    result = m.build(*tradestat)
    assert set(result) == {"by_chapter", "ch_names", "p75_fy26"}
    assert result["ch_names"] is m.CH_NAMES
    assert result["p75_fy26"] == 130
    by = result["by_chapter"]
    assert list(by) == ["28", "29"]
    cl = by["28"][0]
    assert cl["deficit_fy26"] == 83 and cl["exp_fy26"] == 2
    # value -15%, volume -15%, unit price 1.0 -> 1.0 (flat): not scissors, not top quartile
    assert cl["price_growth_pct"] == 0.0 and cl["is_signal"] == "E_NO_CLEAR_SIGNAL"
    et = by["29"][0]
    assert et["price_growth_pct"] == 4.0 and et["is_signal"] == "B_INELASTIC_TARGET"


def test_cli_writes_json_and_reports_count(tradestat, tmp_path):
    imp, exp, qty = tradestat
    out = tmp_path / "out.json"
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(out), "--import-value", imp, "--export-value", exp, "--import-qty", qty],
        capture_output=True,
        text=True,
        check=True,
    )
    assert proc.stdout.strip() == f"2 qualifying codes written to {out}"
    data = json.loads(out.read_text())
    assert sorted(data["by_chapter"]) == ["28", "29"]


def test_cli_help_smoke():
    proc = subprocess.run([sys.executable, str(SCRIPT), "--help"], capture_output=True, text=True)
    assert proc.returncode == 0 and "--deficit-floor" in proc.stdout
