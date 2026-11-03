"""Unit tests for the pure stages of hsn8_signal_analysis."""

from __future__ import annotations

import hsn8_signal_analysis as m


def test_rationale_matches_first_rule_case_insensitively():
    assert m.classify_rationale("urea, prilled").startswith("Fertiliser")
    assert m.classify_rationale("POLYETHYLENE LLDPE").startswith("Base petrochemical")
    assert m.classify_rationale("SULPHURIC ACID").startswith("Industrial acid")


def test_rationale_default_when_nothing_matches():
    assert m.classify_rationale("SOMETHING OBSCURE") == m.RATIONALE_DEFAULT


def test_pct_change():
    assert m.pct_change(110, 100) == 10
    assert m.pct_change(50, 100) == -50
    assert m.pct_change(5, 0) is None


def test_percentile_75_nearest_rank_and_empty():
    assert m.percentile_75([]) == 0
    assert m.percentile_75([1, 2, 3, 4]) == 4  # index int(4*0.75)=3
    assert m.percentile_75([10, 1, 5]) == 10  # index int(3*0.75)=2 of sorted


def _row(imp25, imp26, qg, pg):
    return {"imp_fy25": imp25, "imp_fy26": imp26, "qty_growth_pct": qg, "price_growth_pct": pg}


def test_signal_insufficient_when_any_growth_missing():
    assert m.classify_signal(_row(0, 10, 5, 5), 0) == "C_INSUFFICIENT_DATA"  # vg None
    assert m.classify_signal(_row(10, 10, None, 5), 0) == "C_INSUFFICIENT_DATA"
    assert m.classify_signal(_row(10, 10, 5, None), 0) == "C_INSUFFICIENT_DATA"


def test_signal_scissors():
    assert m.classify_signal(_row(100, 85, -12, 6), 1000) == "A_SCISSORS"


def test_signal_inelastic_needs_top_quartile_value():
    assert m.classify_signal(_row(100, 98, 0, 2), 50) == "B_INELASTIC_TARGET"
    assert m.classify_signal(_row(100, 98, 0, 2), 500) == "E_NO_CLEAR_SIGNAL"


def test_signal_deepening_and_none():
    assert m.classify_signal(_row(100, 130, 20, 0), 1000) == "D_DEEPENING"
    assert m.classify_signal(_row(100, 105, 2, 1), 1000) == "E_NO_CLEAR_SIGNAL"


def test_quantity_fields_with_and_without_qty():
    vrow = {"fy25": 100.0, "fy26": 150.0}
    q = {"unit": "KGS", "qty25": 10, "qty26": 10, "qty_growth": 0}
    out = m.quantity_fields(vrow, q)
    assert out["unit_price_fy25"] == 10.0 and out["unit_price_fy26"] == 15.0
    assert out["price_growth_pct"] == 50.0 and out["unit"] == "KGS"

    empty = m.quantity_fields(vrow, None)
    assert set(empty) == set(out) and all(v is None for v in empty.values())
    assert m.quantity_fields(vrow, {"unit": "KGS", "qty25": 0, "qty26": 5, "qty_growth": 1}) == empty


def test_qualifying_rows_filters_chapter_import_floor_and_deficit():
    imports = [
        {"hs": "28010000", "name": "CHLORINE", "fy25": 10, "fy26": 20, "growth": 100},  # keep
        {"hs": "72010000", "name": "PIG IRON", "fy25": 10, "fy26": 20, "growth": 100},  # non-chem
        {"hs": "29010000", "name": "ETHYLENE", "fy25": 1, "fy26": 0.5, "growth": -50},  # import < 1
        {"hs": "39010000", "name": "POLYETHYLENE", "fy25": 5, "fy26": 5, "growth": 0},  # deficit 0
    ]
    exports = [{"hs": "39010000", "name": "POLYETHYLENE", "fy25": 5, "fy26": 5, "growth": 0}]
    rows = m.qualifying_rows(imports, exports, {})
    assert [r["hs"] for r in rows] == ["28010000"]
    r = rows[0]
    assert r["deficit_fy26"] == 20 and r["exp_fy26"] == 0 and r["chapter"] == "28"
    assert r["rationale"] == m.RATIONALE_DEFAULT and r["qty25"] is None


def test_group_by_chapter_preserves_order():
    rows = [{"chapter": "29", "hs": "a"}, {"chapter": "28", "hs": "b"}, {"chapter": "29", "hs": "c"}]
    g = m.group_by_chapter(rows)
    assert list(g) == ["29", "28"] and [r["hs"] for r in g["29"]] == ["a", "c"]
