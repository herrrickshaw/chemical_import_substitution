"""More coverage for scripts/hsn8_signal_analysis.py: loaders, build() boundaries, every rationale rule, CLI block.
All workbooks are synthetic and written to tmp_path; the real DGCI&S downloads are never read."""
import json

import openpyxl
import pytest

import hsn8_signal_analysis as h
from _mainblock import run_main_block


def _xlsx(path, rows, header_rows=3):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(header_rows):
        ws.append([f"hdr{i}"] * 8)           # full-width header so iter_rows yields 8 columns
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return str(path)


# ---------------------------------- loaders ----------------------------------

def test_load_value_file_keeps_only_eight_digit_codes_and_defaults(tmp_path):
    p = _xlsx(tmp_path / "v.xlsx", [
        (1, "29051100", "  METHANOL  ", 10.0, 0, 12.0, 0, 20.0),
        (2, " 29152100 ", None, None, 0, None, 0, None),          # padded code; None name/values -> '' / 0
        (3, "2915210", "7 digits", 1, 0, 1, 0, 0),
        (4, "291521000", "9 digits", 1, 0, 1, 0, 0),
        (5, "", "empty code", 1, 0, 1, 0, 0),
        (6, None, "none code", 1, 0, 1, 0, 0),
        (7, 38089992, "numeric code", 3, 0, 4, 0, 33.3),          # Excel numeric cell
    ])
    rows = h.load_value_file(p)
    assert rows == [
        {"hs": "29051100", "name": "METHANOL", "fy25": 10.0, "fy26": 12.0, "growth": 20.0},
        {"hs": "29152100", "name": "", "fy25": 0, "fy26": 0, "growth": None},
        {"hs": "38089992", "name": "numeric code", "fy25": 3, "fy26": 4, "growth": 33.3},
    ]


def test_load_value_file_skips_first_three_header_rows(tmp_path):
    p = _xlsx(tmp_path / "v.xlsx", [], header_rows=3)
    ws = openpyxl.load_workbook(p).active
    ws.append([1, "29051100", "A", 1, 0, 1, 0, 0])
    ws.insert_rows(1)                                            # a code in row 3 (header region) must be ignored
    ws.cell(row=3, column=2, value="11111111")
    ws.parent.save(p)
    hs = [r["hs"] for r in h.load_value_file(p)]
    assert "11111111" not in hs and "29051100" in hs


def test_loaders_ignore_codes_sitting_in_the_three_header_rows(tmp_path):
    for loader, width in ((h.load_value_file, 8), (h.load_qty_file, 7)):
        wb = openpyxl.Workbook()
        ws = wb.active
        for code in ("11111111", "22222222", "33333333"):             # header rows 1-3 happen to hold 8-char codes
            ws.append([0, code, "n"] + [1] * (width - 3))
        ws.append([1, "44444444", "n"] + [1] * (width - 3))
        p = tmp_path / f"{loader.__name__}.xlsx"
        wb.save(p)
        got = loader(str(p))
        keys = [r["hs"] for r in got] if isinstance(got, list) else list(got)
        assert keys == ["44444444"], loader.__name__


def test_load_qty_file_maps_units_and_defaults(tmp_path):
    p = _xlsx(tmp_path / "q.xlsx", [
        (1, "29051100", "M", "KGS", 100.0, 60.0, -40.0),
        (2, "29152100", "A", "TON", None, None, None),
        (3, "123", "short", "KGS", 1, 1, 0),
        (4, None, "none", "KGS", 1, 1, 0),
        (5, "29152101", "B", None, 5, 0, -100.0),
    ])
    assert h.load_qty_file(p) == {
        "29051100": {"unit": "KGS", "qty25": 100.0, "qty26": 60.0, "qty_growth": -40.0},
        "29152100": {"unit": "TON", "qty25": 0, "qty26": 0, "qty_growth": None},
        "29152101": {"unit": None, "qty25": 5, "qty26": 0, "qty_growth": -100.0},
    }


# ---------------------------------- rationale rules ----------------------------------

@pytest.mark.parametrize("name,idx", [
    ("UREA", 0), ("rock phosphate", 0), ("NPK mixtures", 0),
    ("GOLD BULLION", 1), ("Platinum", 1), ("SILVER", 1),
    ("POLYETHYLENE", 2), ("PVC resin", 2), ("PROPYLENE", 2),
    ("BENZENE", 3), ("TOLUENE", 3), ("PHENOLS", 3), ("ACRYLIC", 3),
    ("MANCOZEB", 4), ("HERBICIDE", 4), ("AGROCHEMICAL", 4),
    ("PIGMENTS", 5), ("DYES", 5), ("COLOURS", 5),
    ("VITAMINS", 6), ("STEROIDS", 6), ("API", 6), ("PHARMACEUTICAL", 6),
    ("NITRIC ACIDS", 7),
    ("ESSENTIAL OILS", 8), ("FRAGRANCE", 8),
    ("ENZYMES", 9), ("CATALYST", 9), ("DIAGNOSTIC", 9),
    ("WAXES", 10), ("SOAP", 10), ("SURFACTANTS", 10), ("POLISHES", 10),
    ("LATEX", 11), ("RUBBER", 11),
    ("CARBON BLACK", 12), ("ACTIVATED CARBON", 12),
])
def test_every_rationale_rule_is_reachable(name, idx):
    assert h.classify_rationale(name) == h.RATIONALE_RULES[idx][1]


def test_rule_table_shape_and_default():
    assert len(h.RATIONALE_RULES) == 13
    assert len({why for _, why in h.RATIONALE_RULES}) == 13 and h.RATIONALE_DEFAULT not in {w for _, w in h.RATIONALE_RULES}
    assert h.classify_rationale("") == h.RATIONALE_DEFAULT


def test_rule_order_decides_overlaps():
    # "POLYETHYLENE OIL" matches the polymer rule (2) before the oil rule (8)
    assert h.classify_rationale("polyethylene oil") == h.RATIONALE_RULES[2][1]
    # a fertiliser beats a generic acid
    assert h.classify_rationale("PHOSPHORIC ACID FERTILISER") == h.RATIONALE_RULES[0][1]


def test_words_embedded_in_other_words_do_not_match():
    assert h.classify_rationale("BOILERPLATE") == h.RATIONALE_DEFAULT      # contains OIL but not as a word
    assert h.classify_rationale("RAPIDLY") == h.RATIONALE_DEFAULT          # contains API
    assert h.classify_rationale("SPOLISHED") == h.RATIONALE_DEFAULT
    assert h.classify_rationale("MARIGOLD EXTRACT") == h.RATIONALE_DEFAULT  # GOLD inside a word


def test_chapter_tables_are_consistent():
    assert set(h.CHEM_CHAPTERS) == set(h.CH_NAMES) and len(h.CHEM_CHAPTERS) == 8
    assert all(len(c) == 2 for c in h.CHEM_CHAPTERS)


# ---------------------------------- classify_signal boundaries ----------------------------------

def _o(fy25, fy26, qg, pg):
    return {"imp_fy25": fy25, "imp_fy26": fy26, "qty_growth_pct": qg, "price_growth_pct": pg}


@pytest.mark.parametrize("o,p75,sig", [
    (_o(100, 90, -10.01, 5), 1e9, "A_SCISSORS"),
    (_o(100, 90, -9.99, 5), 1e9, "E_NO_CLEAR_SIGNAL"),                # volume not falling enough
    (_o(100, 90, -10, 4.99), 1e9, "E_NO_CLEAR_SIGNAL"),               # price not rising enough
    (_o(100, 95, 0, -8), 95, "E_NO_CLEAR_SIGNAL"),                    # |pg| <= 8 but value exactly -5%: vg > -5 fails
    (_o(100, 95.01, 0, -8), 95.01, "B_INELASTIC_TARGET"),             # imp_fy26 == p75 counts (>=)
    (_o(100, 95.0, 0, 0), 95.01, "E_NO_CLEAR_SIGNAL"),                # just under the 75th percentile
    (_o(100, 115, 15, 0), 1e9, "D_DEEPENING"),
    (_o(100, 114.99, 15, 0), 1e9, "E_NO_CLEAR_SIGNAL"),
    (_o(100, 115, 14.99, 0), 1e9, "E_NO_CLEAR_SIGNAL"),
    (_o(100, 100, 0, -8.5), 100, "E_NO_CLEAR_SIGNAL"),                # a big price DROP is not "flat price"
    (_o(100, 100, 0, -8), 100, "B_INELASTIC_TARGET"),
])
def test_classify_signal_boundaries(o, p75, sig):
    assert h.classify_signal(o, p75) == sig


def test_classify_signal_missing_keys_mean_insufficient():
    assert h.classify_signal({"imp_fy25": 10, "imp_fy26": 9}, 1) == "C_INSUFFICIENT_DATA"
    assert h.classify_signal({"imp_fy25": 10, "imp_fy26": 9, "qty_growth_pct": 1}, 1) == "C_INSUFFICIENT_DATA"
    assert h.classify_signal(_o(10, 9, 0, 0), 1) != "C_INSUFFICIENT_DATA"
    assert h.classify_signal(_o(0, 9, 0, 0), 1) == "C_INSUFFICIENT_DATA"      # zero base: growth undefined


# ---------------------------------- build() ----------------------------------

def _build(tmp_path, imp, exp, qty, **kw):
    return h.build(_xlsx(tmp_path / "i.xlsx", imp), _xlsx(tmp_path / "e.xlsx", exp),
                   _xlsx(tmp_path / "q.xlsx", qty), **kw)


def _flat(res):
    return {o["hs"]: o for v in res["by_chapter"].values() for o in v}


def test_build_import_floor_and_deficit_floor_are_inclusive_at_the_boundary(tmp_path):
    imp = [(1, "29000001", "A", 1, 0, 1.0, 0, 0),        # fy26 == 1.0 and no exports: kept (not < 1)
           (2, "29000002", "B", 1, 0, 0.99, 0, 0),       # fy26 < 1: dropped
           (3, "29000003", "C", 1, 0, 2.0, 0, 0),        # deficit 0.5 exactly: kept
           (4, "29000004", "D", 1, 0, 2.0, 0, 0)]        # deficit 0.49: dropped
    exp = [(1, "29000003", "C", 0, 0, 1.5, 0, 0), (2, "29000004", "D", 0, 0, 1.51, 0, 0)]
    res = _build(tmp_path, imp, exp, [])
    assert set(_flat(res)) == {"29000001", "29000003"}
    res = _build(tmp_path, imp, exp, [], deficit_floor_musd=1.0)
    assert set(_flat(res)) == {"29000001"}                # custom floor: deficit 1.0 passes, 0.5 does not


def test_build_groups_by_chapter_and_rounds_deficit(tmp_path):
    imp = [(1, "28000001", "A", 1, 0, 10.0, 0, 0), (2, "39000001", "B", 1, 0, 10.0, 0, 0),
           (3, "31000001", "C", 1, 0, 10.04, 0, 0), (4, "30000001", "pharma ch30", 1, 0, 10.0, 0, 0)]
    exp = [(1, "31000001", "C", 0, 0, 0.0049, 0, 0)]
    res = _build(tmp_path, imp, exp, [])
    assert set(res["by_chapter"]) == {"28", "39", "31"}          # chapter 30 is not a chemical chapter here
    assert _flat(res)["31000001"]["deficit_fy26"] == 10.0        # round(10.0351, 1)
    assert res["ch_names"] is h.CH_NAMES


def test_build_record_fields_with_quantity(tmp_path):
    imp = [(1, "29051100", "METHANOL", 200.0, 0, 100.0, 0, -50.0)]
    qty = [(1, "29051100", "METHANOL", "KGS", 400.0, 160.0, -60.0)]
    o = _flat(_build(tmp_path, imp, [], qty))["29051100"]
    assert o["chapter"] == "29" and o["name"] == "METHANOL" and o["imp_growth"] == -50.0
    assert (o["imp_fy25"], o["imp_fy26"], o["exp_fy26"]) == (200.0, 100.0, 0)
    assert (o["qty25"], o["qty26"], o["unit"], o["qty_growth_pct"]) == (400.0, 160.0, "KGS", -60.0)
    assert (o["unit_price_fy25"], o["unit_price_fy26"], o["price_growth_pct"]) == (0.5, 0.625, 25.0)
    assert o["is_signal"] == "A_SCISSORS"                         # value -50%, volume -60%, price +25%
    assert o["rationale"] == h.classify_rationale("METHANOL")


@pytest.mark.parametrize("qrow", [
    [],                                                         # no row at all
    [(1, "29000001", "A", "KGS", 0, 5.0, 10.0)],                # zero FY25 quantity
    [(1, "29000001", "A", "KGS", 5.0, 0, 10.0)],                # zero FY26 quantity
])
def test_build_without_usable_quantity_is_insufficient_data(tmp_path, qrow):
    o = _flat(_build(tmp_path, [(1, "29000001", "A", 5.0, 0, 6.0, 0, 20.0)], [], qrow))["29000001"]
    assert o["qty25"] is None and o["unit"] is None and o["unit_price_fy26"] is None
    assert o["price_growth_pct"] is None and o["is_signal"] == "C_INSUFFICIENT_DATA"


def test_build_unit_price_precision_and_missing_quantity_growth(tmp_path):
    imp = [(1, "29000001", "A", 1.0, 0, 2.0, 0, 100.0), (2, "29000002", "B", 1.0, 0, 2.0, 0, 100.0)]
    qty = [(1, "29000001", "A", "KGS", 3.0, 7.0, 133.3), (2, "29000002", "B", "KGS", 3.0, 7.0, None)]
    got = _flat(_build(tmp_path, imp, [], qty))
    a = got["29000001"]
    assert (a["unit_price_fy25"], a["unit_price_fy26"]) == (0.3333, 0.2857)       # 4 decimal places
    assert a["price_growth_pct"] == -14.3
    b = got["29000002"]
    assert b["qty_growth_pct"] is None and b["is_signal"] == "C_INSUFFICIENT_DATA"   # None is kept, not coerced to 0


def test_build_zero_import_base_gives_no_price_growth(tmp_path):
    imp = [(1, "29000001", "A", 0, 0, 6.0, 0, None)]
    qty = [(1, "29000001", "A", "KGS", 5.0, 6.0, 20.0)]
    o = _flat(_build(tmp_path, imp, [], qty))["29000001"]
    assert o["unit_price_fy25"] == 0.0 and o["price_growth_pct"] is None
    assert o["is_signal"] == "C_INSUFFICIENT_DATA"


def test_build_p75_is_the_75th_percentile_index_of_sorted_values(tmp_path):
    imp = [(i, f"2900000{i}", f"N{i}", 1, 0, float(v), 0, 0) for i, v in enumerate([5, 1, 9, 3, 7, 11, 13, 15], 1)]
    res = _build(tmp_path, imp, [], [])
    # sorted: 1 3 5 7 9 11 13 15 ; int(8 * 0.75) = 6 -> 13
    assert res["p75_fy26"] == 13.0
    res1 = _build(tmp_path, imp[:1], [], [])
    assert res1["p75_fy26"] == 5.0                                # one qualifying code: index int(0.75) = 0


def test_build_with_nothing_qualifying(tmp_path):
    res = _build(tmp_path, [(1, "84713010", "LAPTOPS", 1, 0, 100.0, 0, 0)], [], [])
    assert res == {"by_chapter": {}, "ch_names": h.CH_NAMES, "p75_fy26": 0}


def test_build_inelastic_target_uses_p75_of_qualifying_codes(tmp_path):
    imp = [(1, "29000001", "A", 100.0, 0, 100.0, 0, 0), (2, "29000002", "B", 10.0, 0, 10.0, 0, 0),
           (3, "29000003", "C", 10.0, 0, 10.0, 0, 0), (4, "29000004", "D", 10.0, 0, 10.0, 0, 0)]
    qty = [(i, f"2900000{i}", "x", "KGS", 10.0, 10.0, 0) for i in range(1, 5)]
    got = _flat(_build(tmp_path, imp, [], qty))
    assert got["29000001"]["is_signal"] == "B_INELASTIC_TARGET"   # top quartile, flat price, steady value
    assert got["29000002"]["is_signal"] == "E_NO_CLEAR_SIGNAL"


# ---------------------------------- CLI block ----------------------------------

def test_cli_writes_json_and_reports_count(tmp_path, capsys):
    result = {"by_chapter": {"29": [{"hs": "1"}, {"hs": "2"}], "38": [{"hs": "3"}]}, "ch_names": {}, "p75_fy26": 1}
    seen = {}

    def fake_build(*paths):
        seen["paths"] = paths
        return result
    out = tmp_path / "out.json"
    run_main_block(h, [str(out)], build=fake_build)
    assert json.loads(out.read_text()) == result
    assert capsys.readouterr().out == f"3 qualifying codes written to {out}\n"
    assert len(seen["paths"]) == 3 and all(p.endswith(".xlsx") for p in seen["paths"])
    assert seen["paths"][0].endswith("Import-Commodity-wise.xlsx") and seen["paths"][1].endswith("Export-Commodity-wise.xlsx")
    assert seen["paths"][2].endswith("Import-Commodity-wise (1).xlsx")


def test_cli_default_output_name(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    run_main_block(h, [], build=lambda *a: {"by_chapter": {}, "ch_names": {}, "p75_fy26": 0})
    assert (tmp_path / "chem_hsn8_signals.json").exists()
    assert capsys.readouterr().out == "0 qualifying codes written to chem_hsn8_signals.json\n"
