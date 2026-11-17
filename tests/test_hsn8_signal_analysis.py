"""scripts/hsn8_signal_analysis.py: rationale tags, signal classification and the build() pipeline
on small synthetic TradeStat-shaped workbooks (no DGCI&S download)."""
import openpyxl
import pytest

import hsn8_signal_analysis as h

FERT, POLY, AGRO, DYE, PHARMA, OIL, ACID = (h.RATIONALE_RULES[i][1] for i in (0, 2, 4, 5, 6, 8, 7))


@pytest.mark.parametrize("name,why", [
    ("UREA, WHETHER OR NOT IN AQUEOUS SOLUTION", FERT),
    ("DIAMONM HYDRGNORTHPHOSPHT", FERT),
    ("POTASSIUM CHLORIDE", FERT),
    ("POLYPROPYLENE IN PRIMARY FORMS", POLY),
    ("ACETIC ACID", h.RATIONALE_RULES[3][1]),          # organic intermediate wins over generic acid (rule order)
    ("SULPHURIC ACID", ACID),
    ("MANCOZEB TECHNICAL", AGRO),
    ("SOMETHING UNRELATED", h.RATIONALE_DEFAULT),
])
def test_classify_rationale_basic(name, why):
    assert h.classify_rationale(name) == why


@pytest.mark.parametrize("name,why", [
    ("SUPERPHOSPHATES CONTAINING 35 % OR MORE", FERT),     # stem SUPERPHOSPHAT
    ("OTHER FERTILISERS N.E.S.", FERT),                    # stem FERTILIS
    ("MINRL/CHEMCL FERTLSRS CONTNG THREE ELMNTS", FERT),   # TradeStat abbreviation, plural
    ("Fertilizer grade, conforming to IS 5406", FERT),     # US spelling (urea sub-line)
    ("POLYCARBONATES", POLY),                              # stem POLYCARBONAT
    ("OTHERS-OF-HERBICIDES, ANTI-SPROUTING PRODUCTS", AGRO),
    ("PESTICIDES, NOT ELSEWHERE SPECIFIED", AGRO),
    ("INSECTICIDES BASED ON PYRETHROIDS", AGRO),
    ("ORGANIC PIGMENTS", DYE),
    ("ANTIBIOTICS NES", PHARMA),
    ("PERFUMERY COMPOUNDS", OIL),
])
def test_classify_rationale_matches_stems_and_plurals(name, why):
    """The rules list truncated stems (INSECTICID, FUNGICID, SUPERPHOSPHAT, FERTILIS, POLYCARBONAT, PERFUM...)
    but wrapped them in \\b...\\b, so the full word never matched: e.g. 'PESTICIDES' and 'OTHER FERTILISERS'
    in chapters 38/31 fell to the generic default tag."""
    assert h.classify_rationale(name) == why


def _o(fy25, fy26, qg, pg):
    return {"imp_fy25": fy25, "imp_fy26": fy26, "qty_growth_pct": qg, "price_growth_pct": pg}


@pytest.mark.parametrize("o,p75,sig", [
    (_o(100, 80, -15, 10), 1000, "A_SCISSORS"),
    (_o(100, 90, -10, 5), 1000, "A_SCISSORS"),             # all three thresholds inclusive
    (_o(100, 91, -10, 5), 1000, "E_NO_CLEAR_SIGNAL"),      # value fell only 9%
    (_o(100, 100, 0, 8), 100, "B_INELASTIC_TARGET"),       # top quartile, price +/-8, value steady
    (_o(100, 100, 0, 8.1), 100, "E_NO_CLEAR_SIGNAL"),
    (_o(100, 94, 0, 0), 90, "E_NO_CLEAR_SIGNAL"),          # value fell 6% (> -5 required)
    (_o(100, 115, 15, 0), 1000, "D_DEEPENING"),
    (_o(100, 300, 50, 0), 100, "B_INELASTIC_TARGET"),      # B is checked before D
    (_o(0, 50, 10, 10), 10, "C_INSUFFICIENT_DATA"),        # no FY25 base
    (_o(100, 120, None, None), 10, "C_INSUFFICIENT_DATA"),
])
def test_classify_signal(o, p75, sig):
    assert h.classify_signal(o, p75) == sig


def _xlsx(path, rows):
    """TradeStat layout: 3 header rows, data from row 4; col B = HS code."""
    wb = openpyxl.Workbook(); ws = wb.active
    for _ in range(3):
        ws.append(["hdr"])
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return str(path)


def test_build_pipeline(tmp_path):
    # value files: (sno, hs, name, fy25, share, fy26, share, growth)
    imp = _xlsx(tmp_path / "imp.xlsx", [
        (1, "29051100", "METHANOL", 100.0, 0, 50.0, 0, -50.0),          # scissors (qty -40%, price up)
        (2, "29152100", "ACETIC ACID", 10.0, 0, 30.0, 0, 200.0),        # deepening
        (3, "38089992", "PESTICIDES NES", 5.0, 0, 6.0, 0, 20.0),        # no qty row -> insufficient
        (4, "39012000", "POLYETHYLENE", 2.0, 0, 3.0, 0, 50.0),          # exports exceed imports -> excluded
        (5, "28000000", "SMALL LINE", 1.0, 0, 0.8, 0, -20.0),           # imports < $1M -> excluded
        (6, "84713010", "LAPTOPS", 900.0, 0, 999.0, 0, 11.0),           # not a chemical chapter
        (7, "2905", "CHAPTER SUBTOTAL", 1.0, 0, 1.0, 0, 0),             # not 8-digit
        (8, None, "blank", 0, 0, 0, 0, 0),
    ])
    exp = _xlsx(tmp_path / "exp.xlsx", [(1, "39012000", "POLYETHYLENE", 1.0, 0, 5.0, 0, 0),
                                        (2, "29152100", "ACETIC ACID", 1.0, 0, 2.0, 0, 0)])
    # qty file: (sno, hs, name, unit, qty25, qty26, qty_growth)
    qty = _xlsx(tmp_path / "qty.xlsx", [(1, "29051100", "METHANOL", "KGS", 1000.0, 600.0, -40.0),
                                        (2, "29152100", "ACETIC ACID", "KGS", 100.0, 250.0, 150.0)])
    res = h.build(imp, exp, qty)
    got = {o["hs"]: o for v in res["by_chapter"].values() for o in v}
    assert set(got) == {"29051100", "29152100", "38089992"}
    m = got["29051100"]
    assert (m["unit_price_fy25"], m["unit_price_fy26"], m["price_growth_pct"]) == (0.1, 0.0833, -16.7)
    assert m["is_signal"] == "E_NO_CLEAR_SIGNAL"     # value -50 and volume -40 but unit price FELL: not scissors
    a = got["29152100"]
    assert (a["deficit_fy26"], a["exp_fy26"], a["is_signal"]) == (28.0, 2.0, "D_DEEPENING")
    assert got["38089992"]["is_signal"] == "C_INSUFFICIENT_DATA"
    assert res["p75_fy26"] == 50.0
    assert set(res["by_chapter"]) == {"29", "38"}
