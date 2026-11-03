"""
Chemical import-substitution HSN-8 signal analysis.

Inputs: DGCI&S TradeStat commodity-wise Import/Export Value files (US$M) and
the companion Import Quantity file, all at the 8-digit HS code level.

Produces, per qualifying HSN-8 code (FY26 import - FY26 export >= $0.5M):
  - value/volume/price growth FY24-25 -> FY25-26
  - a signal classification against the import-substitution framework:
      A_SCISSORS          value & volume both falling, price rising      -> active substitution
      B_INELASTIC_TARGET  top-quartile value, price flat, value steady   -> PLI/capex candidate
      D_DEEPENING         value & volume both rising fast                -> dependency growing
      E_NO_CLEAR_SIGNAL   no clear 1-FY directional pattern
      C_INSUFFICIENT_DATA no matched quantity row
  - a rule-based qualitative rationale tag (fertiliser/petrochem/agrochem/etc.)

Known limits (see the published article's methodology blockquote for the
user-facing version of this list):
  - only a 1-FY window, not the 3-5yr trend the framework calls for
  - no Import Penetration Ratio (needs domestic production data, not in TradeStat)
  - no Component Shift (Signal C) detection (needs finished-vs-component HSN pairing)
  - no country-of-origin / FTA-rerouting check (TradeStat commodity-wise has no partner country)
  - no systematic tariff/PLI cross-reference per HSN line (only qualitative keyword tags)

Pipeline: load_value_file / load_qty_file (I/O) -> qualifying_rows + classify_signal
(pure) -> group_by_chapter -> write_result (I/O). `build()` wires the stages together.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
from collections.abc import Iterable, Iterator, Sequence
from pathlib import Path
from typing import Any

import openpyxl

log = logging.getLogger(__name__)

# --- Configuration -----------------------------------------------------------

CHEM_CHAPTERS = ["28", "29", "31", "32", "33", "34", "38", "39"]

CH_NAMES = {
    "28": "Inorganic Chemicals & Fertiliser Intermediates",
    "29": "Organic Chemicals",
    "31": "Fertilisers",
    "32": "Dyes, Pigments, Tanning & Colouring Extracts",
    "33": "Essential Oils, Cosmetics & Perfumery",
    "34": "Soaps, Waxes, Polishes & Surface-Active Agents",
    "38": "Miscellaneous Chemical Products",
    "39": "Plastics & Articles Thereof",
}

# TradeStat export layout: three header rows, then one row per HS line.
TRADESTAT_FIRST_DATA_ROW = 4
HS8_LEN = 8

# Qualification: a line must import >= MIN_IMPORT_MUSD and show a trade deficit
# of at least DEFAULT_DEFICIT_FLOOR_MUSD (both US$ million, FY26).
MIN_IMPORT_MUSD = 1
DEFAULT_DEFICIT_FLOOR_MUSD = 0.5

# Signal thresholds (percent growth FY25 -> FY26).
SCISSORS_VALUE_MAX_PCT = -10  # value falling at least this much ...
SCISSORS_QTY_MAX_PCT = -10  # ... and volume falling at least this much ...
SCISSORS_PRICE_MIN_PCT = 5  # ... while unit price rises at least this much
INELASTIC_PRICE_BAND_PCT = 8  # |price growth| within this band = "flat"
INELASTIC_VALUE_MIN_PCT = -5  # value not falling faster than this = "steady"
DEEPENING_VALUE_MIN_PCT = 15
DEEPENING_QTY_MIN_PCT = 15
TOP_QUARTILE = 0.75

SIGNAL_SCISSORS = "A_SCISSORS"
SIGNAL_INELASTIC = "B_INELASTIC_TARGET"
SIGNAL_INSUFFICIENT = "C_INSUFFICIENT_DATA"
SIGNAL_DEEPENING = "D_DEEPENING"
SIGNAL_NONE = "E_NO_CLEAR_SIGNAL"

# Default input locations used by the original one-off run (owner's Downloads).
DEFAULT_DOWNLOADS = Path.home() / "Downloads"
DEFAULT_IMPORT_VALUE = DEFAULT_DOWNLOADS / "TradeStat-Eidb-Import-Commodity-wise.xlsx"
DEFAULT_EXPORT_VALUE = DEFAULT_DOWNLOADS / "TradeStat-Eidb-Export-Commodity-wise.xlsx"
DEFAULT_IMPORT_QTY = DEFAULT_DOWNLOADS / "TradeStat-Eidb-Import-Commodity-wise (1).xlsx"
DEFAULT_OUT_PATH = "chem_hsn8_signals.json"

RATIONALE_RULES = [
    (r'\b(UREA|DAP|DIAMONM|MURIATE|POTASH|POTASSIUM CHLORIDE|SUPERPHOSPHAT|ROCK PHOSPHATE|SULPHATE OF POTASH|NPK|FERTLSR|FERTILIS)\b',
     'Fertiliser feedstock/finished — India lacks phosphate rock & potash reserves; imports are structural, not import-substitutable without mining tie-ups (Morocco/Jordan/Canada/Belarus).'),
    (r'\b(GOLD|PLATINUM|RHODIUM|PALLADIUM|SILVER|NOBLE METAL)\b',
     'Precious/noble metal — natural-resource constrained; India has negligible reserves, substitution means recycling/urban-mining not manufacturing.'),
    (r'\b(POLYPROPYLENE|POLYETHYLENE|PVC|POLYCARBONAT|POLYMER|COPOLYMER|RESIN|ETHYLENE|PROPYLENE)\b',
     'Base petrochemical/polymer — cracker capacity gap; substitutable via new naphtha/gas cracker + downstream polymer trains (PLI/PCPIR candidate).'),
    (r'\b(BENZENE|STYRENE|TOLUENE|XYLENE|METHANOL|ALCOHOL|ACETIC|FORMALDEHYDE|ACRYLIC|PHENOL)\b',
     'Basic organic intermediate — aromatics/methanol value chain; substitutable with refinery-integrated petrochemical capacity.'),
    (r'\b(INSECTICID|FUNGICID|HERBICID|PESTICID|AGROCHEM|CARTAP|MANCOZEB)\b',
     'Agrochemical technical/formulation — patent-expired generics; substitutable via technical-grade manufacturing (China+1 opportunity).'),
    (r'\b(DYE|PIGMENT|COLOUR|COLORANT|PHATHALOCYANINE|PATHALOCYANINE)\b',
     'Specialty dye/pigment — process-technology & environmental-compliance gap (dye intermediates are pollution-intensive); substitutable with clean-tech investment.'),
    (r'\b(VITAMIN|ANTIBIOTIC|BULK DRUG|API|PHARMA|STEROID|HORMONE)\b',
     'Pharma API/intermediate — precursor-chemistry dependency on China; strategic substitution priority (PLI-for-APIs scheme already targets this).'),
    (r'\b(ACID)\b',
     'Industrial acid — commodity intermediate; substitutable via capacity addition, often import-parity priced so investment case is thin unless anti-dumping applies.'),
    (r'\b(OIL|ESSENTIAL OIL|PERFUM|FRAGRANC|AROMATIC)\b',
     'Essential oil/fragrance compound — agro-botanical or specialty-synthesis input; partly substitutable via natural-farming linkages (lavender/lemongrass missions).'),
    (r'\b(ENZYME|CATALYST|REAGENT|DIAGNOSTIC)\b',
     'Catalyst/enzyme/diagnostic reagent — high-IP specialty chemistry; substitution needs licensed tech transfer, not commodity capacity.'),
    (r'\b(WAX|SOAP|SURFACE|SURFACTANT|POLISH)\b',
     'Surfactant/specialty formulation — oleochemical feedstock or IP-licensed formulation; moderate substitution potential via oleochemical integration.'),
    (r'\b(RUBBER|LATEX)\b',
     'Rubber chemical/latex-linked input — natural rubber deficit compounds the synthetic-chemical gap.'),
    (r'\b(CARBON BLACK|ACTIVATED CARBON)\b',
     'Carbon-based industrial input — feedstock (CBFS/coal) linked; substitutable with refinery-integrated carbon black capacity.'),
]
RATIONALE_DEFAULT = 'Chemical intermediate/speciality — deficit driven by feedstock, scale-economics or technology gap; needs case-by-case investment screening.'

_COMPILED_RULES = [(re.compile(pat, re.IGNORECASE), why) for pat, why in RATIONALE_RULES]

Row = dict[str, Any]


# --- Pure helpers ------------------------------------------------------------


def classify_rationale(name: str) -> str:
    """Return the first matching qualitative rationale for an HS line description."""
    for rx, why in _COMPILED_RULES:
        if rx.search(name):
            return why
    return RATIONALE_DEFAULT


def pct_change(new: float, old: float) -> float | None:
    """Percent change old -> new; None when `old` is zero/missing (undefined growth)."""
    if not old:
        return None
    return (new - old) / old * 100


def percentile_75(values: Iterable[float]) -> float:
    """Nearest-rank upper-quartile cut-off used for the 'top-quartile value' test (0 if empty)."""
    vals = sorted(values)
    return vals[int(len(vals) * TOP_QUARTILE)] if vals else 0


def classify_signal(o: Row, p75_value: float) -> str:
    """Map one enriched HS row to a signal code (see module docstring for the legend)."""
    vg = pct_change(o["imp_fy26"], o["imp_fy25"])
    qg = o.get("qty_growth_pct")
    pg = o.get("price_growth_pct")

    if vg is None or qg is None or pg is None:
        return SIGNAL_INSUFFICIENT
    if vg <= SCISSORS_VALUE_MAX_PCT and qg <= SCISSORS_QTY_MAX_PCT and pg >= SCISSORS_PRICE_MIN_PCT:
        return SIGNAL_SCISSORS
    if o["imp_fy26"] >= p75_value and abs(pg) <= INELASTIC_PRICE_BAND_PCT and vg > INELASTIC_VALUE_MIN_PCT:
        return SIGNAL_INELASTIC
    if vg >= DEEPENING_VALUE_MIN_PCT and qg >= DEEPENING_QTY_MIN_PCT:
        return SIGNAL_DEEPENING
    return SIGNAL_NONE


def quantity_fields(value_row: Row, qty_row: Row | None) -> Row:
    """Volume and unit-price fields for a value row; all None when no usable quantity row."""
    if qty_row and qty_row["qty25"] and qty_row["qty26"]:
        up25 = value_row["fy25"] / qty_row["qty25"]
        up26 = value_row["fy26"] / qty_row["qty26"]
        price_growth = pct_change(up26, up25)
        return {
            "qty25": qty_row["qty25"],
            "qty26": qty_row["qty26"],
            "unit": qty_row["unit"],
            "qty_growth_pct": qty_row["qty_growth"],
            "unit_price_fy25": round(up25, 4),
            "unit_price_fy26": round(up26, 4),
            "price_growth_pct": round(price_growth, 1) if price_growth is not None else None,
        }
    return {
        "qty25": None,
        "qty26": None,
        "unit": None,
        "qty_growth_pct": None,
        "unit_price_fy25": None,
        "unit_price_fy26": None,
        "price_growth_pct": None,
    }


def qualifying_rows(
    imports: Sequence[Row],
    exports: Sequence[Row],
    import_qty: dict[str, Row],
    deficit_floor_musd: float = DEFAULT_DEFICIT_FLOOR_MUSD,
) -> list[Row]:
    """Filter import lines to chemical chapters with a material deficit and enrich them."""
    exp_map = {r["hs"]: r for r in exports}
    out: list[Row] = []
    for r in imports:
        ch = r["hs"][:2]
        if ch not in CHEM_CHAPTERS:
            continue
        e = exp_map.get(r["hs"], {"fy26": 0})
        deficit = r["fy26"] - e.get("fy26", 0)
        if r["fy26"] < MIN_IMPORT_MUSD or deficit < deficit_floor_musd:
            continue
        o: Row = {
            "hs": r["hs"],
            "chapter": ch,
            "name": r["name"],
            "imp_fy25": r["fy25"],
            "imp_fy26": r["fy26"],
            "imp_growth": r["growth"],
            "exp_fy26": e.get("fy26", 0),
            "deficit_fy26": round(deficit, 1),
            "rationale": classify_rationale(r["name"]),
        }
        o.update(quantity_fields(r, import_qty.get(r["hs"])))
        out.append(o)
    return out


def group_by_chapter(rows: Iterable[Row]) -> dict[str, list[Row]]:
    """Group rows by their 2-digit HS chapter, preserving input order."""
    by_chapter: dict[str, list[Row]] = {}
    for o in rows:
        by_chapter.setdefault(o["chapter"], []).append(o)
    return by_chapter


# --- I/O ---------------------------------------------------------------------


def _hs8_rows(path: str | Path) -> Iterator[tuple]:
    """Yield raw TradeStat sheet rows whose column B holds an 8-character HS code."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    for r in ws.iter_rows(min_row=TRADESTAT_FIRST_DATA_ROW, values_only=True):
        if not r[1]:
            continue
        if len(str(r[1]).strip()) != HS8_LEN:
            continue
        yield r


def load_value_file(path: str | Path) -> list[Row]:
    """Read a TradeStat Import/Export *Value* export (US$M): hs, name, fy25, fy26, growth."""
    return [
        {
            "hs": str(r[1]).strip(),
            "name": (r[2] or "").strip(),
            "fy25": r[3] or 0,
            "fy26": r[5] or 0,
            "growth": r[7],
        }
        for r in _hs8_rows(path)
    ]


def load_qty_file(path: str | Path) -> dict[str, Row]:
    """Read a TradeStat Import *Quantity* export keyed by HS-8: unit, qty25, qty26, qty_growth."""
    return {
        str(r[1]).strip(): {"unit": r[3], "qty25": r[4] or 0, "qty26": r[5] or 0, "qty_growth": r[6]}
        for r in _hs8_rows(path)
    }


def build(
    import_value_xlsx: str | Path,
    export_value_xlsx: str | Path,
    import_qty_xlsx: str | Path,
    deficit_floor_musd: float = DEFAULT_DEFICIT_FLOOR_MUSD,
) -> dict[str, Any]:
    """Run the full analysis; returns {'by_chapter', 'ch_names', 'p75_fy26'}."""
    imp = load_value_file(import_value_xlsx)
    exp = load_value_file(export_value_xlsx)
    imp_qty = load_qty_file(import_qty_xlsx)
    log.info("loaded %d import, %d export, %d quantity HS-8 lines", len(imp), len(exp), len(imp_qty))

    qualifying = qualifying_rows(imp, exp, imp_qty, deficit_floor_musd)
    p75 = percentile_75(o["imp_fy26"] for o in qualifying)
    for o in qualifying:
        o["is_signal"] = classify_signal(o, p75)

    return {"by_chapter": group_by_chapter(qualifying), "ch_names": CH_NAMES, "p75_fy26": p75}


def write_result(result: dict[str, Any], out_path: str | Path) -> int:
    """Write the result JSON; returns the number of qualifying codes written."""
    with open(out_path, "w") as fh:
        json.dump(result, fh)
    return sum(len(v) for v in result["by_chapter"].values())


# --- CLI ---------------------------------------------------------------------


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Chemical import-substitution HSN-8 signal analysis.")
    p.add_argument("out_path", nargs="?", default=DEFAULT_OUT_PATH, help="output JSON (default: %(default)s)")
    p.add_argument("--import-value", default=DEFAULT_IMPORT_VALUE, help="TradeStat import value xlsx")
    p.add_argument("--export-value", default=DEFAULT_EXPORT_VALUE, help="TradeStat export value xlsx")
    p.add_argument("--import-qty", default=DEFAULT_IMPORT_QTY, help="TradeStat import quantity xlsx")
    p.add_argument(
        "--deficit-floor",
        type=float,
        default=DEFAULT_DEFICIT_FLOOR_MUSD,
        help="minimum FY26 import-export deficit in US$M (default: %(default)s)",
    )
    return p


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    result = build(args.import_value, args.export_value, args.import_qty, args.deficit_floor)
    total = write_result(result, args.out_path)
    print(f"{total} qualifying codes written to {args.out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
