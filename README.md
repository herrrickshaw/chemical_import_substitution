# Chemical import substitution — HSN-8 signal analysis

Screens India's chemical imports at the 8-digit HS code level for import-substitution
signals. It reads DGCI&S TradeStat commodity-wise Import/Export Value files (US$ million)
and the companion Import Quantity file, keeps the chemical chapters (HS 28, 29, 31, 32, 33,
34, 38, 39) with a material FY26 trade deficit, computes value / volume / unit-price growth
FY24-25 → FY25-26, and tags each line with a signal class and a rule-based qualitative
rationale. Output is a JSON registry (`data/chem_hsn8_signals.json`) and a rendered
chapter-wise HTML registry (`data/chem_hsn8_full_hsn_registry.html`) used in a published article.

> Status: research-only snapshot. Figures come from DGCI&S TradeStat exports (public data);
> the classification is a heuristic screen on a one-financial-year window, not investment
> or policy advice.

New here? See [How to use this repo with AI](HOW_TO_USE_WITH_AI.md).

## What it does
- Parses three TradeStat `.xlsx` exports (import value, export value, import quantity) into HS-8 rows (`load_value_file`, `load_qty_file`).
- Keeps lines in the eight chemical chapters with FY26 imports ≥ US$1M and a deficit (import − export) ≥ US$0.5M (`qualifying_rows`; floor adjustable with `--deficit-floor`).
- Derives FY25/FY26 unit prices and price growth when a quantity row exists (`quantity_fields`).
- Classifies each line as `A_SCISSORS`, `B_INELASTIC_TARGET`, `C_INSUFFICIENT_DATA`, `D_DEEPENING` or `E_NO_CLEAR_SIGNAL` (`classify_signal`).
- Attaches a keyword-rule rationale (fertiliser, petrochemical, agrochemical, dye, pharma API, …) per line (`classify_rationale`).
- Writes `{by_chapter, ch_names, p75_fy26}` as JSON; the committed snapshot holds 827 qualifying codes.

## Quick start
```bash
git clone https://github.com/herrrickshaw/chemical_import_substitution.git
cd chemical_import_substitution
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python scripts/hsn8_signal_analysis.py out.json \
  --import-value "TradeStat-Eidb-Import-Commodity-wise.xlsx" \
  --export-value "TradeStat-Eidb-Export-Commodity-wise.xlsx" \
  --import-qty   "TradeStat-Eidb-Import-Commodity-wise (1).xlsx"
```
Without the `--import-*` flags the script looks for those three filenames in `~/Downloads`
(the paths used for the original run). The positional output path defaults to `chem_hsn8_signals.json`.

## How it works
| Stage | Function (`scripts/hsn8_signal_analysis.py`) | Output |
|---|---|---|
| 1. Load | `load_value_file`, `load_qty_file` | HS-8 rows (rows 1–3 of each sheet are headers; non-8-char codes such as chapter subtotals are skipped) |
| 2. Qualify + enrich | `qualifying_rows` → `quantity_fields`, `classify_rationale` | chemical-chapter lines with deficit, unit prices, rationale |
| 3. Threshold | `percentile_75` | nearest-rank upper-quartile FY26 import value (`p75_fy26`) |
| 4. Signal | `classify_signal` | one signal code per line |
| 5. Group + write | `group_by_chapter`, `write_result` | JSON keyed by 2-digit chapter |

Signal rules (percent growth FY25 → FY26, thresholds are named constants at the top of the module):

| Signal | Rule | Reading |
|---|---|---|
| `A_SCISSORS` | value ≤ −10 % and volume ≤ −10 % and unit price ≥ +5 % | active substitution under way |
| `B_INELASTIC_TARGET` | FY26 value in top quartile, \|price growth\| ≤ 8 %, value > −5 % | PLI / capex candidate |
| `D_DEEPENING` | value ≥ +15 % and volume ≥ +15 % | dependency growing |
| `E_NO_CLEAR_SIGNAL` | none of the above | — |
| `C_INSUFFICIENT_DATA` | no matched quantity row, or zero FY25 base | — |

Stages 2–5 are pure functions of in-memory rows, so they are unit-tested without spreadsheets;
file I/O happens only in stage 1 and `write_result`.

## Repository layout
| Path | Purpose |
|---|---|
| `scripts/hsn8_signal_analysis.py` | the whole pipeline + CLI (single module) |
| `tests/` | pytest unit tests for the pure stages, synthetic-workbook end-to-end test, CLI smoke test |
| `data/chem_hsn8_signals.json` | generated output snapshot (827 codes, `p75_fy26` = 61.94); tracked, not rebuilt by tests |
| `data/chem_hsn8_full_hsn_registry.html` | rendered chapter-wise registry with methodology notes, as embedded in the article; tracked |
| `design/design_v2.css` | design tokens / table styles shared with the HTML registry |
| `HOW_TO_USE_WITH_AI.md` | non-programmer guide to using the repo with an AI assistant |
| `pyproject.toml`, `requirements.txt` | ruff + pytest config; runtime dependency is `openpyxl` only |

## Data sources
| Source | What is taken | Access route | Licence / caveat |
|---|---|---|---|
| DGCI&S TradeStat (Export Import Data Bank), commodity-wise, 8-digit | FY2024-25 and FY2025-26 import value, export value (US$M) and import quantity per HS-8 | manual `.xlsx` download from https://tradestat.commerce.gov.in | Government of India open data; the snapshot files are not committed, only the derived JSON/HTML |

## Configuration
No environment variables or secrets. Inputs and the deficit floor are CLI flags
(`--import-value`, `--export-value`, `--import-qty`, `--deficit-floor`, default 0.5 US$M);
signal thresholds and the chapter list are constants in `scripts/hsn8_signal_analysis.py`.

## Testing
```bash
pip install pytest ruff
pytest -q        # 14 tests
ruff check .
```

## Limitations and known issues
- Only a one-financial-year window (FY25 → FY26); the framework the article applies calls for a 3–5 year trend.
- No Import Penetration Ratio: TradeStat has no domestic production data.
- No Component Shift (Signal C) detection — would need finished-vs-component HSN pairing.
- No country-of-origin / FTA-rerouting check: the commodity-wise export has no partner country.
- No systematic tariff / PLI cross-reference per HSN line; rationale tags are keyword rules only.
- `C_INSUFFICIENT_DATA` also covers lines whose FY25 import value is zero (growth undefined).
- The committed JSON carries two extra per-row fields (`is_label`, `is_conf`) that were added
  when the article was rendered; the script does not emit them.
- Unit prices for very high-tonnage lines round to `0.0` at 4 decimals (US$M per kg); price growth is still computed from the unrounded values.

## Licence
No licence file yet — all rights reserved.
