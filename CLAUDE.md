# CLAUDE.md

Guidance for Claude Code sessions working in this repository.

## Working in this repo

- Single source module: `scripts/hsn8_signal_analysis.py` (pure stages `qualifying_rows`,
  `quantity_fields`, `percentile_75`, `classify_signal`, `group_by_chapter`; I/O in
  `load_value_file`, `load_qty_file`, `write_result`; CLI in `main`).
- Run: `python scripts/hsn8_signal_analysis.py out.json --import-value … --export-value … --import-qty …`
  (defaults point at the TradeStat filenames in `~/Downloads`).
- Verify: `pytest -q` and `ruff check .` (deps: `pip install -r requirements.txt pytest ruff`).
  Keep `from __future__ import annotations` — the script must stay Python 3.9 compatible.
- `data/*.json` and `data/*.html` are generated snapshots used by a published article: do not
  regenerate or edit them in a refactor; the HTML carries `is_label`/`is_conf` fields the script
  does not emit.

<!-- source-fetch-toolkit:v1 (managed by ~/bin/add_source_fetch_section.py — edit there, not here) -->
## Source fetching for data collection & literature surveys

When a task needs a primary document (statute, regulator data, government release, paper, filing),
fetch the document itself — never cite a search snippet as if it had been read.

- **On the owner's Mac:** `fetch_source.py "<url | doi:… | celex:… | arxiv:…>" --out ./sources`
  (in `~/bin`; full guide `~/bin/SOURCE_FETCH_TOOLKIT.md`). Scheduled/unattended runs use only
  `bash /Users/umashankar/bin/fetch_source.sh "<target>"` (allowlisted) and then Read the `.txt` it reports.
- **Elsewhere (cloud sessions, CI):** follow the same ladder by hand — direct GET with a browser User-Agent →
  headless browser only if the page needs JavaScript → official alternates (DOI → OpenAlex
  `api.openalex.org/works/doi:<doi>` locations → repository copy; EUR-Lex → EU Publications Office
  `publications.europa.eu/resource/celex/<CELEX>`; arXiv pdf/abs) → Wayback Machine `id_` snapshot.
- **Hard rule:** a CAPTCHA, robot check, Cloudflare/Akamai/WAF challenge or login wall means stop on that
  host. Do not bypass it (no TLS impersonation, stealth browsers, identity rotation or challenge solving);
  look for the same document via another route, or report it as unavailable.
- **Record** for every figure: URL read, title/date, page/table, fetch route; mark abstract-only or
  snippet-only sources. Fetch politely (~1 request/second per host, no bulk crawling).
<!-- /source-fetch-toolkit -->
