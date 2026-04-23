# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Credit Screenr is a workbook-driven credit screening dashboard. A FastAPI backend parses an Excel workbook (`data/Master_File.xlsx` with sheets `master`, `spread_px`, `52w_px` joined on `ID`), computes issuer-level metrics, and serves them as JSON to a vanilla JavaScript SPA. An optional OpenAI integration processes PDF analyst reports into cached sentiment summaries.

## Commands

```bash
# Install dependencies (use a virtual environment)
py -m pip install -r requirements.txt

# Run dev server
py -m uvicorn app.main:app --reload

# Run report processing (standalone, requires OPENAI_API_KEY)
py process_reports.py
```

There is no test suite, linter, or type checker configured.

## Environment Variables

Required: `APP_SECRET`, `APP_USERNAME`, `APP_PASSWORD`

Optional: `OPENAI_API_KEY`, `OPENAI_CA_BUNDLE`, `APP_COOKIE_SECURE` (set `true` in production), `WORKBOOK_PATH`, `REPORT_DIR`, `PROCESSED_REPORT_DIR`, `REPORT_SUMMARY_WORKBOOK`

Supports `.env` file in project root (loaded via `python-dotenv` in `app/config.py`).

## Architecture

**Backend (FastAPI):**
- `app/main.py` — All routes and the `DataStore` singleton. Workbook is lazy-loaded on first dashboard request and held in memory.
- `app/config.py` — `Settings` dataclass from environment variables.
- `app/auth.py` — HMAC-SHA256 session tokens in HTTP-only cookies (12h TTL).
- `app/data_loader.py` — Reads the Excel workbook, joins sheets, computes per-issuer metrics (weighted avg price/yield, distress tier, seniority gap, maturity wall, dislocation). Returns a `WorkbookData` dataclass.
- `app/report_processor.py` — `ReportProcessorService` sends PDFs to OpenAI Responses API, caches results as JSON in `data/processed_reports/`, writes summaries to `data/report_summaries.xlsx`.

**Frontend (vanilla JS, no build step):**
- `app/static/index.html` — SPA shell with login form and tabbed dashboard.
- `app/static/app.js` — All client logic: fetches data, renders tables, manages filters/sorting, handles tab switching. Uses a single `state` object and direct DOM manipulation with `.innerHTML`.
- `app/static/styles.css` — Full styling including distress-tier color coding.

Static assets are served by FastAPI's `StaticFiles` mount at `/static` and cache-busted via query string on the HTML references.

## Data Flow

1. User uploads or pre-places `data/Master_File.xlsx` (3 sheets)
2. `data_loader.load_workbook()` joins, normalizes, and computes issuer aggregates
3. `DataStore` holds the parsed `WorkbookData` in memory
4. Dashboard endpoint enriches issuers with report sentiment from the JSON cache
5. Frontend fetches JSON, filters/sorts client-side, renders tables

## API Endpoints

All protected endpoints require a session cookie (POST `/login` to obtain).

| Method | Path | Purpose |
|--------|------|---------|
| GET | `/api/dashboard` | Issuer table with filters and metadata |
| GET | `/api/issuer-detail?parent_ticker=` | Bond-level detail for one issuer |
| GET | `/api/price-movers` | Securities with >10pt 3M price change |
| GET | `/api/abnormal-prices` | Securities priced >120 |
| POST | `/api/export/issuers` | Excel download of issuer view |
| POST | `/api/export/detail` | Excel download of detail view |
| POST | `/api/admin/upload` | Replace Master_File.xlsx and reload |
| POST | `/api/admin/process-reports` | Trigger OpenAI report processing |

## Key Domain Concepts

- **Issuer** — grouped by `PARENT_TICKER`, aggregated from individual bonds
- **Distress tiers** — classified by weighted average price: Deep Distressed (<40), Distressed (40-70), Stressed (70-90), Near Par (90-100), Par+ (>100)
- **Screening eligibility** — min $200M face, >1Y maturity, max yield threshold
- **Seniority order** — 1st Lien Secured (1) through Junior Subordinated (8), defined in `data_loader.SENIORITY_ORDER`
- **T-90 pricing** — earliest available date in `spread_px` used as prior comparison point

## Sticky Multi-Row Table Headers

When building multi-row sticky `<thead>` headers, do NOT use `rowspan`. Sticky positioning and `rowspan` interact poorly — cells that span rows misalign their bottom edges when scrolling because the browser doesn't stretch rowspan cells during sticky repositioning.

Instead, use two flat `<tr>` rows:
- **Row 1** (`header-group-row`): group labels (e.g. "Face ($BN)", "Coverage") via `colspan`, with empty `<th></th>` placeholders for non-grouped columns.
- **Row 2**: all individual column labels, including those that would have been `rowspan="2"`.

Use `border-collapse: separate; border-spacing: 0` on the table so that `box-shadow` on sticky `<th>` elements renders correctly (borders get clipped under `border-collapse: collapse`). Use `box-shadow` instead of `border-bottom` for header divider lines — they persist during scroll. Body row borders go on `tbody td`.

## Deployment

Deployed to Render.com via `render.yaml`. File storage is ephemeral on Render (uploads don't persist across deploys).
