# Credit Screenr

Minimal authenticated web dashboard for workbook-driven credit screening.

## What it does

- Reads `data/Master_File.xlsx`
- Joins `master`, `spread_px`, and `52w_px` on `ID`
- Groups issuers by `PARENT_TICKER`
- Shows an issuer dashboard with bond-level drilldown
- Supports login and workbook upload refresh
- Scans `Report/` for the latest PDF per company and can cache report summaries and sentiment
- Writes processed report output to `data/report_summaries.xlsx` for dashboard consumption

## Run locally

1. Create and activate a virtual environment.
2. Install dependencies:

```powershell
py -m pip install -r requirements.txt
```

3. Set environment variables or use defaults:

```powershell
$env:APP_SECRET = "change-me"
$env:APP_USERNAME = "admin"
$env:APP_PASSWORD = "change-me"
$env:OPENAI_API_KEY = "sk-..."
```

4. Copy the workbook into `data/Master_File.xlsx`.
5. Start the app:

```powershell
py -m uvicorn app.main:app --reload
```

6. Open `http://127.0.0.1:8000`

## First Render deployment (quick test)

This repo includes a basic `render.yaml` for a first deployment test.

### 1. Push to GitHub

Create a GitHub repo and push this project.

### 2. Create a Render Web Service

- Sign in to Render
- Create a new **Blueprint** or **Web Service** from your GitHub repo
- Render can read the included `render.yaml`

### 3. Set environment variables in Render

At minimum set:

- `APP_SECRET`
- `APP_USERNAME`
- `APP_PASSWORD`

Optional:

- `OPENAI_API_KEY`
- `OPENAI_CA_BUNDLE`

### 4. Deploy

Render will build with:

```bash
pip install -r requirements.txt
```

and start with:

```bash
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

### 5. Test the live URL

- Open the Render URL
- Log in with the credentials you set
- Confirm the dashboard loads
- Test workbook upload

### Important note for first deployment

This first hosted version is fine for testing, but uploaded files and generated report files should not be treated as permanent production storage yet. Later you can move those to persistent cloud storage.

## Process reports

Run report processing separately from the dashboard server:

```powershell
py process_reports.py
```

This reads the latest report per company from `Report/`, calls the OpenAI API, and writes the processed results to `data/report_summaries.xlsx`.

## Notes

- The leading blank/index column in the workbook is dropped automatically.
- The app uses the latest date in `spread_px` as anchor date and the earliest available date as T-90.
- This is the MVP foundation; Bloomberg/BQL refresh automation can be added later.
- Report processing uses the OpenAI Responses API and stores cached JSON in `data/processed_reports`.
- The dashboard itself reads the Excel summary cache and does not make live OpenAI calls during page load.
