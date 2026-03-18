from __future__ import annotations

import shutil
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .auth import SESSION_COOKIE, create_session_token, decode_session_token
from .config import DATA_DIR, load_settings
from .data_loader import WorkbookData, load_workbook
from .report_processor import ReportProcessorService, normalize_company_name


app = FastAPI(title="Credit Screenr")
settings = load_settings()
DATA_DIR.mkdir(exist_ok=True)
settings.processed_report_dir.mkdir(parents=True, exist_ok=True)


class DataStore:
    def __init__(self) -> None:
        self.dataset: WorkbookData | None = None

    def refresh(self, workbook_path: Path) -> WorkbookData:
        self.dataset = load_workbook(workbook_path)
        return self.dataset


store = DataStore()
report_service = ReportProcessorService(
    api_key=settings.openai_api_key,
    report_dir=settings.report_dir,
    processed_dir=settings.processed_report_dir,
    summary_workbook_path=settings.report_summary_workbook,
    ca_bundle=settings.openai_ca_bundle or None,
)


def get_current_user(request: Request) -> str:
    token = request.cookies.get(SESSION_COOKIE)
    session = decode_session_token(settings.app_secret, token)
    if not session:
        raise HTTPException(status_code=401, detail="Unauthorized")
    return session.username


def ensure_data_loaded() -> WorkbookData:
    if store.dataset is None:
        workbook_path = settings.workbook_path
        if not workbook_path.exists():
            raise HTTPException(status_code=500, detail=f"Workbook not found: {workbook_path}")
        store.refresh(workbook_path)
    return store.dataset


def enrich_issuers_with_reports(dataset: WorkbookData) -> dict:
    report_records = report_service.snapshot()
    issuers = []
    for row in dataset.issuer_table:
        enriched = dict(row)
        report = report_records.get(normalize_company_name(str(row.get("Issuer", ""))))
        if report:
            enriched["REPORT_SENTIMENT_SCORE"] = report.get("sentiment_score")
            enriched["REPORT_SENTIMENT_LABEL"] = report.get("sentiment_label")
            enriched["REPORT_SENTIMENT_COLOR"] = report.get("sentiment_color")
            enriched["REPORT_SUMMARY_BULLETS"] = report.get("summary_bullets", [])
            enriched["REPORT_DATE"] = report.get("report_date")
            enriched["REPORT_SOURCE_FILE"] = report.get("source_file")
        else:
            enriched["REPORT_SENTIMENT_SCORE"] = None
            enriched["REPORT_SENTIMENT_LABEL"] = ""
            enriched["REPORT_SENTIMENT_COLOR"] = ""
            enriched["REPORT_SUMMARY_BULLETS"] = []
            enriched["REPORT_DATE"] = None
            enriched["REPORT_SOURCE_FILE"] = None
        issuers.append(enriched)
    return {
        "issuers": issuers,
        "filters": dataset.filters,
        "metadata": {**dataset.metadata, "processed_report_count": len(report_records)},
    }


@app.get("/", response_class=FileResponse)
def root() -> FileResponse:
    return FileResponse(Path(__file__).parent / "static" / "index.html")


@app.post("/login")
def login(username: str = Form(...), password: str = Form(...)) -> JSONResponse:
    if username != settings.username or password != settings.password:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    token = create_session_token(settings.app_secret, username)
    response = JSONResponse({"ok": True, "username": username})
    response.set_cookie(
        key=SESSION_COOKIE,
        value=token,
        httponly=True,
        samesite="lax",
        secure=settings.cookie_secure,
        max_age=60 * 60 * 12,
    )
    return response


@app.post("/logout")
def logout() -> JSONResponse:
    response = JSONResponse({"ok": True})
    response.delete_cookie(SESSION_COOKIE)
    return response


@app.get("/api/session")
def session(request: Request) -> JSONResponse:
    token = request.cookies.get(SESSION_COOKIE)
    user = decode_session_token(settings.app_secret, token)
    if not user:
        return JSONResponse({"authenticated": False})
    return JSONResponse({"authenticated": True, "username": user.username})


@app.get("/api/dashboard")
def dashboard(_: str = Depends(get_current_user)) -> JSONResponse:
    dataset = ensure_data_loaded()
    return JSONResponse(enrich_issuers_with_reports(dataset))


@app.get("/api/issuers/{parent_ticker}")
def issuer_detail(parent_ticker: str, _: str = Depends(get_current_user)) -> JSONResponse:
    dataset = ensure_data_loaded()
    rows = [row for row in dataset.instrument_rows if row.get("PARENT_TICKER") == parent_ticker]
    if not rows:
        raise HTTPException(status_code=404, detail="Issuer not found")
    issuer_name = next((row.get("Issuer") for row in dataset.issuer_table if row.get("PARENT_TICKER") == parent_ticker), "")
    report = report_service.snapshot().get(normalize_company_name(str(issuer_name)))
    return JSONResponse({"issuer": parent_ticker, "rows": rows, "report": report})


@app.post("/api/admin/upload")
def admin_upload(workbook: UploadFile = File(...), _: str = Depends(get_current_user)) -> JSONResponse:
    if not workbook.filename or not workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx workbook")

    target_path = DATA_DIR / "Master_File.xlsx"
    with target_path.open("wb") as target:
        shutil.copyfileobj(workbook.file, target)

    dataset = store.refresh(target_path)
    return JSONResponse({"ok": True, "metadata": dataset.metadata})


@app.post("/api/admin/process-reports")
def admin_process_reports(_: str = Depends(get_current_user)) -> JSONResponse:
    try:
        records = report_service.refresh_once()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Report processing failed: {exc}") from exc
    return JSONResponse({"ok": True, "processed_report_count": len(records)})


app.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")
