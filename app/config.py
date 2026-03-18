from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_WORKBOOK = DATA_DIR / "Master_File.xlsx"

load_dotenv(BASE_DIR / ".env")


@dataclass(frozen=True)
class Settings:
    app_secret: str
    username: str
    password: str
    workbook_path: Path
    cookie_secure: bool
    openai_api_key: str
    openai_ca_bundle: str
    report_dir: Path
    processed_report_dir: Path
    report_summary_workbook: Path


def load_settings() -> Settings:
    cookie_secure = os.getenv("APP_COOKIE_SECURE", "false").strip().lower() in {"1", "true", "yes", "on"}
    return Settings(
        app_secret=os.getenv("APP_SECRET", "dev-secret-change-me"),
        username=os.getenv("APP_USERNAME", "admin"),
        password=os.getenv("APP_PASSWORD", "admin"),
        workbook_path=Path(os.getenv("WORKBOOK_PATH", DEFAULT_WORKBOOK)),
        cookie_secure=cookie_secure,
        openai_api_key=os.getenv("OPENAI_API_KEY", ""),
        openai_ca_bundle=os.getenv("OPENAI_CA_BUNDLE", ""),
        report_dir=Path(os.getenv("REPORT_DIR", BASE_DIR / "Report")),
        processed_report_dir=Path(os.getenv("PROCESSED_REPORT_DIR", DATA_DIR / "processed_reports")),
        report_summary_workbook=Path(os.getenv("REPORT_SUMMARY_WORKBOOK", DATA_DIR / "report_summaries.xlsx")),
    )
