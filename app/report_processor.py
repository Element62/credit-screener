from __future__ import annotations

import base64
import json
import re
import ssl
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import certifi


OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses"
REPORT_NAME_RE = re.compile(r"^[^_]+_(?P<company>.+?)_(?P<date>\d{4}-\d{2}-\d{2})_(?P<id>\d+)\.pdf$", re.IGNORECASE)


@dataclass
class ReportRecord:
    company_name: str
    company_key: str
    report_date: str
    source_file: str
    summary_bullets: list[str]
    sentiment_score: int
    sentiment_label: str
    sentiment_color: str
    processed_at: str


def normalize_company_name(value: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
    suffixes = {"inc", "corp", "corporation", "company", "co", "l p", "lp", "l p.", "ltd", "plc", "sarl", "sa", "holdings", "holding"}
    tokens = [token for token in normalized.split() if token not in suffixes]
    return " ".join(tokens)


def sentiment_color(score: int) -> str:
    if score >= 65:
        return "good"
    if score >= 40:
        return "warn"
    return "bad"


def parse_report_filename(path: Path) -> dict[str, str] | None:
    match = REPORT_NAME_RE.match(path.name)
    if not match:
        return None
    company_name = match.group("company").replace("_", " ").strip()
    return {
        "company_name": company_name,
        "company_key": normalize_company_name(company_name),
        "report_date": match.group("date"),
        "source_file": str(path),
    }


def load_cached_reports(processed_dir: Path) -> dict[str, dict[str, Any]]:
    processed_dir.mkdir(parents=True, exist_ok=True)
    records: dict[str, dict[str, Any]] = {}
    for file in processed_dir.glob("*.json"):
        try:
            payload = json.loads(file.read_text(encoding="utf-8"))
        except Exception:
            continue
        company_key = payload.get("company_key")
        if company_key:
            records[company_key] = payload
    return records


def write_report_summary_workbook(records: dict[str, dict[str, Any]], workbook_path: Path) -> None:
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    for record in records.values():
        rows.append(
            {
                "Issuer Name": record.get("company_name"),
                "Bullet Point Summary": "\n".join(record.get("summary_bullets", [])) if record.get("summary_bullets") else "",
                "Sentiment": record.get("sentiment_label") or "",
                "Sentiment Score": record.get("sentiment_score"),
                "Sentiment Color": record.get("sentiment_color") or "",
                "Report Date": record.get("report_date") or "",
                "Source File": record.get("source_file") or "",
                "Company Key": record.get("company_key"),
                "Processing Status": record.get("processing_status") or "",
                "Processing Error": record.get("processing_error") or "",
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(by=["Issuer Name", "Report Date"], ascending=[True, False])
    df.to_excel(workbook_path, index=False)


def load_report_summary_workbook(workbook_path: Path) -> dict[str, dict[str, Any]]:
    if not workbook_path.exists():
        return {}
    df = pd.read_excel(workbook_path)

    def clean_scalar(value: Any) -> Any:
        if pd.isna(value):
            return None
        return value

    records: dict[str, dict[str, Any]] = {}
    for row in df.to_dict(orient="records"):
        issuer_name = str(clean_scalar(row.get("Issuer Name")) or "").strip()
        if not issuer_name:
            continue
        company_key = str(clean_scalar(row.get("Company Key")) or normalize_company_name(issuer_name))
        summary_text = str(clean_scalar(row.get("Bullet Point Summary")) or "")
        bullets = [line.strip() for line in summary_text.splitlines() if line.strip()]
        records[company_key] = {
            "company_name": issuer_name,
            "company_key": company_key,
            "report_date": clean_scalar(row.get("Report Date")),
            "source_file": clean_scalar(row.get("Source File")),
            "summary_bullets": bullets,
            "sentiment_score": int(clean_scalar(row.get("Sentiment Score")) or 0) if clean_scalar(row.get("Sentiment Score")) is not None else None,
            "sentiment_label": clean_scalar(row.get("Sentiment")) or "",
            "sentiment_color": clean_scalar(row.get("Sentiment Color")) or "",
            "processing_status": clean_scalar(row.get("Processing Status")) or "",
            "processing_error": clean_scalar(row.get("Processing Error")) or "",
        }
    return records


def latest_reports_by_company(report_dir: Path) -> dict[str, dict[str, str]]:
    report_dir.mkdir(parents=True, exist_ok=True)
    latest: dict[str, dict[str, str]] = {}
    for path in report_dir.glob("*.pdf"):
        parsed = parse_report_filename(path)
        if not parsed:
            continue
        current = latest.get(parsed["company_key"])
        if not current or (parsed["report_date"], path.stat().st_mtime) > (current["report_date"], Path(current["source_file"]).stat().st_mtime):
            latest[parsed["company_key"]] = parsed
    return latest


def build_prompt(metadata: dict[str, str]) -> str:
    return (
        "You are a credit analyst. Read the attached sell-side or research report and return a concise JSON object. "
        "Focus on credit-relevant takeaways only. "
        "Return exactly this schema: "
        '{"company_name": string, "report_date": string, "summary_bullets": [string, string, string, string, string], '
        '"sentiment_score": integer, "sentiment_label": string}. '
        "Sentiment score must be from 0 to 100 where 0 is very negative credit view, 50 is neutral, and 100 is very positive. "
        "Sentiment label must be one of: Negative, Mixed, Positive. "
        f"Use the company name {metadata['company_name']} and report date {metadata['report_date']} unless the document clearly indicates otherwise."
    )


def _build_ssl_context(ca_bundle: str | None = None) -> ssl.SSLContext:
    cafile = ca_bundle or certifi.where()
    return ssl.create_default_context(cafile=cafile)


def call_openai_for_report(api_key: str, path: Path, metadata: dict[str, str], ca_bundle: str | None = None) -> dict[str, Any]:
    pdf_b64 = base64.b64encode(path.read_bytes()).decode("utf-8")
    body = {
        "model": "gpt-4o-mini",
        "input": [
            {
                "role": "user",
                "content": [
                    {"type": "input_file", "filename": path.name, "file_data": f"data:application/pdf;base64,{pdf_b64}"},
                    {"type": "input_text", "text": build_prompt(metadata)},
                ],
            }
        ],
        "text": {"format": {"type": "json_object"}},
    }
    request = urllib.request.Request(
        OPENAI_RESPONSES_URL,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180, context=_build_ssl_context(ca_bundle)) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"OpenAI request failed: {exc.code} {detail}") from exc

    output_text = payload.get("output_text", "")
    if not output_text:
        raise RuntimeError("OpenAI response did not include output_text")
    return json.loads(output_text)


def persist_report(processed_dir: Path, metadata: dict[str, str], model_output: dict[str, Any]) -> dict[str, Any]:
    processed_dir.mkdir(parents=True, exist_ok=True)
    score = int(model_output.get("sentiment_score", 50))
    record = {
        "company_name": model_output.get("company_name") or metadata["company_name"],
        "company_key": metadata["company_key"],
        "report_date": model_output.get("report_date") or metadata["report_date"],
        "source_file": metadata["source_file"],
        "summary_bullets": [str(item).strip() for item in model_output.get("summary_bullets", []) if str(item).strip()][:5],
        "sentiment_score": max(0, min(100, score)),
        "sentiment_label": model_output.get("sentiment_label") or "Mixed",
        "sentiment_color": sentiment_color(score),
        "processed_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "processing_status": "processed",
        "processing_error": "",
    }
    target = processed_dir / f"{metadata['company_key']}.json"
    target.write_text(json.dumps(record, indent=2), encoding="utf-8")
    return record


def process_latest_reports(
    api_key: str,
    report_dir: Path,
    processed_dir: Path,
    summary_workbook_path: Path,
    ca_bundle: str | None = None,
) -> dict[str, dict[str, Any]]:
    latest = latest_reports_by_company(report_dir)
    cached = load_cached_reports(processed_dir)
    if not api_key:
        write_report_summary_workbook(cached, summary_workbook_path)
        return cached

    for company_key, metadata in latest.items():
        existing = cached.get(company_key)
        if existing and existing.get("report_date") == metadata["report_date"] and existing.get("source_file") == metadata["source_file"]:
            if "processing_status" not in existing:
                existing["processing_status"] = "processed"
            if "processing_error" not in existing:
                existing["processing_error"] = ""
            continue
        try:
            model_output = call_openai_for_report(api_key, Path(metadata["source_file"]), metadata, ca_bundle=ca_bundle)
            cached[company_key] = persist_report(processed_dir, metadata, model_output)
        except Exception as exc:
            cached[company_key] = {
                "company_name": metadata["company_name"],
                "company_key": metadata["company_key"],
                "report_date": metadata["report_date"],
                "source_file": metadata["source_file"],
                "summary_bullets": [],
                "sentiment_score": None,
                "sentiment_label": "",
                "sentiment_color": "",
                "processed_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                "processing_status": "error",
                "processing_error": str(exc),
            }
    write_report_summary_workbook(cached, summary_workbook_path)
    return cached


class ReportProcessorService:
    def __init__(self, api_key: str, report_dir: Path, processed_dir: Path, summary_workbook_path: Path, ca_bundle: str | None = None) -> None:
        self.api_key = api_key
        self.report_dir = report_dir
        self.processed_dir = processed_dir
        self.summary_workbook_path = summary_workbook_path
        self.ca_bundle = ca_bundle

    def refresh_once(self) -> dict[str, dict[str, Any]]:
        return process_latest_reports(
            self.api_key,
            self.report_dir,
            self.processed_dir,
            self.summary_workbook_path,
            ca_bundle=self.ca_bundle,
        )

    def snapshot(self) -> dict[str, dict[str, Any]]:
        return load_report_summary_workbook(self.summary_workbook_path)
