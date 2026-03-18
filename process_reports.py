from __future__ import annotations

from app.config import load_settings
from app.report_processor import process_latest_reports


def main() -> None:
    settings = load_settings()
    records = process_latest_reports(
        api_key=settings.openai_api_key,
        report_dir=settings.report_dir,
        processed_dir=settings.processed_report_dir,
        summary_workbook_path=settings.report_summary_workbook,
        ca_bundle=settings.openai_ca_bundle or None,
    )
    print(f"Processed report records: {len(records)}")
    print(f"Summary workbook: {settings.report_summary_workbook}")
    failures = [record for record in records.values() if record.get("processing_status") == "error"]
    if failures:
        print(f"Report processing completed with {len(failures)} error record(s).")
        for failure in failures:
            print(f"- {failure.get('company_name')}: {failure.get('processing_error')}")


if __name__ == "__main__":
    main()
