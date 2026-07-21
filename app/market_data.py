"""Parser for the Market_Data workbook (J.P. Morgan HY & Leveraged Loan snapshot).

The workbook holds five loosely-structured sheets. This module normalizes each
into JSON-friendly dicts consumed by the Market tab. Parsing is defensive:
sheets are located by keyword and sections by row labels rather than fixed
indices, so minor template edits don't break it.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

DATE_RE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")
CHANGE_LABELS = ("Week", "MTD", "YTD")


def _num(v):
    if pd.isna(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _str(v) -> str:
    return "" if pd.isna(v) else str(v).strip()


def _find_sheet(xl: pd.ExcelFile, *keywords: str):
    for name in xl.sheet_names:
        low = name.lower()
        if all(k in low for k in keywords):
            return name
    return None


def _parse_spread_sheet(df: pd.DataFrame) -> dict:
    title = _str(df.iloc[0, 0])
    subtitle = _str(df.iloc[1, 0])
    # Header row: first row with an empty col0 but populated col1.
    header_row = next(
        (i for i in range(len(df)) if pd.isna(df.iloc[i, 0]) and pd.notna(df.iloc[i, 1])),
        3,
    )
    columns = [_str(c) for c in df.iloc[header_row, 1:].tolist() if pd.notna(c)]
    ncols = len(columns)

    latest = None
    changes: list[dict] = []
    source = ""
    for i in range(header_row + 1, len(df)):
        label = df.iloc[i, 0]
        if pd.isna(label):
            continue
        label = str(label).strip()
        if label.lower().startswith("source"):
            source = label
            continue
        values = [_num(df.iloc[i, 1 + j]) for j in range(ncols)]
        if DATE_RE.match(label) and any(v is not None for v in values):
            latest = {"date": label, "values": values}  # keep last => most recent
        elif label in CHANGE_LABELS:
            changes.append({"label": label, "values": values})
    changes.sort(key=lambda c: CHANGE_LABELS.index(c["label"]) if c["label"] in CHANGE_LABELS else 9)
    return {
        "title": title,
        "subtitle": subtitle,
        "columns": columns,
        "latest": latest,
        "changes": changes,
        "source": source,
    }


def _parse_perf_sheet(df: pd.DataFrame) -> dict:
    title = _str(df.iloc[0, 0])
    header_row = next(
        (i for i in range(len(df)) if _str(df.iloc[i, 0]) == "Sector"),
        3,
    )
    week: dict[str, float] = {}
    ytd: dict[str, float] = {}
    source = ""
    for i in range(header_row + 1, len(df)):
        s_l, v_l = df.iloc[i, 0], df.iloc[i, 1]
        s_r, v_r = df.iloc[i, 3], df.iloc[i, 4]
        if pd.notna(s_l):
            name = _str(s_l)
            if name.lower().startswith("source"):
                source = name
            elif _num(v_l) is not None:
                week[name] = _num(v_l)
        if pd.notna(s_r):
            name = _str(s_r)
            if not name.lower().startswith("source") and _num(v_r) is not None:
                ytd[name] = _num(v_r)

    # Merge on sector; keep YTD ordering first, then any week-only names.
    ordered = list(dict.fromkeys(list(ytd.keys()) + list(week.keys())))
    sectors = [
        {
            "sector": name,
            "week": week.get(name),
            "ytd": ytd.get(name),
            "is_index": "jpm" in name.lower() or "index" in name.lower(),
        }
        for name in ordered
    ]
    sectors.sort(key=lambda x: (x["ytd"] is None, -(x["ytd"] or 0)))
    return {"title": title, "sectors": sectors, "source": source}


def _parse_defaults_sheet(df: pd.DataFrame) -> dict:
    header_row = next(
        (i for i in range(len(df)) if _str(df.iloc[i, 0]) == "Date"),
        3,
    )
    rows: list[dict] = []
    source = ""
    note = ""
    for i in range(header_row + 1, len(df)):
        first = df.iloc[i, 0]
        if pd.isna(first):
            continue
        text = str(first).strip()
        if text.lower().startswith("source"):
            source = text
            continue
        if text.startswith("*"):
            note = text
            continue
        dt = pd.to_datetime(first, errors="coerce")
        date_str = dt.strftime("%d-%b-%y") if pd.notna(dt) else text
        rows.append(
            {
                "date": date_str,
                "issuer": _str(df.iloc[i, 1]),
                "bonds": _num(df.iloc[i, 2]),
                "loans": _num(df.iloc[i, 3]),
                "total": _num(df.iloc[i, 4]),
                "industry": _str(df.iloc[i, 5]),
                "action": _str(df.iloc[i, 6]),
            }
        )
    return {"rows": rows, "source": source, "note": note}


def load_market_data(path: Path | str) -> dict:
    xl = pd.ExcelFile(path)

    def read(name):
        return pd.read_excel(xl, sheet_name=name, header=None)

    loan_spreads_name = _find_sheet(xl, "loan", "spread")
    hy_spreads_name = _find_sheet(xl, "hy", "spread") or _find_sheet(xl, "high yield", "spread")
    hy_perf_name = _find_sheet(xl, "hy", "performance") or _find_sheet(xl, "high yield", "performance")
    loan_perf_name = _find_sheet(xl, "loan", "performance")
    defaults_name = _find_sheet(xl, "default")

    hy_spreads = _parse_spread_sheet(read(hy_spreads_name)) if hy_spreads_name else None
    loan_spreads = _parse_spread_sheet(read(loan_spreads_name)) if loan_spreads_name else None

    as_of = ""
    for block in (hy_spreads, loan_spreads):
        if block and block.get("latest"):
            as_of = block["latest"].get("date", "")
            break

    return {
        "as_of": as_of,
        "hy_spreads": hy_spreads,
        "loan_spreads": loan_spreads,
        "hy_perf": _parse_perf_sheet(read(hy_perf_name)) if hy_perf_name else None,
        "loan_perf": _parse_perf_sheet(read(loan_perf_name)) if loan_perf_name else None,
        "defaults": _parse_defaults_sheet(read(defaults_name)) if defaults_name else None,
    }
