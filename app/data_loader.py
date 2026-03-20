from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pandas as pd


SENIORITY_ORDER = {
    "1st Lien Secured": 1,
    "2nd Lien Secured": 2,
    "Sr Secured": 3,
    "Secured": 4,
    "Sr Unsecured": 5,
    "Senior Unsecured": 5,
    "Subordinated Unsecured": 6,
    "Subordinated": 7,
    "Junior Subordinated": 8,
}


@dataclass
class WorkbookData:
    issuer_table: list[dict]
    instrument_rows: list[dict]
    filters: dict
    metadata: dict


def _make_unique_columns(columns: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    resolved: list[str] = []
    for raw_col in columns:
        col = str(raw_col).strip()
        if not col or col.lower().startswith("unnamed:"):
            col = ""
        if col == "":
            resolved.append("")
            continue
        counts[col] = counts.get(col, 0) + 1
        resolved.append(col if counts[col] == 1 else f"{col}_{counts[col]}")
    return resolved


def _clean_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = _make_unique_columns([str(c) for c in df.columns])
    if len(df.columns) > 0 and df.columns[0] == "":
        df = df.iloc[:, 1:]
    return df


def _to_date_string(series: pd.Series, date_system: str = "excel") -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        converted = pd.to_datetime(series, errors="coerce")
    else:
        numeric = pd.to_numeric(series, errors="coerce")
        if date_system == "excel" and numeric.notna().mean() > 0.8:
            converted = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
        else:
            converted = pd.to_datetime(series, errors="coerce")
    return converted.dt.strftime("%Y-%m-%d")


def _distress_tier(px_val: float | int | None) -> str:
    if pd.isna(px_val):
        return ""
    if px_val < 60:
        return "Deep Distressed"
    if px_val < 80:
        return "Distressed"
    if px_val < 90:
        return "Stressed"
    if px_val <= 100:
        return "Par-Adjacent"
    return "Premium"


def _gap_label(gap: float | int | None) -> str:
    if pd.isna(gap):
        return ""
    if gap < 0:
        return "Inverted"
    if gap < 5:
        return "Normal"
    if gap < 15:
        return "Elevated"
    if gap < 25:
        return "Recovery Cliff"
    return "Severe"


def _is_secured(rank_str: str | None) -> bool:
    rank = str(rank_str or "").lower()
    return "secured" in rank and "unsecured" not in rank


def _normalize_master(df_master: pd.DataFrame) -> pd.DataFrame:
    rename_map = {"CPN_2": "CPN_VALUE", "CPN": "CPN_TYPE"}
    df = df_master.rename(columns=rename_map).copy()
    for col in ["AMT_OUTSTANDING", "CPN_VALUE"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "MATURITY" in df.columns:
        df["MATURITY"] = _to_date_string(df["MATURITY"])
    return df


def _normalize_spread(df_spread: pd.DataFrame) -> pd.DataFrame:
    df = df_spread.copy()
    for col in ["PX_MID", "YIELD", "OAS", "I_SPREAD", "Z_SPREAD"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["DATE"] = _to_date_string(df["DATE"])
    return df


def _normalize_52w(df_52w: pd.DataFrame) -> pd.DataFrame:
    df = df_52w.copy()
    for col in ["PX_HIGH_52W", "PX_LOW_52W"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["WINDOW_START"] = _to_date_string(df["WINDOW_START"])
    df["WINDOW_END"] = _to_date_string(df["WINDOW_END"])
    df["DATE_OF_HIGH"] = _to_date_string(df["DATE_OF_HIGH"])
    df["DATE_OF_LOW"] = _to_date_string(df["DATE_OF_LOW"])
    return df


def _issuer_metrics(df: pd.DataFrame, anchor_date: str) -> pd.DataFrame:
    today = pd.Timestamp(anchor_date)
    max_known_rank = max(SENIORITY_ORDER.values())

    def compute(group: pd.DataFrame) -> pd.Series:
        face = pd.to_numeric(group["AMT_OUTSTANDING"], errors="coerce")
        px = pd.to_numeric(group["PX_MID"], errors="coerce")
        bond_yield = pd.to_numeric(group["YIELD"], errors="coerce") if "YIELD" in group.columns else pd.Series(index=group.index, dtype="float64")
        maturity = pd.to_datetime(group["MATURITY"], errors="coerce")
        rank_text = group["PAYMENT_RANK"].fillna("").astype(str).str.lower()
        eligible_metric_mask = face.notna() & (face >= 200_000_000) & (maturity > today + pd.DateOffset(years=1)) & ~rank_text.str.contains("subordinated")
        face_sum = face.sum()
        px_mask = px.notna() & (px > 0) & eligible_metric_mask
        px_face_sum = face[px_mask].sum()
        wtavg_px = (px[px_mask] * face[px_mask]).sum() / px_face_sum if px_face_sum > 0 else pd.NA
        yield_mask = bond_yield.notna() & (bond_yield > 0) & eligible_metric_mask
        yield_face_sum = face[yield_mask].sum()
        wtavg_yield = (bond_yield[yield_mask] * face[yield_mask]).sum() / yield_face_sum if yield_face_sum > 0 else pd.NA
        max_yield = bond_yield[yield_mask].max() if yield_mask.any() else pd.NA
        secured_face = face[group["_IS_SECURED"]].sum()
        secured_pct = secured_face / face_sum * 100 if face_sum > 0 else pd.NA

        ranks = group["PAYMENT_RANK"].fillna("").astype(str)
        rank_order = ranks.map(SENIORITY_ORDER).fillna(max_known_rank + 1)
        rank_lower = ranks.str.lower()
        first_lien_mask = rank_lower.eq("1st lien secured") & px.notna() & (px > 0)
        senior_unsecured_mask = rank_lower.isin(["sr unsecured", "senior unsecured"]) & px.notna() & (px > 0)
        if first_lien_mask.any() and senior_unsecured_mask.any():
            px_most_senior = px[first_lien_mask].median()
            px_most_junior = px[senior_unsecured_mask].median()
            seniority_gap = px_most_senior - px_most_junior
        else:
            seniority_gap = pd.NA

        def wall(start: pd.Timestamp, end: pd.Timestamp) -> tuple[float, float]:
            mask = maturity.between(start, end)
            amount_mm = face[mask].sum() / 1e6
            pct = amount_mm / (face_sum / 1e6) * 100 if face_sum > 0 else 0
            return amount_mm, pct

        wall_lt1y_mm, wall_lt1y_pct = wall(today, today + pd.DateOffset(years=1))
        wall_1_3y_mm, wall_1_3y_pct = wall(today + pd.DateOffset(years=1), today + pd.DateOffset(years=3))
        wall_3_5y_mm, wall_3_5y_pct = wall(today + pd.DateOffset(years=3), today + pd.DateOffset(years=5))
        issuer_oas_delta = (pd.to_numeric(group["OAS_DELTA"], errors="coerce") * face).sum() / face_sum if face_sum > 0 else pd.NA

        return pd.Series(
            {
                "ISSUER_FACE_MM": face_sum / 1e6,
                "ISSUER_WTAVG_PX": wtavg_px,
                "ISSUER_WTAVG_YIELD": wtavg_yield,
                "ISSUER_MAX_YIELD": max_yield,
                "ISSUER_DIST_TO_PAR": 100.0 - wtavg_px if pd.notna(wtavg_px) else pd.NA,
                "ISSUER_DISLOCATION_MM": group["DISLOCATION_MM"].sum(),
                "ISSUER_DISLOCATION_52W_MM": group["DISLOCATION_52W_MM"].sum(),
                "ISSUER_DISLOCATION_52W_SEC_MM": group.loc[group["_IS_SECURED"], "DISLOCATION_52W_MM"].sum(),
                "ISSUER_DISLOCATION_52W_UNSEC_MM": group.loc[~group["_IS_SECURED"], "DISLOCATION_52W_MM"].sum(),
                "SECURED_PCT": secured_pct,
                "SENIORITY_GAP": seniority_gap,
                "GAP_SIGNAL": _gap_label(seniority_gap),
                "ISSUER_OAS_DELTA": issuer_oas_delta,
                "WALL_LT1Y_MM": wall_lt1y_mm,
                "WALL_LT1Y_PCT": wall_lt1y_pct,
                "WALL_1_3Y_MM": wall_1_3y_mm,
                "WALL_1_3Y_PCT": wall_1_3y_pct,
                "WALL_3_5Y_MM": wall_3_5y_mm,
                "WALL_3_5Y_PCT": wall_3_5y_pct,
            }
        )

    return df.groupby("PARENT_TICKER", dropna=False).apply(compute, include_groups=False).reset_index()


def _wall_label(mm: float | int | None, pct: float | int | None) -> str:
    if pd.isna(mm) or not mm:
        return "-"
    return f"${mm / 1000:.2f}B | {pct:.0f}%"


def _records_from_df(df: pd.DataFrame) -> list[dict]:
    cleaned = df.astype(object).where(pd.notna(df), None)
    records = cleaned.to_dict(orient="records")
    for record in records:
        for key, value in list(record.items()):
            if isinstance(value, float):
                record[key] = round(value, 4)
    return records


def _load_xlsx_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    ns_main = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ns_rel_doc = {"rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    ns_pkg = {"pkg": "http://schemas.openxmlformats.org/package/2006/relationships"}

    with ZipFile(path) as zf:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst_root.findall("main:si", ns_main):
                parts = [node.text or "" for node in si.findall(".//main:t", ns_main)]
                shared_strings.append("".join(parts))

        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall("pkg:Relationship", ns_pkg)
        }

        target = None
        for sheet in workbook_root.findall("main:sheets/main:sheet", ns_main):
            if sheet.attrib.get("name") == sheet_name:
                rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target = rel_map.get(rid)
                break
        if not target:
            raise KeyError(f"Sheet not found: {sheet_name}")

        sheet_root = ET.fromstring(zf.read(f"xl/{target}"))
        rows_data: list[list[str | None]] = []

        def col_index(cell_ref: str) -> int:
            letters = "".join(ch for ch in cell_ref if ch.isalpha())
            result = 0
            for ch in letters:
                result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
            return result - 1

        for row in sheet_root.findall("main:sheetData/main:row", ns_main):
            cells = row.findall("main:c", ns_main)
            row_values: dict[int, str | None] = {}
            max_col = -1
            for cell in cells:
                ref = cell.attrib.get("r", "")
                idx = col_index(ref) if ref else max_col + 1
                max_col = max(max_col, idx)
                cell_type = cell.attrib.get("t")
                value = None
                if cell_type == "inlineStr":
                    node = cell.find("main:is/main:t", ns_main)
                    value = node.text if node is not None else None
                else:
                    node = cell.find("main:v", ns_main)
                    if node is not None:
                        raw = node.text
                        if cell_type == "s" and raw is not None:
                            value = shared_strings[int(raw)]
                        else:
                            value = raw
                row_values[idx] = value
            if max_col >= 0:
                rows_data.append([row_values.get(i) for i in range(max_col + 1)])

    if not rows_data:
        return pd.DataFrame()
    header = rows_data[0]
    data = rows_data[1:]
    return pd.DataFrame(data, columns=header)


def load_workbook(path: Path) -> WorkbookData:
    df_master = _normalize_master(_clean_sheet(_load_xlsx_sheet(path, "master")))
    df_spread = _normalize_spread(_clean_sheet(_load_xlsx_sheet(path, "spread_px")))
    df_52w = _normalize_52w(_clean_sheet(_load_xlsx_sheet(path, "52w_px")))

    df_master = df_master[pd.to_numeric(df_master["AMT_OUTSTANDING"], errors="coerce") > 0].copy()
    spread_dates = sorted(d for d in df_spread["DATE"].dropna().unique().tolist())
    anchor_date = spread_dates[-1] if spread_dates else None
    t90_date = spread_dates[0] if len(spread_dates) > 1 else anchor_date

    anchor_rename = {"PX_MID": "PX_MID", "OAS": "OAS", "I_SPREAD": "I_SPREAD", "Z_SPREAD": "Z_SPREAD"}
    t90_rename = {"PX_MID": "PX_MID_T90", "OAS": "OAS_T90", "I_SPREAD": "I_SPREAD_T90", "Z_SPREAD": "Z_SPREAD_T90"}
    if "YIELD" in df_spread.columns:
        anchor_rename["YIELD"] = "YIELD"
        t90_rename["YIELD"] = "YIELD_T90"

    df_anchor = df_spread[df_spread["DATE"] == anchor_date].drop(columns=["DATE"]).rename(columns=anchor_rename)
    df_t90 = df_spread[df_spread["DATE"] == t90_date].drop(columns=["DATE"]).rename(columns=t90_rename)
    df = (
        df_master.merge(df_anchor, on="ID", how="left")
        .merge(df_t90, on="ID", how="left")
        .merge(df_52w[["ID", "PX_HIGH_52W", "DATE_OF_HIGH", "PX_LOW_52W", "DATE_OF_LOW"]], on="ID", how="left")
    )

    px = pd.to_numeric(df["PX_MID"], errors="coerce")
    face = pd.to_numeric(df["AMT_OUTSTANDING"], errors="coerce")
    px_high = pd.to_numeric(df["PX_HIGH_52W"], errors="coerce")
    px_low = pd.to_numeric(df["PX_LOW_52W"], errors="coerce")
    oas = pd.to_numeric(df["OAS"], errors="coerce")
    oas_t90 = pd.to_numeric(df["OAS_T90"], errors="coerce")

    df["DIST_TO_PAR"] = 100.0 - px
    df["DISLOCATION_MM"] = (df["DIST_TO_PAR"] / 100.0) * (face / 1e6)
    df["DIST_TO_52W_HIGH"] = px_high - px
    df["DIST_TO_52W_LOW"] = px - px_low
    df["DISLOCATION_52W_MM"] = (df["DIST_TO_52W_HIGH"] / 100.0) * (face / 1e6)
    df["OAS_DELTA"] = oas - oas_t90
    df["DISTRESS_TIER"] = px.apply(_distress_tier)
    df["_IS_SECURED"] = df["PAYMENT_RANK"].apply(_is_secured)

    issuer_agg = _issuer_metrics(df, anchor_date) if anchor_date else pd.DataFrame(columns=["PARENT_TICKER"])
    tranche_count = df.groupby("PARENT_TICKER", dropna=False)["ID"].count().rename("# Tranches").reset_index()
    if anchor_date:
        with_mat = df.assign(MAT_DT=pd.to_datetime(df["MATURITY"], errors="coerce"))
        nearest_mat = (
            with_mat[with_mat["MAT_DT"] >= pd.Timestamp(anchor_date)]
            .groupby("PARENT_TICKER", dropna=False)["MAT_DT"]
            .min()
            .dt.strftime("%Y-%m-%d")
            .rename("Nearest Maturity")
            .reset_index()
        )
    else:
        nearest_mat = pd.DataFrame(columns=["PARENT_TICKER", "Nearest Maturity"])

    issuer_display = (
        issuer_agg.merge(tranche_count, on="PARENT_TICKER", how="left")
        .merge(nearest_mat, on="PARENT_TICKER", how="left")
        .merge(df[["PARENT_TICKER", "SECTOR"]].drop_duplicates("PARENT_TICKER"), on="PARENT_TICKER", how="left")
        .rename(
            columns={
                "SECTOR": "Sector",
                "ISSUER_FACE_MM": "Face ($MM)",
                "ISSUER_WTAVG_PX": "WA Price",
                "ISSUER_WTAVG_YIELD": "WA Yield",
                "ISSUER_MAX_YIELD": "Max Yield",
                "ISSUER_DISLOCATION_52W_SEC_MM": "52W PEAK UPSIDE SECURED ($MM)",
                "ISSUER_DISLOCATION_52W_UNSEC_MM": "52W PEAK UPSIDE UNSECURED ($MM)",
                "SECURED_PCT": "Secured (%)",
                "SENIORITY_GAP": "Seniority Basis (pts)",
                "GAP_SIGNAL": "Gap Signal",
                "ISSUER_OAS_DELTA": "OAS Delta (bps)",
            }
        )
    )
    if "ISSUER_NAME" in df.columns:
        issuer_names = df[["PARENT_TICKER", "ISSUER_NAME"]].drop_duplicates("PARENT_TICKER")
        issuer_display = issuer_display.merge(issuer_names, on="PARENT_TICKER", how="left")
        issuer_display["Issuer"] = issuer_display["ISSUER_NAME"].fillna(issuer_display["PARENT_TICKER"])
    else:
        issuer_display["Issuer"] = issuer_display["PARENT_TICKER"]
    issuer_display["Face ($BN)"] = issuer_display["Face ($MM)"] / 1000.0
    issuer_display["<1Y"] = issuer_display.apply(lambda r: _wall_label(r.get("WALL_LT1Y_MM"), r.get("WALL_LT1Y_PCT")), axis=1)
    issuer_display["1-3Y"] = issuer_display.apply(lambda r: _wall_label(r.get("WALL_1_3Y_MM"), r.get("WALL_1_3Y_PCT")), axis=1)
    issuer_display["3-5Y"] = issuer_display.apply(lambda r: _wall_label(r.get("WALL_3_5Y_MM"), r.get("WALL_3_5Y_PCT")), axis=1)
    issuer_display["COVERAGE PRIMARY"] = pd.NA
    issuer_display["COVERAGE SECONDARY"] = pd.NA
    issuer_display["Adj Unsecured Sprd Movement (bps)"] = pd.NA
    issuer_display["Sprd Movement Label"] = ""

    issuer_columns = [
        "PARENT_TICKER",
        "Issuer", "Sector", "Face ($BN)", "WA Price", "WA Yield",
        "52W PEAK UPSIDE SECURED ($MM)",
        "52W PEAK UPSIDE UNSECURED ($MM)", "Secured (%)", "Seniority Basis (pts)", "Gap Signal",
        "<1Y", "1-3Y", "3-5Y", "Nearest Maturity", "OAS Delta (bps)",
        "COVERAGE PRIMARY", "COVERAGE SECONDARY",
        "Adj Unsecured Sprd Movement (bps)", "Sprd Movement Label", "# Tranches", "Face ($MM)", "Max Yield",
    ]
    issuer_display = issuer_display[[col for col in issuer_columns if col in issuer_display.columns]].sort_values(
        by=["Face ($MM)", "52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)"],
        ascending=[False, False, False],
        na_position="last"
    )

    instrument_columns = [
        "ID", "PARENT_TICKER", "NAME", "PAYMENT_RANK", "MATURITY", "AMT_OUTSTANDING",
        "PX_MID", "PX_MID_T90", "YIELD", "YIELD_T90", "DIST_TO_PAR", "DISLOCATION_MM", "PX_HIGH_52W", "DATE_OF_HIGH",
        "PX_LOW_52W", "DATE_OF_LOW", "DISLOCATION_52W_MM", "OAS", "OAS_T90", "OAS_DELTA",
        "DISTRESS_TIER", "SECTOR", "COUNTRY", "CPN_TYPE", "CPN_VALUE",
    ]
    instruments = df[[col for col in instrument_columns if col in df.columns]].copy()
    instruments["AMT_OUTSTANDING_MM"] = pd.to_numeric(instruments["AMT_OUTSTANDING"], errors="coerce") / 1e6
    instruments = instruments.sort_values(by=["PARENT_TICKER", "DISLOCATION_MM"], ascending=[True, False], na_position="last")

    filters = {
        "sectors": sorted([s for s in issuer_display["Sector"].dropna().astype(str).unique().tolist() if s and s != "Unknown"]),
        "distress_tiers": ["Deep Distressed", "Distressed", "Stressed", "Par-Adjacent", "Premium"],
    }
    metadata = {
        "anchor_date": anchor_date,
        "t90_date": t90_date,
        "issuer_count": int(len(issuer_display)),
        "instrument_count": int(len(instruments)),
        "workbook_path": str(path),
    }

    return WorkbookData(
        issuer_table=_records_from_df(issuer_display),
        instrument_rows=_records_from_df(instruments),
        filters=filters,
        metadata=metadata,
    )
