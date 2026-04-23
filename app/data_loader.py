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
    abnormal_price_rows: list[dict]
    excluded_rows: list[dict]
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
            numeric = numeric.where(numeric.abs() <= 2_958_465, other=pd.NA)
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
    for col in ["PX_MID", "YIELD", "OAS", "I_SPREAD", "Z_SPREAD", "LAST_MONTH_VOLUME"]:
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
        abnormal_price_mask = px > 120
        normal_mask = ~abnormal_price_mask
        eligible_metric_mask = face.notna() & (face >= 200_000_000) & (maturity > today + pd.DateOffset(years=1)) & ~rank_text.str.contains("subordinated") & normal_mask
        eligible_yield_screen_mask = eligible_metric_mask & px.notna() & (px <= 95)
        face_sum = face.sum()
        summary_face = face[normal_mask].sum()
        is_preferred = group["_IS_PREFERRED"]
        not_defaulted = ~group["_IS_DEFAULTED"]
        bond_criteria = (
            normal_mask
            & not_defaulted
            & face.notna()
            & (face >= 200_000_000)
            & bond_yield.notna()
            & (bond_yield >= 7)
            & px.notna()
            & (px < 90)
            & (maturity > today + pd.DateOffset(years=1))
            & ~rank_text.str.contains("subordinated")
            & ~is_preferred
        )
        pref_criteria = (
            normal_mask
            & not_defaulted
            & face.notna()
            & (face >= 200_000_000)
            & bond_yield.notna()
            & (bond_yield >= 7)
            & px.notna()
            & (px < 90)
            & (maturity > today + pd.DateOffset(years=1))
            & is_preferred
        )
        criteria_face_mask = bond_criteria | pref_criteria
        criteria_face_sum = face[criteria_face_mask].sum()
        screen_mask = criteria_face_mask & px.notna() & (px > 0)
        px_face_sum = face[screen_mask].sum()
        wtavg_px = (px[screen_mask] * face[screen_mask]).sum() / px_face_sum if px_face_sum > 0 else pd.NA
        wtavg_yield = (bond_yield[screen_mask] * face[screen_mask]).sum() / px_face_sum if px_face_sum > 0 else pd.NA
        issuer_screen_yield_mask = bond_yield.notna() & (bond_yield > 0) & eligible_yield_screen_mask
        max_yield = bond_yield[issuer_screen_yield_mask].max() if issuer_screen_yield_mask.any() else pd.NA
        ranks = group["PAYMENT_RANK"].fillna("").astype(str)
        rank_order = ranks.map(SENIORITY_ORDER).fillna(max_known_rank + 1)
        rank_lower = ranks.str.lower()
        first_lien_mask = rank_lower.eq("1st lien secured") & px.notna() & (px > 0) & normal_mask
        senior_unsecured_mask = rank_lower.isin(["sr unsecured", "senior unsecured"]) & px.notna() & (px > 0) & normal_mask
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
        issuer_oas_delta = (pd.to_numeric(group["OAS_DELTA"], errors="coerce")[normal_mask] * face[normal_mask]).sum() / summary_face if summary_face > 0 else pd.NA

        move_3m = pd.to_numeric(group.get("PRICE_MOVE_3M"), errors="coerce") if "PRICE_MOVE_3M" in group.columns else pd.Series(dtype="float64")
        move_3m_valid = move_3m[screen_mask & move_3m.notna()]
        if not move_3m_valid.empty:
            max_3m_idx = move_3m_valid.abs().idxmax()
            max_3m = move_3m_valid.loc[max_3m_idx]
        else:
            max_3m = pd.NA

        move_7d = pd.to_numeric(group.get("PRICE_MOVE_7D"), errors="coerce") if "PRICE_MOVE_7D" in group.columns else pd.Series(dtype="float64")
        move_7d_valid = move_7d[screen_mask & move_7d.notna()]
        if not move_7d_valid.empty:
            max_7d_idx = move_7d_valid.abs().idxmax()
            max_7d = move_7d_valid.loc[max_7d_idx]
        else:
            max_7d = pd.NA

        return pd.Series(
            {
                "ISSUER_FACE_MM": face_sum / 1e6,
                "ISSUER_SECURED_FACE_MM": face[criteria_face_mask & group["_IS_SECURED"]].sum() / 1e6,
                "ISSUER_UNSECURED_FACE_MM": face[criteria_face_mask & ~group["_IS_SECURED"] & ~group["_IS_PREFERRED"]].sum() / 1e6,
                "ISSUER_PREFERRED_FACE_MM": face[criteria_face_mask & group["_IS_PREFERRED"]].sum() / 1e6,
                "ISSUER_SCREEN_FACE_MM": criteria_face_sum / 1e6,
                "ISSUER_SCREEN_FACE_PCT": (criteria_face_sum / face_sum * 100) if face_sum > 0 else pd.NA,
                "ISSUER_WTAVG_PX": wtavg_px,
                "ISSUER_WTAVG_YIELD": wtavg_yield,
                "ISSUER_MAX_YIELD": max_yield,
                "ISSUER_DIST_TO_PAR": 100.0 - wtavg_px if pd.notna(wtavg_px) else pd.NA,
                "ISSUER_DISLOCATION_MM": group.loc[screen_mask, "DISLOCATION_MM"].sum(),
                "ISSUER_DISLOCATION_SEC_MM": group.loc[screen_mask & group["_IS_SECURED"], "DISLOCATION_MM"].sum(),
                "ISSUER_DISLOCATION_UNSEC_MM": group.loc[screen_mask & ~group["_IS_SECURED"] & ~group["_IS_PREFERRED"], "DISLOCATION_MM"].sum(),
                "ISSUER_DISLOCATION_PREF_MM": group.loc[screen_mask & group["_IS_PREFERRED"], "DISLOCATION_MM"].sum(),
                "ISSUER_DISLOCATION_52W_MM": group.loc[screen_mask, "DISLOCATION_52W_MM"].sum(),
                "ISSUER_DISLOCATION_52W_SEC_MM": group.loc[screen_mask & group["_IS_SECURED"], "DISLOCATION_52W_MM"].sum(),
                "ISSUER_DISLOCATION_52W_UNSEC_MM": group.loc[screen_mask & ~group["_IS_SECURED"] & ~group["_IS_PREFERRED"], "DISLOCATION_52W_MM"].sum(),
                "ISSUER_DISLOCATION_52W_PREF_MM": group.loc[screen_mask & group["_IS_PREFERRED"], "DISLOCATION_52W_MM"].sum(),
                "SENIORITY_GAP": seniority_gap,
                "GAP_SIGNAL": _gap_label(seniority_gap),
                "ISSUER_OAS_DELTA": issuer_oas_delta,
                "WALL_LT1Y_MM": wall_lt1y_mm,
                "WALL_LT1Y_PCT": wall_lt1y_pct,
                "HAS_DEFAULTED": bool(group["_IS_DEFAULTED"].any()),
                "ISSUER_MAX_PRICE_MOVE_3M": max_3m,
                "ISSUER_MAX_PRICE_MOVE_7D": max_7d,
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
    t7_date = spread_dates[-2] if len(spread_dates) >= 3 else None

    anchor_rename = {"PX_MID": "PX_MID", "OAS": "OAS", "I_SPREAD": "I_SPREAD", "Z_SPREAD": "Z_SPREAD"}
    t90_rename = {"PX_MID": "PX_MID_T90", "OAS": "OAS_T90", "I_SPREAD": "I_SPREAD_T90", "Z_SPREAD": "Z_SPREAD_T90"}
    t7_rename = {"PX_MID": "PX_MID_T7"}
    if "YIELD" in df_spread.columns:
        anchor_rename["YIELD"] = "YIELD"
        t90_rename["YIELD"] = "YIELD_T90"
    if "LAST_MONTH_VOLUME" in df_spread.columns:
        anchor_rename["LAST_MONTH_VOLUME"] = "LAST_MONTH_VOLUME"
    if "DEFAULTED" in df_spread.columns:
        anchor_rename["DEFAULTED"] = "DEFAULTED"

    df_anchor = df_spread[df_spread["DATE"] == anchor_date].drop(columns=["DATE"]).rename(columns=anchor_rename)
    non_anchor_drop = [c for c in ["LAST_MONTH_VOLUME", "DEFAULTED"] if c in df_spread.columns]
    df_t90 = df_spread[df_spread["DATE"] == t90_date].drop(columns=["DATE"] + non_anchor_drop).rename(columns=t90_rename)
    merge_chain = (
        df_master.merge(df_anchor, on="ID", how="left")
        .merge(df_t90, on="ID", how="left")
    )
    if t7_date:
        df_t7 = df_spread[df_spread["DATE"] == t7_date][["ID", "PX_MID"]].rename(columns=t7_rename)
        merge_chain = merge_chain.merge(df_t7, on="ID", how="left")
    df = merge_chain.merge(df_52w[["ID", "PX_HIGH_52W", "DATE_OF_HIGH", "PX_LOW_52W", "DATE_OF_LOW"]], on="ID", how="left")

    # Exclude bonds that matured more than 3 months ago
    if anchor_date and "MATURITY" in df.columns:
        maturity_cutoff = pd.Timestamp(anchor_date) - pd.DateOffset(months=3)
        mat = pd.to_datetime(df["MATURITY"], errors="coerce")
        df = df[mat.isna() | (mat >= maturity_cutoff)].copy()

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
    px_t90 = pd.to_numeric(df["PX_MID_T90"], errors="coerce")
    df["PRICE_MOVE_3M"] = px - px_t90
    if "PX_MID_T7" in df.columns:
        px_t7 = pd.to_numeric(df["PX_MID_T7"], errors="coerce")
        df["PRICE_MOVE_7D"] = px - px_t7
    df["DISTRESS_TIER"] = px.apply(_distress_tier)
    df["_IS_SECURED"] = df["PAYMENT_RANK"].apply(_is_secured)
    df["_IS_PREFERRED"] = df["PAYMENT_RANK"].fillna("").astype(str).str.lower().eq("preferred")
    df["_IS_DEFAULTED"] = df["DEFAULTED"].fillna("").astype(str).eq("1") if "DEFAULTED" in df.columns else False

    issuer_agg = _issuer_metrics(df, anchor_date) if anchor_date else pd.DataFrame(columns=["PARENT_TICKER"])
    summary_df = df[~(pd.to_numeric(df["PX_MID"], errors="coerce") > 120)].copy()
    issuer_display = (
        issuer_agg
        .merge(df[["PARENT_TICKER", "SECTOR"]].drop_duplicates("PARENT_TICKER"), on="PARENT_TICKER", how="left")
        .rename(
            columns={
                "SECTOR": "Sector",
                "ISSUER_FACE_MM": "Total Face ($MM)",
                "ISSUER_SCREEN_FACE_MM": "Face Meets Criteria ($MM)",
                "ISSUER_SCREEN_FACE_PCT": "% Face Meets Criteria",
                "ISSUER_WTAVG_PX": "Price",
                "ISSUER_WTAVG_YIELD": "Yield",
                "ISSUER_MAX_YIELD": "Max Yield",
                "ISSUER_SECURED_FACE_MM": "Secured Face ($MM)",
                "ISSUER_UNSECURED_FACE_MM": "Unsecured Face ($MM)",
                "ISSUER_PREFERRED_FACE_MM": "Preferred Face ($MM)",
                "ISSUER_DISLOCATION_52W_SEC_MM": "52W PEAK UPSIDE SECURED ($MM)",
                "ISSUER_DISLOCATION_52W_UNSEC_MM": "52W PEAK UPSIDE UNSECURED ($MM)",
                "ISSUER_DISLOCATION_52W_PREF_MM": "52W PEAK UPSIDE PREFERRED ($MM)",
                "ISSUER_DISLOCATION_SEC_MM": "RETURN TO PAR SECURED ($MM)",
                "ISSUER_DISLOCATION_UNSEC_MM": "RETURN TO PAR UNSECURED ($MM)",
                "ISSUER_DISLOCATION_PREF_MM": "RETURN TO PAR PREFERRED ($MM)",
                "SENIORITY_GAP": "Seniority Basis (pts)",
                "GAP_SIGNAL": "Gap Signal",
                "ISSUER_OAS_DELTA": "OAS Delta (bps)",
                "ISSUER_MAX_PRICE_MOVE_3M": "3M Price Move",
                "ISSUER_MAX_PRICE_MOVE_7D": "7D Price Move",
            }
        )
    )
    if "ISSUER_NAME" in df.columns:
        issuer_names = df[["PARENT_TICKER", "ISSUER_NAME"]].drop_duplicates("PARENT_TICKER")
        issuer_display = issuer_display.merge(issuer_names, on="PARENT_TICKER", how="left")
        issuer_display["Issuer"] = issuer_display["ISSUER_NAME"].fillna(issuer_display["PARENT_TICKER"])
    else:
        issuer_display["Issuer"] = issuer_display["PARENT_TICKER"]
    issuer_display["Secured Face ($BN)"] = issuer_display["Secured Face ($MM)"] / 1000.0
    issuer_display["Unsecured Face ($BN)"] = issuer_display["Unsecured Face ($MM)"] / 1000.0
    issuer_display["Preferred Face ($BN)"] = issuer_display["Preferred Face ($MM)"] / 1000.0
    issuer_display["Maturity Within 1Y"] = issuer_display.apply(lambda r: _wall_label(r.get("WALL_LT1Y_MM"), r.get("WALL_LT1Y_PCT")), axis=1)
    issuer_display["COVERAGE PRIMARY"] = pd.NA
    issuer_display["COVERAGE SECONDARY"] = pd.NA
    issuer_display["Adj Unsecured Sprd Movement (bps)"] = pd.NA
    issuer_display["Sprd Movement Label"] = ""

    issuer_columns = [
        "PARENT_TICKER",
        "Issuer", "Sector", "Secured Face ($BN)", "Unsecured Face ($BN)", "Preferred Face ($BN)",
        "Price", "Yield", "3M Price Move", "7D Price Move",
        "52W PEAK UPSIDE SECURED ($MM)",
        "52W PEAK UPSIDE UNSECURED ($MM)",
        "52W PEAK UPSIDE PREFERRED ($MM)",
        "RETURN TO PAR SECURED ($MM)",
        "RETURN TO PAR UNSECURED ($MM)",
        "RETURN TO PAR PREFERRED ($MM)",
        "Seniority Basis (pts)", "Gap Signal",
        "Maturity Within 1Y", "OAS Delta (bps)",
        "COVERAGE PRIMARY", "COVERAGE SECONDARY",
        "Adj Unsecured Sprd Movement (bps)", "Sprd Movement Label",
        "Total Face ($MM)", "Secured Face ($MM)", "Unsecured Face ($MM)", "Preferred Face ($MM)", "Max Yield",
        "HAS_DEFAULTED",
    ]
    issuer_display = issuer_display[[col for col in issuer_columns if col in issuer_display.columns]].sort_values(
        by=["Total Face ($MM)", "52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)"],
        ascending=[False, False, False],
        na_position="last"
    )

    # Exclude issuers where no bonds meet screening criteria
    has_qualifying = (
        issuer_display["Secured Face ($BN)"].fillna(0).gt(0)
        | issuer_display["Unsecured Face ($BN)"].fillna(0).gt(0)
        | issuer_display["Preferred Face ($BN)"].fillna(0).gt(0)
        | issuer_display["HAS_DEFAULTED"].fillna(False).astype(bool)
    )
    issuer_display = issuer_display[has_qualifying].copy()

    instrument_columns = [
        "ID", "PARENT_TICKER", "NAME", "PAYMENT_RANK", "MATURITY", "AMT_OUTSTANDING",
        "PX_MID", "PX_MID_T7", "PX_MID_T90", "PRICE_MOVE_7D", "PRICE_MOVE_3M",
        "YIELD", "YIELD_T90", "DIST_TO_PAR", "DISLOCATION_MM", "PX_HIGH_52W", "DATE_OF_HIGH",
        "PX_LOW_52W", "DATE_OF_LOW", "DISLOCATION_52W_MM", "OAS", "OAS_T90", "OAS_DELTA",
        "DISTRESS_TIER", "SECTOR", "COUNTRY", "CPN_TYPE", "CPN_VALUE", "LAST_MONTH_VOLUME", "_IS_DEFAULTED",
    ]
    instruments = summary_df[[col for col in instrument_columns if col in summary_df.columns]].copy()
    instruments["AMT_OUTSTANDING_MM"] = pd.to_numeric(instruments["AMT_OUTSTANDING"], errors="coerce") / 1e6
    if "LAST_MONTH_VOLUME" in instruments.columns:
        raw_vol = pd.to_numeric(instruments["LAST_MONTH_VOLUME"], errors="coerce")
        pref_mask = instruments["PAYMENT_RANK"].fillna("").astype(str).str.lower().eq("preferred")
        instruments["LAST_30D_VOLUME_MM"] = raw_vol / 1e3
        instruments.loc[pref_mask, "LAST_30D_VOLUME_MM"] = raw_vol[pref_mask] / 1e6
    # Null out yield and price movement for defaulted bonds
    if "_IS_DEFAULTED" in instruments.columns:
        dmask = instruments["_IS_DEFAULTED"].fillna(False).astype(bool)
        for col in ["YIELD", "YIELD_T90", "PX_MID_T90", "PX_MID_T7", "PRICE_MOVE_3M", "PRICE_MOVE_7D", "OAS_DELTA"]:
            if col in instruments.columns:
                instruments.loc[dmask, col] = pd.NA
    instruments = instruments.sort_values(by=["PARENT_TICKER", "DISLOCATION_MM"], ascending=[True, False], na_position="last")

    abnormal_prices = df[pd.to_numeric(df["PX_MID"], errors="coerce") > 120].copy()
    if "ISSUER_NAME" in abnormal_prices.columns:
        abnormal_prices["Issuer"] = abnormal_prices["ISSUER_NAME"].fillna(abnormal_prices["PARENT_TICKER"])
    else:
        abnormal_prices["Issuer"] = abnormal_prices["PARENT_TICKER"]
    abnormal_prices["AMT_OUTSTANDING_MM"] = pd.to_numeric(abnormal_prices["AMT_OUTSTANDING"], errors="coerce") / 1e6
    abnormal_prices["PRICE_MOVE_3M"] = pd.to_numeric(abnormal_prices["PX_MID"], errors="coerce") - pd.to_numeric(abnormal_prices["PX_MID_T90"], errors="coerce")
    abnormal_price_columns = [
        "Issuer", "PARENT_TICKER", "ID", "NAME", "SECTOR", "PAYMENT_RANK", "MATURITY",
        "AMT_OUTSTANDING_MM", "PX_MID", "PX_MID_T90", "PRICE_MOVE_3M", "YIELD", "YIELD_T90",
        "PX_LOW_52W", "PX_HIGH_52W", "COUNTRY", "CPN_TYPE", "CPN_VALUE",
    ]
    abnormal_prices = abnormal_prices[[col for col in abnormal_price_columns if col in abnormal_prices.columns]].sort_values(
        by=["PX_MID", "AMT_OUTSTANDING_MM"],
        ascending=[False, False],
        na_position="last",
    )

    # Build excluded bonds: in summary_df but not meeting screening criteria
    exc = summary_df.copy()
    exc_face = pd.to_numeric(exc["AMT_OUTSTANDING"], errors="coerce")
    exc_yield = pd.to_numeric(exc["YIELD"], errors="coerce") if "YIELD" in exc.columns else pd.Series(index=exc.index, dtype="float64")
    exc_px = pd.to_numeric(exc["PX_MID"], errors="coerce")
    exc_mat = pd.to_datetime(exc["MATURITY"], errors="coerce")
    exc_rank = exc["PAYMENT_RANK"].fillna("").astype(str).str.lower()
    exc_defaulted = exc["_IS_DEFAULTED"] if "_IS_DEFAULTED" in exc.columns else pd.Series(False, index=exc.index)
    exc_preferred = exc["_IS_PREFERRED"] if "_IS_PREFERRED" in exc.columns else pd.Series(False, index=exc.index)
    today_ts = pd.Timestamp(anchor_date) if anchor_date else pd.Timestamp.now()
    mat_cutoff = today_ts + pd.DateOffset(years=1)
    common_criteria = (
        exc_face.notna() & (exc_face >= 200_000_000)
        & exc_yield.notna() & (exc_yield >= 7)
        & exc_px.notna() & (exc_px < 90)
        & (exc_mat > mat_cutoff)
        & ~exc_defaulted
    )
    bond_pass = common_criteria & ~exc_rank.str.contains("subordinated") & ~exc_preferred
    pref_pass = common_criteria & exc_preferred
    screen_pass = bond_pass | pref_pass
    excluded_df = exc[~screen_pass].copy()
    if "ISSUER_NAME" in excluded_df.columns:
        excluded_df["Issuer"] = excluded_df["ISSUER_NAME"].fillna(excluded_df["PARENT_TICKER"])
    else:
        excluded_df["Issuer"] = excluded_df["PARENT_TICKER"]
    excluded_df["AMT_OUTSTANDING_MM"] = pd.to_numeric(excluded_df["AMT_OUTSTANDING"], errors="coerce") / 1e6
    # Tag exclusion reason
    reasons = []
    for _, r in exc.iterrows():
        r_reasons = []
        f = pd.to_numeric(r.get("AMT_OUTSTANDING"), errors="coerce")
        y = pd.to_numeric(r.get("YIELD"), errors="coerce")
        p = pd.to_numeric(r.get("PX_MID"), errors="coerce")
        m = pd.to_datetime(r.get("MATURITY"), errors="coerce")
        if pd.isna(f) or f < 200_000_000:
            r_reasons.append("Face < $200M")
        if pd.isna(y):
            r_reasons.append("No Yield")
        elif y < 7:
            r_reasons.append("Yield < 7%")
        if pd.isna(p):
            r_reasons.append("No Price")
        elif p >= 90:
            r_reasons.append("Price >= 90")
        if pd.notna(m) and m <= mat_cutoff:
            r_reasons.append("Maturity <= 1Y")
        rk = str(r.get("PAYMENT_RANK") or "").lower()
        if "subordinated" in rk:
            r_reasons.append("Subordinated")
        if r.get("_IS_DEFAULTED"):
            r_reasons.append("Defaulted")
        reasons.append("; ".join(r_reasons) if r_reasons else "")
    exc["EXCLUSION_REASON"] = reasons
    excluded_df["Exclusion Reason"] = exc.loc[excluded_df.index, "EXCLUSION_REASON"]
    exclusion_columns = [
        "Issuer", "PARENT_TICKER", "NAME", "SECTOR", "PAYMENT_RANK", "MATURITY",
        "AMT_OUTSTANDING_MM", "PX_MID", "YIELD", "Exclusion Reason",
    ]
    excluded_df = excluded_df[[col for col in exclusion_columns if col in excluded_df.columns]].sort_values(
        by=["AMT_OUTSTANDING_MM"], ascending=[False], na_position="last",
    )

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
        abnormal_price_rows=_records_from_df(abnormal_prices),
        excluded_rows=_records_from_df(excluded_df),
        filters=filters,
        metadata=metadata,
    )
