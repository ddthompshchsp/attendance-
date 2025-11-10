import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import calendar
import re
import zipfile

st.set_page_config(page_title="ADA Monthly + Cumulative Reports", layout="wide")

# ---- THEME / COLORS ----
BLUE = "#2E75B6"
RED = "#C00000"
GREY = "#E6E6E6"

# ---- GROUP DEFINITIONS ----
# NOTE: Guzman is NOT in Group 1; it is split into two dedicated sheets (4yo / 3yo)
GROUPS = {
    "Group 1": [
        "Donna","Mission","Monte Alto","Sam Fordyce","San Carlos","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    "Group 2": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    "Group 3": ["Mercedes","Edinburg"],
    # Guzman handled separately
}

# Simple alias fixes for token matching (keeps display names untouched)
SIMPLE_ALIASES = {"fairs":"farias","sequin":"seguin","chaps":"chapa"}

# ---- HELPERS ----
def _month_name(m: int) -> str:
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _canon(x: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', str(x).lower())

def _norm_simple(s: str) -> str:
    k = _canon(s)
    return SIMPLE_ALIASES.get(k, k)

def _weighted_avg_rate(df: pd.DataFrame) -> float:
    """Weighted Attendance Rate (0..100) by Current."""
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    w = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    r = df["Attendance Rate"].fillna(0).astype(float)
    tot = float(w.sum())
    return float((r*w).sum()/tot) if tot > 0 else np.nan

def _center_block(center_df: pd.DataFrame) -> pd.DataFrame:
    """Return class rows + one TOTAL row per center/month."""
    body = center_df.sort_values("Class Name").copy()
    w = _weighted_avg_rate(center_df)
    return pd.concat([
        body,
        pd.DataFrame([{
            "Year": center_df['Year'].dropna().iloc[0] if center_df['Year'].notna().any() else np.nan,
            "Month": center_df['Month'].dropna().iloc[0] if center_df['Month'].notna().any() else np.nan,
            "Center Name": center_df['Center Name'].dropna().iloc[0] if center_df['Center Name'].notna().any() else np.nan,
            "Class Name": "TOTAL",
            "Funded": center_df['Funded'].sum(min_count=1),
            "Current": center_df['Current'].sum(min_count=1),
            "Attendance Rate": float(w) if pd.notna(w) else np.nan,
        }])
    ], ignore_index=True)

def _detect_age(text: str):
    t = str(text).lower()
    if re.search(r'(^|[^0-9])4($|[^0-9])|4yo|4-?yr|4\s*year', t):
        return "4yo"
    if re.search(r'(^|[^0-9])3($|[^0-9])|3yo|3-?yr|3\s*year', t):
        return "3yo"
    return None

def _append_guzman_age(df: pd.DataFrame) -> pd.DataFrame:
    """Add Age Tag for Guzman classes and annotate Class Name with (3yo)/(4yo)."""
    out = df.copy()
    if "Center Name" not in out.columns or "Class Name" not in out.columns:
        return out
    mask = out["Center Name"].astype(str).str.lower().str.contains("guzman")
    out.loc[mask, "Age Tag"] = out.loc[mask, "Class Name"].astype(str).map(_detect_age)
    out.loc[mask, "Class Name"] = out.loc[mask].apply(
        lambda r: f"{r['Class Name']} ({r['Age Tag']})" if pd.notna(r.get("Age Tag")) else r["Class Name"],
        axis=1
    )
    return out

def _load_enrollment(file_bytes: bytes, preferred_sheet="V12POP_ERSEA_Enrollment") -> pd.DataFrame:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    use_sheet = preferred_sheet if preferred_sheet in xls.sheet_names else xls.sheet_names[0]
    df = None
    for hdr in (1, 0):
        try:
            tmp = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=hdr)
            if not tmp.empty:
                df = tmp
                break
        except Exception:
            pass
    if df is None or df.empty:
        raise ValueError("Could not read a non-empty sheet from the workbook.")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={"Unnamed: 6":"Funded","Unnamed: 8":"Current","Unnamed: 9":"Attendance Rate"})
    keep = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
    df = df[[c for c in keep if c in df.columns]].copy()
    for c in ['Year','Month','Funded','Current','Attendance Rate']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    if 'Class Name' in df.columns:
        mask_total = df['Class Name'].astype(str).str.upper().eq('TOTAL')
        mask_blank = df['Class Name'].astype(str).str.strip().eq('') | df['Class Name'].isna()
        df = df[~(mask_total | mask_blank)].copy()
    return df

def _match_centers(simple_list, present_full):
    matched = []
    for simple in simple_list:
        tok = _norm_simple(simple)
        for full in present_full:
            if tok in _canon(full):
                if _canon(full) not in {_canon(x) for x in matched}:
                    matched.append(full)
    return matched

# ---------- UI ----------
st.title("ADA Monthly & Cumulative Reports (Groups + Guzman Split)")

uploaded = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Upload the enrollment workbook to begin.")
    st.stop()

file_bytes = uploaded.read()
if not file_bytes:
    st.error("Uploaded file is empty.")
    st.stop()

try:
    df_all = _load_enrollment(file_bytes)
except Exception as e:
    st.error(f"Failed to read workbook: {e}")
    st.stop()

# Month controls
months_present = sorted({int(m) for m in df_all['Month'].dropna().astype(int).unique()})
if not months_present:
    st.error("No valid Month values found in the file.")
    st.stop()

colA, colB, colC = st.columns([1,1,2])
with colA:
    # Default to Sept & Oct if present
    default_months = [m for m in [9,10] if m in months_present] or [months_present[-1]]
    sel_months = st.multiselect(
        "Monthly Report Months",
        options=months_present,
        default=default_months,
        format_func=_month_name
    )
with colB:
    # Cumulative month range (inclusive)
    start_m = st.selectbox("Cumulative Start Month", options=months_present, index=months_present.index(min(default_months)))
    end_m = st.selectbox("Cumulative End Month", options=months_present, index=months_present.index(max(default_months)))

if not sel_months:
    st.warning("Select at least one month for the Monthly report.")
    st.stop()

# ---------- BUILD MONTHLY REPORT ----------
def build_monthly_excel(df_all: pd.DataFrame, sel_months: list[int]) -> bytes:
    grouped_src = _append_guzman_age(df_all[df_all['Month'].isin(sel_months)].copy())
    ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y%m%d_%H%M")
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        pct_fmt    = wb.add_format({"num_format": "0.00%"})
        bold_fmt   = wb.add_format({"bold": True})
        s_head     = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})
        s_cell     = wb.add_format({"border": 1, "font_size": 9})
        s_pct      = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})

        # ADA overview for latest selected month
        latest_m = max(sel_months)
        m_df = grouped_src[grouped_src["Month"] == latest_m].copy()
        if not m_df.empty:
            blocks = [_center_block(g) for _, g in m_df.groupby("Center Name", dropna=False)]
            ada = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(columns	m_df.columns)
            agency_overall = _weighted_avg_rate(m_df)
            if not ada.empty:
                ada = pd.concat([ada, pd.DataFrame([{
                    "Year": m_df['Year'].dropna().iloc[0] if m_df['Year'].notna().any() else np.nan,
                    "Month": latest_m,
                    "Center Name": "HCHSP (Overall)",
                    "Class Name": "TOTAL",
                    "Funded": m_df['Funded'].sum(min_count=1),
                    "Current": m_df['Current'].sum(min_count=1),
                    "Attendance Rate": float(agency_overall) if pd.notna(agency_overall) else np.nan,
                }])], ignore_index=True)
        else:
            ada = pd.DataFrame(columns=['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate'])

        # ADA sheet
        ada.to_excel(writer, index=False, sheet_name="ADA", startrow=3)
        ws = writer.sheets["ADA"]
        header_row = 3
        for c, name in enumerate(ada.columns):
            ws.write(header_row, c, name, header_fmt)
        ws.set_row(header_row, 26)
        for i, col in enumerate(ada.columns):
            width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2)) if not ada.empty else 16
            ws.set_column(i, i, width)
        if not ada.empty and 'Attendance Rate' in ada.columns:
            ci = ada.columns.get_loc('Attendance Rate')
            ki = ada.columns.get_loc('Class Name')
            for r in range(len(ada)):
                excel_r = r + header_row + 1
                v = ada.iloc[r, ci]
                if pd.isna(v):
                    ws.write_blank(excel_r, ci, None, pct_fmt)
                else:
                    ws.write_number(excel_r, ci, float(v)/100.0, pct_fmt)
                if str(ada.iloc[r, ki]).upper() == "TOTAL":
                    ws.set_row(excel_r, None, bold_fmt)
        ws.freeze_panes(header_row + 1, 0)
        last_col = max(0, len(ada.columns)-1)
        last_row = header_row + len(ada)
        ws.autofilter(header_row, 0, last_row, last_col)
        ws.write(0, 1, f"Daily Attendance Rate — {_month_name(latest_m)}", bold_fmt)

        # Per-group writer
        def write_group_sheet(sheet_name: str, df_src: pd.DataFrame):
            ws = wb.add_worksheet(sheet_name[:31])
            start_row = 3
            # ADA blocks per center per month
            for fn, g_center in df_src.groupby("Center Name"):
                for m in sorted(g_center["Month"].dropna().unique().astype(int)):
                    g = g_center[g_center["Month"] == m].copy()
                    if g.empty:
                        continue
                    block = _center_block(g)
                    ws.write(start_row - 2, 0, f"{fn} — {_month_name(m)}", bold_fmt)
                    ws.set_row(start_row, 26)
                    # headers
                    for c, name in enumerate(block.columns):
                        ws.write(start_row, c, name, header_fmt)
                    # values
                    for r in range(len(block)):
                        for c, col in enumerate(block.columns):
                            val = block.iloc[r, c]
                            if col == "Attendance Rate":
                                if pd.isna(val):
                                    ws.write_blank(start_row + 1 + r, c, None, pct_fmt)
                                else:
                                    ws.write_number(start_row + 1 + r, c, float(val)/100.0, pct_fmt)
                            else:
                                if pd.isna(val):
                                    ws.write_blank(start_row + 1 + r, c, None)
                                else:
                                    ws.write(start_row + 1 + r, c, val)
                        if str(block.iloc[r]["Class Name"]).upper() == "TOTAL":
                            ws.set_row(start_row + 1 + r, None, bold_fmt)
                    # autosize
                    for i, col in enumerate(block.columns):
                        width = max(12, min(44, int(block[col].astype(str).map(len).max()) + 2))
                        ws.set_column(i, i, width)
                    start_row += len(block) + 3
            ws.freeze_panes(4, 0)

            # Compact summary (Month, Center, %)
            sums = (
                df_src.groupby(["Month","Center Name"], dropna=False)
                      .apply(_weighted_avg_rate)
                      .reset_index(name="Attendance %")
                      .sort_values(["Month","Attendance %"], ascending=[True, False])
            )
            small_r, small_c = start_row, 0
            ws.write(small_r,   small_c,   "Month",        wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9}))
            ws.write(small_r,   small_c+1, "Center Name",  wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9}))
            ws.write(small_r,   small_c+2, "Attendance %", wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9}))
            for i, row in sums.iterrows():
                ws.write_string(small_r + 1 + i, small_c,   _month_name(int(row["Month"])) if pd.notna(row["Month"]) else "", s_cell)
                ws.write_string(small_r + 1 + i, small_c+1, str(row["Center Name"]), s_cell)
                if pd.isna(row["Attendance %"]):
                    ws.write_blank(small_r + 1 + i, small_c+2, None, s_pct)
                else:
                    ws.write_number(small_r + 1 + i, small_c+2, float(row["Attendance %"])/100.0, s_pct)
            last_row_small = small_r + len(sums)
            ws.autofilter(small_r, small_c, last_row_small, small_c + 2)
            ws.set_column(small_c, small_c, 10)
            ws.set_column(small_c+1, small_c+1, 36)
            ws.set_column(small_c+2, small_c+2, 14)

            # Charts: one per month (e.g., Sept / Oct)
            months_for_charts = sorted(df_src["Month"].dropna().unique().astype(int).tolist())
            chart_col_offset = 5
            for idx, m in enumerate(months_for_charts):
                rows_for_m = [i for i, (_, r) in enumerate(sums.iterrows()) if int(r["Month"]) == int(m)]
                if not rows_for_m:
                    continue
                start_off, end_off = rows_for_m[0], rows_for_m[-1]
                ch = wb.add_chart({"type": "column"})
                ch.add_series({
                    "name": f"{_month_name(m)} Attendance %",
                    "categories": [sheet_name[:31], small_r + 1 + start_off, small_c + 1, small_r + 1 + end_off, small_c + 1],
                    "values":     [sheet_name[:31], small_r + 1 + start_off, small_c + 2, small_r + 1 + end_off, small_c + 2],
                    "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
                })
                ch.set_y_axis({"num_format": "0.00%"})
                ch.set_title({"name": f"{_month_name(m)} — Attendance %"})
                ws.insert_chart(small_r, small_c + chart_col_offset + idx*8, ch, {"x_scale": 1.1, "y_scale": 1.0})

        # Build Group 1–3 (no Guzman)
        present_full = grouped_src['Center Name'].dropna().unique().tolist()
        for sheet_name, simple_list in GROUPS.items():
            matched = _match_centers(simple_list, present_full)
            df_use = grouped_src[grouped_src["Center Name"].isin(matched)].copy()
            write_group_sheet(sheet_name, df_use)

        # Guzman split
        guz = grouped_src[grouped_src["Center Name"].astype(str).str.lower().str.contains("guzman")].copy()
        if not guz.empty:
            guz["Age Tag"] = guz["Class Name"].astype(str).map(_detect_age)
            for tag, title in [("4yo","Group 4 — Guzman (4yo)"),("3yo","Group 4 — Guzman (3yo)")]:
                sub = guz[guz["Age Tag"]==tag].copy()
                if sub.empty: 
                    continue
                write_group_sheet(title, sub)

    return bio.getvalue()

# ---------- BUILD CUMULATIVE REPORT ----------
def build_cumulative_excel(df_all: pd.DataFrame, start_m: int, end_m: int) -> bytes:
    months = [m for m in df_all['Month'].dropna().astype(int).unique() if start_m <= m <= end_m]
    if not months:
        months = [int(df_all['Month'].dropna().astype(int).max())]
    df_rng = _append_guzman_age(df_all[df_all['Month'].isin(months)].copy())

    def _cumulative_class(dfcls: pd.DataFrame) -> pd.Series:
        w = dfcls['Current'].fillna(0).astype(float)
        r = dfcls['Attendance Rate'].fillna(0).astype(float)
        wsum = w.sum()
        rate = float((w*r).sum()/wsum) if wsum>0 else np.nan
        funded = float(dfcls['Funded'].max(skipna=True)) if 'Funded' in dfcls else np.nan
        current = float(dfcls['Current'].max(skipna=True)) if 'Current' in dfcls else np.nan
        return pd.Series({
            "Month Range": f"{_month_name(start_m)}–{_month_name(end_m)}",
            "Funded": funded,
            "Current": current,
            "Attendance Rate": rate
        })

    cum_class = (
        df_rng.groupby(["Center Name","Class Name"], dropna=False)
              .apply(_cumulative_class)
              .reset_index()
    )

    def _center_total_from_rows(center_rows: pd.DataFrame) -> pd.DataFrame:
        w = center_rows['Current'].fillna(0).astype(float)
        r = center_rows['Attendance Rate'].fillna(0).astype(float)
        wsum = w.sum()
        rate = float((w*r).sum()/wsum) if wsum>0 else np.nan
        return pd.DataFrame([{
            "Center Name": center_rows['Center Name'].iloc[0],
            "Class Name": "TOTAL",
            "Month Range": center_rows['Month Range'].iloc[0],
            "Funded": center_rows['Funded'].sum(min_count=1),
            "Current": center_rows['Current'].sum(min_count=1),
            "Attendance Rate": rate
        }])

    centers_totals = cum_class.groupby("Center Name", dropna=False).apply(_center_total_from_rows).reset_index(drop=True)
    cum_full = pd.concat([cum_class, centers_totals], ignore_index=True)

    # Build workbook
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        pct_fmt    = wb.add_format({"num_format": "0.00%"})
        s_head     = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})
        s_cell     = wb.add_format({"border": 1, "font_size": 9})
        s_pct      = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})

        # Dashboard: ranked centers cumulative
        centers_rank = (
            cum_full[cum_full["Class Name"]=="TOTAL"][["Center Name","Attendance Rate"]]
            .sort_values("Attendance Rate", ascending=False)
            .reset_index(drop=True)
        )
        ws0 = wb.add_worksheet("Dashboard")
        ws0.write(0, 1, f"Cumulative ADA — {_month_name(start_m)} to {_month_name(end_m)}", wb.add_format({"bold": True, "font_size": 22, "font_color": BLUE}))
        ws0.write(2, 1, "Center Name", s_head); ws0.write(2, 2, "Attendance %", s_head)
        for i, row in centers_rank.iterrows():
            ws0.write_string(3+i, 1, str(row["Center Name"]), s_cell)
            if pd.isna(row["Attendance Rate"]):
                ws0.write_blank(3+i, 2, None, s_pct)
            else:
                ws0.write_number(3+i, 2, float(row["Attendance Rate"])/100.0, s_pct)
        # chart
        if not centers_rank.empty:
            ch = wb.add_chart({"type":"column"})
            ch.add_series({
                "name": "Cumulative ADA",
                "categories": ["Dashboard", 3, 1, 3+len(centers_rank)-1, 1],
                "values":     ["Dashboard", 3, 2, 3+len(centers_rank)-1, 2],
                "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
            })
            ch.set_y_axis({"num_format": "0.00%"})
            ch.set_title({"name": "Centers — Cumulative ADA"})
            ws0.insert_chart(2, 5, ch, {"x_scale": 1.2, "y_scale": 1.1})

        # Per-group sheets
        present_full2 = df_rng['Center Name'].dropna().unique().tolist()
        for sheet_name, simple_list in GROUPS.items():
            matched = _match_centers(simple_list, present_full2)
            view = cum_full[cum_full["Center Name"].isin(matched)].copy()
            ws = wb.add_worksheet(sheet_name[:31])

            # class rows
            class_rows = view[view["Class Name"]!="TOTAL"][["Center Name","Class Name","Month Range","Funded","Current","Attendance Rate"]].copy()
            head_row = 3
            for c, name in enumerate(class_rows.columns):
                ws.write(head_row, c, name, header_fmt)
            for r in range(len(class_rows)):
                for c, col in enumerate(class_rows.columns):
                    val = class_rows.iloc[r, c]
                    if col == "Attendance Rate":
                        ws.write_number(head_row + 1 + r, c, 0 if pd.isna(val) else float(val)/100.0, pct_fmt)
                    else:
                        if pd.isna(val):
                            ws.write_blank(head_row + 1 + r, c, None)
                        else:
                            ws.write(head_row + 1 + r, c, val)
            for i, col in enumerate(class_rows.columns):
                width = max(12, min(44, int(class_rows[col].astype(str).map(len).max()) + 2))
                ws.set_column(i, i, width)

            # center totals ranked
            start_r = head_row + 2 + len(class_rows)
            totals = view[view["Class Name"]=="TOTAL"][["Center Name","Month Range","Funded","Current","Attendance Rate"]].copy().sort_values("Attendance Rate", ascending=False)
            headers = ["Center Name","Month Range","Funded","Current","Attendance %"]
            for c, name in enumerate(headers):
                ws.write(start_r, c, name, header_fmt)
            for i, (_, row) in enumerate(totals.iterrows()):
                ws.write_string(start_r + 1 + i, 0, str(row["Center Name"]))
                ws.write_string(start_r + 1 + i, 1, str(row["Month Range"]))
                ws.write_number(start_r + 1 + i, 2, float(row["Funded"]) if pd.notna(row["Funded"]) else 0)
                ws.write_number(start_r + 1 + i, 3, float(row["Current"]) if pd.notna(row["Current"]) else 0)
                ws.write_number(start_r + 1 + i, 4, 0 if pd.isna(row["Attendance Rate"]) else float(row["Attendance Rate"])/100.0, pct_fmt)

            # chart
            if not totals.empty:
                ch2 = wb.add_chart({"type":"column"})
                ch2.add_series({
                    "name": "Cumulative ADA",
                    "categories": [sheet_name[:31], start_r + 1, 0, start_r + len(totals), 0],
                    "values":     [sheet_name[:31], start_r + 1, 4, start_r + len(totals), 4],
                    "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
                })
                ch2.set_y_axis({"num_format":"0.00%"})
                ch2.set_title({"name":"Cumulative ADA by Center"})
                ws.insert_chart(start_r, 6, ch2, {"x_scale": 1.0, "y_scale": 1.0})

            ws.freeze_panes(4, 0)

        # Guzman split cumulative
        guz = df_rng[df_rng["Center Name"].astype(str).str.lower().str.contains("guzman")].copy()
        if not guz.empty:
            guz["Age Tag"] = guz["Class Name"].astype(str).map(_detect_age)
            for tag, title in [("4yo","Guzman (4yo) — Cumulative"),("3yo","Guzman (3yo) — Cumulative")]:
                sub = guz[guz["Age Tag"]==tag].copy()
                if sub.empty:
                    continue
                # cumulative by class
                cum_g = (
                    sub.groupby(["Center Name","Class Name"], dropna=False)
                       .apply(lambda dfcls: pd.Series({
                           "Month Range": f"{_month_name(start_m)}–{_month_name(end_m)}",
                           "Funded": dfcls['Funded'].max(skipna=True),
                           "Current": dfcls['Current'].max(skipna=True),
                           "Attendance Rate": (dfcls['Attendance Rate'].fillna(0).to_numpy() * dfcls['Current'].fillna(0).to_numpy()).sum() /
                                              max(1e-9, dfcls['Current'].fillna(0).sum())
                       }))
                       .reset_index()
                )
                # center total
                w = cum_g['Current'].fillna(0).astype(float)
                r = cum_g['Attendance Rate'].fillna(0).astype(float)
                wsum = w.sum()
                rate = float((w*r).sum()/wsum) if wsum>0 else np.nan
                tot = pd.DataFrame([{
                    "Center Name": "Guzman",
                    "Class Name": "TOTAL",
                    "Month Range": f"{_month_name(start_m)}–{_month_name(end_m)}",
                    "Funded": cum_g['Funded'].sum(min_count=1),
                    "Current": cum_g['Current'].sum(min_count=1),
                    "Attendance Rate": rate
                }])
                view = pd.concat([cum_g, tot], ignore_index=True)
                ws = wb.add_worksheet(title[:31])
                head = ["Center Name","Class Name","Month Range","Funded","Current","Attendance Rate"]
                for c, name in enumerate(head):
                    ws.write(3, c, name, header_fmt)
                for r_i in range(len(view)):
                    for c, col in enumerate(head):
                        val = view.iloc[r_i][col]
                        if col == "Attendance Rate":
                            ws.write_number(4 + r_i, c, 0 if pd.isna(val) else float(val)/100.0, pct_fmt)
                        else:
                            if pd.isna(val):
                                ws.write_blank(4 + r_i, c, None)
                            else:
                                ws.write(4 + r_i, c, val)
                for i, col in enumerate(head):
                    width = max(12, min(44, int(view[col].astype(str).map(len).max()) + 2))
                    ws.set_column(i, i, width)
                ws.freeze_panes(4, 0)

    return bio.getvalue()

# ---------- GENERATE BOTH & ZIP ----------
monthly_bytes = build_monthly_excel(df_all, sel_months)
cumulative_bytes = build_cumulative_excel(df_all, start_m, end_m)

ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y%m%d_%H%M")
zip_buf = BytesIO()
with zipfile.ZipFile(zip_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    zf.writestr(f"ADA_Monthly_Groups_GuzmanSplit_{ts}.xlsx", monthly_bytes)
    zf.writestr(f"ADA_Cumulative_{_month_name(start_m)}_{_month_name(end_m)}_{ts}.xlsx", cumulative_bytes)
zip_buf.seek(0)

st.success("Reports ready! Monthly groups (with Guzman split) + Cumulative range report.")
st.download_button(
    "Download ZIP (Monthly + Cumulative)",
    data=zip_buf.getvalue(),
    file_name=f"ADA_Reports_{ts}.zip",
    mime="application/zip"
)

# Optional: also expose the individual files
with st.expander("Download individual files"):
    st.download_button(
        "Download Monthly Groups (xlsx)",
        data=monthly_bytes,
        file_name=f"ADA_Monthly_Groups_GuzmanSplit_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="mon"
    )
    st.download_button(
        "Download Cumulative (xlsx)",
        data=cumulative_bytes,
        file_name=f"ADA_Cumulative_{_month_name(start_m)}_{_month_name(end_m)}_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="cum"
    )


