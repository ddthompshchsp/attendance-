
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
import re

src_path = Path("/mnt/data/V12POP_ERSEA_Enrollment-1122-2.xlsx")
assert src_path.exists(), "Source Excel not found at /mnt/data/V12POP_ERSEA_Enrollment-1122-2.xlsx"

# ---------- Helpers ----------
def _month_name(m: int) -> str:
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _weighted_avg_rate(df: pd.DataFrame) -> float:
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    weights = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    rates   = df["Attendance Rate"].fillna(0).astype(float)
    total_w = float(weights.sum())
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

def _center_block(center_df: pd.DataFrame) -> pd.DataFrame:
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

def _canon(x: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', str(x).lower())

# User simple names for grouping (we'll match by token containment, preserving full names in output)
SIMPLE_ALIASES = {"fairs":"farias","sequin":"seguin","chaps":"chapa"}
def _norm_simple(s: str) -> str:
    k = _canon(s)
    return SIMPLE_ALIASES.get(k, k)

CENTER_GROUPS_SIMPLE = {
    "Group 1": [
        "Donna","Mission","Monte Alto","Sam Fordyce","San Carlos","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    "Group 2": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    "Group 3": ["Mercedes","Edinburg"],
    "Group 4": ["Guzman"]
}

def center_matches_simple(full_name: str, simple: str) -> bool:
    return _norm_simple(simple) in _canon(full_name)

def _detect_age_tag(text: str) -> str:
    t = str(text).lower()
    if re.search(r'(^|[^0-9])3($|[^0-9])|3yo|3-?yr|3\s*year', t):
        return " (3yo)"
    if re.search(r'(^|[^0-9])4($|[^0-9])|4yo|4-?yr|4\s*year', t):
        return " (4yo)"
    return ""

def _append_guzman_age(df: pd.DataFrame) -> pd.DataFrame:
    if "Center Name" not in df.columns or "Class Name" not in df.columns:
        return df
    out = df.copy()
    mask = out["Center Name"].astype(str).str.lower().str.contains("guzman")
    out.loc[mask, "Class Name"] = (
        out.loc[mask, "Class Name"]
        .astype(str)
        .apply(lambda s: s + _detect_age_tag(s) if _detect_age_tag(s) else s)
    )
    return out

# ---------- Load & clean ----------
xls = pd.ExcelFile(src_path)
use_sheet = "V12POP_ERSEA_Enrollment" if "V12POP_ERSEA_Enrollment" in xls.sheet_names else xls.sheet_names[0]

# try header rows
df = None
for hdr in (1, 0):
    try:
        tmp = pd.read_excel(src_path, sheet_name=use_sheet, header=hdr)
        if not tmp.empty:
            df = tmp
            break
    except Exception:
        pass

assert df is not None and not df.empty, "Could not read a non-empty sheet from the workbook."

df.columns = [str(c).strip() for c in df.columns]
df = df.rename(columns={
    "Unnamed: 6": "Funded",
    "Unnamed: 8": "Current",
    "Unnamed: 9": "Attendance Rate",
})

base_cols = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
df = df[[c for c in base_cols if c in df.columns]].copy()

for c in ['Year', 'Month', 'Funded', 'Current', 'Attendance Rate']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce')

# remove totals/blank class
if 'Class Name' in df.columns:
    mask_total = df['Class Name'].astype(str).str.upper().eq('TOTAL')
    mask_blank = df['Class Name'].astype(str).str.strip().eq('') | df['Class Name'].isna()
    df = df[~(mask_total | mask_blank)].copy()

# months present; we will keep both 9 and 10 if available
months_present = sorted({int(m) for m in df['Month'].dropna().astype(int).unique()})
target_months = [m for m in [9,10] if m in months_present] or [months_present[-1]]
sub_df = df[df['Month'].isin(target_months)].copy()

# ADA for latest month (for ADA & Dashboard)
latest_m = max(target_months)
m_df = df[df['Month'] == latest_m].copy()

# ADA build (latest month only, like before)
blocks = [_center_block(g) for _, g in m_df.groupby("Center Name", dropna=False)]
ada = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(columns=base_cols)
agency_overall = _weighted_avg_rate(m_df)
ada = pd.concat([ada, pd.DataFrame([{
    "Year": m_df['Year'].dropna().iloc[0] if m_df['Year'].notna().any() else np.nan,
    "Month": latest_m,
    "Center Name": "HCHSP (Overall)",
    "Class Name": "TOTAL",
    "Funded": m_df['Funded'].sum(min_count=1),
    "Current": m_df['Current'].sum(min_count=1),
    "Attendance Rate": float(agency_overall) if pd.notna(agency_overall) else np.nan,
}])], ignore_index=True)

# ---------- Write Excel with updated group naming and condensed summary tables ----------
out_path = Path(f"/mnt/data/ADA_Dashboard_GROUPS_FILTER_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx")

with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
    wb = writer.book
    BLUE = "#2E75B6"; RED = "#C00000"; GREY_HEADER = "#E6E6E6"; LIGHT_GRID = "#D9D9D9"

    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
    pct_fmt    = wb.add_format({"num_format": "0.00%"})
    bold_fmt   = wb.add_format({"bold": True})
    title_fmt  = wb.add_format({"bold": True, "font_size": 16})
    s_head     = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1, "font_size": 9})
    s_cell     = wb.add_format({"border": 1, "font_size": 9})
    s_pct      = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})

    # ADA sheet (latest month)
    ada.to_excel(writer, index=False, sheet_name="ADA", startrow=3)
    ws = writer.sheets["ADA"]
    header_row = 3
    for c, name in enumerate(ada.columns):
        ws.write(header_row, c, name, header_fmt)
    ws.set_row(header_row, 26)
    for i, col in enumerate(ada.columns):
        width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2))
        ws.set_column(i, i, width)
    if 'Attendance Rate' in ada.columns:
        ci = ada.columns.get_loc('Attendance Rate')
        ki = ada.columns.get_loc('Class Name')
        for r in range(len(ada)):
            excel_r = r + header_row + 1
            v = ada.iloc[r, ci]
            ws.write_number(excel_r, ci, (float(v)/100.0) if pd.notna(v) else 0, pct_fmt)
            if str(ada.iloc[r, ki]).upper() == "TOTAL":
                ws.set_row(excel_r, None, bold_fmt)
    last_col = len(ada.columns) - 1
    last_row = header_row + len(ada)
    ws.freeze_panes(header_row + 1, 0)
    ws.autofilter(header_row, 0, last_row, last_col)
    ws.merge_range(0, 1, 1, last_col, f"Daily Attendance Rate 25-26 — {_month_name(latest_m)}", title_fmt)

    # Dashboard (latest month same as before)
    dash = wb.add_worksheet("Dashboard")
    dash.write(0, 1, "Daily Attendance Dashboard", wb.add_format({"bold": True, "font_size": 24, "font_color": BLUE}))
    dash.write(2, 1, f"Month: {_month_name(latest_m)}")
    dash.write(3, 1, f"Exported {datetime.now(ZoneInfo('America/Chicago')).strftime('%m/%d/%Y %I:%M %p %Z')}")

    # Ranked table
    centers = (
        m_df.groupby("Center Name", dropna=False)
            .apply(_weighted_avg_rate)
            .reset_index(name="Attendance %")
            .sort_values("Attendance %", ascending=False)
    )
    left_r, left_c = 6, 1
    dash.write(left_r, left_c,   "Center Name",   s_head)
    dash.write(left_r, left_c+1, "Attendance %",  s_head)
    for i, row in centers.iterrows():
        dash.write_string(left_r + 1 + i, left_c, "" if pd.isna(row["Center Name"]) else str(row["Center Name"]), s_cell)
        dash.write_number(left_r + 1 + i, left_c + 1, (row["Attendance %"] / 100.0) if pd.notna(row["Attendance %"]) else 0, s_pct)
    last_tab_row = left_r + len(centers)
    dash.autofilter(left_r, left_c, last_tab_row, left_c + 1)
    dash.set_column(left_c, left_c, 30)
    dash.set_column(left_c + 1, left_c + 1, 16)

    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name": f"Attendance Rate — {_month_name(latest_m)}",
        "categories": ["Dashboard", left_r + 1, left_c, last_tab_row, left_c],
        "values":     ["Dashboard", left_r + 1, left_c + 1, last_tab_row, left_c + 1],
        "data_labels": {"value": True, "num_format": "0.00%", "font": {"bold": True, "size": 9}},
    })
    bar.set_y_axis({"num_format": "0.00%"})
    bar.set_title({"name": f"Attendance Rate — {_month_name(latest_m)}"})
    dash.insert_chart(5, 5, bar, {"x_scale": 1.25, "y_scale": 1.15})

    dash.write(last_tab_row + 3, left_c,   "Agency Overall %", s_head)
    dash.write_number(last_tab_row + 3, left_c+1, (agency_overall/100.0) if pd.notna(agency_overall) else 0, s_pct)

    # ---------- Grouped campus sheets with condensed tables & month filter ----------
    USED_FULL = set()
    grouped_src = _append_guzman_age(sub_df)  # include both Sept/Oct where available

    def write_group_sheet(sheet_name: str, simples: list, df_src: pd.DataFrame):
        ws_grp = wb.add_worksheet(sheet_name[:31])

        # Match full center names by tokens
        present_full = df_src['Center Name'].dropna().unique().tolist()
        matched_full = []
        for simple in simples:
            token = _norm_simple(simple)
            for full in present_full:
                if center_matches_simple(full, token):
                    key = _canon(full)
                    if key not in { _canon(x) for x in matched_full } and key not in USED_FULL:
                        matched_full.append(full)

        if not matched_full:
            ws_grp.write(0, 0, f"No data for months {', '.join(_month_name(m) for m in target_months)} for: {', '.join(simples)}")
            return

        # Build ADA-style blocks for each month separately, stacking them all together
        # (We keep ADA tables per center per month, so filtering remains natural)
        start_row = 3
        for fn in matched_full:
            for m in target_months:
                g = df_src[(df_src["Center Name"].astype(str) == fn) & (df_src["Month"] == m)]
                if g.empty:
                    continue
                block = _center_block(g)

                # header for each block
                ws_grp.write(start_row - 2, 0, f"{fn} — {_month_name(m)}", bold_fmt)
                for c, name in enumerate(block.columns):
                    ws_grp.write(start_row, c, name, header_fmt)
                ws_grp.set_row(start_row, 26)

                # values
                for r in range(len(block)):
                    for c, col in enumerate(block.columns):
                        val = block.iloc[r, c]
                        if col == "Attendance Rate":
                            ws_grp.write_number(start_row + 1 + r, c, (float(val)/100.0) if pd.notna(val) else 0, pct_fmt)
                        else:
                            ws_grp.write(start_row + 1 + r, c, val)
                    if str(block.iloc[r]["Class Name"]).upper() == "TOTAL":
                        ws_grp.set_row(start_row + 1 + r, None, bold_fmt)

                # autosize
                for i, col in enumerate(block.columns):
                    width = max(12, min(44, int(block[col].astype(str).map(len).max()) + 2))
                    ws_grp.set_column(i, i, width)

                start_row = start_row + 1 + len(block) + 2  # spacing between blocks

        # freeze top after first header row
        ws_grp.freeze_panes(4, 0)

        # ---------- Condensed summary table (for charts + filter by Month) ----------
        # Build Month-Center totals from df_src filtered to matched_full
        filt = df_src[df_src["Center Name"].isin(matched_full)].copy()
        sums = (
            filt.groupby(["Month","Center Name"], dropna=False)
                .apply(_weighted_avg_rate)
                .reset_index(name="Attendance %")
                .sort_values(["Month","Attendance %"], ascending=[True, False])
        )

        # Small, compact table
        small_r, small_c = start_row, 0
        ws_grp.write(small_r,   small_c,   "Month",        s_head)
        ws_grp.write(small_r,   small_c+1, "Center Name",  s_head)
        ws_grp.write(small_r,   small_c+2, "Attendance %", s_head)
        for i, row in sums.iterrows():
            ws_grp.write_string(small_r + 1 + i, small_c,   _month_name(int(row["Month"])) if pd.notna(row["Month"]) else "", s_cell)
            ws_grp.write_string(small_r + 1 + i, small_c+1, str(row["Center Name"]), s_cell)
            ws_grp.write_number(small_r + 1 + i, small_c+2, (row["Attendance %"]/100.0 if pd.notna(row["Attendance %"]) else 0), s_pct)

        last_row_small = small_r + len(sums)
        ws_grp.autofilter(small_r, small_c, last_row_small, small_c + 2)
        ws_grp.set_column(small_c, small_c, 10)   # Month
        ws_grp.set_column(small_c+1, small_c+1, 36)  # Center
        ws_grp.set_column(small_c+2, small_c+2, 14)  # %

        # Two charts: one for Sept, one for Oct (if present)
        months_for_charts = [m for m in target_months if m in sums["Month"].unique()]
        chart_col_offset = 5
        for idx, m in enumerate(months_for_charts):
            m_sums = sums[sums["Month"] == m]
            if m_sums.empty:
                continue
            # place chart to the right of the condensed table
            chart = wb.add_chart({"type": "column"})
            # categories and values reference the current small table area filtered by month block
            # compute relative ranges
            start_idx = sums[sums["Month"] == m].index.min()
            end_idx   = sums[sums["Month"] == m].index.max()
            # We can't rely on absolute index, so rebuild start/end row offsets
            # Let's compute row offsets by enumerating
            rows_for_m = [i for i, (_, r) in enumerate(sums.iterrows()) if int(r["Month"]) == int(m)]
            if not rows_for_m:
                continue
            start_off = rows_for_m[0]
            end_off   = rows_for_m[-1]
            chart.add_series({
                "name": f"{_month_name(m)} Attendance %",
                "categories": [sheet_name, small_r + 1 + start_off, small_c + 1, small_r + 1 + end_off, small_c + 1],
                "values":     [sheet_name, small_r + 1 + start_off, small_c + 2, small_r + 1 + end_off, small_c + 2],
                "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
            })
            chart.set_y_axis({"num_format": "0.00%"})
            chart.set_title({"name": f"{_month_name(m)} — Attendance %"})
            ws_grp.insert_chart(small_r, small_c + chart_col_offset + idx*8, chart, {"x_scale": 1.1, "y_scale": 1.0})

        # Mark used centers
        for fn in matched_full:
            USED_FULL.add(_canon(fn))

    # Create sheets in requested order, now named Group 1..4, using both Sept/Oct data
    for sheet_name, simple_list in CENTER_GROUPS_SIMPLE.items():
        write_group_sheet(sheet_name, simple_list, grouped_src)

out_path

