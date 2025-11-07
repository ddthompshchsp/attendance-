
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

# User's simple campus names, with typo aliases normalized
SIMPLE_ALIASES = {
    "fairs": "farias",
    "sequin": "seguin",
    "chaps": "chapa",
}
def _norm_simple(s: str) -> str:
    k = _canon(s)
    return SIMPLE_ALIASES.get(k, k)

CENTER_GROUPS_SIMPLE = {
    "Grp1_Donna_to_SanJuan": [
        "Donna","Mission","Monte Alto","Sam Fordyce","San Carlos","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    "Grp2_Chapa_to_Longoria": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    "Grp3_Mercedes_Edinburg": ["Mercedes","Edinburg"],
    "Grp4_Guzman": ["Guzman"]
}

def center_matches_simple(full_name: str, simple: str) -> bool:
    """Return True if the normalized token of 'simple' appears inside the normalized full center string."""
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

# month select: pick the latest month number present
months_present = []
if 'Month' in df.columns:
    parsed = set()
    for m in df['Month'].dropna().unique():
        try:
            parsed.add(int(float(m)))
        except:
            pass
    months_present = sorted(parsed)

assert months_present, "No valid Month values found."
sel_m = months_present[-1]

m_df = df[df['Month'] == sel_m].copy()

# ADA build
blocks = [_center_block(g) for _, g in m_df.groupby("Center Name", dropna=False)]
ada = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(columns=base_cols)
agency_overall = _weighted_avg_rate(m_df)
ada = pd.concat([ada, pd.DataFrame([{
    "Year": m_df['Year'].dropna().iloc[0] if m_df['Year'].notna().any() else np.nan,
    "Month": sel_m,
    "Center Name": "HCHSP (Overall)",
    "Class Name": "TOTAL",
    "Funded": m_df['Funded'].sum(min_count=1),
    "Current": m_df['Current'].sum(min_count=1),
    "Attendance Rate": float(agency_overall) if pd.notna(agency_overall) else np.nan,
}])], ignore_index=True)

# ---------- Write Excel with updated grouping ----------
out_path = Path(f"/mnt/data/ADA_Dashboard_FINAL_MATCHED_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx")

with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
    wb = writer.book
    BLUE = "#2E75B6"; RED = "#C00000"; GREY_HEADER = "#E6E6E6"; LIGHT_GRID = "#D9D9D9"; DARK_TEXT = "#1F2937"

    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
    pct_fmt    = wb.add_format({"num_format": "0.00%"})
    bold_fmt   = wb.add_format({"bold": True})
    title_fmt  = wb.add_format({"bold": True, "font_size": 16})
    s_head     = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1})
    s_cell     = wb.add_format({"border": 1})
    s_pct      = wb.add_format({"border": 1, "num_format": "0.00%"})

    # ADA sheet
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
    ws.merge_range(0, 1, 1, last_col, f"Daily Attendance Rate 25-26 — {_month_name(sel_m)}", title_fmt)

    # Dashboard quick
    dash = wb.add_worksheet("Dashboard")
    dash.write(0, 1, "Daily Attendance Dashboard", wb.add_format({"bold": True, "font_size": 24, "font_color": BLUE}))
    dash.write(2, 1, f"Month: {_month_name(sel_m)}")
    dash.write(3, 1, f"Exported {datetime.now(ZoneInfo('America/Chicago')).strftime('%m/%d/%Y %I:%M %p %Z')}")

    left_r, left_c = 6, 1
    dash.write(left_r, left_c,   "Center Name",   s_head)
    dash.write(left_r, left_c+1, "Attendance %",  s_head)

    centers = (
        m_df.groupby("Center Name", dropna=False)
            .apply(_weighted_avg_rate)
            .reset_index(name="Attendance %")
            .sort_values("Attendance %", ascending=False)
    )

    for i, row in centers.iterrows():
        dash.write_string(left_r + 1 + i, left_c, "" if pd.isna(row["Center Name"]) else str(row["Center Name"]), s_cell)
        dash.write_number(left_r + 1 + i, left_c + 1, (row["Attendance %"] / 100.0) if pd.notna(row["Attendance %"]) else 0, s_pct)

    last_tab_row = left_r + len(centers)
    dash.autofilter(left_r, left_c, last_tab_row, left_c + 1)
    dash.set_column(left_c, left_c, 30)
    dash.set_column(left_c + 1, left_c + 1, 16)

    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name": f"Attendance Rate — {_month_name(sel_m)}",
        "categories": ["Dashboard", left_r + 1, left_c, last_tab_row, left_c],
        "values":     ["Dashboard", left_r + 1, left_c + 1, last_tab_row, left_c + 1],
        "data_labels": {"value": True, "num_format": "0.00%", "font": {"bold": True, "size": 9}},
    })
    bar.set_y_axis({"num_format": "0.00%"})
    bar.set_title({"name": f"Attendance Rate — {_month_name(sel_m)}"})
    dash.insert_chart(5, 5, bar, {"x_scale": 1.25, "y_scale": 1.15})

    dash.write(last_tab_row + 3, left_c,   "Agency Overall %", s_head)
    dash.write_number(last_tab_row + 3, left_c+1, (agency_overall/100.0) if pd.notna(agency_overall) else 0, s_pct)

    # ---------- Grouped campus sheets (match by token, keep full names) ----------
    USED_FULL = set()  # track normalized full center names already placed
    grouped_src = _append_guzman_age(m_df)

    def write_group_sheet(sheet_name: str, simples: list, df_src: pd.DataFrame):
        ws_name = sheet_name[:31]
        ws_grp = wb.add_worksheet(ws_name)

        # Map simple tokens to matching full-center names present in data, excluding already used
        present_full = df_src['Center Name'].dropna().unique().tolist()
        matched_full = []
        for simple in simples:
            token = _norm_simple(simple)
            # find all centers whose normalized name contains this token
            for full in present_full:
                if center_matches_simple(full, token):
                    k = _canon(full)
                    if k not in { _canon(x) for x in matched_full } and k not in USED_FULL:
                        matched_full.append(full)

        if not matched_full:
            ws_grp.write(0, 0, f"No data for this month for: {', '.join(simples)}")
            return

        # Build ADA-style blocks in the order of matched_full
        out_tables = []
        for fn in matched_full:
            g = df_src[df_src["Center Name"].astype(str) == fn]
            if g.empty:
                continue
            out_tables.append(_center_block(g))

        if not out_tables:
            ws_grp.write(0, 0, f"No rows after filtering for: {', '.join(simples)}")
            return

        grp_df = pd.concat(out_tables, ignore_index=True)

        # header (ADA style)
        start_row = 3
        for c, name in enumerate(grp_df.columns):
            ws_grp.write(start_row, c, name, header_fmt)
        ws_grp.set_row(start_row, 26)

        # values
        for r in range(len(grp_df)):
            for c, col in enumerate(grp_df.columns):
                val = grp_df.iloc[r, c]
                if col == "Attendance Rate":
                    ws_grp.write_number(start_row + 1 + r, c, (float(val)/100.0) if pd.notna(val) else 0, pct_fmt)
                else:
                    ws_grp.write(start_row + 1 + r, c, val)
            if str(grp_df.iloc[r]["Class Name"]).upper() == "TOTAL":
                ws_grp.set_row(start_row + 1 + r, None, bold_fmt)

        # autosize
        for i, col in enumerate(grp_df.columns):
            width = max(12, min(44, int(grp_df[col].astype(str).map(len).max()) + 2))
            ws_grp.set_column(i, i, width)

        # freeze + filter
        last_col = len(grp_df.columns) - 1
        last_row = start_row + len(grp_df)
        ws_grp.freeze_panes(start_row + 1, 0)
        ws_grp.autofilter(start_row, 0, last_row, last_col)

        # Summary (TOTAL rows) + chart
        centers_sum = (
            grp_df[grp_df["Class Name"].astype(str).str.upper().eq("TOTAL")]
            .copy()
            .sort_values("Attendance Rate", ascending=False)
        )
        if not centers_sum.empty:
            sum_r, sum_c = last_row + 2, 0
            ws_grp.write(sum_r,   sum_c,   "Center Name", s_head)
            ws_grp.write(sum_r,   sum_c+1, "Attendance %", s_head)
            for i, row in centers_sum.iterrows():
                ws_grp.write_string(sum_r + 1 + i, sum_c,   str(row["Center Name"]), s_cell)
                ws_grp.write_number(sum_r + 1 + i, sum_c+1, (float(row["Attendance Rate"])/100.0 if pd.notna(row["Attendance Rate"]) else 0), s_pct)

            chart = wb.add_chart({"type": "column"})
            chart.add_series({
                "name": f"{sheet_name} — Attendance %",
                "categories": [ws_name, sum_r + 1, sum_c,   sum_r + len(centers_sum), sum_c],
                "values":     [ws_name, sum_r + 1, sum_c+1, sum_r + len(centers_sum), sum_c+1],
                "data_labels": {"value": True, "num_format": "0.00%", "font": {"bold": True, "size": 9}},
            })
            chart.set_y_axis({"num_format": "0.00%"})
            chart.set_legend({"none": True})
            chart.set_title({"name": f"{sheet_name} — Attendance %"})
            ws_grp.insert_chart(sum_r, sum_c + 3, chart, {"x_scale": 1.15, "y_scale": 1.05})

        # Mark used full names so they won't appear again on later sheets
        for fn in matched_full:
            USED_FULL.add(_canon(fn))

    # Create sheets in order using token matching, preserving full names in output
    write_group_sheet("Grp1_Donna_to_SanJuan", CENTER_GROUPS_SIMPLE["Grp1_Donna_to_SanJuan"], _append_guzman_age(m_df))
    write_group_sheet("Grp2_Chapa_to_Longoria", CENTER_GROUPS_SIMPLE["Grp2_Chapa_to_Longoria"], _append_guzman_age(m_df))
    write_group_sheet("Grp3_Mercedes_Edinburg", CENTER_GROUPS_SIMPLE["Grp3_Mercedes_Edinburg"], _append_guzman_age(m_df))
    write_group_sheet("Grp4_Guzman",            CENTER_GROUPS_SIMPLE["Grp4_Guzman"],            _append_guzman_age(m_df))

out_path

