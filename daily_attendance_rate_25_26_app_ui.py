import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
from pathlib import Path
from io import BytesIO

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

# ---- COLORS ----
BLUE = "#2E75B6"
RED = "#C00000"
GREY_HEADER = "#E6E6E6"
LIGHT_GRID = "#D9D9D9"
DARK_TEXT = "#1F2937"

PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"
logo_path = Path("header_logo.png")

def _month_name(m: int) -> str:
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _month_sort_key(name: str) -> int:
    # School-year order: Aug..Dec + Jan..Jun
    order = list(range(8, 13)) + list(range(1, 7))
    try:
        m = list(calendar.month_name).index(name)
        return order.index(m)
    except Exception:
        return 99

def _weighted_avg_rate(df: pd.DataFrame) -> float:
    """Weighted avg of Attendance Rate (0..100) by Current."""
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    weights = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    rates   = df["Attendance Rate"].fillna(0).astype(float)
    total_w = float(weights.sum())
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

# ---------- HEADER ----------
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    if logo_path.exists():
        st.image(str(logo_path), width=160)
    st.markdown("<h2 style='text-align:center;margin:8px 0;'>Daily Attendance Rate 25-26</h2>", unsafe_allow_html=True)
st.divider()

uploaded = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("Upload the Enrollment workbook to begin.")
    st.stop()

file_bytes = uploaded.read()
if not file_bytes:
    st.error("Uploaded file is empty.")
    st.stop()

# Inspect sheets
try:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Unable to read workbook. Error: {e}")
    st.stop()

# Choose sheet (auto-use preferred if present)
if len(sheet_names) == 1:
    use_sheet = sheet_names[0]
    st.success(f"Using sheet: **{use_sheet}**")
else:
    use_sheet = PREFERRED_SHEET if PREFERRED_SHEET in sheet_names else st.selectbox("Choose sheet", options=sheet_names, index=0)

# Read with auto header detection: try header=1 then header=0 (no UI control)
df = None
read_err = None
for hdr in (1, 0):
    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=hdr)
        if not df.empty:
            break
    except Exception as e:
        read_err = e
if df is None or df.empty:
    st.error(f"Failed to read sheet '{use_sheet}'. Error: {read_err}")
    st.stop()

# ---------- CLEAN ----------
df.columns = [str(c).strip() for c in df.columns]
df = df.rename(columns={
    "Unnamed: 6": "Funded",
    "Unnamed: 8": "Current",
    "Unnamed: 9": "Attendance Rate",
})
base_cols = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
df = df[[c for c in base_cols if c in df.columns]]

# numeric coercions
for c in ['Year', 'Month', 'Funded', 'Current', 'Attendance Rate']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce')

# Remove pre-existing subtotals/totals from source:
#  - drop rows where Class Name == TOTAL
#  - drop rows where Class Name is blank/NaN (green subtotal lines)
if 'Class Name' in df.columns:
    mask_total = df['Class Name'].astype(str).str.upper().eq('TOTAL')
    mask_blank = df['Class Name'].astype(str).str.strip().eq('') | df['Class Name'].isna()
    df = df[~(mask_total | mask_blank)].copy()

# Month selection
months_present = []
if 'Month' in df.columns:
    parsed = set()
    for m in df['Month'].dropna().unique():
        try: parsed.add(int(float(m)))
        except: pass
    months_present = sorted(parsed)

class_rows = df[df['Month'].isin(months_present)].copy()
month_names = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_sort_key)
if not month_names:
    st.error("No valid Month values found.")
    st.stop()

sel_month = st.selectbox("Month", options=month_names, index=len(month_names)-1)
sel_m = list(calendar.month_name).index(sel_month) if sel_month in list(calendar.month_name) else None
m_df = class_rows[class_rows['Month'] == sel_m].copy()

# --- Build ADA: class rows + ONE TOTAL per center + agency overall ---
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

# ---------- EXPORT (Excel with Dashboard) ----------
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    now = datetime.now(ZoneInfo("America/Chicago"))

    # Sheet 1: ADA
    ada_sheet = "ADA"
    ada.to_excel(writer, index=False, sheet_name=ada_sheet, startrow=3)
    wb = writer.book
    ws = writer.sheets[ada_sheet]

    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
    pct_fmt    = wb.add_format({"num_format": "0.00%"})
    bold_fmt   = wb.add_format({"bold": True})
    title_fmt  = wb.add_format({"bold": True, "font_size": 16})
    ts_fmt     = wb.add_format({"italic": True, "font_color": "#7F7F7F"})

    header_row = 3
    for c, name in enumerate(ada.columns):
        ws.write(header_row, c, name, header_fmt)
    ws.set_row(header_row, 26)

    # Autosize
    for i, col in enumerate(ada.columns):
        width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2))
        ws.set_column(i, i, width)

    # Attendance Rate -> decimal + bold TOTAL rows
    if 'Attendance Rate' in ada.columns:
        ci = ada.columns.get_loc('Attendance Rate')
        ki = ada.columns.get_loc('Class Name')
        for r in range(len(ada)):
            excel_r = r + header_row + 1
            v = ada.iloc[r, ci]
            ws.write_number(excel_r, ci, (float(v)/100.0) if pd.notna(v) else 0, pct_fmt)
            if str(ada.iloc[r, ki]).upper() == "TOTAL":
                ws.set_row(excel_r, None, bold_fmt)

    # Freeze + AutoFilter over full table
    ws.freeze_panes(header_row + 1, 0)
    last_col = len(ada.columns) - 1
    last_row = header_row + len(ada)
    ws.autofilter(header_row, 0, last_row, last_col)

    # Title + timestamp
    ws.merge_range(0, 1, 1, last_col, f"Daily Attendance Rate 25-26 — {sel_month}", title_fmt)
    ws.merge_range(2, 1, 2, last_col, f"Exported {now.strftime('%m/%d/%Y %I:%M %p %Z')}", ts_fmt)
    if logo_path.exists():
        ws.insert_image(0, 0, str(logo_path), {"x_scale": 0.45, "y_scale": 0.45, "x_offset": 4, "y_offset": 4})

    # Sheet 2: Dashboard
    dash = wb.add_worksheet("Dashboard")
    dash.set_zoom(125)

    # Title block
    big = wb.add_format({"bold": True, "font_size": 28, "font_color": BLUE})
    dash.write(0, 1, "Daily Attendance Dashboard", big)
    if logo_path.exists():
        dash.insert_image(0, 0, str(logo_path), {"x_scale": 0.40, "y_scale": 0.40, "x_offset": 4, "y_offset": 2})
    dash.write(2, 1, f"Month: {sel_month}")
    dash.write(3, 1, f"Exported {now.strftime('%m/%d/%Y %I:%M %p %Z')}")

    # Left table (sorted high->low) with AutoFilter
    left_r, left_c = 6, 1
    head    = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1, "align": "left"})
    cell    = wb.add_format({"border": 1})
    cell_b  = wb.add_format({"border": 1, "bold": True})
    cell_pct= wb.add_format({"border": 1, "num_format": "0.00%"})

    dash.write(left_r, left_c,   "Center Name",   head)
    dash.write(left_r, left_c+1, "Attendance %",  head)

    centers = (
        m_df.groupby("Center Name", dropna=False)
            .apply(_weighted_avg_rate)
            .reset_index(name="Attendance %")
            .sort_values("Attendance %", ascending=False)
    )

    for i, row in centers.iterrows():
        dash.write_string(left_r + 1 + i, left_c, "" if pd.isna(row["Center Name"]) else str(row["Center Name"]), cell_b if i == 0 else cell)
        dash.write_number(left_r + 1 + i, left_c + 1, (row["Attendance %"] / 100.0) if pd.notna(row["Attendance %"]) else 0, cell_pct)

    last_tab_row = left_r + len(centers)
    dash.autofilter(left_r, left_c, last_tab_row, left_c + 1)
    dash.set_column(left_c, left_c, 30)
    dash.set_column(left_c + 1, left_c + 1, 16)

    # Month/Agency mini table (a few rows below, detached)
    s_head = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1})
    s_cell = wb.add_format({"border": 1})
    s_pct  = wb.add_format({"border": 1, "num_format": "0.00%"})

    mini_r = last_tab_row + 3  # a few blank rows after the center table
    dash.write(mini_r,   left_c,   "Month",             s_head)
    dash.write(mini_r,   left_c+1, "Agency Overall %",  s_head)
    dash.write(mini_r+1, left_c,   sel_month,           s_cell)
    dash.write_number(mini_r+1, left_c+1, (agency_overall/100.0) if pd.notna(agency_overall) else 0, s_pct)

    # Bar chart (>=95% red) — title "Attendance Rate — {Month}"
    vals = [(float(v) / 100.0 if pd.notna(v) else 0.0) for v in centers["Attendance %"].tolist()]
    points = [{"fill": {"color": (RED if v >= 0.95 else BLUE)}} for v in vals]

    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name": f"Attendance Rate — {sel_month}",
        "categories": ["Dashboard", left_r + 1, left_c, last_tab_row, left_c],
        "values":     ["Dashboard", left_r + 1, left_c + 1, last_tab_row, left_c + 1],
        "data_labels": {"value": True, "num_format": "0.00%", "font": {"bold": True, "size": 9}},
        "points": points,
    })
    bar.set_y_axis({
        "num_format": "0.00%",
        "min": 0.86, "max": 0.98,
        "major_gridlines": {"visible": True, "line": {"color": LIGHT_GRID}},
        "minor_gridlines": {"visible": False},
        "font": {"size": 9, "color": DARK_TEXT},
    })
    bar.set_x_axis({
        "label_position": "low",
        "num_font": {"size": 9, "rotation": -45},
        "major_tick_mark": "none",
        "minor_tick_mark": "none",
    })
    bar.set_legend({"none": True})
    bar.set_title({"name": f"Attendance Rate — {sel_month}"})
    bar.set_chartarea({"border": {"none": True}})
    bar.set_plotarea({"border": {"none": True}})
    dash.insert_chart(5, 5, bar, {"x_scale": 1.35, "y_scale": 1.22})

    # Agency Overall KPI — rounded rectangle with centered title + percent
    kpi_text  = f"{agency_overall:.2f}%"
    kpi_title = "Agency Overall"
    try:
        dash.insert_shape(1, 12, {
            "type": "rounded_rectangle",
            "width": 260, "height": 100,
            "text": f"{kpi_title}\n{kpi_text}",
            "fill": {"color": RED},
            "line": {"color": RED},
            "font": {"bold": True, "color": "white", "size": 20},
            "align": {"vertical": "vcenter", "horizontal": "center"},
        })
    except Exception:
        dash.insert_textbox(1, 12, f"{kpi_title}\n{kpi_text}", {
            "width": 260, "height": 100,
            "fill": {"color": RED},
            "line": {"color": RED},
            "font": {"bold": True, "color": "white", "size": 20},
            "align": {"vertical": "vcenter", "horizontal": "center"},
        })

    # Trend table & chart (kept, with callout-like data labels)
    # Build monthly agency-overall series across the school year
    mo = []
    for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
        mm = int(m)
        mdf = class_rows[class_rows['Month'] == mm]
        mo.append({"Month Name": _month_name(mm), "Agency Overall %": _weighted_avg_rate(mdf) / 100.0})
    mo_df = pd.DataFrame(mo)

    tr_r = mini_r + 5  # a bit below the mini table
    tr_c = left_c

    dash.write(tr_r,   tr_c,   "Month",             s_head)
    dash.write(tr_r,   tr_c+1, "Agency Overall %",  s_head)
    for i, row in mo_df.iterrows():
        dash.write_string(tr_r + 1 + i, tr_c,   str(row["Month Name"]), s_cell)
        dash.write_number(tr_r + 1 + i, tr_c+1, float(row["Agency Overall %"]) if pd.notna(row["Agency Overall %"]) else 0, s_pct)

    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Agency Overall %",
        "categories": ["Dashboard", tr_r + 1, tr_c,   tr_r + len(mo_df), tr_c],
        "values":     ["Dashboard", tr_r + 1, tr_c+1, tr_r + len(mo_df), tr_c+1],
        "marker": {"type": "circle", "size": 6},
        "line": {"color": RED, "width": 2},
        # callout-like labels (white fill + light border)
        "data_labels": {"value": True, "num_format": "0.00%",
                        "fill": {"color": "#FFFFFF"},
                        "border": {"color": LIGHT_GRID},
                        "font": {"size": 9}},
    })
    line.set_y_axis({
        "num_format": "0.00%",
        "min": 0.86, "max": 0.98,
        "major_gridlines": {"visible": True, "line": {"color": LIGHT_GRID}},
        "font": {"size": 9, "color": DARK_TEXT},
    })
    line.set_x_axis({"label_position": "low", "num_font": {"size": 9}})
    line.set_legend({"none": True})
    line.set_title({"name": "Agency Attendance Trend — 2025–2026"})
    dash.insert_chart(tr_r, 5, line, {"x_scale": 1.35, "y_scale": 1.12})

# Download
st.download_button(
    "Download Report",
    data=output.getvalue(),
    file_name=f"ADA_Dashboard_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
