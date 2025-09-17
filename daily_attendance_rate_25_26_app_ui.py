# Write an updated app that mimics the screenshot’s dashboard styling more closely.
# Key styling:
# - Big left-aligned "Daily Attendance Dashboard" title in dark blue
# - Category labels angled (~-45 degrees)
# - First bar red, others blue
# - Chart y-axis 86% .. 98% with light grey gridlines
# - KPI red textbox with white text on top-right
# - White banner-like textbox for "Attendance Rate - {Month}" above the bar chart
# - Bottom line chart in red with markers and data labels, same axis style
# - Left grey table with bold top center and zebra rows
# - Logos placed similarly (top-left in title block and small over bar chart)
#
# No drops anywhere. ADA is class rows + per-campus TOTAL + Agency TOTAL.
from pathlib import Path

code = r'''
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
from pathlib import Path
from io import BytesIO

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

# ---- COLORS / FONTS ----
BLUE = "#2E75B6"
RED = "#C00000"
GREY_HEADER = "#E6E6E6"
LIGHT_GRID = "#D9D9D9"
DARK_TEXT = "#1F2937"

PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"
logo_path = Path("header_logo.png")

def _month_name(m: int) -> str:
    import calendar
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _weighted_avg_rate(df: pd.DataFrame) -> float:
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    weights = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    rates   = df["Attendance Rate"].fillna(0).astype(float)  # 0..100
    total_w = float(weights.sum())
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

# ---------- HEADER ----------
h1, h2, h3 = st.columns([1,2,1])
with h2:
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
    st.markdown("<h2 style='text-align:center;margin:8px 0;'>Daily Attendance Rate 25-26</h2>", unsafe_allow_html=True)
st.divider()

uploaded = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("Upload the Enrollment workbook to begin.")
    st.stop()

file_bytes = uploaded.read()
if not file_bytes:
    st.error("Uploaded file is empty."); st.stop()

# Inspect sheets
try:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Unable to read workbook. Error: {e}")
    st.stop()

# Choose sheet
if len(sheet_names) == 1:
    use_sheet = sheet_names[0]
    st.success(f"Using sheet: **{use_sheet}**")
else:
    if PREFERRED_SHEET in sheet_names:
        use_sheet = PREFERRED_SHEET
        st.success(f"Using preferred sheet: **{use_sheet}**")
    else:
        use_sheet = st.selectbox("Choose sheet to read", options=sheet_names, index=0)

header_row = st.number_input("Header row (0-indexed). Use 1 if headers are on the 2nd row.", min_value=0, value=1, step=1)

try:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=int(header_row))
except Exception as e:
    st.error(f"Failed to read sheet '{use_sheet}'. Error: {e}")
    st.stop()

if df.empty:
    st.error("Selected sheet appears empty."); st.stop()

# ---------- CLEAN ----------
df.columns = [str(c).strip() for c in df.columns]
df = df.rename(columns={"Unnamed: 6":"Funded", "Unnamed: 8":"Current", "Unnamed: 9":"Attendance Rate"})
base_cols = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
df = df[[c for c in base_cols if c in df.columns]]
for c in ['Year','Month','Funded','Current','Attendance Rate']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce')

# Month select
months_present = []
if 'Month' in df.columns:
    parsed = set()
    for m in df['Month'].dropna().unique():
        try: parsed.add(int(float(m)))
        except: pass
    months_present = sorted(parsed)
class_rows = df[df['Month'].isin(months_present)]
month_names = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=lambda n: (list(range(8,13))+list(range(1,7))).index(list(calendar.month_name).index(n)) if n in list(calendar.month_name) else 99)
sel_month = st.selectbox("Dashboard month", options=month_names, index=len(month_names)-1 if month_names else 0)
sel_m = list(calendar.month_name).index(sel_month) if sel_month in list(calendar.month_name) else None
m_df = class_rows[class_rows['Month']==sel_m].copy()

# Build ADA with totals
def _center_block(center_df):
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

blocks = [ _center_block(g) for _, g in m_df.groupby("Center Name", dropna=False) ]
ada = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(columns=base_cols)

overall = _weighted_avg_rate(m_df)
ada = pd.concat([ada, pd.DataFrame([{
    "Year": m_df['Year'].dropna().iloc[0] if m_df['Year'].notna().any() else np.nan,
    "Month": sel_m,
    "Center Name": "HCHSP (Overall)",
    "Class Name": "TOTAL",
    "Funded": m_df['Funded'].sum(min_count=1),
    "Current": m_df['Current'].sum(min_count=1),
    "Attendance Rate": float(overall) if pd.notna(overall) else np.nan,
}])], ignore_index=True)

# ---------- EXPORT ----------
from xlsxwriter.utility import xl_rowcol_to_cell

output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    now = datetime.now(ZoneInfo("America/Chicago"))
    # ADA
    ada_sheet = "ADA"
    ada.to_excel(writer, index=False, sheet_name=ada_sheet, startrow=3)
    wb = writer.book
    ws = writer.sheets[ada_sheet]

    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align":"center","valign":"vcenter"})
    pct_fmt    = wb.add_format({"num_format":"0.00%"})
    bold_fmt   = wb.add_format({"bold": True})
    title_fmt  = wb.add_format({"bold": True, "font_size": 16})
    ts_fmt     = wb.add_format({"italic": True, "font_color": "#7F7F7F"})

    header_row = 3
    for c, name in enumerate(ada.columns):
        ws.write(header_row, c, name, header_fmt)
    ws.set_row(header_row, 26)

    for i, col in enumerate(ada.columns):
        width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2))
        ws.set_column(i, i, width)

    # Attendance Rate decimal
    if 'Attendance Rate' in ada.columns:
        ci = ada.columns.get_loc('Attendance Rate')
        ki = ada.columns.get_loc('Class Name')
        for r in range(len(ada)):
            excel_r = r+header_row+1
            v = ada.iloc[r, ci]
            ws.write_number(excel_r, ci, (float(v)/100.0) if pd.notna(v) else 0, pct_fmt)
            if str(ada.iloc[r, ki]).upper() == "TOTAL":
                ws.set_row(excel_r, None, bold_fmt)

    ws.freeze_panes(header_row+1, 0)
    last_col = len(ada.columns)-1
    ws.merge_range(0,1,1,last_col, f"Daily Attendance Rate 25-26 — {sel_month}", title_fmt)
    ws.merge_range(2,1,2,last_col, f"Exported {now.strftime('%m/%d/%Y %I:%M %p %Z')}", ts_fmt)
    if logo_path.exists():
        ws.insert_image(0,0,str(logo_path), {"x_scale":0.6,"y_scale":0.6,"x_offset":4,"y_offset":4})

    # DASHBOARD
    dash = wb.add_worksheet("Dashboard")
    dash.set_zoom(125)

    # Title block (left aligned)
    big = wb.add_format({"bold": True, "font_size": 28, "font_color": BLUE})
    dash.write(0,1,"Daily Attendance Dashboard", big)
    if logo_path.exists():
        dash.insert_image(0,0,str(logo_path), {"x_scale":0.5,"y_scale":0.5,"x_offset":4,"y_offset":2})
    dash.write(2,1, f"Month for bar chart: {sel_month}")
    dash.write(3,1, f"Exported {now.strftime('%m/%d/%Y %I:%M %p %Z')}")

    # Left table
    left_r, left_c = 6, 1
    head = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border":1, "align":"left"})
    cell = wb.add_format({"border":1})
    cell_b = wb.add_format({"border":1,"bold":True})
    cell_pct = wb.add_format({"border":1,"num_format":"0.00%"})
    dash.write(left_r, left_c, "Center Name", head); dash.write(left_r, left_c+1, "Attendance %", head)

    centers = m_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %").sort_values("Attendance %", ascending=False)
    for i, row in centers.iterrows():
        dash.write(left_r+1+i, left_c, row["Center Name"], cell_b if i==0 else cell)
        dash.write_number(left_r+1+i, left_c+1, (row["Attendance %"]/100.0) if pd.notna(row["Attendance %"]) else 0, cell_pct)

    dash.set_column(left_c, left_c, 30); dash.set_column(left_c+1, left_c+1, 16)

    # Bar chart (right) – mimic styles
    bar = wb.add_chart({"type":"column"})
    bar.add_series({
        "name": f"Attendance Rate - {sel_month}",
        "categories": ["Dashboard", left_r+1, left_c, left_r+len(centers), left_c],
        "values":     ["Dashboard", left_r+1, left_c+1, left_r+len(centers), left_c+1],
        "data_labels": {"value": True, "num_format":"0.00%","font":{"bold":True,"size":9}},
        "fill":{"color":BLUE},
        "border":{"color":BLUE},
        "points":[{"fill":{"color":RED}}] + [{}]*(max(len(centers)-1,0))
    })
    # Axes styling
    bar.set_y_axis({
        "num_format":"0.00%",
        "min":0.86, "max":0.98,
        "major_gridlines":{"visible":True, "line":{"color":LIGHT_GRID}},
        "minor_gridlines":{"visible":False},
        "font":{"size":9, "color":DARK_TEXT},
    })
    bar.set_x_axis({
        "label_position":"low",
        "num_font":{"size":9},
        "text_axis": True,
        "minor_tick_mark": "none",
        "major_tick_mark": "none",
        "name_font":{"size":10},
        "num_format": "General",
        "label_align": 1,  # left
        "rotation": -45,
    })
    bar.set_legend({"none": True})
    bar.set_title({"name": ""})
    # Plot area clean
    bar.set_chartarea({"border":{"none":True}})
    bar.set_plotarea({"border":{"none":True}})

    # Insert bar
    dash.insert_chart(5, 5, bar, {"x_scale":1.4,"y_scale":1.25})

    # White banner-like textbox for the chart title (shadow-like via border)
    dash.insert_textbox(4, 7, f"Attendance Rate - {sel_month}", {
        "width": 420, "height": 36,
        "font": {"bold": True, "size": 18, "color": DARK_TEXT},
        "align": {"vertical":"vcenter","horizontal":"center"},
        "fill": {"color": "#FFFFFF"},
        "line": {"color": LIGHT_GRID},
    })

    # Small logo overlay on chart area (optional)
    if logo_path.exists():
        dash.insert_image(5, 7, str(logo_path), {"x_scale":0.25,"y_scale":0.25,"x_offset":10,"y_offset":-16})

    # KPI red textbox
    if not np.isnan(overall):
        dash.insert_textbox(1, 12, f"Agency Overall\n{overall:.2f}%", {
            "width": 220, "height": 80,
            "font": {"bold": True, "color": "white", "size": 20},
            "align": {"vertical":"vcenter","horizontal":"center"},
            "fill": {"color": RED},
            "line": {"color": RED},
        })

    # Bottom trend table + chart
    mo = []
    for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
        mm = int(m)
        mdf = class_rows[class_rows['Month']==mm]
        mo.append({"Month Name": _month_name(mm), "Agency Overall % (dec)": _weighted_avg_rate(mdf)/100.0})
    mo_df = pd.DataFrame(mo)
    lr, lc = 30, 1
    dash.write(lr, lc, "Month", head); dash.write(lr, lc+1, "Agency Overall % (dec)", head)
    for i, row in mo_df.iterrows():
        dash.write(lr+1+i, lc, row["Month Name"], cell)
        dash.write_number(lr+1+i, lc+1, row["Agency Overall % (dec)"], cell_pct)

    line = wb.add_chart({"type":"line"})
    line.add_series({
        "name":"Agency Overall %",
        "categories": ["Dashboard", lr+1, lc, lr+len(mo_df), lc],
        "values": ["Dashboard", lr+1, lc+1, lr+len(mo_df), lc+1],
        "marker":{"type":"circle","size":6},
        "line":{"color":RED,"width":2},
        "data_labels":{"value":True,"num_format":"0.00%","font":{"size":9}},
    })
    line.set_y_axis({
        "num_format":"0.00%",
        "min":0.86, "max":0.98,
        "major_gridlines":{"visible":True, "line":{"color":LIGHT_GRID}},
        "font":{"size":9, "color":DARK_TEXT},
    })
    line.set_x_axis({"label_position":"low", "num_font":{"size":9}})
    line.set_legend({"none": True})
    line.set_title({"name":"Agency Attendance Trend — 2025"})
    dash.insert_chart(28, 5, line, {"x_scale":1.4,"y_scale":1.15})

# Download
st.download_button(
    "⬇️ Download Excel (Template-style Dashboard, v2)",
    data=output.getvalue(),
    file_name=f"ADA_Dashboard_template_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
'''

reqs = """streamlit
pandas
numpy
XlsxWriter
"""

base = Path("/mnt/data")
(base / "daily_attendance_rate_25_26_app_template_dash_v2.py").write_text(code, encoding="utf-8")
(base / "requirements.txt").write_text(reqs, encoding="utf-8")

print("Files written:")
print("- /mnt/data/daily_attendance_rate_25_26_app_template_dash_v2.py")
print("- /mnt/data/requirements.txt")

