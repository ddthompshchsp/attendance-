
import calendar
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Daily Attendance Rate 25–26", layout="wide")

# --- Branding ---
PRIMARY_BLUE = "#2E75B6"
PRIMARY_RED  = "#C00000"
TEXT_MUTED   = "#6B7280"
logo_path = Path("header_logo.png")

# --- Config ---
PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"

# ---------- Helpers ----------
def _month_name(m):
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _month_order_key(name):
    try:
        m = list(calendar.month_name).index(name)
    except Exception:
        return 99
    order = list(range(8,13)) + list(range(1,7))  # Aug..Dec + Jan..Jun
    return order.index(m) if m in order else 99

def _weighted_avg_rate(df):
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    weights = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    rates   = df["Attendance Rate"].fillna(0).astype(float)  # 0..100
    total_w = weights.sum()
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

# ---------- Header ----------
st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
h1, h2, h3 = st.columns([1,2,1])
with h2:
    if logo_path.exists():
        st.image(str(logo_path), width=220)
    st.markdown(
        f"""
        <div style="text-align:center; margin-top:8px;">
            <div style="font-weight:800; font-size:32px; line-height:1.15;">HCHSP Daily Attendance (2025–2026)</div>
            <div style="color:{TEXT_MUTED}; font-size:15px; margin-top:6px;">
                Upload your Enrollment workbook, pick a month, and download a formatted ADA file.<br/>
                Charts live on the Dashboard page.
            </div>
        </div>
        """, unsafe_allow_html=True
    )
st.divider()

# ---------- Sidebar Nav ----------
page = st.sidebar.radio("Pages", ["Export", "Dashboard"], index=0)

# ---------- Upload ----------
uploaded = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Upload the Enrollment workbook to begin.")
    st.stop()

# Read bytes once
file_bytes = uploaded.read()
if not file_bytes:
    st.error("Uploaded file is empty.")
    st.stop()

# Inspect sheets
try:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Unable to read workbook: {e}")
    st.stop()

# Pick sheet
if len(sheet_names) == 1:
    use_sheet = sheet_names[0]
    st.success(f"Using sheet: **{use_sheet}**")
else:
    if PREFERRED_SHEET in sheet_names:
        use_sheet = PREFERRED_SHEET
        st.success(f"Using preferred sheet: **{use_sheet}**")
    else:
        use_sheet = st.selectbox("Choose sheet to read", options=sheet_names, index=0)

# Header row
header_row = st.number_input("Header row (0-indexed). Use 1 if headers are on the 2nd row.", min_value=0, value=1, step=1)

# Load chosen sheet
try:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=int(header_row))
except Exception as e:
    st.error(f"Failed to read sheet '{use_sheet}': {e}")
    st.stop()

if df.empty:
    st.error("Selected sheet appears empty.")
    st.stop()

# Normalize
df.columns = [str(c).strip() for c in df.columns]
df = df.rename(columns={"Unnamed: 6":"Funded", "Unnamed: 8":"Current", "Unnamed: 9":"Attendance Rate"})
base_cols = ['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate']
df = df[[c for c in base_cols if c in df.columns] + [c for c in df.columns if c not in base_cols]]
for c in ['Year','Month','Funded','Current','Attendance Rate']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")

# Month options
months_present = []
if 'Month' in df.columns:
    parsed = set()
    for m in df['Month'].dropna().unique():
        try: parsed.add(int(float(m)))
        except: pass
    months_present = sorted(parsed)

work = df[df['Month'].isin(months_present)].copy() if months_present else df.copy()
class_rows = work[work.get('Class Name').notna()].copy()

month_choices = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_order_key)
if not month_choices:
    st.error("No Month values found in the sheet.")
    st.stop()

sel_month_name = st.selectbox("Select month", options=month_choices, index=len(month_choices)-1)
sel_month_num = list(calendar.month_name).index(sel_month_name)
month_df = class_rows[class_rows['Month']==sel_month_num].copy()

# ---------- Build ADA with totals ----------
def _center_block(center_df):
    """Return the rows for a single center: class rows + a TOTAL row (weighted rate)."""
    if center_df.empty:
        return center_df
    # Class rows sorted by Class Name
    body = center_df.sort_values(["Class Name"], na_position="last").copy()
    # TOTAL row
    funded_sum = center_df['Funded'].sum(min_count=1) if 'Funded' in center_df else np.nan
    current_sum = center_df['Current'].sum(min_count=1) if 'Current' in center_df else np.nan
    w_avg = _weighted_avg_rate(center_df)
    total_row = {
        "Year": center_df['Year'].dropna().iloc[0] if 'Year' in center_df and center_df['Year'].notna().any() else np.nan,
        "Month": center_df['Month'].dropna().iloc[0] if 'Month' in center_df and center_df['Month'].notna().any() else np.nan,
        "Center Name": center_df['Center Name'].dropna().iloc[0] if 'Center Name' in center_df and center_df['Center Name'].notna().any() else np.nan,
        "Class Name": "TOTAL",
        "Funded": funded_sum,
        "Current": current_sum,
        "Attendance Rate": float(w_avg) if pd.notna(w_avg) else np.nan,
    }
    return pd.concat([body, pd.DataFrame([total_row])], ignore_index=True)

final_blocks = []
for center, g in month_df.groupby("Center Name", dropna=False):
    final_blocks.append(_center_block(g))

ada_df = pd.concat(final_blocks, ignore_index=True) if final_blocks else pd.DataFrame(columns=base_cols)

# Agency overall TOTAL row
overall_w = _weighted_avg_rate(month_df)
overall_row = {
    "Year": month_df['Year'].dropna().iloc[0] if 'Year' in month_df and month_df['Year'].notna().any() else np.nan,
    "Month": sel_month_num,
    "Center Name": "HCHSP (Overall)",
    "Class Name": "TOTAL",
    "Funded": month_df['Funded'].sum(min_count=1) if 'Funded' in month_df else np.nan,
    "Current": month_df['Current'].sum(min_count=1) if 'Current' in month_df else np.nan,
    "Attendance Rate": float(overall_w) if pd.notna(overall_w) else np.nan,
}
ada_df = pd.concat([ada_df, pd.DataFrame([overall_row])], ignore_index=True)

# ---------- EXPORT PAGE ----------
if page == "Export":
    # Preview
    st.markdown(f"#### ADA Preview — {sel_month_name}")
    st.dataframe(
        ada_df.style.format({"Attendance Rate": "{:.2f}%"}).set_properties(subset=["Class Name"], **{"font-weight":"bold"}),
        use_container_width=True, height=380
    )

    # Build Excel (ADA + Dashboard)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        chicago_now = datetime.now(ZoneInfo("America/Chicago"))

        # Sheet 1: ADA
        ada_sheet = "ADA"
        ada_df.to_excel(writer, index=False, sheet_name=ada_sheet, startrow=3)
        wb = writer.book
        ws = writer.sheets[ada_sheet]

        header_fmt   = wb.add_format({"bold": True, "font_color": "white", "bg_color": PRIMARY_BLUE, "align": "center", "valign": "vcenter", "text_wrap": True})
        percent_fmt  = wb.add_format({"num_format": "0.00%"})
        bold_fmt     = wb.add_format({"bold": True})

        header_row = 3
        for col_num, col_name in enumerate(ada_df.columns):
            ws.write(header_row, col_num, col_name, header_fmt)
        ws.set_row(header_row, 26)

        # Autosize
        for i, col in enumerate(ada_df.columns):
            width = max(12, min(44, int(ada_df[col].astype(str).map(len).max()) + 2))
            ws.set_column(i, i, width)

        last_row_index = len(ada_df) + header_row
        last_col_index = len(ada_df.columns) - 1
        ws.autofilter(header_row, 0, last_row_index, last_col_index)
        ws.freeze_panes(header_row + 1, 0)

        # % format & bold TOTAL rows
        if 'Attendance Rate' in ada_df.columns:
            col_idx = ada_df.columns.get_loc('Attendance Rate')
            cls_idx = ada_df.columns.get_loc('Class Name')
            for r in range(len(ada_df)):
                excel_row = r + header_row + 1
                v = ada_df.iloc[r, col_idx]
                # write as decimal if number
                if pd.isna(v):
                    ws.write_blank(excel_row, col_idx, None, percent_fmt)
                else:
                    ws.write_number(excel_row, col_idx, float(v)/100.0, percent_fmt)
                # bold entire row if TOTAL
                if str(ada_df.iloc[r, cls_idx]).upper() == "TOTAL":
                    ws.set_row(excel_row, None, bold_fmt)

        # Title + timestamp + logo
        title_fmt    = wb.add_format({"bold": True, "align": "left", "valign": "vcenter", "font_size": 16})
        timestamp_fmt= wb.add_format({"italic": True, "align": "left", "valign": "vcenter", "font_color": "#7F7F7F"})
        ws.merge_range(0, 1, 1, last_col_index, f"Daily Attendance Rate 25–26 — {sel_month_name}", title_fmt)
        ws.merge_range(2, 1, 2, last_col_index, f"(Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')})", timestamp_fmt)
        if logo_path.exists():
            ws.insert_image(0, 0, str(logo_path), {"x_scale":0.6, "y_scale":0.6, "x_offset":4, "y_offset":4})
        ws.set_row(0, 44); ws.set_row(1, 24); ws.set_row(2, 18)

        # Sheet 2: Dashboard data for Excel charts (optional static)
        dash_sheet = "Dashboard"
        # centers data
        centers_tbl = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %").sort_values("Attendance %", ascending=False)
        centers_out = centers_tbl.copy()
        centers_out["Attendance % (dec)"] = centers_out["Attendance %"]/100.0
        centers_out.to_excel(writer, index=False, sheet_name=dash_sheet, startrow=0)
        ws2 = writer.sheets[dash_sheet]
        ws2.set_column(0, 0, 28)
        ws2.set_column(1, 2, 18)

        # monthly trend
        monthly_overall = []
        for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
            mdf = class_rows[class_rows['Month']==int(m)]
            monthly_overall.append({"Month Name": _month_name(m), "Agency Overall % (dec)": _weighted_avg_rate(mdf)/100.0})
        mo_df = pd.DataFrame(monthly_overall)
        start_r = len(centers_out) + 3
        if not mo_df.empty:
            ws2.write(start_r, 0, "Month")
            ws2.write(start_r, 1, "Agency Overall % (dec)")
            for i, row in mo_df.iterrows():
                ws2.write(start_r+1+i, 0, row["Month Name"])
                ws2.write(start_r+1+i, 1, row["Agency Overall % (dec)"])

    st.download_button(
        label="⬇️ Download Excel (ADA + Dashboard)",
        data=output.getvalue(),
        file_name=f"ADA_ByCampus_Classes_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- DASHBOARD PAGE (interactive charts) ----------
else:
    # Filters
    month_choices = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_order_key)
    d_month = st.selectbox("Dashboard month", options=month_choices, index=len(month_choices)-1, key="dash_month")
    d_month_num = list(calendar.month_name).index(d_month)
    d_month_df = class_rows[class_rows['Month']==d_month_num].copy()

    centers = sorted([c for c in d_month_df['Center Name'].dropna().unique()])
    sel_centers = st.multiselect("Filter centers", options=centers, default=centers)
    d_month_df = d_month_df[d_month_df['Center Name'].isin(sel_centers)] if sel_centers else d_month_df.iloc[0:0]

    # Bar: Attendance by center (weighted)
    if d_month_df.empty:
        st.info("No data for selected filters.")
    else:
        bar_df = d_month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %")
        bar_df = bar_df.sort_values("Attendance %", ascending=False)
        fig = px.bar(
            bar_df, x="Center Name", y="Attendance %",
            text=bar_df["Attendance %"].map(lambda v: f"{v:.2f}%"),
            color_discrete_sequence=[PRIMARY_BLUE]
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(title=f"Attendance Rate by Centers — {d_month}", yaxis_title="Attendance (%)", xaxis_title="", margin=dict(l=20,r=20,t=60,b=80))
        st.plotly_chart(fig, use_container_width=True)

    # Line: Agency overall trend (Aug–Jun)
    trend = []
    for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
        mdf = class_rows[class_rows['Month']==int(m)]
        trend.append({"Month Name": _month_name(m), "Agency Overall %": _weighted_avg_rate(mdf)})
    trend_df = pd.DataFrame(trend)
    if not trend_df.empty:
        fig2 = px.line(trend_df, x="Month Name", y="Agency Overall %", markers=True, color_discrete_sequence=[PRIMARY_RED])
        fig2.update_traces(text=[f"{v:.2f}%" for v in trend_df["Agency Overall %"]], textposition="top center")
        fig2.update_layout(title="Agency Trend — Aug to Jun", yaxis_title="Attendance (%)", xaxis_title="", margin=dict(l=20,r=20,t=60,b=40))
        st.plotly_chart(fig2, use_container_width=True)
