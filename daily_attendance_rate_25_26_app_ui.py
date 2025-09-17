# Build the app "like two versions ago" but:
# - Remove Drop Trends chart entirely
# - Remove any data after Attendance Rate (no drops columns)
# - Keep Dashboard sheet (with bar + line charts) and Export page
# - Keep month picker and branding

from pathlib import Path

code = r'''
import re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
from pathlib import Path
from io import BytesIO

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

# ---- THEME / COLORS ----
PRIMARY_BLUE = "#2E75B6"
PRIMARY_RED = "#C00000"
BG_LIGHT = "#F7F9FC"
TEXT_MUTED = "#6B7280"

# CONFIG
PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"
logo_path = Path("header_logo.png")

# ---------- HEADER ----------
with st.container():
    cols = st.columns([1,2,1])
    with cols[0]:
        if logo_path.exists():
            st.image(str(logo_path), use_column_width=True)
    with cols[1]:
        st.markdown("<h2 style='text-align:center;margin:6px 0;'>Daily Attendance Rate 25-26</h2>", unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center;color:{TEXT_MUTED}'>Export on Sheet 1, dashboard on Sheet 2. Only through Attendance Rate.</div>", unsafe_allow_html=True)
    with cols[2]:
        st.write("")

st.divider()

# ---------- SIDEBAR NAV ----------
st.sidebar.markdown("### Navigation")
page = st.sidebar.radio("Go to", ["Export", "Dashboard (Preview)"], index=0)

# Upload up front so both pages can see data
uploaded_file = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])

def _month_name(m):
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _month_order_key(name):
    # Aug (8) -> Jun (6)
    try:
        m = list(calendar.month_name).index(name)
    except Exception:
        return 99
    order = list(range(8,13)) + list(range(1,7))
    return order.index(m) if m in order else 99

def _weighted_avg_rate(df):
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    weights = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    rates = df["Attendance Rate"].fillna(0).astype(float)
    total_w = weights.sum()
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    if not file_bytes:
        st.error("Uploaded file is empty. Please upload a valid .xlsx file.")
        st.stop()

    # inspect sheets
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
        sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"Unable to read workbook. Error: {e}")
        st.stop()

    # pick sheet
    if len(sheet_names) == 1:
        use_sheet = sheet_names[0]
        st.success(f"Only one sheet found. Using sheet: **{use_sheet}**")
    else:
        if PREFERRED_SHEET in sheet_names:
            use_sheet = PREFERRED_SHEET
            st.success(f"Using preferred sheet: **{use_sheet}**")
        else:
            st.warning(f"Preferred sheet '{PREFERRED_SHEET}' not found. Please choose a sheet.")
            use_sheet = st.selectbox("Choose sheet to read", options=sheet_names, index=0)

    # header row select
    header_row = st.number_input("Header row (0-indexed). Use 1 if headers are on the 2nd row.", min_value=0, value=1, step=1)

    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=int(header_row))
    except Exception as e:
        st.error(f"Failed to read sheet '{use_sheet}'. Error: {e}")
        st.stop()

    if df.empty:
        st.error(f"The selected sheet ('{use_sheet}') appears empty.")
        st.stop()

    # ---------- CLEAN / NORMALIZE ----------
    df.columns = [str(c).strip() for c in df.columns]

    # rename expected Unnamed columns if present
    df = df.rename(columns={
        "Unnamed: 6": "Funded",
        "Unnamed: 8": "Current",
        "Unnamed: 9": "Attendance Rate"
    })

    # keep only relevant columns (stop at Attendance Rate)
    base_cols = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
    use_cols = [c for c in base_cols if c in df.columns]
    df = df[use_cols]

    # coerce numeric columns
    for c in ['Year', 'Month', 'Funded', 'Current', 'Attendance Rate']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # month parsing
    months_present = []
    if 'Month' in df.columns:
        raw_months = df['Month'].dropna().unique()
        parsed = set()
        for m in raw_months:
            try:
                parsed.add(int(float(m)))
            except Exception:
                continue
        months_present = sorted(parsed)

    work = df[df['Month'].isin(months_present)].copy() if months_present else df.copy()
    class_rows = work[work.get('Class Name').notna()].copy()

    # Month choices & selection
    month_choices = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_order_key)
    sel_month_name = st.selectbox("Select month", options=month_choices, index=len(month_choices)-1 if month_choices else 0)
    sel_month_num = list(calendar.month_name).index(sel_month_name) if sel_month_name in list(calendar.month_name) else None
    month_df = class_rows[class_rows['Month']==sel_month_num] if sel_month_num is not None else class_rows.copy()

    # center totals (weighted by 'Current')
    def center_totals(g):
        weights = g.get('Current', pd.Series([0]*len(g))).fillna(0.0).astype(float)
        rates = g.get('Attendance Rate', pd.Series([np.nan]*len(g))).fillna(0.0).astype(float)
        w_avg = (rates * weights).sum() / weights.sum() if weights.sum() > 0 else np.nan
        return pd.Series({
            'Year': g['Year'].dropna().iloc[0] if 'Year' in g and g['Year'].notna().any() else np.nan,
            'Month': g['Month'].dropna().iloc[0] if 'Month' in g and g['Month'].notna().any() else np.nan,
            'Center Name': g['Center Name'].dropna().iloc[0] if 'Center Name' in g and g['Center Name'].notna().any() else np.nan,
            'Class Name': 'TOTAL',
            'Funded': g['Funded'].sum(min_count=1) if 'Funded' in g else np.nan,
            'Current': g['Current'].sum(min_count=1) if 'Current' in g else np.nan,
            'Attendance Rate': float(w_avg) if pd.notna(w_avg) else np.nan
        })

    if not month_df.empty:
        center_total_df = month_df.groupby(['Year','Month','Center Name'], dropna=False).apply(center_totals).reset_index(drop=True)
    else:
        center_total_df = pd.DataFrame(columns=base_cols)

    # combine details + totals per center
    combined_parts = []
    for (yr, mo, center), g in month_df.groupby(['Year','Month','Center Name']):
        combined_parts.append(g.sort_values(['Class Name']))
        combined_parts.append(center_total_df[(center_total_df['Year']==yr)&(center_total_df['Month']==mo)&(center_total_df['Center Name']==center)])
    final_df = pd.concat(combined_parts, ignore_index=True) if combined_parts else pd.DataFrame(columns=center_total_df.columns)

    # overall weighted avg for selected month
    overall_weighted = _weighted_avg_rate(month_df)
    overall_df = pd.DataFrame([{
        'Year': month_df['Year'].dropna().iloc[0] if ('Year' in month_df and month_df['Year'].notna().any()) else np.nan,
        'Month': sel_month_num if sel_month_num is not None else np.nan,
        'Center Name': 'HCHSP (Overall)',
        'Class Name': 'TOTAL',
        'Funded': month_df['Funded'].sum(min_count=1) if 'Funded' in month_df else np.nan,
        'Current': month_df['Current'].sum(min_count=1) if 'Current' in month_df else np.nan,
        'Attendance Rate': float(overall_weighted) if pd.notna(overall_weighted) else np.nan
    }])

    final_df = pd.concat([final_df, overall_df], ignore_index=True)

    # ---------- EXPORT PAGE ----------
    if page == "Export":
        # Excel export (xlsxwriter)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheet = "ADA"
            final_df.to_excel(writer, index=False, sheet_name=sheet, startrow=3)
            wb = writer.book
            ws = writer.sheets[sheet]

            # formats
            header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": PRIMARY_BLUE, "align": "center", "valign": "vcenter", "text_wrap": True})
            bold_fmt = wb.add_format({"bold": True})
            bold_pct_fmt = wb.add_format({"bold": True, "num_format": "0.00%"})
            percent_fmt = wb.add_format({"num_format": "0.00%"})
            title_fmt = wb.add_format({"bold": True, "align": "left", "valign": "vcenter", "font_size": 16})
            timestamp_fmt = wb.add_format({"italic": True, "align": "left", "valign": "vcenter", "font_color": "#7F7F7F"})

            # header row
            header_row = 3
            for col_num, col_name in enumerate(final_df.columns):
                ws.write(header_row, col_num, col_name, header_fmt)
            ws.set_row(header_row, 30)

            # autosize columns
            for i, col in enumerate(final_df.columns):
                width = max(12, min(40, int(final_df[col].astype(str).map(len).max()) + 2))
                ws.set_column(i, i, width)

            # autofilter & freeze
            last_row_index = len(final_df) + header_row
            last_col_index = len(final_df.columns) - 1
            ws.autofilter(header_row, 0, last_row_index, last_col_index)
            ws.freeze_panes(header_row + 1, 0)

            # percent formatting for Attendance Rate (0..100 -> Excel decimal)
            if 'Attendance Rate' in final_df.columns:
                col_idx = final_df.columns.get_loc('Attendance Rate')
                cls_idx = final_df.columns.get_loc('Class Name')
                for r in range(len(final_df)):
                    value_0_to_100 = final_df.iloc[r, col_idx]
                    excel_row = r + header_row + 1
                    if pd.isna(value_0_to_100):
                        ws.write_blank(excel_row, col_idx, None, percent_fmt)
                    else:
                        val_decimal = float(value_0_to_100) / 100.0
                        fmt = bold_pct_fmt if str(final_df.iloc[r].get('Class Name', '')).upper() == 'TOTAL' else percent_fmt
                        ws.write_number(excel_row, col_idx, val_decimal, fmt)
                    # Bold TOTAL rows
                    if str(final_df.iloc[r, cls_idx]).upper() == 'TOTAL':
                        ws.set_row(excel_row, None, bold_fmt)

            # title + timestamp
            ws.merge_range(0, 1, 1, last_col_index, f"Daily Attendance Rate 25-26 — {sel_month_name}", title_fmt)
            chicago_now = datetime.now(ZoneInfo("America/Chicago"))
            timestamp_text = f"(Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')})"
            ws.merge_range(2, 1, 2, last_col_index, timestamp_text, timestamp_fmt)

            ws.set_row(0, 44)
            ws.set_row(1, 24)
            ws.set_row(2, 18)

            # insert logo top-left
            if logo_path.exists():
                ws.insert_image(0, 0, str(logo_path), {"x_scale": 0.6, "y_scale": 0.6, "x_offset": 4, "y_offset": 4})

            # ---------------- Sheet 2: Dashboard (bar + line only) ----------------
            dash = wb.add_worksheet("Dashboard")
            dash.set_zoom(115)

            # Title & timestamp
            dash_title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": PRIMARY_BLUE})
            dash.write(0, 1, "Daily Attendance Dashboard", dash_title_fmt)
            ts_fmt = wb.add_format({"italic": True, "font_color": "#7F7F7F"})
            dash.write(1, 1, f"Month for bar chart: {sel_month_name}")
            dash.write(2, 1, f"Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')}", ts_fmt)
            # Logo (if exists)
            if logo_path.exists():
                dash.insert_image(0, 0, str(logo_path), {"x_scale": 0.6, "y_scale": 0.6, "x_offset": 2, "y_offset": 2})

            # A) Centers bar (selected month)
            centers_tbl_start = (5, 1)   # row, col
            bar_df = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %")
            bar_df = bar_df.sort_values("Attendance %", ascending=False)
            dash.write_row(centers_tbl_start[0], centers_tbl_start[1], ["Center Name", "Attendance %"])
            for i, row in bar_df.iterrows():
                dash.write(centers_tbl_start[0]+1+i, centers_tbl_start[1],     row["Center Name"])
                dash.write(centers_tbl_start[0]+1+i, centers_tbl_start[1]+1,   (row["Attendance %"]/100.0) if pd.notna(row["Attendance %"]) else None)

            bar = wb.add_chart({"type":"column"})
            bar.add_series({
                "name":       f"Attendance Rate by Centers — {sel_month_name}",
                "categories": ["Dashboard", centers_tbl_start[0]+1, centers_tbl_start[1], centers_tbl_start[0]+len(bar_df), centers_tbl_start[1]],
                "values":     ["Dashboard", centers_tbl_start[0]+1, centers_tbl_start[1]+1, centers_tbl_start[0]+len(bar_df), centers_tbl_start[1]+1],
                "data_labels": {"value": True, "num_format":"0.00%"},
                "fill": {"color": PRIMARY_BLUE},
                "border": {"color": PRIMARY_BLUE},
            })
            bar.set_y_axis({"num_format":"0.00%"})
            bar.set_title({"name": f"Attendance Rate by Centers — {sel_month_name}"})
            dash.insert_chart(4, 5, bar, {"x_scale":1.1, "y_scale":1.15})

            # B) Agency trend line (Aug–Jun)
            trend_tbl_start = (5, 9)
            monthly_overall = []
            if not class_rows.empty and 'Month' in class_rows:
                for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
                    mdf = class_rows[class_rows['Month']==int(m)]
                    monthly_overall.append({
                        "Month": int(m),
                        "Month Name": _month_name(m),
                        "Agency Overall %": _weighted_avg_rate(mdf)
                    })
            tdf = pd.DataFrame(monthly_overall).sort_values("Month", key=lambda s: s.astype(int)) if monthly_overall else pd.DataFrame(columns=["Month Name","Agency Overall %"])

            dash.write_row(trend_tbl_start[0], trend_tbl_start[1], ["Month", "Agency Overall %"])
            for i, row in tdf.iterrows():
                dash.write(trend_tbl_start[0]+1+i, trend_tbl_start[1],   row["Month Name"])
                dash.write(trend_tbl_start[0]+1+i, trend_tbl_start[1]+1, (row["Agency Overall %"]/100.0) if pd.notna(row["Agency Overall %"]) else None)

            line = wb.add_chart({"type":"line"})
            line.add_series({
                "name": "Agency Overall %",
                "categories": ["Dashboard", trend_tbl_start[0]+1, trend_tbl_start[1], trend_tbl_start[0]+len(tdf), trend_tbl_start[1]],
                "values": ["Dashboard", trend_tbl_start[0]+1, trend_tbl_start[1]+1, trend_tbl_start[0]+len(tdf), trend_tbl_start[1]+1],
                "marker": {"type":"circle", "size":5},
                "line": {"color": PRIMARY_RED, "width": 2},
            })
            line.set_title({"name":"Agency Trend — Aug to Jun"})
            line.set_y_axis({"num_format":"0.00%"})
            dash.insert_chart(20, 1, line, {"x_scale":1.1, "y_scale":1.15})

        st.download_button(
            label="⬇️ Download Excel (ADA + Dashboard)",
            data=output.getvalue(),
            file_name=f"ADA_ByCampus_Classes_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------- DASHBOARD PREVIEW (inside app with Plotly) ----------
    else:
        # Preview month selector
        month_choices = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_order_key)
        sel_month_name = st.selectbox("Preview month", options=month_choices, index=len(month_choices)-1 if month_choices else 0)
        sel_month_num = list(calendar.month_name).index(sel_month_name) if sel_month_name in list(calendar.month_name) else None
        month_df = class_rows[class_rows['Month']==sel_month_num] if sel_month_num is not None else class_rows.copy()

        # KPI
        k1, k2, k3 = st.columns(3)
        with k1:
            st.metric(f"Agency Overall ({sel_month_name})", f"{_weighted_avg_rate(month_df):.2f}%")
        with k2:
            tmp = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).sort_values(ascending=False)
            st.metric("Top Center (weighted)", tmp.index[0] if len(tmp) else "—", delta=f"{tmp.iloc[0]:.2f}%" if len(tmp) else None)
        with k3:
            tmp2 = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).sort_values(ascending=True)
            st.metric("Lowest Center (weighted)", tmp2.index[0] if len(tmp2) else "—", delta=f"{tmp2.iloc[0]:.2f}%" if len(tmp2) else None)

        st.markdown("---")
        tab1, tab2 = st.tabs(["Attendance Rate by Centers (Bar)", "Agency Trend Aug–Jun (Line)"])

        with tab1:
            if month_df.empty:
                st.info("No data for the selected month.")
            else:
                bar_df = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %")
                bar_df = bar_df.sort_values("Attendance %", ascending=False)
                fig = px.bar(
                    bar_df, x="Center Name", y="Attendance %",
                    text=bar_df["Attendance %"].map(lambda v: f"{v:.2f}%"),
                    color_discrete_sequence=[PRIMARY_BLUE]
                )
                fig.update_traces(textposition="outside")
                fig.update_layout(yaxis_title="Attendance (%)", xaxis_title="", margin=dict(l=20,r=20,t=30,b=60), title=f"Attendance Rate by Centers — {sel_month_name}")
                st.plotly_chart(fig, use_container_width=True)

        with tab2:
            monthly_overall = []
            if not class_rows.empty and 'Month' in class_rows:
                for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
                    mdf = class_rows[class_rows['Month']==int(m)]
                    monthly_overall.append({"Month": int(m), "Month Name": _month_name(m), "Agency Overall %": _weighted_avg_rate(mdf)})
            trend_df = pd.DataFrame(monthly_overall).sort_values("Month", key=lambda s: s.astype(int))
            if trend_df.empty:
                st.info("No monthly data available.")
            else:
                fig2 = px.line(
                    trend_df, x="Month Name", y="Agency Overall %",
                    markers=True, color_discrete_sequence=[PRIMARY_RED]
                )
                fig2.update_traces(text=[f"{v:.2f}%" for v in trend_df["Agency Overall %"]], textposition="top center")
                fig2.update_layout(yaxis_title="Attendance (%)", xaxis_title="", margin=dict(l=20,r=20,t=30,b=60), title="Agency Trend — Aug to Jun")
                st.plotly_chart(fig2, use_container_width=True)

else:
    st.info("Upload the Enrollment workbook to enable Export and Dashboard views.")
'''

reqs = """streamlit
pandas
numpy
XlsxWriter
plotly
"""

base = Path("/mnt/data")
(base / "daily_attendance_rate_25_26_app_no_drops.py").write_text(code, encoding="utf-8")
(base / "requirements.txt").write_text(reqs, encoding="utf-8")

print("Files saved:")
print("- /mnt/data/daily_attendance_rate_25_26_app_no_drops.py")
print("- /mnt/data/requirements.txt")
