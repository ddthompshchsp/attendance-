
import re
import calendar
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Daily Attendance Rate 25–26", layout="wide")

# ---------- Branding / Theme ---------
PRIMARY_BLUE = "#2E75B6"
PRIMARY_RED  = "#C00000"
TEXT_MUTED   = "#6B7280"

logo_path = Path("header_logo.png")  # your blue/red County logo

# ---------- Hero Header (like your screenshot) ----------
st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    if logo_path.exists():
        st.image(str(logo_path), width=220)
    st.markdown(
        f"""
        <div style="text-align:center; margin-top:8px;">
            <div style="font-weight:800; font-size:34px; line-height:1.15;">HCHSP Daily Attendance (2025–2026)</div>
            <div style="color:{TEXT_MUTED}; font-size:16px; margin-top:6px;">
                Upload your <b>Enrollment.xlsx</b> file to generate a formatted report.<br/>
                Optionally include a Dashboard sheet and Drops analysis.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
st.divider()

# ---------- Controls row ----------
left, right = st.columns([1, 1])
with left:
    include_dashboard = st.toggle("➕ Add Dashboard sheet", value=True,
                                  help="Adds a second sheet with KPI and charts in HCHSP colors.")
with right:
    include_drops = st.toggle("📉 Include Drops", value=True,
                              help="Maps 'Drops' columns (M & N or duplicate 'Drops' headers) and adds a Drops chart when Dashboard is enabled.")

# ---------- Upload ----------
uploaded = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

# ---------- Helpers ----------
PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"

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
    rates   = df["Attendance Rate"].fillna(0).astype(float)
    total_w = weights.sum()
    if total_w <= 0:
        return np.nan
    return float((rates*weights).sum()/total_w)

def _detect_drops(df):
    """Return (cur_drop, cum_drop, note). Tries duplicate 'Drops' headers first, then M & N fallback."""
    cols = list(df.columns)
    # duplicate Drops (e.g., Drops, Drops.1)
    idxs = [i for i,c in enumerate(cols) if re.fullmatch(r"(?i)drops(\.\d+)?", str(c).strip())]
    if len(idxs) >= 2:
        return cols[idxs[0]], cols[idxs[1]], "Detected duplicate 'Drops' headers."
    # fallback to M/N if present
    if len(cols) >= 14:
        return cols[12], cols[13], "Using columns M (current) & N (cumulative)."
    return None, None, "No drops columns found; add or rename to 'Drops'."

# ---------- Main ----------
if uploaded:
    file_bytes = uploaded.read()
    if not file_bytes:
        st.error("Uploaded file is empty."); st.stop()

    # sheet pick
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
        sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"Unable to read workbook: {e}"); st.stop()

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
        st.error(f"Failed to read sheet '{use_sheet}': {e}"); st.stop()

    if df.empty:
        st.error("Selected sheet appears empty."); st.stop()

    # Normalize
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={"Unnamed: 6":"Funded", "Unnamed: 8":"Current", "Unnamed: 9":"Attendance Rate"})
    base_cols = ['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate']
    df = df[[c for c in base_cols if c in df.columns] + [c for c in df.columns if c not in base_cols]]
    for c in ['Year','Month','Funded','Current','Attendance Rate']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Month parsing
    months_present = []
    if 'Month' in df.columns:
        parsed = set()
        for m in df['Month'].dropna().unique():
            try: parsed.add(int(float(m)))
            except: pass
        months_present = sorted(parsed)
    work = df[df['Month'].isin(months_present)].copy() if months_present else df.copy()
    class_rows = work[work.get('Class Name').notna()].copy()

    # Drops mapping (only if requested)
    cur_drop_col = cum_drop_col = None; drop_note = ""
    if include_drops:
        cur_drop_col, cum_drop_col, drop_note = _detect_drops(class_rows)
        st.caption(f"Drops mapping: {drop_note}")
        # Manual override
        all_cols = list(class_rows.columns)
        col1, col2 = st.columns(2)
        with col1:
            cur_drop_col = st.selectbox("Current Drops column", options=[None]+all_cols,
                                        index=(all_cols.index(cur_drop_col)+1 if cur_drop_col in all_cols else 0))
        with col2:
            cum_drop_col = st.selectbox("Cumulative Drops column", options=[None]+all_cols,
                                        index=(all_cols.index(cum_drop_col)+1 if cum_drop_col in all_cols else 0))

    # Build center totals and final export table (ADA)
    def center_totals(g):
        weights = g.get('Current', pd.Series([0]*len(g))).fillna(0.0).astype(float)
        rates   = g.get('Attendance Rate', pd.Series([np.nan]*len(g))).fillna(0.0).astype(float)
        w_avg   = (rates*weights).sum()/weights.sum() if weights.sum() > 0 else np.nan
        return pd.Series({
            'Year': g['Year'].dropna().iloc[0] if 'Year' in g and g['Year'].notna().any() else np.nan,
            'Month': g['Month'].dropna().iloc[0] if 'Month' in g and g['Month'].notna().any() else np.nan,
            'Center Name': g['Center Name'].dropna().iloc[0] if 'Center Name' in g and g['Center Name'].notna().any() else np.nan,
            'Class Name': 'TOTAL',
            'Funded': g['Funded'].sum(min_count=1) if 'Funded' in g else np.nan,
            'Current': g['Current'].sum(min_count=1) if 'Current' in g else np.nan,
            'Attendance Rate': float(w_avg) if pd.notna(w_avg) else np.nan
        })

    if not class_rows.empty:
        center_total_df = class_rows.groupby(['Year','Month','Center Name'], dropna=False).apply(center_totals).reset_index(drop=True)
    else:
        center_total_df = pd.DataFrame(columns=['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate'])

    combined_parts = []
    for (yr, mo, center), g in class_rows.groupby(['Year','Month','Center Name']):
        combined_parts.append(g.sort_values(['Class Name']))
        combined_parts.append(center_total_df[(center_total_df['Year']==yr)&(center_total_df['Month']==mo)&(center_total_df['Center Name']==center)])
    combined = pd.concat(combined_parts, ignore_index=True) if combined_parts else pd.DataFrame(columns=center_total_df.columns)

    # Agency monthly trend
    monthly_overall = []
    if not class_rows.empty and 'Month' in class_rows:
        for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
            mdf = class_rows[class_rows['Month']==m]
            monthly_overall.append({"Month": int(m), "Month Name": _month_name(m), "Agency Overall %": _weighted_avg_rate(mdf)})
    monthly_overall = pd.DataFrame(monthly_overall)

    latest_m = int(class_rows['Month'].dropna().max()) if not class_rows.empty else None
    latest_df = class_rows[class_rows['Month']==latest_m] if latest_m is not None else class_rows.iloc[0:0]
    latest_overall = _weighted_avg_rate(latest_df)

    overall_df = pd.DataFrame([{
        'Year': class_rows['Year'].dropna().iloc[0] if ('Year' in class_rows and class_rows['Year'].notna().any()) else np.nan,
        'Month': latest_m if latest_m is not None else np.nan,
        'Center Name': 'HCHSP (Overall)',
        'Class Name': 'TOTAL',
        'Funded': class_rows['Funded'].sum(min_count=1) if 'Funded' in class_rows else np.nan,
        'Current': class_rows['Current'].sum(min_count=1) if 'Current' in class_rows else np.nan,
        'Attendance Rate': float(latest_overall) if pd.notna(latest_overall) else np.nan
    }])

    final_df = pd.concat([combined, overall_df], ignore_index=True)

    # Month selector for Dashboard bars
    month_choices = sorted({_month_name(m) for m in class_rows['Month'].dropna().unique()}, key=_month_order_key)
    sel_month_name = st.selectbox("Choose month for Dashboard charts (in Excel)", options=month_choices, index=len(month_choices)-1 if month_choices else 0)
    sel_month_num = list(calendar.month_name).index(sel_month_name) if sel_month_name in list(calendar.month_name) else None
    month_df = class_rows[class_rows['Month']==sel_month_num] if sel_month_num is not None else class_rows.copy()

    # ---------- Build Excel (single workbook, conditional sheets) ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Sheet 1: ADA
        sheet = "ADA"
        final_df.to_excel(writer, index=False, sheet_name=sheet, startrow=3)
        wb = writer.book
        ws = writer.sheets[sheet]

        header_fmt   = wb.add_format({"bold": True, "font_color": "white", "bg_color": PRIMARY_BLUE, "align": "center", "valign": "vcenter", "text_wrap": True})
        bold_fmt     = wb.add_format({"bold": True})
        bold_pct_fmt = wb.add_format({"bold": True, "num_format": "0.00%"})
        percent_fmt  = wb.add_format({"num_format": "0.00%"})
        title_fmt    = wb.add_format({"bold": True, "align": "left", "valign": "vcenter", "font_size": 16})
        timestamp_fmt= wb.add_format({"italic": True, "align": "left", "valign": "vcenter", "font_color": "#7F7F7F"})

        header_row = 3
        for col_num, col_name in enumerate(final_df.columns):
            ws.write(header_row, col_num, col_name, header_fmt)
        ws.set_row(header_row, 30)

        for i, col in enumerate(final_df.columns):
            width = max(12, min(40, int(final_df[col].astype(str).map(len).max()) + 2))
            ws.set_column(i, i, width)

        last_row_index = len(final_df) + header_row
        last_col_index = len(final_df.columns) - 1
        ws.autofilter(header_row, 0, last_row_index, last_col_index)
        ws.freeze_panes(header_row + 1, 0)

        if 'Attendance Rate' in final_df.columns:
            col_idx = final_df.columns.get_loc('Attendance Rate')
            for r in range(len(final_df)):
                val_0_100 = final_df.iloc[r, col_idx]
                excel_row = r + header_row + 1
                if pd.isna(val_0_100):
                    ws.write_blank(excel_row, col_idx, None, percent_fmt)
                else:
                    ws.write_number(excel_row, col_idx, float(val_0_100)/100.0,
                                    bold_pct_fmt if str(final_df.iloc[r].get('Class Name','')).upper()=='TOTAL' else percent_fmt)

        month_labels = sorted({_month_name(m) for m in work['Month'].dropna().unique()}, key=_month_order_key)
        month_label_text = ", ".join([m for m in month_labels if m])

        chicago_now = datetime.now(ZoneInfo("America/Chicago"))
        ws.merge_range(0, 1, 1, last_col_index, f"Daily Attendance Rate 25–26 ({month_label_text})", title_fmt)
        ws.merge_range(2, 1, 2, last_col_index, f"(Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')})", timestamp_fmt)
        ws.set_row(0, 44); ws.set_row(1, 24); ws.set_row(2, 18)
        if logo_path.exists():
            ws.insert_image(0, 0, str(logo_path), {"x_scale":0.6, "y_scale":0.6, "x_offset":4, "y_offset":4})

        # Sheet 2: Dashboard (only if requested)
        if include_dashboard:
            dash = wb.add_worksheet("Dashboard")
            dash.set_zoom(115)

            # Title & timestamp
            dash_title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": PRIMARY_BLUE})
            dash.write(0, 1, "Daily Attendance Dashboard", dash_title_fmt)
            ts_fmt = wb.add_format({"italic": True, "font_color": TEXT_MUTED})
            dash.write(1, 1, f"Month for bar charts: {sel_month_name}")
            dash.write(2, 1, f"Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')}", ts_fmt)
            if logo_path.exists():
                dash.insert_image(0, 0, str(logo_path), {"x_scale":0.6, "y_scale":0.6, "x_offset":2, "y_offset":2})

            # KPI
            kpi_fmt_label = wb.add_format({"bold": True, "align":"center", "valign":"vcenter", "bg_color":"#EEF3FA", "border":1})
            kpi_fmt_val   = wb.add_format({"bold": True, "align":"center", "valign":"vcenter", "font_size":16, "font_color":PRIMARY_BLUE, "border":1})
            dash.merge_range(4, 1, 5, 3, f"Agency Overall ({sel_month_name})", kpi_fmt_label)
            kpi_val = _weighted_avg_rate(month_df)
            if pd.isna(kpi_val):
                dash.merge_range(6, 1, 7, 3, "—", kpi_fmt_val)
            else:
                pct_fmt = wb.add_format({"bold": True, "align":"center", "valign":"vcenter", "font_size":16, "font_color":PRIMARY_BLUE, "num_format":"0.00%", "border":1})
                dash.merge_range(6, 1, 7, 3, kpi_val/100.0, pct_fmt)

            # Data tables for charts
            centers_tbl_start = (10, 1)
            cmdf = month_df.groupby("Center Name", dropna=False).apply(_weighted_avg_rate).reset_index(name="Attendance %")
            cmdf = cmdf.sort_values("Attendance %", ascending=False)
            dash.write_row(centers_tbl_start[0], centers_tbl_start[1], ["Center Name","Attendance %"])
            for i,row in cmdf.iterrows():
                dash.write(centers_tbl_start[0]+1+i, centers_tbl_start[1],   row["Center Name"])
                dash.write(centers_tbl_start[0]+1+i, centers_tbl_start[1]+1, row["Attendance %"]/100.0)

            trend_tbl_start = (10, 6)
            tdf = monthly_overall.sort_values("Month", key=lambda s: s.astype(int)) if not monthly_overall.empty else monthly_overall
            dash.write_row(trend_tbl_start[0], trend_tbl_start[1], ["Month","Agency Overall %"])
            for i,row in tdf.iterrows():
                dash.write(trend_tbl_start[0]+1+i, trend_tbl_start[1],   row["Month Name"])
                dash.write(trend_tbl_start[0]+1+i, trend_tbl_start[1]+1, row["Agency Overall %"]/100.0)

            drops_tbl_start = (10, 11)
            ddf = pd.DataFrame(columns=["Center Name","Drops"])
            if include_drops:
                drop_col = cur_drop_col or cum_drop_col
                if drop_col and drop_col in month_df.columns:
                    tmp = month_df.copy()
                    tmp[drop_col] = pd.to_numeric(tmp[drop_col], errors="coerce")
                    ddf = tmp.groupby("Center Name", dropna=False)[drop_col].sum(min_count=1).reset_index()
                    ddf = ddf.sort_values(drop_col, ascending=False).rename(columns={drop_col:"Drops"})
            dash.write_row(drops_tbl_start[0], drops_tbl_start[1], ["Center Name","Drops"])
            for i,row in ddf.iterrows():
                dash.write(drops_tbl_start[0]+1+i, drops_tbl_start[1],   row["Center Name"])
                dash.write(drops_tbl_start[0]+1+i, drops_tbl_start[1]+1, row["Drops"])

            # Charts
            bar = wb.add_chart({"type":"column"})
            bar.add_series({
                "name":"Attendance %",
                "categories":["Dashboard", centers_tbl_start[0]+1, centers_tbl_start[1], centers_tbl_start[0]+len(cmdf), centers_tbl_start[1]],
                "values":["Dashboard", centers_tbl_start[0]+1, centers_tbl_start[1]+1, centers_tbl_start[0]+len(cmdf), centers_tbl_start[1]+1],
                "data_labels":{"value":True, "num_format":"0.00%"},
                "fill":{"color":PRIMARY_BLUE}, "border":{"color":PRIMARY_BLUE},
            })
            bar.set_title({"name": f"Centers by Attendance — {sel_month_name}"})
            bar.set_y_axis({"num_format":"0.00%"})
            dash.insert_chart(4, 5, bar, {"x_scale":1.1, "y_scale":1.1})

            line = wb.add_chart({"type":"line"})
            line.add_series({
                "name":"Agency Overall %",
                "categories":["Dashboard", trend_tbl_start[0]+1, trend_tbl_start[1], trend_tbl_start[0]+len(tdf), trend_tbl_start[1]],
                "values":["Dashboard", trend_tbl_start[0]+1, trend_tbl_start[1]+1, trend_tbl_start[0]+len(tdf), trend_tbl_start[1]+1],
                "marker":{"type":"circle","size":5},
                "line":{"color":PRIMARY_RED,"width":2},
            })
            line.set_title({"name":"Agency Trend — Aug to Jun"})
            line.set_y_axis({"num_format":"0.00%"})
            dash.insert_chart(20, 1, line, {"x_scale":1.1, "y_scale":1.1})

            if include_drops:
                drop_chart = wb.add_chart({"type":"column"})
                drop_chart.add_series({
                    "name":"Drops",
                    "categories":["Dashboard", drops_tbl_start[0]+1, drops_tbl_start[1], drops_tbl_start[0]+len(ddf), drops_tbl_start[1]],
                    "values":["Dashboard", drops_tbl_start[0]+1, drops_tbl_start[1]+1, drops_tbl_start[0]+len(ddf), drops_tbl_start[1]+1],
                    "data_labels":{"value":True},
                    "fill":{"color":PRIMARY_RED}, "border":{"color":PRIMARY_RED},
                })
                drop_chart.set_title({"name": f"Drop Trends — {sel_month_name}"})
                dash.insert_chart(20, 5, drop_chart, {"x_scale":1.1, "y_scale":1.1})

    st.download_button(
        "⬇️ Download Excel (ADA{}{})".format(
            " + Dashboard" if include_dashboard else "",
            " + Drops" if include_drops else ""
        ),
        data=output.getvalue(),
        file_name=f"ADA_ByCampus_Classes_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload the Enrollment workbook to generate your report.")
