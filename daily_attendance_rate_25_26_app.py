

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
from pathlib import Path
from io import BytesIO

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

# CONFIG
PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"
logo_path = Path("header_logo.png")

# HEADER UI (small preview inside the app)
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_l:
    if logo_path.exists():
        st.image(str(logo_path), width=140)
with hdr_c:
    st.markdown("## Daily Attendance Rate 25-26 Export Tool")
with hdr_r:
    st.write("")

uploaded_file = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # read the file bytes once (do not reuse uploaded_file directly)
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

    # If the workbook has only one sheet, use it automatically
    if len(sheet_names) == 1:
        use_sheet = sheet_names[0]
        st.success(f"Only one sheet found. Using sheet: **{use_sheet}**")
    else:
        # Try preferred sheet name first (exact match)
        if PREFERRED_SHEET in sheet_names:
            use_sheet = PREFERRED_SHEET
            st.success(f"Using preferred sheet: **{use_sheet}**")
        else:
            # fallback: show selectbox for user to pick
            st.warning(f"Preferred sheet '{PREFERRED_SHEET}' not found. Please choose a sheet.")
            use_sheet = st.selectbox("Choose sheet to read", options=sheet_names, index=0)

    # allow header row selection (0-indexed). Default 1 to match your original file
    header_row = st.number_input("Header row (0-indexed). Use 1 if headers are on the 2nd row.", min_value=0, value=1, step=1)

    # read the chosen sheet
    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=int(header_row))
    except Exception as e:
        st.error(f"Failed to read sheet '{use_sheet}'. Error: {e}")
        st.stop()

    if df.empty:
        st.error(f"The selected sheet ('{use_sheet}') appears empty.")
        st.stop()

    # --- PROCESS DATAFRAME (keeps your logic, Normalized removed) ---
    df.columns = [str(c).strip() for c in df.columns]

    # rename expected Unnamed columns if present
    df = df.rename(columns={
        "Unnamed: 6": "Funded",
        "Unnamed: 8": "Current",
        "Unnamed: 9": "Attendance Rate"
    })

    # keep only relevant columns if present
    cols = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
    df = df[[c for c in cols if c in df.columns]]

    # coerce numeric columns
    for c in ['Year', 'Month', 'Funded', 'Current', 'Attendance Rate']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # parse months robustly into ints
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

    # rows with class names
    class_rows = work[work.get('Class Name').notna()].copy()

    # center totals (weighted by 'Current')
    def center_totals(g):
        weights = g['Current'].fillna(0.0).astype(float)
        rates = g['Attendance Rate'].fillna(0.0).astype(float)
        w_avg = (rates * weights).sum() / weights.sum() if weights.sum() > 0 else np.nan
        return pd.Series({
            'Year': g['Year'].dropna().iloc[0] if g['Year'].notna().any() else np.nan,
            'Month': g['Month'].dropna().iloc[0] if g['Month'].notna().any() else np.nan,
            'Center Name': g['Center Name'].dropna().iloc[0] if g['Center Name'].notna().any() else np.nan,
            'Class Name': 'TOTAL',
            'Funded': g['Funded'].sum(min_count=1) if 'Funded' in g else np.nan,
            'Current': g['Current'].sum(min_count=1) if 'Current' in g else np.nan,
            'Attendance Rate': float(w_avg) if pd.notna(w_avg) else np.nan
        })

    if not class_rows.empty:
        center_total_df = class_rows.groupby('Center Name', dropna=False).apply(center_totals).reset_index(drop=True)
    else:
        center_total_df = pd.DataFrame(columns=['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate'])

    # combine details + totals per center
    combined_parts = []
    for center, g in class_rows.groupby('Center Name'):
        combined_parts.append(g.sort_values('Class Name'))
        combined_parts.append(center_total_df[center_total_df['Center Name'] == center])
    combined = pd.concat(combined_parts, ignore_index=True) if combined_parts else pd.DataFrame(columns=center_total_df.columns)

    # overall weighted average
    weights_overall = class_rows['Current'].fillna(0.0).astype(float) if 'Current' in class_rows else pd.Series(dtype=float)
    rates_overall = class_rows['Attendance Rate'].fillna(0.0).astype(float) if 'Attendance Rate' in class_rows else pd.Series(dtype=float)
    overall_weighted = (rates_overall * weights_overall).sum() / weights_overall.sum() if weights_overall.sum() > 0 else np.nan

    overall_df = pd.DataFrame([{
        'Year': class_rows['Year'].dropna().iloc[0] if ('Year' in class_rows and class_rows['Year'].notna().any()) else np.nan,
        'Month': class_rows['Month'].dropna().iloc[0] if ('Month' in class_rows and class_rows['Month'].notna().any()) else np.nan,
        'Center Name': 'HCHSP (Overall)',
        'Class Name': 'TOTAL',
        'Funded': class_rows['Funded'].sum(min_count=1) if 'Funded' in class_rows else np.nan,
        'Current': class_rows['Current'].sum(min_count=1) if 'Current' in class_rows else np.nan,
        'Attendance Rate': float(overall_weighted) if pd.notna(overall_weighted) else np.nan
    }])

    final_df = pd.concat([combined, overall_df], ignore_index=True)

    # month labels
    def month_name(m):
        try:
            return calendar.month_name[int(m)]
        except Exception:
            return str(m)

    month_labels = sorted({month_name(m) for m in work['Month'].dropna().unique()} , key=lambda x: list(calendar.month_name).index(x) if x in list(calendar.month_name) else 0)
    month_label_text = ", ".join([m for m in month_labels if m])

    # Excel export (xlsxwriter)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet = "ADA"
        final_df.to_excel(writer, index=False, sheet_name=sheet, startrow=3)
        wb = writer.book
        ws = writer.sheets[sheet]

        # formats
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#2E75B6", "align": "center", "valign": "vcenter", "text_wrap": True})
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
            for r in range(len(final_df)):
                value_0_to_100 = final_df.iloc[r, col_idx]
                excel_row = r + header_row + 1
                if pd.isna(value_0_to_100):
                    ws.write_blank(excel_row, col_idx, None, percent_fmt)
                else:
                    val_decimal = float(value_0_to_100) / 100.0
                    fmt = bold_pct_fmt if str(final_df.iloc[r].get('Class Name', '')).upper() == 'TOTAL' else percent_fmt
                    ws.write_number(excel_row, col_idx, val_decimal, fmt)

        # bold TOTAL rows
        for r_idx, row in final_df.iterrows():
            if str(row.get('Class Name', '')).upper() == 'TOTAL':
                ws.set_row(r_idx + header_row + 1, None, bold_fmt)

        # title + timestamp; leave space for logo in top-left
        ws.merge_range(0, 1, 1, last_col_index, f"Daily Attendance Rate 25-26 ({month_label_text})", title_fmt)
        chicago_now = datetime.now(ZoneInfo("America/Chicago"))
        timestamp_text = f"(Exported {chicago_now.strftime('%m/%d/%Y %I:%M %p %Z')})"
        ws.merge_range(2, 1, 2, last_col_index, timestamp_text, timestamp_fmt)

        ws.set_row(0, 44)
        ws.set_row(1, 24)
        ws.set_row(2, 18)

        # insert logo top-left
        if logo_path.exists():
            ws.insert_image(0, 0, str(logo_path), {"x_scale": 0.6, "y_scale": 0.6, "x_offset": 4, "y_offset": 4})

    # Download button
    st.download_button(
        label="Download Daily Attendance Report",
        data=output.getvalue(),
        file_name=f"ADA_ByCampus_Classes_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
