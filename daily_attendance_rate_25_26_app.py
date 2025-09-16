import pandas as pd
import numpy as np
from datetime import datetime
import calendar
from pathlib import Path
import streamlit as st
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown("### Daily Attendance Rate 25-26 Export Tool")

uploaded_file = st.file_uploader("Upload Enrollment Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="V12POP_ERSEA_Enrollment", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={
        "Unnamed: 6": "Funded",
        "Unnamed: 8": "Current",
        "Unnamed: 9": "Attendance Rate"
    })
    cols = ['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate']
    df = df[[c for c in cols if c in df.columns]]
    for c in ['Year','Month','Funded','Current','Attendance Rate']:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    months_selected = sorted(set(int(m) for m in df['Month'].dropna().unique()))
    work = df[df['Month'].isin(months_selected)].copy()
    class_rows = work[work['Class Name'].notna()].copy()
    class_rows['Normalized Average'] = class_rows['Attendance Rate'].astype(float)
    def center_totals(g):
        weights = g['Current'].fillna(0.0).astype(float)
        rates = g['Attendance Rate'].fillna(0.0).astype(float)
        w_avg = (rates * weights).sum() / weights.sum() if weights.sum() > 0 else np.nan
        return pd.Series({
            'Year': g['Year'].dropna().iloc[0] if g['Year'].notna().any() else np.nan,
            'Month': g['Month'].dropna().iloc[0] if g['Month'].notna().any() else np.nan,
            'Center Name': g['Center Name'].dropna().iloc[0],
            'Class Name': 'TOTAL',
            'Funded': g['Funded'].sum(min_count=1),
            'Current': g['Current'].sum(min_count=1),
            'Attendance Rate': float(w_avg) if pd.notna(w_avg) else np.nan,
            'Normalized Average': float(w_avg) if pd.notna(w_avg) else np.nan
        })
    center_total_df = class_rows.groupby('Center Name', dropna=False).apply(center_totals).reset_index(drop=True)
    combined_parts = []
    for center, g in class_rows.groupby('Center Name'):
        combined_parts.append(g.sort_values('Class Name'))
        combined_parts.append(center_total_df[center_total_df['Center Name'] == center])
    combined = pd.concat(combined_parts, ignore_index=True)
    weights_overall = class_rows['Current'].fillna(0.0).astype(float)
    rates_overall = class_rows['Attendance Rate'].fillna(0.0).astype(float)
    overall_weighted = (rates_overall * weights_overall).sum() / weights_overall.sum() if weights_overall.sum() > 0 else np.nan
    overall_df = pd.DataFrame([{
        'Year': class_rows['Year'].dropna().iloc[0] if class_rows['Year'].notna().any() else np.nan,
        'Month': class_rows['Month'].dropna().iloc[0] if class_rows['Month'].notna().any() else np.nan,
        'Center Name': 'HCHSP (Overall)',
        'Class Name': 'TOTAL',
        'Funded': class_rows['Funded'].sum(min_count=1),
        'Current': class_rows['Current'].sum(min_count=1),
        'Attendance Rate': float(overall_weighted) if pd.notna(overall_weighted) else np.nan,
        'Normalized Average': float(overall_weighted) if pd.notna(overall_weighted) else np.nan
    }])
    final_df = pd.concat([combined, overall_df], ignore_index=True)
    def month_name(m):
        try:
            return calendar.month_name[int(m)]
        except Exception:
            return str(m)
    month_labels = sorted({month_name(m) for m in work['Month'].dropna().unique()}, key=lambda x: list(calendar.month_name).index(x) if x in list(calendar.month_name) else 0)
    month_label_text = ", ".join([m for m in month_labels if m])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet = "ADA"
        final_df.to_excel(writer, index=False, sheet_name=sheet, startrow=2)
        wb = writer.book
        ws = writer.sheets[sheet]
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#2E75B6", "align": "center", "valign": "vcenter", "text_wrap": True})
        bold_fmt = wb.add_format({"bold": True})
        bold_pct_fmt = wb.add_format({"bold": True, "num_format": "0.00%"})
        percent_fmt = wb.add_format({"num_format": "0.00%"})
        title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 14})
        red_fmt = wb.add_format({"bold": True, "font_color": "red", "align": "center", "valign": "vcenter", "font_size": 12})
        for col_num, col_name in enumerate(final_df.columns):
            ws.write(2, col_num, col_name, header_fmt)
        ws.set_row(2, 30)
        last_row = len(final_df) + 2
        last_col = len(final_df.columns) - 1
        ws.autofilter(2, 0, last_row, last_col)
        ws.freeze_panes(3, 0)
        for i, col in enumerate(final_df.columns):
            width = max(12, min(32, int(final_df[col].astype(str).map(len).max()) + 2))
            ws.set_column(i, i, width)
        pct_cols = ['Attendance Rate','Normalized Average']
        for col_name in pct_cols:
            if col_name in final_df.columns:
                col_idx = final_df.columns.get_loc(col_name)
                for r in range(len(final_df)):
                    value_0_to_100 = final_df.iloc[r, col_idx]
                    val_decimal = float(value_0_to_100) / 100.0 if pd.notna(value_0_to_100) else None
                    excel_row = r + 3
                    fmt = bold_pct_fmt if str(final_df.iloc[r]['Class Name']).upper() == 'TOTAL' else percent_fmt
                    if val_decimal is None or pd.isna(val_decimal):
                        ws.write_blank(excel_row, col_idx, None, fmt)
                    else:
                        ws.write_number(excel_row, col_idx, val_decimal, fmt)
        for r_idx, row in final_df.iterrows():
            if str(row['Class Name']).upper() == 'TOTAL':
                ws.set_row(r_idx + 3, None, bold_fmt)
        title_text = f"Daily Attendance Rate 25-26 ({month_label_text})"
        ws.merge_range(0, 0, 0, last_col, title_text, title_fmt)
        timestamp = f"(Exported {datetime.now().strftime('%m/%d/%Y %I:%M %p')} CST)"
        ws.merge_range(1, 0, 1, last_col, timestamp, red_fmt)
        if logo_path.exists():
            ws.insert_image(0, 0, str(logo_path), {"x_scale": 0.4, "y_scale": 0.4})
    st.download_button(
        label="Download Daily Attendance Report",
        data=output.getvalue(),
        file_name=f"ADA_ByCampus_Classes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
