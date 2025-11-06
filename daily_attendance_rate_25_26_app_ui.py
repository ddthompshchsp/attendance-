import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo
import calendar
from pathlib import Path
from io import BytesIO
import re

st.set_page_config(page_title="Daily Attendance Rate 25-26", layout="wide")

# ---- COLORS ----
BLUE = "#2E75B6"
RED = "#C00000"
GREY_HEADER = "#E6E6E6"
LIGHT_GRID = "#D9D9D9"
DARK_TEXT = "#1F2937"

PREFERRED_SHEET = "V12POP_ERSEA_Enrollment"
logo_path = Path("header_logo.png")

# ---------- HELPERS ----------
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

def _center_block(center_df: pd.DataFrame) -> pd.DataFrame:
    """Return class rows for one center + a single TOTAL row (weighted %)."""
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

# ---------- CANONICALIZATION / GROUPING ----------
def _canon(x: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', str(x).lower())

# alias map to canonical campus names
CENTER_ALIASES = {
    "samfordyce":"Sam Fordyce",
    "sanfelipe":"San Felipe",
    "sancarlos":"San Carlos",
    "sanhouston":"Sam Houston",
    "samhouston":"Sam Houston",
    "segin":"Seguin",
    "sequin":"Seguin",
    "farias":"Farias",
    "fairs":"Farias",
    "escandon":"Escandon",
    "esandon":"Escandon",
    "edinburgnorth":"Edinburg North",
    "edinburg":"Edinburg",
    "zavala":"Zavala",
    "salinas":"Salinas",
    "sanjuan":"San Juan",
    "mission":"Mission",
    "montealto":"Monte Alto",
    "donna":"Donna",
    "wilson":"Wilson",
    "thigpen":"Thigpen",
    "palacios":"Palacios",
    "singleterry":"Singleterry",
    "guerra":"Guerra",
    "alvarez":"Alvarez",
    "longoria":"Longoria",
    "guzman":"Guzman",
    "chapa":"Chapa", "chaps":"Chapa",
    "mercedes":"Mercedes",
    # passthroughs:
    "samfordyce":"Sam Fordyce",
    "sancarlos":"San Carlos",
    "samhouston":"Sam Houston",
}

def canon_center(name: str) -> str:
    key = _canon(name)
    return CENTER_ALIASES.get(key, str(name))

# Group definitions (sheet titles -> list of centers). Case/spacing doesn’t matter.
CENTER_GROUPS = {
    # Table 1
    "Grp1_Donna_to_SanJuan": [
        "Donna","Mission","Monte Alto","Sam Fordyce","San Carlos","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    # Table 2
    "Grp2_Chapa_to_Longoria": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    # Table 3
    "Grp3_Mercedes_Edinburg": ["Mercedes","Edinburg"],
    # Table 4 (Guzman special)
    "Grp4_Guzman": ["Guzman"]
}

# Tracks campuses already placed on a group sheet (canonicalized)
USED_CENTERS = set()

# ---------- GUZMAN AGE TAGGING ----------
def _detect_age_tag(text: str) -> str:
    """Return '(3yo)' or '(4yo)' if class name hints; else ''."""
    t = str(text).lower()
    if re.search(r'(^|[^0-9])3($|[^0-9])|3yo|3-?yr|3\s*year', t):
        return " (3yo)"
    if re.search(r'(^|[^0-9])4($|[^0-9])|4yo|4-?yr|4\s*year', t):
        return " (4yo)"
    return ""

def _append_guzman_age(df: pd.DataFrame) -> pd.DataFrame:
    """For Guzman rows, append (3yo)/(4yo) to Class Name when detectable."""
    if "Center Name" not in df.columns or "Class Name" not in df.columns:
        return df
    out = df.copy()
    mask = out["Center Name"].astype(str).str.strip().str.lower().eq("guzman")
    out.loc[mask, "Class Name"] = (
        out.loc[mask, "Class Name"]
        .astype(str)
        .apply(lambda s: s + _detect_age_tag(s) if _detect_age_tag(s) else s)
    )
    return out

# ---------- UI HEADER ----------
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    if logo_path.exists():
        st.image(str(logo_path), width=160)
    st.markdown("<h2 style='text-align:center;margin:8px 0;'>Daily Attendance Rate 25-26</h2>", unsafe_allow_html=True)
st.divider()

# ---------- FILE INPUT ----------
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

# Remove pre-existing subtotals/totals from source
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

# ---------- BUILD ADA (class rows + ONE TOTAL per center + agency overall) ----------
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

# ---------- EXPORT (Excel with Dashboard + Group Sheets) ----------
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

    # Month/Agency mini table
    s_head = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1})
    s_cell = wb.add_format({"border": 1})
    s_pct  = wb.add_format({"border": 1, "num_format": "0.00%"})

    mini_r = last_tab_row + 3
    dash.write(mini_r,   left_c,   "Month",             s_head)
    dash.write(mini_r,   left_c+1, "Agency Overall %",  s_head)
    dash.write(mini_r+1, left_c,   sel_month,           s_cell)
    dash.write_number(mini_r+1, left_c+1, (agency_overall/100.0) if pd.notna(agency_overall) else 0, s_pct)

    # Bar chart — title "Attendance Rate — {Month}"
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

    # Agency Overall KPI
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

    # Trend table & chart (agency overall by month)
    mo = []
    for m in sorted(class_rows['Month'].dropna().unique(), key=lambda x: int(x)):
        mm = int(m)
        mdf = class_rows[class_rows['Month'] == mm]
        mo.append({"Month Name": _month_name(mm), "Agency Overall %": _weighted_avg_rate(mdf) / 100.0})
    mo_df = pd.DataFrame(mo)

    tr_r = mini_r + 5
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

    # ---------------- GROUPED CAMPUS TABLES + CHARTS ----------------
    # Work on the current month’s class rows (m_df) and add age tags for Guzman.
    grouped_src = _append_guzman_age(m_df)

    def write_group_sheet(sheet_name: str, centers_list: list, df_src: pd.DataFrame):
        """
        Writes one grouped sheet. Skips any centers already placed on earlier sheets.
        Ensures no duplicate campuses within or across groups.
        """
        # 1) Canonicalize intended center list and remove duplicates locally
        wanted_order = []
        seen_local = set()
        for c in centers_list:
            k = _canon(c)
            if k not in seen_local:
                seen_local.add(k)
                wanted_order.append(c)

        # 2) Exclude centers that were already used on a prior sheet
        allowed = [c for c in wanted_order if _canon(c) not in USED_CENTERS]

        # Always create a sheet so the file structure is stable
        ws_name = sheet_name[:31]
        ws_grp = wb.add_worksheet(ws_name)

        if not allowed:
            ws_grp.write(0, 0, f"All campuses in this group were already placed on other sheets.")
            return

        # 3) Pull rows for allowed centers (respect aliases) and drop any dup rows
        df2 = df_src.copy()
        df2["__cn"] = df2["Center Name"].astype(str).map(canon_center)
        allowed_keys = {_canon(c) for c in allowed}
        keep_mask = df2["__cn"].apply(lambda x: _canon(x) in allowed_keys)
        sub = (
            df2.loc[keep_mask]
               .drop_duplicates(subset=["Center Name", "Class Name", "Month", "Year"], keep="first")
               .drop(columns="__cn")
               .copy()
        )

        if sub.empty:
            ws_grp.write(0, 0, f"No data for the selected month for: {', '.join(allowed)}")
            return

        # GUZMAN SPECIAL: If this group is *only* Guzman (by definition in CENTER_GROUPS),
        # split into two 8-class tables and build per-table charts.
        just_guzman = len(allowed) == 1 and _canon(allowed[0]) == "guzman"
        if just_guzman:
            gz = sub[sub["Center Name"].astype(str).str.strip().str.lower().eq("guzman")].copy()
            if gz.empty:
                ws_grp.write(0, 0, "No Guzman rows for this month.")
                return

            # Prepare per-class rows (exclude any TOTAL from source, we're rebuilding)
            gz_class = gz[~gz["Class Name"].astype(str).str.upper().eq("TOTAL")].copy()
            # Build two chunks of up to 8 classes each
            gz_class = gz_class.sort_values(["Class Name"])
            chunks = [gz_class.iloc[:8].copy(), gz_class.iloc[8:16].copy()]

            title_fmt_g = wb.add_format({"bold": True, "font_size": 16})
            ws_grp.write(0, 1, "Guzman — Split (Two Tables of up to 8 Classes Each)", title_fmt_g)

            start_rows = [3, 3]     # both tables start near top; we'll offset columns
            start_cols = [1, 8]     # second table placed to the right
            chart_offsets = [(0, 0), (0, 0)]

            present_keys = set()

            for idx, part in enumerate(chunks):
                if part.empty:
                    continue

                # Build TOTAL row for this subset (weighted by Current)
                part_total = _center_block(part)
                # Keep only that center's rows (it's all Guzman) and its TOTAL
                # _center_block returns body + TOTAL (we want both)
                out_df = part_total.copy()

                # Write table
                sr, sc = start_rows[idx], start_cols[idx]
                # Header
                for c, name in enumerate(out_df.columns):
                    ws_grp.write(sr, sc + c, name, header_fmt)
                ws_grp.set_row(sr, 26)

                # Values
                for r in range(len(out_df)):
                    for c, col in enumerate(out_df.columns):
                        val = out_df.iloc[r, c]
                        if col == "Attendance Rate":
                            ws_grp.write_number(sr + 1 + r, sc + c, (float(val)/100.0) if pd.notna(val) else 0, pct_fmt)
                        else:
                            ws_grp.write(sr + 1 + r, sc + c, val)
                    # Bold the TOTAL row
                    if str(out_df.iloc[r]["Class Name"]).upper() == "TOTAL":
                        ws_grp.set_row(sr + 1 + r, None, bold_fmt)

                # Autosize a bit (limited to that block)
                for c, col in enumerate(out_df.columns):
                    width = max(12, min(44, int(out_df[col].astype(str).map(len).max()) + 2))
                    ws_grp.set_column(sc + c, sc + c, width)

                # Freeze just above data for first table only
                if idx == 0:
                    ws_grp.freeze_panes(sr + 1, sc)

                # Center summary for chart: pick TOTAL row
                centers_sum = (
                    out_df[out_df["Class Name"].astype(str).str.upper().eq("TOTAL")]
                    .copy()
                )
                if not centers_sum.empty:
                    # Write a compact summary block under this table
                    sum_r = sr + len(out_df) + 3
                    sum_c = sc
                    ws_grp.write(sum_r,   sum_c,   "Center Name", s_head)
                    ws_grp.write(sum_r,   sum_c+1, "Attendance %", s_head)
                    for i, row in centers_sum.iterrows():
                        ws_grp.write_string(sum_r + 1 + i, sum_c,   str(row["Center Name"]), s_cell)
                        ws_grp.write_number(sum_r + 1 + i, sum_c+1, (float(row["Attendance Rate"])/100.0 if pd.notna(row["Attendance Rate"]) else 0), s_pct)

                    # Chart for this chunk
                    chart = wb.add_chart({"type": "column"})
                    chart.add_series({
                        "name": f"Guzman Part {idx+1} — Attendance %",
                        "categories": [ws_name, sum_r + 1, sum_c,   sum_r + len(centers_sum), sum_c],
                        "values":     [ws_name, sum_r + 1, sum_c+1, sum_r + len(centers_sum), sum_c+1],
                        "data_labels": {"value": True, "num_format": "0.00%", "font": {"bold": True, "size": 9}},
                    })
                    chart.set_y_axis({"num_format": "0.00%"})
                    chart.set_legend({"none": True})
                    chart.set_title({"name": f"Guzman Part {idx+1} — Attendance %"})
                    ws_grp.insert_chart(sum_r, sum_c + 3, chart, {"x_scale": 1.15, "y_scale": 1.0})

                present_keys.add("guzman")

            # Mark Guzman as used
            USED_CENTERS.update(present_keys)
            return

        # ---- Standard (non-Guzman) group handling: one big table with per-center TOTALs + 1 chart ----
        blocks = []
        present_keys = set()
        for cn, g in sub.groupby("Center Name", dropna=False):
            blocks.append(_center_block(g))
            present_keys.add(_canon(str(cn)))
        grp_df = pd.concat(blocks, ignore_index=True)

        # Write title
        ws_grp.write(0, 1, sheet_name.replace("_", " "), title_fmt)

        start_row = 3
        # Write header
        for c, name in enumerate(grp_df.columns):
            ws_grp.write(start_row, 1 + c, name, header_fmt)
        ws_grp.set_row(start_row, 26)

        # Write values
        for r in range(len(grp_df)):
            for c, col in enumerate(grp_df.columns):
                val = grp_df.iloc[r, c]
                if col == "Attendance Rate":
                    ws_grp.write_number(start_row + 1 + r, 1 + c, (float(val)/100.0) if pd.notna(val) else 0, pct_fmt)
                else:
                    ws_grp.write(start_row + 1 + r, 1 + c, val)
            if str(grp_df.iloc[r]["Class Name"]).upper() == "TOTAL":
                ws_grp.set_row(start_row + 1 + r, None, bold_fmt)

        # Autosize this block
        for i, col in enumerate(grp_df.columns):
            width = max(12, min(44, int(grp_df[col].astype(str).map(len).max()) + 2))
            ws_grp.set_column(1 + i, 1 + i, width)

        # Freeze + filter
        last_col = 1 + len(grp_df.columns) - 1
        last_row = start_row + len(grp_df)
        ws_grp.freeze_panes(start_row + 1, 1)
        ws_grp.autofilter(start_row, 1, last_row, last_col)

        # Center-level summary for chart (TOTAL rows only)
        centers_sum = (
            grp_df[grp_df["Class Name"].astype(str).str.upper().eq("TOTAL")]
            .copy()
            .sort_values("Attendance Rate", ascending=False)
        )
        if not centers_sum.empty:
            sum_r, sum_c = last_row + 2, 1
            ws_grp.write(sum_r,   sum_c,   "Center Name", s_head)
            ws_grp.write(sum_r,   sum_c+1, "Attendance %", s_head)
            for i, row in centers_sum.iterrows():
                ws_grp.write_string(sum_r + 1 + i, sum_c,   str(row["Center Name"]), s_cell)
                ws_grp.write_number(sum_r + 1 + i, sum_c+1, (float(row["Attendance Rate"])/100.0 if pd.notna(row["Attendance Rate"]) else 0), s_pct)

            # Chart for the group
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
            ws_grp.insert_chart(sum_r, sum_c + 3, chart, {"x_scale": 1.2, "y_scale": 1.1})

        # Mark centers as used
        USED_CENTERS.update(present_keys)

    # IMPORTANT: First matched group gets the campus; later groups will skip it.
    write_group_sheet("Grp1_Donna_to_SanJuan", CENTER_GROUPS["Grp1_Donna_to_SanJuan"], grouped_src)
    write_group_sheet("Grp2_Chapa_to_Longoria", CENTER_GROUPS["Grp2_Chapa_to_Longoria"], grouped_src)
    write_group_sheet("Grp3_Mercedes_Edinburg", CENTER_GROUPS["Grp3_Mercedes_Edinburg"], grouped_src)
    write_group_sheet("Grp4_Guzman",            CENTER_GROUPS["Grp4_Guzman"],            grouped_src)

# Download
st.download_button(
    "Download Report",
    data=output.getvalue(),
    file_name=f"ADA_Dashboard_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

