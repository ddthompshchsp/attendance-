import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import calendar, re, zipfile

st.set_page_config(page_title="ADA Monthly + Cumulative Reports", layout="wide")

# ---- THEME / COLORS ----
BLUE = "#2E75B6"
RED  = "#C00000"
GREY = "#E6E6E6"

# ---- GROUP DEFINITIONS ----
# NOTE: Guzman is NOT in Group 1; handled separately (split 3yo/4yo).
# Per your change: Group 3 ADD Camarena.
GROUPS = {
    "Group 1": [
        "Donna","Mission","Monte Alto","Sam Fordyce","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    "Group 2": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    "Group 3": ["Mercedes","Edinburg","Camarena"],  # Camarena added
    # Guzman handled separately (split by 3yo/4yo)
}

# Optional typos/aliases
SIMPLE_ALIASES = {
    "fairs":"farias","sequin":"seguin","chaps":"chapa",
    "camerana":"camarena","camarano":"camarena"
}

# ---- HELPERS ----
def _month_name(m: int) -> str:
    try:
        return calendar.month_name[int(m)]
    except Exception:
        return str(m)

def _canon(x: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', str(x).lower())

def _norm_simple(s: str) -> str:
    k = _canon(s)
    return SIMPLE_ALIASES.get(k, k)

def _weighted_avg_rate(df: pd.DataFrame) -> float:
    """Center/class monthly ADA weighted by Current (0..100, not fraction)."""
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    w = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    r = df["Attendance Rate"].fillna(0).astype(float)
    tot = float(w.sum())
    return float((r*w).sum()/tot) if tot > 0 else np.nan

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

def _detect_age(text: str):
    t = str(text).lower()
    if re.search(r'(^|[^0-9])4($|[^0-9])|4yo|4-?yr|4\s*year', t): return "4yo"
    if re.search(r'(^|[^0-9])3($|[^0-9])|3yo|3-?yr|3\s*year', t): return "3yo"
    return None

def _append_guzman_age(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Center Name" not in out.columns or "Class Name" not in out.columns:
        return out
    mask = out["Center Name"].astype(str).str.lower().str.contains("guzman")
    out.loc[mask, "Age Tag"] = out.loc[mask, "Class Name"].astype(str).map(_detect_age)
    out.loc[mask, "Class Name"] = out.loc[mask].apply(
        lambda r: f"{r['Class Name']} ({r['Age Tag']})" if pd.notna(r.get("Age Tag")) else r["Class Name"],
        axis=1
    )
    return out

def _load_enrollment(file_bytes: bytes, preferred_sheet="V12POP_ERSEA_Enrollment") -> pd.DataFrame:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    use_sheet = preferred_sheet if preferred_sheet in xls.sheet_names else xls.sheet_names[0]
    df = None
    for hdr in (1, 0):
        try:
            tmp = pd.read_excel(BytesIO(file_bytes), sheet_name=use_sheet, header=hdr)
            if not tmp.empty:
                df = tmp
                break
        except Exception:
            pass
    if df is None or df.empty:
        raise ValueError("Could not read a non-empty sheet from the workbook.")
    df.columns = [str(c).strip() for c in df.columns]
    # remap common unnamed numerics
    df = df.rename(columns={"Unnamed: 6":"Funded","Unnamed: 8":"Current","Unnamed: 9":"Attendance Rate"})
    keep = ['Year', 'Month', 'Center Name', 'Class Name', 'Funded', 'Current', 'Attendance Rate']
    df = df[[c for c in keep if c in df.columns]].copy()
    for c in ['Year','Month','Funded','Current','Attendance Rate']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    if 'Class Name' in df.columns:
        mask_total = df['Class Name'].astype(str).str.upper().eq('TOTAL')
        mask_blank = df['Class Name'].astype(str).str.strip().eq('') | df['Class Name'].isna()
        df = df[~(mask_total | mask_blank)].copy()
    return df

# STRICT matcher: compares SHORT LABELS (before '-' or '('). Prevents "Edinburg" matching "Edinburg North".
def _match_centers(simple_list, present_full):
    def _short_label(full_name: str) -> str:
        return re.split(r'\s*[\-\(–—]\s*', str(full_name), maxsplit=1)[0].strip()
    # Map of normalized short labels -> list of full names present
    full_map = {}
    for f in present_full:
        key = _norm_simple(_short_label(f))
        full_map.setdefault(key, []).append(f)
    matched, seen = [], set()
    for simple in simple_list:
        k = _norm_simple(simple)
        for f in full_map.get(k, []):
            c = _canon(f)
            if c not in seen:
                matched.append(f); seen.add(c)
    return matched

# ---------- UI ----------
st.title("ADA Monthly & Cumulative Reports (Groups + Guzman Split)")

uploaded = st.file_uploader("Upload Enrollment Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Upload the enrollment workbook to begin.")
    st.stop()

file_bytes = uploaded.read()
if not file_bytes:
    st.error("Uploaded file is empty.")
    st.stop()

try:
    df_all = _load_enrollment(file_bytes)
except Exception as e:
    st.error(f"Failed to read workbook: {e}")
    st.stop()

months_present = sorted({int(m) for m in df_all['Month'].dropna().astype(int).unique()})
if not months_present:
    st.error("No valid Month values found in the file.")
    st.stop()

colA, colB, colC = st.columns([1,1,2])
with colA:
    default_months = [m for m in [9,10] if m in months_present] or [months_present[-1]]
    sel_months = st.multiselect(
        "Monthly Report Months",
        options=months_present,
        default=default_months,
        format_func=_month_name
    )
with colB:
    start_m = st.selectbox("Cumulative Start Month", options=months_present,
                           index=months_present.index(min(default_months)))
    end_m = st.selectbox("Cumulative End Month", options=months_present,
                         index=months_present.index(max(default_months)))

if not sel_months:
    st.warning("Select at least one month for the Monthly report.")
    st.stop()

# ---- As-of date + class-day weights (for cumulative) ----
st.subheader("Cumulative Days Weighting")
as_of = st.date_input("As-of date (for partial current month)", value=date.today())
default_full_month_days = st.number_input("Default class days for FULL months", min_value=1, max_value=31, value=20, step=1)

cum_months = [m for m in months_present if start_m <= m <= end_m]

def _business_days_up_to(d: date) -> int:
    start = np.datetime64(date(d.year, d.month, 1))
    end   = np.datetime64(d + timedelta(days=1))
    return int(np.busday_count(start, end))

days_per_month = {}
for m in cum_months:
    pref = _business_days_up_to(as_of) if (m == as_of.month) else default_full_month_days
    days_per_month[m] = st.number_input(
        f"Class days for {_month_name(m)}",
        min_value=0, max_value=31, value=int(pref), step=1, key=f"days_{m}"
    )

# ---------- BUILD MONTHLY REPORT ----------
def build_monthly_excel(df_all: pd.DataFrame, sel_months: list[int]) -> bytes:
    grouped_src = _append_guzman_age(df_all[df_all['Month'].isin(sel_months)].copy())
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        FMT_HEADER = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        FMT_PCT    = wb.add_format({"num_format": "0.00%"})
        FMT_BOLD   = wb.add_format({"bold": True})
        FMT_CELL   = wb.add_format({})
        FMT_GRID   = wb.add_format({"border": 1, "font_size": 9})
        FMT_GRID_P = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})
        FMT_HEAD_S = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})

        # --- ADA overview for latest month ---
        latest_m = max(sel_months)
        m_df = grouped_src[grouped_src["Month"] == latest_m].copy()
        if not m_df.empty:
            blocks = [_center_block(g) for _, g in m_df.groupby("Center Name", dropna=False)]
            ada = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(columns=m_df.columns)
            agency_overall = _weighted_avg_rate(m_df)
            if not ada.empty:
                ada = pd.concat([ada, pd.DataFrame([{
                    "Year": m_df['Year'].dropna().iloc[0] if m_df['Year'].notna().any() else np.nan,
                    "Month": latest_m,
                    "Center Name": "HCHSP (Overall)",
                    "Class Name": "TOTAL",
                    "Funded": m_df['Funded'].sum(min_count=1),
                    "Current": m_df['Current'].sum(min_count=1),
                    "Attendance Rate": float(agency_overall) if pd.notna(agency_overall) else np.nan,
                }])], ignore_index=True)
        else:
            ada = pd.DataFrame(columns=['Year','Month','Center Name','Class Name','Funded','Current','Attendance Rate'])

        # Write ADA sheet
        ada.to_excel(writer, index=False, sheet_name="ADA", startrow=3)
        ws = writer.sheets["ADA"]
        header_row = 3
        # style header row
        for c, name in enumerate(ada.columns):
            ws.write(header_row, c, name, FMT_HEADER)
        ws.set_row(header_row, 26)
        # widths
        for i, col in enumerate(ada.columns):
            width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2)) if not ada.empty else 16
            ws.set_column(i, i, width)
        # write cells and bold TOTAL per-cell (avoid set_row formatting)
        if not ada.empty and 'Attendance Rate' in ada.columns:
            ci = ada.columns.get_loc('Attendance Rate')
            ki = ada.columns.get_loc('Class Name')
            for r in range(len(ada)):
                excel_r = r + header_row + 1
                is_total = str(ada.iloc[r, ki]).upper() == "TOTAL"
                for c, colname in enumerate(ada.columns):
                    val = ada.iloc[r, c]
                    # choose bold fmt for TOTAL cells (including %)
                    if c == ci:
                        if pd.isna(val):
                            ws.write_blank(excel_r, c, None, FMT_PCT if not is_total else wb.add_format({"num_format":"0.00%","bold":True}))
                        else:
                            ws.write_number(excel_r, c, float(val)/100.0, FMT_PCT if not is_total else wb.add_format({"num_format":"0.00%","bold":True}))
                    else:
                        if pd.isna(val):
                            ws.write_blank(excel_r, c, None, FMT_BOLD if is_total else FMT_CELL)
                        else:
                            ws.write(excel_r, c, val, FMT_BOLD if is_total else FMT_CELL)

        ws.freeze_panes(header_row + 1, 0)
        last_col = max(0, len(ada.columns)-1)
        last_row = header_row + len(ada)
        ws.autofilter(header_row, 0, last_row, last_col)
        ws.write(0, 1, f"Daily Attendance Rate — {_month_name(latest_m)}", FMT_BOLD)

        # ---- per-group writer ----
        def write_group_sheet(sheet_name: str, df_src: pd.DataFrame):
            wsg = wb.add_worksheet(sheet_name[:31])
            start_row = 3
            for fn, g_center in df_src.groupby("Center Name"):
                for m in sorted(g_center["Month"].dropna().unique().astype(int)):
                    g = g_center[g_center["Month"] == m].copy()
                    if g.empty:
                        continue
                    block = _center_block(g)
                    wsg.write(start_row - 2, 0, f"{fn} — {_month_name(m)}", FMT_BOLD)
                    # headers
                    for c, name in enumerate(block.columns):
                        wsg.write(start_row, c, name, FMT_HEADER)
                    # values
                    for r in range(len(block)):
                        is_total = str(block.iloc[r]["Class Name"]).upper() == "TOTAL"
                        for c, col in enumerate(block.columns):
                            val = block.iloc[r, c]
                            if col == "Attendance Rate":
                                if pd.isna(val):
                                    wsg.write_blank(start_row + 1 + r, c, None, FMT_GRID_P if not is_total else wb.add_format({"border":1,"num_format":"0.00%","font_size":9,"bold":True}))
                                else:
                                    wsg.write_number(start_row + 1 + r, c, float(val)/100.0, FMT_GRID_P if not is_total else wb.add_format({"border":1,"num_format":"0.00%","font_size":9,"bold":True}))
                            else:
                                if pd.isna(val):
                                    wsg.write_blank(start_row + 1 + r, c, None, FMT_GRID if not is_total else wb.add_format({"border":1,"font_size":9,"bold":True}))
                                else:
                                    wsg.write(start_row + 1 + r, c, val, FMT_GRID if not is_total else wb.add_format({"border":1,"font_size":9,"bold":True}))
                    # autosize
                    for i, col in enumerate(block.columns):
                        width = max(12, min(44, int(block[col].astype(str).map(len).max()) + 2))
                        wsg.set_column(i, i, width)
                    start_row += len(block) + 3
            wsg.freeze_panes(4, 0)

            # Compact summary (Month, Center, %)
            sums = (
                df_src.groupby(["Month","Center Name"], dropna=False)
                      .apply(_weighted_avg_rate)
                      .reset_index(name="Attendance %")
                      .sort_values(["Month","Attendance %"], ascending=[True, False])
            )
            small_r, small_c = start_row, 0
            wsg.write(small_r,   small_c,   "Month",        FMT_HEAD_S)
            wsg.write(small_r,   small_c+1, "Center Name",  FMT_HEAD_S)
            wsg.write(small_r,   small_c+2, "Attendance %", FMT_HEAD_S)
            for i, row in sums.iterrows():
                wsg.write_string(small_r + 1 + i, small_c,   _month_name(int(row["Month"])) if pd.notna(row["Month"]) else "", FMT_GRID)
                wsg.write_string(small_r + 1 + i, small_c+1, str(row["Center Name"]), FMT_GRID)
                if pd.isna(row["Attendance %"]):
                    wsg.write_blank(small_r + 1 + i, small_c+2, None, FMT_GRID_P)
                else:
                    wsg.write_number(small_r + 1 + i, small_c+2, float(row["Attendance %"])/100.0, FMT_GRID_P)
            last_row_small = small_r + len(sums)
            wsg.autofilter(small_r, small_c, last_row_small, small_c + 2)

            # Charts: one per month
            months_for_charts = sorted(df_src["Month"].dropna().unique().astype(int).tolist())
            chart_col_offset = 5
            for idx, m in enumerate(months_for_charts):
                rows_for_m = [i for i, (_, r) in enumerate(sums.iterrows()) if int(r["Month"]) == int(m)]
                if not rows_for_m:
                    continue
                start_off, end_off = rows_for_m[0], rows_for_m[-1]
                ch = wb.add_chart({"type": "column"})
                ch.add_series({
                    "name": f"{_month_name(m)} Attendance %",
                    "categories": [sheet_name[:31], small_r + 1 + start_off, small_c + 1, small_r + 1 + end_off, small_c + 1],
                    "values":     [sheet_name[:31], small_r + 1 + start_off, small_c + 2, small_r + 1 + end_off, small_c + 2],
                    "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
                })
                ch.set_y_axis({"num_format": "0.00%"})
                ch.set_title({"name": f"{_month_name(m)} — Attendance %"})
                wsg.insert_chart(small_r, small_c + chart_col_offset + idx*8, ch, {"x_scale": 1.1, "y_scale": 1.0})

        # Build Group 1–3 (no Guzman)
        present_full = grouped_src['Center Name'].dropna().unique().tolist()
        for sheet_name, simple_list in GROUPS.items():
            matched = _match_centers(simple_list, present_full)
            df_use = grouped_src[grouped_src["Center Name"].isin(matched)].copy()
            write_group_sheet(sheet_name, df_use)

        # Guzman split
        guz = grouped_src[grouped_src["Center Name"].astype(str).str.lower().str.contains("guzman")].copy()
        if not guz.empty:
            guz["Age Tag"] = guz["Class Name"].astype(str).map(_detect_age)
            for tag, title in [("4yo","Group 4 — Guzman (4yo)"),("3yo","Group 4 — Guzman (3yo)")]:
                sub = guz[guz["Age Tag"]==tag].copy()
                if sub.empty: 
                    continue
                write_group_sheet(title, sub)

    return bio.getvalue()

# ---------- BUILD CUMULATIVE (DAYS-WEIGHTED) ----------
def build_cumulative_excel(df_all: pd.DataFrame, start_m: int, end_m: int, days_per_month: dict[int,int]) -> bytes:
    months = [m for m in df_all['Month'].dropna().astype(int).unique() if start_m <= m <= end_m]
    if not months:
        months = [int(df_all['Month'].dropna().astype(int).max())]
    df_rng = _append_guzman_age(df_all[df_all['Month'].isin(months)].copy())

    # CENTER x MONTH monthly ADA (weighted by Current across classes)
    center_month = (
        df_rng.groupby(["Center Name","Month"], dropna=False)
              .apply(_weighted_avg_rate)
              .reset_index(name="Monthly Rate")   # 0..100
    )

    # CENTER days-weighted cumulative
    def _cum_days_weight(g: pd.DataFrame) -> pd.Series:
        weights = g["Month"].map(lambda m: float(days_per_month.get(int(m), 0)))
        rates   = g["Monthly Rate"].astype(float).fillna(0.0)
        wsum = float(weights.sum())
        cum_rate = float((rates * weights).sum() / wsum) if wsum > 0 else np.nan
        return pd.Series({"Attendance Rate": cum_rate, "Days Sum": wsum})

    centers_totals = (
        center_month.groupby("Center Name", dropna=False)
                    .apply(_cum_days_weight)
                    .reset_index()
    )

    # Build workbook
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        FMT_HEADER = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        FMT_PCT    = wb.add_format({"num_format": "0.00%"})
        FMT_GRID   = wb.add_format({"border": 1, "font_size": 9})
        FMT_GRID_P = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})
        FMT_HEAD_S = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})

        centers_rank = centers_totals[["Center Name","Attendance Rate"]].sort_values("Attendance Rate", ascending=False).reset_index(drop=True)
        ws0 = wb.add_worksheet("Dashboard")
        ws0.write(0, 1, f"Cumulative ADA (Days-weighted) — {_month_name(start_m)} to {_month_name(end_m)}", wb.add_format({"bold": True, "font_size": 22, "font_color": BLUE}))
        ws0.write(2, 1, "Center Name", FMT_HEAD_S); ws0.write(2, 2, "Attendance %", FMT_HEAD_S)
        for i, row in centers_rank.iterrows():
            ws0.write_string(3+i, 1, str(row["Center Name"]), FMT_GRID)
            val = row["Attendance Rate"]
            if pd.isna(val): ws0.write_blank(3+i, 2, None, FMT_GRID_P)
            else:            ws0.write_number(3+i, 2, float(val)/100.0, FMT_GRID_P)
        if not centers_rank.empty:
            ch = wb.add_chart({"type":"column"})
            ch.add_series({
                "name": "Cumulative ADA",
                "categories": ["Dashboard", 3, 1, 3+len(centers_rank)-1, 1],
                "values":     ["Dashboard", 3, 2, 3+len(centers_rank)-1, 2],
                "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
            })
            ch.set_y_axis({"num_format": "0.00%"})
            ch.set_title({"name": "Centers — Cumulative ADA (Days-weighted)"})
            ws0.insert_chart(2, 5, ch, {"x_scale": 1.2, "y_scale": 1.1})

        # Per-group sheets (center totals only, days-weighted)
        present_full2 = df_rng['Center Name'].dropna().unique().tolist()
        for sheet_name, simple_list in GROUPS.items():
            matched = _match_centers(simple_list, present_full2)
            ws = wb.add_worksheet(sheet_name[:31])

            view_tot = centers_totals[centers_totals["Center Name"].isin(matched)][["Center Name","Attendance Rate","Days Sum"]].copy()
            head_row = 3
            headers = ["Center Name","Attendance % (days-weighted)","Days Sum"]
            for c, name in enumerate(headers):
                ws.write(head_row, c, name, FMT_HEADER)
            for i, (_, row) in enumerate(view_tot.iterrows()):
                ws.write_string(head_row + 1 + i, 0, str(row["Center Name"]), FMT_GRID)
                ws.write_number(head_row + 1 + i, 1, 0 if pd.isna(row["Attendance Rate"]) else float(row["Attendance Rate"])/100.0, FMT_GRID_P)
                ws.write_number(head_row + 1 + i, 2, float(row["Days Sum"]) if pd.notna(row["Days Sum"]) else 0, FMT_GRID)

            ws.freeze_panes(4, 0)

            if not view_tot.empty:
                ch2 = wb.add_chart({"type":"column"})
                ch2.add_series({
                    "name": "Cumulative ADA",
                    "categories": [sheet_name[:31], head_row + 1, 0, head_row + len(view_tot), 0],
                    "values":     [sheet_name[:31], head_row + 1, 1, head_row + len(view_tot), 1],
                    "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
                })
                ch2.set_y_axis({"num_format":"0.00%"})
                ch2.set_title({"name":"Cumulative ADA by Center (Days-weighted)"})
                ws.insert_chart(head_row, 5, ch2, {"x_scale": 1.0, "y_scale": 1.0})

        # Transparency: center-month rates
        center_month.to_excel(writer, index=False, sheet_name="Center_Monthly_ADA")

    return bio.getvalue()

# ---------- GENERATE BOTH & ZIP ----------
monthly_bytes = build_monthly_excel(df_all, sel_months)
cumulative_bytes = build_cumulative_excel(df_all, start_m, end_m, days_per_month)

ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y%m%d_%H%M")
zip_buf = BytesIO()
with zipfile.ZipFile(zip_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    zf.writestr(f"ADA_Monthly_Groups_GuzmanSplit_{ts}.xlsx", monthly_bytes)
    zf.writestr(f"ADA_Cumulative_{_month_name(start_m)}_{_month_name(end_m)}_{ts}.xlsx", cumulative_bytes)
zip_buf.seek(0)

st.success("Reports ready! Monthly groups (with Guzman split) + Cumulative (days-weighted).")
st.download_button(
    "Download ZIP (Monthly + Cumulative)",
    data=zip_buf.getvalue(),
    file_name=f"ADA_Reports_{ts}.zip",
    mime="application/zip"
)

with st.expander("Download individual files"):
    st.download_button(
        "Download Monthly Groups (xlsx)",
        data=monthly_bytes,
        file_name=f"ADA_Monthly_Groups_GuzmanSplit_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="mon"
    )
    st.download_button(
        "Download Cumulative (xlsx)",
        data=cumulative_bytes,
        file_name=f"ADA_Cumulative_{_month_name(start_m)}_{_month_name(end_m)}_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="cum"
    )
