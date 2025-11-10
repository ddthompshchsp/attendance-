import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import calendar
import re
import zipfile

st.set_page_config(page_title="ADA Monthly + Cumulative Reports", layout="wide")

# ---- THEME / COLORS ----
BLUE = "#2E75B6"
GREY = "#E6E6E6"

# ---- GROUP DEFINITIONS (your requested changes) ----
# NOTE: Guzman is NOT in Group 1; it is split into two dedicated sheets (4yo / 3yo)
GROUPS = {
    "Group 1": [
        "Donna","Mission","Monte Alto","Sam Fordyce","Seguin","Sam Houston",
        "Wilson","Thigpen","Zavala","Salinas","San Juan"
    ],
    "Group 2": [
        "Chapa","Escandon","Guerra","Palacios","Singleterry","Edinburg North",
        "Alvarez","Farias","Longoria"
    ],
    "Group 3": [
        "Mercedes","Edinburg","Camarena"  # Added Camarena; removed Edinburg North & San Carlos
    ],
    # Guzman handled separately
}

# Simple alias fixes for token matching (keeps display names untouched)
SIMPLE_ALIASES = {
    "fairs":"farias",
    "sequin":"seguin",
    "chaps":"chapa",
    "camerana":"camarena",
    "camarano":"camarena"
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
    """Weighted Attendance Rate (0..100) by Current across CLASS rows."""
    if df.empty or "Attendance Rate" not in df.columns:
        return np.nan
    w = df.get("Current", pd.Series([0]*len(df))).fillna(0).astype(float)
    r = df["Attendance Rate"].fillna(0).astype(float)
    tot = float(w.sum())
    return float((r*w).sum()/tot) if tot > 0 else np.nan

def _center_block(center_df: pd.DataFrame) -> pd.DataFrame:
    """Return class rows + one TOTAL row per center/month (Monthly workbook)."""
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
    if re.search(r'(^|[^0-9])4($|[^0-9])|4yo|4-?yr|4\s*year', t):
        return "4yo"
    if re.search(r'(^|[^0-9])3($|[^0-9])|3yo|3-?yr|3\s*year', t):
        return "3yo"
    return None

def _append_guzman_age(df: pd.DataFrame) -> pd.DataFrame:
    """Add Age Tag for Guzman classes and annotate Class Name with (3yo)/(4yo)."""
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

# <<< CHANGED: robust matcher (prevents 'Edinburg' matching 'Edinburg North')
def _match_centers(simple_list, present_full):
    """
    Strictly map GROUPS short names (e.g., 'Edinburg') to the file's full names like
    'Edinburg (serves ages 3–4)' by comparing SHORT LABELS (text before the first
    '-' or '('). Prevents accidental matches like 'Edinburg' -> 'Edinburg North'.
    """
    def _short_label(full_name: str) -> str:
        return re.split(r'\s*[\-\(–—]\s*', str(full_name), maxsplit=1)[0].strip()

    # Build lookup from CANON(short_label(full)) -> list of full names
    full_map = {}
    for f in present_full:
        sl = _short_label(f)
        key = _norm_simple(sl)
        full_map.setdefault(key, []).append(f)

    matched, seen = [], set()
    for simple in simple_list:
        k = _norm_simple(simple)
        for f in full_map.get(k, []):
            c = _canon(f)
            if c not in seen:
                matched.append(f)
                seen.add(c)
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

# Month controls
months_present = sorted({int(m) for m in df_all['Month'].dropna().astype(int).unique()})
if not months_present:
    st.error("No valid Month values found in the file.")
    st.stop()

colA, colB, colC = st.columns([1,1,2])
with colA:
    # Default to Sept & Oct if present
    default_months = [m for m in [9,10] if m in months_present] or [months_present[-1]]
    sel_months = st.multiselect(
        "Monthly Report Months",
        options=months_present,
        default=default_months,
        format_func=_month_name
    )
with colB:
    # Cumulative month range (inclusive)
    start_m = st.selectbox("Cumulative Start Month", options=months_present, index=months_present.index(min(default_months)))
    end_m = st.selectbox("Cumulative End Month", options=months_present, index=months_present.index(max(default_months)))

if not sel_months:
    st.warning("Select at least one month for the Monthly report.")
    st.stop()

# <<< CHANGED: As-of date & class-days inputs
st.subheader("Cumulative Days Weighting")
as_of = st.date_input("As-of date (for partial current month)", value=date.today())
default_full_month_days = st.number_input("Default class days for FULL months", min_value=1, max_value=31, value=20, step=1)

# Build the list of months in the cumulative range
cum_months = [m for m in months_present if start_m <= m <= end_m]
days_per_month = {}

def _business_days_up_to(d: date) -> int:
    # Mon-Fri business days inclusive of start month first day, exclusive of next day
    start = np.datetime64(date(d.year, d.month, 1))
    end = np.datetime64(d + timedelta(days=1))
    return int(np.busday_count(start, end))

for m in cum_months:
    if m == as_of.month:
        default_days = _business_days_up_to(as_of)
    else:
        default_days = default_full_month_days
    days_per_month[m] = st.number_input(f"Class days for {_month_name(m)}", min_value=0, max_value=31, value=int(default_days), step=1, key=f"days_{m}")

# ---------- BUILD MONTHLY REPORT (unchanged output logic) ----------
def build_monthly_excel(df_all: pd.DataFrame, sel_months: list[int]) -> bytes:
    grouped_src = _append_guzman_age(df_all[df_all['Month'].isin(sel_months)].copy())
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        pct_fmt    = wb.add_format({"num_format": "0.00%"})
        bold_fmt   = wb.add_format({"bold": True})
        s_head     = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})

        # ADA overview for latest selected month
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

        # ADA sheet
        ada.to_excel(writer, index=False, sheet_name="ADA", startrow=3)
        ws = writer.sheets["ADA"]
        header_row = 3
        for c, name in enumerate(ada.columns):
            ws.write(header_row, c, name, header_fmt)
        ws.set_row(header_row, 26)
        for i, col in enumerate(ada.columns):
            width = max(12, min(44, int(ada[col].astype(str).map(len).max()) + 2)) if not ada.empty else 16
            ws.set_column(i, i, width)
        if not ada.empty and 'Attendance Rate' in ada.columns:
            ci = ada.columns.get_loc('Attendance Rate')
            ki = ada.columns.get_loc('Class Name')
            for r in range(len(ada)):
                excel_r = r + header_row + 1
                v = ada.iloc[r, ci]
                if pd.isna(v):
                    ws.write_blank(excel_r, ci, None, pct_fmt)
                else:
                    ws.write_number(excel_r, ci, float(v)/100.0, pct_fmt)
                if str(ada.iloc[r, ki]).upper() == "TOTAL":
                    ws.set_row(excel_r, None, bold_fmt)
        ws.freeze_panes(header_row + 1, 0)
        last_col = max(0, len(ada.columns)-1)
        last_row = header_row + len(ada)
        ws.autofilter(header_row, 0, last_row, last_col)
        ws.write(0, 1, f"Daily Attendance Rate — {_month_name(latest_m)}", wb.add_format({"bold": True}))

        # Per-group writer
        def write_group_sheet(sheet_name: str, df_src: pd.DataFrame):
            ws = wb.add_worksheet(sheet_name[:31])
            start_row = 3
            for fn, g_center in df_src.groupby("Center Name"):
                for m in sorted(g_center["Month"].dropna().unique().astype(int)):
                    g = g_center[g_center["Month"] == m].copy()
                    if g.empty:
                        continue
                    block = _center_block(g)
                    ws.write(start_row - 2, 0, f"{fn} — {_month_name(m)}", wb.add_format({"bold": True}))
                    ws.set_row(start_row, 26)
                    # headers
                    for c, name in enumerate(block.columns):
                        ws.write(start_row, c, name, header_fmt)
                    # values
                    for r in range(len(block)):
                        for c, col in enumerate(block.columns):
                            val = block.iloc[r, c]
                            if col == "Attendance Rate":
                                if pd.isna(val):
                                    ws.write_blank(start_row + 1 + r, c, None, pct_fmt)
                                else:
                                    ws.write_number(start_row + 1 + r, c, float(val)/100.0, pct_fmt)
                            else:
                                if pd.isna(val):
                                    ws.write_blank(start_row + 1 + r, c, None)
                                else:
                                    ws.write(start_row + 1 + r, c, val)
                        if str(block.iloc[r]["Class Name"]).upper() == "TOTAL":
                            ws.set_row(start_row + 1 + r, None, wb.add_format({"bold": True}))
                    # autosize
                    for i, col in enumerate(block.columns):
                        width = max(12, min(44, int(block[col].astype(str).map(len).max()) + 2))
                        ws.set_column(i, i, width)
                    start_row += len(block) + 3
            ws.freeze_panes(4, 0)

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
                # reuse same writer
                write_group_sheet(title, sub)

    return bio.getvalue()

# ---------- BUILD CUMULATIVE REPORT (DAYS-WEIGHTED) ----------
# <<< CHANGED: days-weighted cumulative per CENTER
def build_cumulative_excel(df_all: pd.DataFrame, start_m: int, end_m: int, days_per_month: dict[int,int]) -> bytes:
    months = [m for m in df_all['Month'].dropna().astype(int).unique() if start_m <= m <= end_m]
    if not months:
        months = [int(df_all['Month'].dropna().astype(int).max())]
    df_rng = _append_guzman_age(df_all[df_all['Month'].isin(months)].copy())

    # Step 1: compute CENTER x MONTH monthly rate (weighted by Current across classes)
    center_month = (
        df_rng.groupby(["Center Name","Month"], dropna=False)
              .apply(_weighted_avg_rate)
              .reset_index(name="Monthly Rate")   # 0..100
    )

    # Step 2: compute days-weighted cumulative rate per CENTER
    def _cum_days_weight(g: pd.DataFrame) -> pd.Series:
        # Use provided days_per_month; if missing, treat as 0 (ignored)
        weights = g["Month"].map(lambda m: float(days_per_month.get(int(m), 0)))
        rates = g["Monthly Rate"].astype(float).fillna(0.0)
        wsum = float(weights.sum())
        cum_rate = float((rates * weights).sum() / wsum) if wsum > 0 else np.nan
        return pd.Series({"Attendance Rate": cum_rate, "Days Sum": wsum})

    centers_totals = (
        center_month.groupby("Center Name", dropna=False)
                    .apply(_cum_days_weight)
                    .reset_index()
    )

    # Optional: per-class cumulative (kept simple: show last known Funded/Current)
    def _class_stub(dfcls: pd.DataFrame) -> pd.Series:
        # Not days-weighted at class level (your ask is center-level), but we keep a stub table
        return pd.Series({
            "Month Range": f"{_month_name(start_m)}–{_month_name(end_m)}",
            "Funded": dfcls["Funded"].max(skipna=True) if "Funded" in dfcls else np.nan,
            "Current": dfcls["Current"].max(skipna=True) if "Current" in dfcls else np.nan,
        })

    cum_class = (
        df_rng.groupby(["Center Name","Class Name"], dropna=False)
              .apply(_class_stub)
              .reset_index()
    )

    # Build workbook
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": BLUE, "align": "center", "valign": "vcenter"})
        pct_fmt    = wb.add_format({"num_format": "0.00%"})
        s_head     = wb.add_format({"bold": True, "bg_color": GREY, "border": 1, "font_size": 9})
        s_cell     = wb.add_format({"border": 1, "font_size": 9})
        s_pct      = wb.add_format({"border": 1, "num_format": "0.00%", "font_size": 9})

        # Dashboard: ranked centers cumulative (days-weighted)
        centers_rank = centers_totals[["Center Name","Attendance Rate"]].sort_values("Attendance Rate", ascending=False).reset_index(drop=True)
        ws0 = wb.add_worksheet("Dashboard")
        ws0.write(0, 1, f"Cumulative ADA (Days-weighted) — {_month_name(start_m)} to {_month_name(end_m)}", wb.add_format({"bold": True, "font_size": 22, "font_color": BLUE}))
        ws0.write(2, 1, "Center Name", s_head); ws0.write(2, 2, "Attendance %", s_head)
        for i, row in centers_rank.iterrows():
            ws0.write_string(3+i, 1, str(row["Center Name"]), s_cell)
            val = row["Attendance Rate"]
            if pd.isna(val):
                ws0.write_blank(3+i, 2, None, s_pct)
            else:
                ws0.write_number(3+i, 2, float(val)/100.0, s_pct)
        # chart
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

        # Per-group sheets: class table + center totals (days-weighted)
        present_full2 = df_rng['Center Name'].dropna().unique().tolist()
        for sheet_name, simple_list in GROUPS.items():
            matched = _match_centers(simple_list, present_full2)
            ws = wb.add_worksheet(sheet_name[:31])

            # classes (informational)
            view_cls = cum_class[cum_class["Center Name"].isin(matched)].copy()
            class_rows = view_cls[["Center Name","Class Name","Month Range","Funded","Current"]].copy()
            head_row = 3
            headers1 = list(class_rows.columns) + []  # no rate here by design
            for c, name in enumerate(headers1):
                ws.write(head_row, c, name, header_fmt)
            for r in range(len(class_rows)):
                for c, col in enumerate(headers1):
                    val = class_rows.iloc[r, c]
                    if pd.isna(val):
                        ws.write_blank(head_row + 1 + r, c, None)
                    else:
                        ws.write(head_row + 1 + r, c, val)
            for i, col in enumerate(headers1):
                width = max(12, min(44, int(class_rows[col].astype(str).map(len).max()) + 2)) if not class_rows.empty else 16
                ws.set_column(i, i, width)

            # center totals (days-weighted)
            start_r = head_row + 2 + len(class_rows)
            view_tot = centers_totals[centers_totals["Center Name"].isin(matched)][["Center Name","Attendance Rate","Days Sum"]].copy()
            headers2 = ["Center Name","Attendance % (days-weighted)","Days Sum"]
            for c, name in enumerate(headers2):
                ws.write(start_r, c, name, header_fmt)
            for i, (_, row) in enumerate(view_tot.iterrows()):
                ws.write_string(start_r + 1 + i, 0, str(row["Center Name"]))
                ws.write_number(start_r + 1 + i, 1, 0 if pd.isna(row["Attendance Rate"]) else float(row["Attendance Rate"])/100.0, pct_fmt)
                ws.write_number(start_r + 1 + i, 2, float(row["Days Sum"]) if pd.notna(row["Days Sum"]) else 0)

            # chart
            if not view_tot.empty:
                ch2 = wb.add_chart({"type":"column"})
                ch2.add_series({
                    "name": "Cumulative ADA",
                    "categories": [sheet_name[:31], start_r + 1, 0, start_r + len(view_tot), 0],
                    "values":     [sheet_name[:31], start_r + 1, 1, start_r + len(view_tot), 1],
                    "data_labels": {"value": True, "num_format": "0.00%", "font": {"size": 9}},
                })
                ch2.set_y_axis({"num_format":"0.00%"})
                ch2.set_title({"name":"Cumulative ADA by Center (Days-weighted)"})
                ws.insert_chart(start_r, 6, ch2, {"x_scale": 1.0, "y_scale": 1.0})

            ws.freeze_panes(4, 0)

        # Guzman split cumulative (optional days-weighting by center name 'Guzman')
        guz_names = [n for n in present_full2 if "guzman" in str(n).lower()]
        if guz_names:
            # Filter center_month to Guzman only
            guz_cm = center_month[center_month["Center Name"].isin(guz_names)].copy()
            if not guz_cm.empty:
                # days-weighted cumulative
                w = guz_cm["Month"].map(lambda m: float(days_per_month.get(int(m), 0))).to_numpy()
                r = guz_cm["Monthly Rate"].astype(float).to_numpy()
                wsum = w.sum()
                rate = float((r*w).sum()/wsum) if wsum>0 else np.nan

                ws = wb.add_worksheet("Guzman — Cumulative")
                ws.write(3, 0, "Center Name", header_fmt)
                ws.write(3, 1, "Attendance Rate", header_fmt)
                ws.write(3, 2, "Days Sum", header_fmt)
                ws.write(4, 0, "Guzman")
                if pd.isna(rate):
                    ws.write_blank(4, 1, None, pct_fmt)
                else:
                    ws.write_number(4, 1, rate/100.0, pct_fmt)
                ws.write_number(4, 2, wsum)

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
