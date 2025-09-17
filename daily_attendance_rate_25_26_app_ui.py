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

    # --- Month/Overall mini-table (moved down below main table) ---
    s_head = wb.add_format({"bold": True, "bg_color": GREY_HEADER, "border": 1})
    s_cell = wb.add_format({"border": 1})
    s_pct  = wb.add_format({"border": 1, "num_format": "0.00%"})

    mini_r = last_tab_row + 3  # a few blank rows
    dash.write(mini_r, left_c,   "Month",             s_head)
    dash.write(mini_r, left_c+1, "Agency Overall %",  s_head)
    dash.write(mini_r+1, left_c, sel_month, s_cell)
    dash.write_number(mini_r+1, left_c+1, (agency_overall/100.0) if pd.notna(agency_overall) else 0, s_pct)

    # --- Bar chart (>=95% red) ---
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
    bar.set_y_axis({"num_format": "0.00%", "min": 0.86, "max": 0.98})
    bar.set_x_axis({"label_position": "low", "num_font": {"size": 9, "rotation": -45}})
    bar.set_legend({"none": True})
    bar.set_title({"name": f"Attendance Rate — {sel_month}"})
    dash.insert_chart(5, 5, bar, {"x_scale": 1.35, "y_scale": 1.22})

    # --- Red KPI rounded rectangle ---
    kpi_text = f"{agency_overall:.2f}%"
    kpi_title = "Agency Overall"
    dash.insert_shape(1, 12, {
        "type": "rounded_rectangle",
        "width": 240, "height": 90,
        "text": f"{kpi_title}\n{kpi_text}",
        "fill": {"color": RED},
        "line": {"color": RED},
        "font": {"bold": True, "color": "white", "size": 20},
        "align": {"vertical": "vcenter", "horizontal": "center"},
    })

    # --- Trend chart (kept) ---
    trend = wb.add_chart({"type": "line"})
    trend.add_series({
        "name": "Agency Overall %",
        "categories": ["Dashboard", mini_r+1, left_c, mini_r+1, left_c],
        "values":     ["Dashboard", mini_r+1, left_c+1, mini_r+1, left_c+1],
        "data_labels": {"value": True, "num_format": "0.00%"},
    })
    trend.set_title({"name": "Agency Attendance Trend — 2025-2026"})
    dash.insert_chart(mini_r+5, left_c, trend, {"x_scale": 1.2, "y_scale": 1.0})
