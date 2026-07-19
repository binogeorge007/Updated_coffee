"""
refresh_dashboard.py - benchmark-only dashboard refresh.

coffee_briefing.py runs once a day and logs Kirehalli's local Karnataka
prices to the Google Sheet. But ICE Arabica (New York) and ICE Robusta
(London) settle at different times of day, and a single fixed daily run
will always be stale for at least one of them. This script re-fetches just
the global benchmark numbers (ICE Arabica/Robusta + AUD/INR + USD/INR) and
refreshes the dashboard, without touching the Sheet's row history.

Runs on two schedules (see .github/workflows/refresh-benchmarks.yml):
  - shortly after ICE Robusta (London) settles
  - shortly after ICE Arabica (New York) settles

It never appends a new Sheet row - Kirehalli only posts local Karnataka
prices once a day, so a second append would create a duplicate row for the
same date/grade and corrupt the Tracked High/Low + 3/6-month-flag formulas,
which assume exactly one row per grade per calendar day. Instead:
  - if today's row already exists for a grade (written by coffee_briefing.py
    earlier that day), this updates just the ICE/FX/landed-cost/benchmark
    cells on that row in place
  - the dashboard (docs/index.html) is always regenerated from whatever is
    freshest, whether or not today's Sheet row exists yet

HONEST LIMITATIONS - read before relying on the timing:
  - Kirehalli bundles both ICE Arabica and ICE Robusta into a single daily
    post. There is no independent live/free Robusta feed coded here, so an
    "after London close" run only helps if Kirehalli's own page happened to
    update by then - it is not a true independent London-session feed.
  - Trading Economics gives an independent Arabica cross-check, which is
    more likely to reflect a fresher/live print.
  - GitHub Actions cron is fixed UTC and does not shift for daylight saving
    in London or New York, so the cron times in the workflow will drift up
    to an hour off the real market close twice a year (Mar/Nov in the US,
    Mar/Oct in the UK) until manually adjusted.
  - Never fabricates a number: if a fetch fails, existing cells and the
    dashboard are left showing the last successfully fetched values.
"""
import os
import datetime

import gspread

from coffee_briefing import (
    fetch_kirehalli, fetch_ice_arabica_benchmark, fetch_rate,
    get_gsheet_client, get_or_create_worksheets, read_assumptions,
    read_price_history, GRADES, PRICE_LOG_HEADERS,
    milled_equivalent_usd_kg, india_landed_cost, global_benchmark_usd_per_kg,
    tracked_stats, window_flag, build_signal, brazil_png_landed_cost_aud_per_kg,
    render_dashboard_html,
)


def main():
    today = datetime.date.today()
    print(f"=== Benchmark-only refresh for {today.isoformat()} ===")

    print("Fetching Kirehalli (for its embedded ICE prices)...")
    kirehalli = fetch_kirehalli()
    print("Fetching ICE Arabica cross-check (Trading Economics)...")
    ice_arabica_te = fetch_ice_arabica_benchmark()
    print("Fetching AUD/INR and USD/INR (xe.com)...")
    aud_inr = fetch_rate("AUD", "INR")
    usd_inr = fetch_rate("USD", "INR")

    ice_arabica = kirehalli.get("ice_arabica_cents_lb") or ice_arabica_te
    ice_robusta = kirehalli.get("ice_robusta_usd_tonne")

    if not any([ice_arabica, ice_robusta, aud_inr, usd_inr]):
        print("All fetches failed this run - leaving Sheet and dashboard untouched.")
        return

    client = get_gsheet_client()
    ws_log, ws_assump = get_or_create_worksheets(client)
    assumptions = read_assumptions(ws_assump)
    history = read_price_history(ws_log)

    all_values = ws_log.get_all_values()
    if not all_values:
        print("Sheet has no rows yet - run coffee_briefing.py at least once first.")
        return
    headers = all_values[0]

    def hcol(name):
        return headers.index(name) if name in headers else None

    date_col, variety_col, grade_col = hcol("Date"), hcol("Variety"), hcol("Grade")
    mid_col, low_col, high_col = hcol("Mid Price (Rs/50kg)"), hcol("Price Low (Rs/50kg)"), hcol("Price High (Rs/50kg)")
    chg_col = hcol("Change vs prior")

    grade_data, grade_changes = {}, {}
    range_statuses = {"Arabica": {}, "Robusta": {}}
    flags_by_grade, usd_kg_by_grade, milled_by_grade, landed_by_grade = {}, {}, {}, {}
    cell_updates = []

    for variety, grade in GRADES:
        matching = [
            (i, row) for i, row in enumerate(all_values[1:], start=2)
            if len(row) > grade_col and row[variety_col] == variety and row[grade_col] == grade
        ]
        if not matching:
            continue
        todays = [(i, row) for i, row in matching if row[date_col] == today.isoformat()]
        row_i, row = todays[-1] if todays else matching[-1]

        try:
            mid = float(row[mid_col])
        except (ValueError, IndexError, TypeError):
            continue
        try:
            grade_data[(variety, grade)] = (float(row[low_col]), float(row[high_col]))
        except (ValueError, IndexError, TypeError):
            grade_data[(variety, grade)] = None
        grade_changes[(variety, grade)] = row[chg_col] if chg_col is not None and len(row) > chg_col else ""

        weeks, tr_high, tr_low, status = tracked_stats(history, variety, grade, today, mid)
        flag_text, flag_color = window_flag(history, variety, grade, today, mid)
        range_statuses[variety][grade] = status
        flags_by_grade[(variety, grade)] = (flag_text, flag_color)

        ice_price = ice_arabica if variety == "Arabica" else ice_robusta
        if ice_price is None or aud_inr is None or usd_inr is None:
            continue

        landed = india_landed_cost(mid, aud_inr, assumptions, variety, grade)
        usd_kg = (mid / 50) / usd_inr
        milled_usd_kg = milled_equivalent_usd_kg(usd_kg, variety, grade)
        usd_kg_by_grade[(variety, grade)] = usd_kg
        milled_by_grade[(variety, grade)] = milled_usd_kg
        landed_by_grade[(variety, grade)] = landed

        if todays:
            ice_unit = "US cents/lb (ICE Arabica)" if variety == "Arabica" else "USD/tonne (ICE Robusta)"
            bench = global_benchmark_usd_per_kg(variety, ice_price, variety == "Arabica")
            values_by_col = {
                "ICE Futures Price": ice_price,
                "ICE Futures Unit": ice_unit,
                "AUD/INR": aud_inr,
                "USD/INR": usd_inr,
                "Est. AU Landed Cost (AUD/kg)": round(landed, 2),
                "Price (USD/kg)": round(usd_kg, 2),
                "Global Benchmark (USD/kg)": round(bench, 2) if bench else "N/A",
                "Milled/Clean Equivalent (USD/kg)": round(milled_usd_kg, 2),
            }
            for col_name, value in values_by_col.items():
                idx = hcol(col_name)
                if idx is not None:
                    a1 = gspread.utils.rowcol_to_a1(row_i, idx + 1)
                    cell_updates.append({"range": a1, "values": [[value]]})

    if cell_updates:
        ws_log.batch_update(cell_updates, value_input_option="USER_ENTERED")
        print(f"Updated {len(cell_updates)} cell(s) on today's already-logged row(s).")
    else:
        print("No logged row for today yet for any grade - Sheet left untouched; "
              "dashboard below uses the most recently logged local prices with fresh global numbers.")

    bench_usd_kg_arabica = global_benchmark_usd_per_kg("Arabica", ice_arabica, True)
    bench_usd_kg_robusta = global_benchmark_usd_per_kg("Robusta", ice_robusta, False)
    brazil_png_landed = brazil_png_landed_cost_aud_per_kg(ice_arabica, usd_inr, aud_inr, assumptions)
    arabica_signal, arabica_reason = build_signal("Arabica", range_statuses["Arabica"])
    robusta_signal, robusta_reason = build_signal("Robusta", range_statuses["Robusta"])

    print("Rendering refreshed dashboard HTML...")
    html = render_dashboard_html(
        today, grade_data, grade_changes, range_statuses, flags_by_grade,
        milled_by_grade, landed_by_grade, usd_kg_by_grade, ice_arabica,
        ice_robusta, aud_inr, usd_inr, bench_usd_kg_arabica,
        bench_usd_kg_robusta, arabica_signal, arabica_reason,
        robusta_signal, robusta_reason, brazil_png_landed, history,
    )
    os.makedirs("docs", exist_ok=True)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Wrote docs/index.html")
    print("\nDone.")


if __name__ == "__main__":
    main()
