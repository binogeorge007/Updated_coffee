"""
Daily coffee market briefing - cloud version.

Fetches Karnataka local prices (Kirehalli), ICE Arabica/Robusta benchmarks
(Trading Economics), and AUD/INR + USD/INR rates (xe.com), appends a row per
grade to a Google Sheet, computes a tracked-history high/low + buy/wait signal
per variety, and sends the result over Telegram and email.

Designed to run unattended on GitHub Actions - no local machine required.

CAVEATS - read before relying on this:
  - Kirehalli/Trading Economics/xe.com are scraped with regex/BeautifulSoup.
    They provide no stable API and can change their HTML at any time, which
    will silently break parsing here. Check the Action's run logs after the
    first few runs.
  - Nothing here fabricates a number: if a fetch fails, that field is recorded
    as None / "N/A" and flagged in the message rather than guessed.
  - Brazil and PNG landed-cost figures use the same ICE NY Arabica benchmark
    (there's no live scraped Brazil-specific or PNG-specific spot price) -
    they are "same benchmark, different freight lane" estimates, not real
    origin quotes.
  - AU wholesale (Sydney/Newcastle) rates aren't published anywhere publicly;
    this script does not attempt to fetch them. Log real supplier quotes
    directly in the Google Sheet's manual columns.

Required environment variables (set as GitHub Secrets, injected via
`env:` in the workflow - never commit real values):
  GOOGLE_SERVICE_ACCOUNT_JSON  - full JSON key content for a Google Cloud
                                 service account with Sheets API access
  GOOGLE_SHEET_ID              - the target spreadsheet's ID (from its URL)
  TELEGRAM_BOT_TOKEN           - from @BotFather
  TELEGRAM_CHAT_ID             - your chat/user/group ID
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, EMAIL_TO
                                - SMTP creds for outbound email (e.g. Gmail
                                  with an App Password, or any SMTP provider)

Any of the notification blocks (Telegram / email) are skipped gracefully if
their env vars aren't set, so you can enable just one if you prefer.
"""
import os
import re
import json
import smtplib
import datetime
from email.mime.text import MIMEText

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; CoffeeBriefingBot/1.0)"}

PRICE_LOG_HEADERS = [
    "Date", "Variety", "Grade", "Price Low (Rs/50kg)", "Price High (Rs/50kg)",
    "Change vs prior", "Source", "ICE Futures Price", "ICE Futures Unit",
    "AUD/INR", "USD/INR", "Est. AU Landed Cost (AUD/kg)",
    "AU Wholesale Quote - Sydney (manual)", "AU Wholesale Quote - Newcastle (manual)",
    "Bean Origin (manual)", "Notes", "Mid Price (Rs/50kg)", "Weeks Tracked",
    "Tracked High (Rs/50kg)", "Tracked Low (Rs/50kg)", "Range Status",
    "Price (USD/kg)", "Global Benchmark (USD/kg)",
    "3/6-Month Flag", "Flag Color",
    "Milled/Clean Equivalent (USD/kg)", "Outturn % Used",
]

ASSUMPTIONS_DEFAULTS = {
    "Freight (AUD/kg)": 0.55,
    "Customs & Clearance (% of value)": 0.05,
    "Warehouse (AUD/kg)": 0.15,
    "Delivery (AUD/kg)": 0.10,
    "Desired Margin (% of landed cost)": 0.12,
}

GRADES = [
    ("Arabica", "Parchment (AP)"),
    ("Arabica", "Cherry (AC)"),
    ("Robusta", "Parchment (RP)"),
    ("Robusta", "Cherry (RC)"),
]

# Milling outturn ratios: how much clean/green bean weight comes out of a kg
# of dried parchment or cherry. These are industry rule-of-thumb figures used
# by Indian curers/exporters and in Coffee Board of India crop-estimation
# methodology - NOT a lab-measured figure for any specific lot. Actual outturn
# varies (+/- ~5 points) with bean moisture, size, and processing quality.
# I could not pull a single authoritative numeric citation via live web fetch
# this session (several Indian coffee-trade/government pages returned empty),
# so treat these as reasonable industry estimates - ask your curer for their
# actual outturn percentage if you need a precise figure.
OUTTURN_PCT = {
    ("Arabica", "Parchment"): 0.80,   # ~1.25 kg parchment -> 1 kg clean
    ("Arabica", "Cherry"): 0.48,      # ~2.08 kg cherry -> 1 kg clean
    ("Robusta", "Parchment"): 0.82,   # ~1.22 kg parchment -> 1 kg clean
    ("Robusta", "Cherry"): 0.50,      # ~2.00 kg cherry -> 1 kg clean
}


def outturn_pct_for(variety, grade):
    stage = "Parchment" if "Parchment" in grade else "Cherry"
    return OUTTURN_PCT[(variety, stage)]


def milled_equivalent_usd_kg(unmilled_usd_kg, variety, grade):
    """Converts an unmilled (parchment/cherry weight-basis) USD/kg price into
    an estimated milled/clean-green-bean-equivalent USD/kg price, so it's
    comparable on the same weight basis as the global ICE benchmark (which is
    clean green bean, FOB)."""
    pct = outturn_pct_for(variety, grade)
    return unmilled_usd_kg / pct


def market_grade_name(variety, grade):
    """Indian coffee trade naming: washed (parchment) Arabica is sold as
    'Plantation' once milled - Plantation A/B/C, PB. Robusta's washed/milled
    form keeps the name 'Parchment' (no separate 'Plantation' naming for
    Robusta). Natural/dry-processed coffee - Arabica or Robusta - stays
    'Cherry' both before and after milling. Kirehalli only ever publishes the
    raw, unmilled farm-gate/curing-works price (confirmed by checking its
    live site directly - there's no free public feed for actual traded
    Plantation-grade prices), so this label is attached to our *estimated*
    milled-equivalent figure, not a live quote."""
    if variety == "Arabica" and "Parchment" in grade:
        return "Arabica Plantation (estimated milled)"
    if variety == "Arabica":
        return "Arabica Cherry (estimated milled)"
    if "Parchment" in grade:
        return "Robusta Parchment (estimated milled)"
    return "Robusta Cherry (estimated milled)"


# --------------------------------------------------------------------------
# Fetching (same sources/logic as the local version, regex bug already fixed)
# --------------------------------------------------------------------------

def http_get(url, timeout=15):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
        return resp.text
    except Exception as e:
        print(f"  [warn] fetch failed for {url}: {e}")
        return None


def fetch_kirehalli():
    """Find the latest Kirehalli 'Coffee Prices (Karnataka)' post and parse it."""
    html = http_get("https://kirehalli.com/coffee-prices-daily/")
    if not html:
        return {}
    soup = BeautifulSoup(html, "html.parser")
    link = None
    for a in soup.find_all("a", href=True):
        if "coffee-prices-karnataka" in a["href"]:
            link = a["href"]
            break
    if not link:
        print("  [warn] could not find a Kirehalli post link")
        return {}

    post_html = http_get(link)
    if not post_html:
        return {}
    text = BeautifulSoup(post_html, "html.parser").get_text("\n")

    result = {"source_url": link}

    def find_range(label):
        m = re.search(rf"{label}.*?Rs\s*([\d,]+)\s*[–-]\s*(?:Rs\s*)?([\d,]+)",
                      text, re.IGNORECASE | re.DOTALL)
        if not m:
            return None, None
        return int(m.group(1).replace(",", "")), int(m.group(2).replace(",", ""))

    def find_change(label):
        m = re.search(rf"{label}.*?(No Change|▲\s*\+?Rs\s*[\d,]+|▼\s*-?Rs\s*[\d,]+)",
                      text, re.IGNORECASE | re.DOTALL)
        return m.group(1).strip() if m else ""

    result["arabica_parchment"] = find_range("Arabica Parchment")
    result["arabica_parchment_chg"] = find_change("Arabica Parchment")
    result["arabica_cherry"] = find_range("Arabica Cherry")
    result["arabica_cherry_chg"] = find_change("Arabica Cherry")
    result["robusta_parchment"] = find_range("Robusta Parchment")
    result["robusta_parchment_chg"] = find_change("Robusta Parchment")
    result["robusta_cherry"] = find_range("Robusta Cherry")
    result["robusta_cherry_chg"] = find_change("Robusta Cherry")

    m = re.search(r"arabica coffee.*?([\d]+\.[\d]+)\s*(?:US\s*)?cents?/lb", text, re.IGNORECASE)
    result["ice_arabica_cents_lb"] = float(m.group(1)) if m else None

    m = re.search(r"robusta coffee.*?US\$?\s*([\d,]+)\s*/?\s*tonne", text, re.IGNORECASE)
    result["ice_robusta_usd_tonne"] = float(m.group(1).replace(",", "")) if m else None

    return result


def fetch_ice_arabica_benchmark():
    html = http_get("https://tradingeconomics.com/commodity/coffee")
    if not html:
        return None
    m = re.search(r"rose to ([\d.]+) USd/Lbs|fell to ([\d.]+) USd/Lbs|to ([\d.]+) USd/Lbs", html)
    if not m:
        return None
    return float(next(g for g in m.groups() if g))


def fetch_rate(from_ccy, to_ccy):
    html = http_get(f"https://www.xe.com/currencyconverter/convert/?Amount=1&From={from_ccy}&To={to_ccy}")
    if not html:
        return None
    m = re.search(rf"1\.00 {from_ccy} = ([\d.]+) {to_ccy}", html)
    return float(m.group(1)) if m else None


# --------------------------------------------------------------------------
# Google Sheets
# --------------------------------------------------------------------------

def get_gsheet_client():
    """Authenticate with a service account. GOOGLE_SERVICE_ACCOUNT_JSON holds
    the full JSON key content (paste the whole key file as the secret value)."""
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not set")
    info = json.loads(raw)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def get_or_create_worksheets(client):
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        raise RuntimeError("GOOGLE_SHEET_ID is not set")
    sh = client.open_by_key(sheet_id)

    try:
        ws_log = sh.worksheet("Price Log")
    except gspread.WorksheetNotFound:
        ws_log = sh.add_worksheet(title="Price Log", rows=2000, cols=len(PRICE_LOG_HEADERS))
        ws_log.append_row(PRICE_LOG_HEADERS)

    try:
        ws_assump = sh.worksheet("Assumptions")
    except gspread.WorksheetNotFound:
        ws_assump = sh.add_worksheet(title="Assumptions", rows=10, cols=3)
        ws_assump.append_row(["Item", "Value", "Note"])
        for label, val in ASSUMPTIONS_DEFAULTS.items():
            ws_assump.append_row([label, val, "Placeholder - update with your real number"])

    return ws_log, ws_assump


def read_assumptions(ws_assump):
    rows = ws_assump.get_all_records()
    values = {row["Item"]: float(row["Value"]) for row in rows if row.get("Item")}
    return {
        "freight": values.get("Freight (AUD/kg)", ASSUMPTIONS_DEFAULTS["Freight (AUD/kg)"]),
        "customs_pct": values.get("Customs & Clearance (% of value)", ASSUMPTIONS_DEFAULTS["Customs & Clearance (% of value)"]),
        "warehouse": values.get("Warehouse (AUD/kg)", ASSUMPTIONS_DEFAULTS["Warehouse (AUD/kg)"]),
        "delivery": values.get("Delivery (AUD/kg)", ASSUMPTIONS_DEFAULTS["Delivery (AUD/kg)"]),
        "margin_pct": values.get("Desired Margin (% of landed cost)", ASSUMPTIONS_DEFAULTS["Desired Margin (% of landed cost)"]),
    }


def read_price_history(ws_log):
    """Returns list of dicts for every existing row (empty list if sheet is new)."""
    records = ws_log.get_all_records()
    history = []
    for row in records:
        try:
            d = datetime.datetime.strptime(str(row["Date"]), "%Y-%m-%d").date()
        except (ValueError, KeyError):
            continue
        history.append({
            "date": d,
            "variety": row.get("Variety"),
            "grade": row.get("Grade"),
            "mid_price": row.get("Mid Price (Rs/50kg)"),
        })
    return history


# --------------------------------------------------------------------------
# Calculations (Python equivalents of the xlsx formulas from the local version)
# --------------------------------------------------------------------------

def landed_cost_aud_per_kg(base_price_aud_per_kg, assumptions):
    a = assumptions
    return (base_price_aud_per_kg * (1 + a["customs_pct"])
            + a["freight"] + a["warehouse"] + a["delivery"]
            + base_price_aud_per_kg * a["margin_pct"])


def india_landed_cost(mid_price_rs_per_50kg, aud_inr, assumptions, variety, grade):
    """Landed cost is computed off the MILLED-equivalent price, not the raw
    unmilled parchment/cherry price - what actually gets exported and cleared
    through customs is milled clean green bean (Plantation/Cherry grade), not
    raw parchment. Using the raw price here would understate true cost by
    skipping the milling step entirely."""
    unmilled_aud_per_kg = (mid_price_rs_per_50kg / 50) / aud_inr
    milled_aud_per_kg = unmilled_aud_per_kg / outturn_pct_for(variety, grade)
    return landed_cost_aud_per_kg(milled_aud_per_kg, assumptions)


def global_benchmark_usd_per_kg(variety, ice_price, ice_unit_is_cents_lb):
    if ice_price is None:
        return None
    if ice_unit_is_cents_lb:
        return ice_price / 100 / 0.45359237
    return ice_price / 1000


def brazil_png_landed_cost_aud_per_kg(ice_arabica_cents_lb, usd_inr, aud_inr, assumptions):
    if ice_arabica_cents_lb is None or not usd_inr or not aud_inr:
        return None
    usd_per_kg = ice_arabica_cents_lb / 100 / 0.45359237
    usd_to_aud = usd_inr / aud_inr
    base_aud_per_kg = usd_per_kg * usd_to_aud
    return landed_cost_aud_per_kg(base_aud_per_kg, assumptions)


def tracked_stats(history, variety, grade, today, today_mid_price):
    """Mirrors the xlsx Tracked High/Low/Range Status/Weeks Tracked formulas."""
    same = [h for h in history if h["variety"] == variety and h["grade"] == grade and h["mid_price"] is not None]
    if not same:
        weeks = 0.0
        tracked_high = tracked_low = today_mid_price
    else:
        first_date = min(h["date"] for h in same)
        weeks = (today - first_date).days / 7
        prices = [h["mid_price"] for h in same] + [today_mid_price]
        tracked_high = max(prices)
        tracked_low = min(prices)

    if weeks < 1:
        status = "Day 1 - building history"
    elif today_mid_price >= tracked_high:
        status = f"{round(weeks)}-wk high (since tracking began)"
    elif today_mid_price <= tracked_low:
        status = f"{round(weeks)}-wk low (since tracking began)"
    else:
        status = f"{round(weeks)}-wk range, mid-range"

    return weeks, tracked_high, tracked_low, status


def window_flag(history, variety, grade, today, today_mid_price):
    """Flags when today's price is at (or beyond) the 3-month or 6-month
    high/low, computed from the sheet's own accumulated history.

    Green  = at a 3- or 6-month LOW  -> historically cheap, safer to buy.
    Red    = at a 3- or 6-month HIGH -> historically expensive, riskier to buy.
    A 6-month extreme implies the 3-month one too (nested window), so the
    longer window "wins" for the label.

    Honesty note: this only means something once real history has
    accumulated. With less than ~90/180 days of logged data, "N/A - insufficient
    history yet" is returned rather than a fabricated flag.
    """
    same = [h for h in history if h["variety"] == variety and h["grade"] == grade and h["mid_price"] is not None]
    if not same:
        return "N/A - insufficient history yet", "none"

    oldest = min(h["date"] for h in same)
    days_of_history = (today - oldest).days

    def prices_within(days):
        cutoff = today - datetime.timedelta(days=days)
        vals = [h["mid_price"] for h in same if h["date"] >= cutoff]
        vals.append(today_mid_price)
        return vals

    six_mo = prices_within(180) if days_of_history >= 180 else None
    three_mo = prices_within(90) if days_of_history >= 90 else None

    if six_mo is not None:
        if today_mid_price <= min(six_mo):
            return "6-MONTH LOW - historically cheap, safer to buy", "green"
        if today_mid_price >= max(six_mo):
            return "6-MONTH HIGH - historically expensive, riskier to buy", "red"
    if three_mo is not None:
        if today_mid_price <= min(three_mo):
            return "3-MONTH LOW - historically cheap, safer to buy", "green"
        if today_mid_price >= max(three_mo):
            return "3-MONTH HIGH - historically expensive, riskier to buy", "red"

    if three_mo is None:
        return f"N/A - only {days_of_history} day(s) of history, need 90+ for a 3-month read", "none"
    return "within 3/6-month range", "none"


def build_signal(variety, grade_statuses):
    """grade_statuses: dict of grade -> range status string, for this variety."""
    at_high = any("high" in s for s in grade_statuses.values())
    at_low = any("low" in s and "high" not in s for s in grade_statuses.values())
    month = datetime.date.today().month
    if 6 <= month <= 8:
        season_note = "Brazil frost-risk season (Jun-Aug) - Arabica historically volatile/rising" if variety == "Arabica" else "Brazil frost-risk season - less directly relevant to Robusta"
    elif month in (9, 10):
        season_note = "Brazil flowering period - crop-formation risk"
    elif month in (11, 12, 1):
        season_note = "Vietnam Robusta harvest window - potential Robusta softness" if variety == "Robusta" else "Vietnam Robusta harvest window - less directly relevant to Arabica"
    elif 1 <= month <= 3:
        season_note = "Karnataka's typical seasonal low window (Jan-Mar harvest)"
    else:
        season_note = "no major seasonal trigger this month"

    grades_desc = ", ".join(f"{g}: {s}" for g, s in grade_statuses.items())
    if at_high and not at_low:
        headline = "WAIT"
    elif at_low and not at_high:
        headline = "BUY"
    else:
        headline = "MIXED"
    reasoning = f"{grades_desc}. {season_note}. Directional guidance only, not a trading signal."
    return headline, reasoning


# --------------------------------------------------------------------------
# Notifications
# --------------------------------------------------------------------------

GREEN_BG = {"red": 0.71, "green": 0.88, "blue": 0.71}
RED_BG = {"red": 0.96, "green": 0.71, "blue": 0.71}
FLAG_COL_LETTER = gspread.utils.rowcol_to_a1(1, PRICE_LOG_HEADERS.index("3/6-Month Flag") + 1)[:-1]


def color_flag_cells(ws_log, first_new_row, new_rows):
    """Colors the '3/6-Month Flag' cell green (3/6-month low - safer to buy)
    or red (3/6-month high - riskier) for each just-appended row, based on
    the 'Flag Color' value computed alongside it. Leaves the cell unstyled
    ("none") when there isn't enough history yet or price is mid-range."""
    color_idx = PRICE_LOG_HEADERS.index("Flag Color")
    for i, row in enumerate(new_rows):
        color = row[color_idx]
        if color not in ("green", "red"):
            continue
        cell = f"{FLAG_COL_LETTER}{first_new_row + i}"
        bg = GREEN_BG if color == "green" else RED_BG
        try:
            ws_log.format(cell, {"backgroundColor": bg})
        except Exception as e:
            print(f"  [warn] could not color-format {cell}: {e}")


# --------------------------------------------------------------------------
# Static HTML dashboard (published to GitHub Pages by the workflow, so the
# same look as the Cowork artifact is viewable from any device/browser -
# no Chrome extension, no Cowork session required).
# --------------------------------------------------------------------------

GRADE_DISPLAY = {
    ("Arabica", "Parchment (AP)"): ("Arabica parchment (AP)", "Plantation", "~80%"),
    ("Arabica", "Cherry (AC)"): ("Arabica cherry (AC)", "Cherry", "~48%"),
    ("Robusta", "Parchment (RP)"): ("Robusta parchment (RP)", "Parchment", "~82%"),
    ("Robusta", "Cherry (RC)"): ("Robusta cherry (RC)", "Cherry", "~50%"),
}


def _change_style(chg):
    if not chg:
        return "#6b6b68", "&mdash;"
    low = chg.lower()
    if "▲" in chg or "up" in low:
        return "#3B6D11", chg
    if "▼" in chg or "down" in low:
        return "#791F1F", chg
    return "#6b6b68", chg


def _status_style(status):
    low = (status or "").lower()
    if "high" in low:
        return "#791F1F"
    if "low" in low:
        return "#3B6D11"
    return "#888780"


def _flag_pill(flag_text, flag_color):
    if flag_color == "green":
        label = flag_text.split(" -")[0] if flag_text else "LOW"
        return f'<span style="background: #DCEFDA; color: #205C1F; font-size: 11px; padding: 2px 8px; border-radius: 10px;">{label}</span>'
    if flag_color == "red":
        label = flag_text.split(" -")[0] if flag_text else "HIGH"
        return f'<span style="background: #F7DCDC; color: #791F1F; font-size: 11px; padding: 2px 8px; border-radius: 10px;">{label}</span>'
    label = "building history" if "insufficient" in (flag_text or "").lower() else "within range"
    return f'<span style="background: #ECE9E0; color: #6b6b68; font-size: 11px; padding: 2px 8px; border-radius: 10px;">{label}</span>'


def _seasonality_overview(month):
    if 6 <= month <= 8:
        return ("We're inside Brazil's Jun&ndash;Aug frost-risk window (historically pushes Arabica higher "
                "within 1&ndash;2 months), and Karnataka's own seasonal low is typically Jan&ndash;Mar during peak "
                "harvest &mdash; several months away.")
    if month in (9, 10):
        return "Brazil is in its flowering period &mdash; a crop-formation risk window that can move Arabica."
    if month in (11, 12, 1):
        return "Vietnam's Robusta harvest (Nov&ndash;Jan) is underway, which can soften Robusta prices globally."
    if 1 <= month <= 3:
        return "Karnataka is in its typical seasonal low window (Jan&ndash;Mar harvest) &mdash; local prices are often at their cheapest."
    return "No major seasonal trigger this month for either variety."


def render_dashboard_html(today, grade_data, grade_changes, range_statuses, flags_by_grade,
                           milled_by_grade, landed_by_grade, usd_kg_by_grade, ice_arabica,
                           ice_robusta, aud_inr, usd_inr, bench_usd_kg_arabica,
                           bench_usd_kg_robusta, arabica_signal, arabica_reason,
                           robusta_signal, robusta_reason, brazil_png_landed, history):
    """Renders the same visual layout as the Cowork 'Coffee Buying Dashboard'
    artifact as a single static HTML file, so it can be published to GitHub
    Pages and viewed from any device without Chrome or Cowork."""

    if history:
        oldest = min(h["date"] for h in history)
        days_tracked = (today - oldest).days
        history_note = (
            f"Tracking real Kirehalli history since {oldest.isoformat()} ({days_tracked} days logged so far). "
            "Range Status and the 3/6-month flag are computed off this genuine accumulated history, not placeholders "
            "- both become fully meaningful once 90/180+ days exist."
        )
    else:
        history_note = "No prior history yet - this is the first logged day, so Range Status will read \"Day 1\"."

    local_rows = ""
    for variety, grade in GRADES:
        display_name, milled_name, outturn_pct = GRADE_DISPLAY[(variety, grade)]
        price_range = grade_data.get((variety, grade))
        chg = grade_changes.get((variety, grade), "")
        status = range_statuses.get(variety, {}).get(grade, "N/A")
        flag_text, flag_color = flags_by_grade.get((variety, grade), ("N/A", "none"))
        chg_color, chg_label = _change_style(chg)
        status_color = _status_style(status)
        if price_range:
            low, high = price_range
            outturn = float(outturn_pct.strip("~%")) / 100
            milled_low, milled_high = round(low / outturn), round(high / outturn)
            range_txt = f"{low:,} &ndash; {high:,}"
            milled_txt = f"{milled_low:,}&ndash;{milled_high:,}<br><span style=\"color: #888780;\">= {milled_name}, {outturn_pct} outturn</span>"
        else:
            range_txt = '<span style="color: #888780;">data unavailable today</span>'
            milled_txt = '<span style="color: #888780;">N/A</span>'
        local_rows += f"""
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">{display_name}</td>
      <td style="padding: 8px 10px; text-align: right;">{range_txt}</td>
      <td style="padding: 8px 10px; text-align: right; color: {chg_color};">{chg_label}</td>
      <td style="padding: 8px 10px; text-align: right; color: {status_color}; font-size: 12px;">{status}</td>
      <td style="padding: 8px 10px; text-align: right;">{_flag_pill(flag_text, flag_color)}</td>
      <td style="padding: 8px 10px; text-align: right; font-size: 12px;">{milled_txt}</td>
    </tr>"""

    def fmt(v, prefix="$", suffix="", nd=2):
        return f"{prefix}{v:,.{nd}f}{suffix}" if v is not None else "N/A"

    bench_rows = ""
    for variety, grade in GRADES:
        name = market_grade_name(variety, grade)
        unmilled = usd_kg_by_grade.get((variety, grade))
        milled = milled_by_grade.get((variety, grade))
        bench_rows += f"""
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">India {name.replace(' (estimated milled)', '')} (Karnataka, est. milled)</td>
      <td style="padding: 8px 10px; text-align: right; color: #888780;">{fmt(unmilled)}</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(milled)}</td>
    </tr>"""
        if variety == "Arabica" and grade == "Parchment (AP)":
            bench_rows += f"""
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">Global Arabica benchmark (ICE NY, Brazil-dominant)</td>
      <td style="padding: 8px 10px; text-align: right; color: #c9c7bd;">&mdash;</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(bench_usd_kg_arabica)}</td>
    </tr>
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">PNG Arabica (same global benchmark)</td>
      <td style="padding: 8px 10px; text-align: right; color: #c9c7bd;">&mdash;</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(bench_usd_kg_arabica)}</td>
    </tr>"""
    bench_rows += f"""
    <tr>
      <td style="padding: 8px 10px;">Global Robusta benchmark (ICE London)</td>
      <td style="padding: 8px 10px; text-align: right; color: #c9c7bd;">&mdash;</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(bench_usd_kg_robusta)}</td>
    </tr>"""

    landed_rows = ""
    for variety, grade in GRADES:
        name = market_grade_name(variety, grade).replace(" (estimated milled)", " (est. milled)")
        landed_rows += f"""
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">{name}</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(landed_by_grade.get((variety, grade)))}</td>
    </tr>"""

    fetch_ok = any(grade_data.get((v, g)) for v, g in GRADES) and aud_inr and usd_inr
    banner_class = "" if fetch_ok else " (some sources failed today - see N/A fields below)"

    html = f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Coffee Buying Dashboard</title></head>
<body>
<div style="color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 720px; margin: 0 auto; padding: 24px 16px; color: #1a1a1a;">

  <div style="margin-bottom: 1.25rem;">
    <div style="font-size: 13px; color: #6b6b68; margin-bottom: 2px;">Coffee buying dashboard &middot; last updated</div>
    <div style="font-size: 20px; font-weight: 500;">{today.isoformat()}{banner_class}</div>
  </div>

  <div style="font-size: 12px; color: #888780; background: #f1efe8; border-radius: 8px; padding: 8px 12px; margin-bottom: 1rem;">
    {history_note} Published automatically by GitHub Actions &mdash; no browser extension or Cowork session needed to view this.
  </div>

  <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 1.25rem;">
    <div style="background: #FAEEDA; border: 0.5px solid #EF9F27; border-radius: 12px; padding: 1rem 1.1rem;">
      <div style="font-size: 13px; color: #854F0B;">Arabica signal</div>
      <div style="font-size: 16px; font-weight: 500; color: #633806; margin-top: 4px;">{arabica_signal}</div>
      <div style="font-size: 13px; color: #633806; margin-top: 6px; line-height: 1.5;">{arabica_reason}</div>
    </div>
    <div style="background: #FAEEDA; border: 0.5px solid #EF9F27; border-radius: 12px; padding: 1rem 1.1rem;">
      <div style="font-size: 13px; color: #854F0B;">Robusta signal</div>
      <div style="font-size: 16px; font-weight: 500; color: #633806; margin-top: 4px;">{robusta_signal}</div>
      <div style="font-size: 13px; color: #633806; margin-top: 6px; line-height: 1.5;">{robusta_reason}</div>
    </div>
  </div>

  <div style="font-size: 13px; color: #6b6b68; margin-bottom: 1rem; line-height: 1.5;">
    Overall: {_seasonality_overview(today.month)}
  </div>

  <div style="font-size: 16px; font-weight: 500; margin: 1.5rem 0 0.75rem;">Karnataka local prices (&#8377; per 50kg, raw curing-works price &mdash; unmilled)</div>
  <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
    <tr style="background: #f1efe8;">
      <td style="padding: 8px 10px; font-weight: 500;">Variety / grade (raw, unmilled)</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Range</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Change</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Range status</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">3/6-mo flag</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Milled-equiv. (est.)</td>
    </tr>{local_rows}
  </table>
  <div style="font-size: 12px; color: #888780; margin-top: 6px;">Source: Kirehalli daily report, Sakleshpur / Chikmagalur / Hassan region. "Milled-equiv." divides the raw range by the industry outturn ratio to estimate the milled grade's price; this is an estimate, not a live Plantation-grade quote (none exists publicly).</div>

  <div style="font-size: 16px; font-weight: 500; margin: 1.5rem 0 0.75rem;">Global benchmarks</div>
  <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px;">
    <div style="background: #f1efe8; border-radius: 8px; padding: 1rem;">
      <div style="font-size: 12px; color: #6b6b68;">ICE Arabica (NY)</div>
      <div style="font-size: 20px; font-weight: 500; margin-top: 4px;">{ice_arabica if ice_arabica else 'N/A'}&cent;/lb</div>
    </div>
    <div style="background: #f1efe8; border-radius: 8px; padding: 1rem;">
      <div style="font-size: 12px; color: #6b6b68;">ICE Robusta (London)</div>
      <div style="font-size: 20px; font-weight: 500; margin-top: 4px;">${f'{ice_robusta:,.0f}' if ice_robusta else 'N/A'}/t</div>
    </div>
    <div style="background: #f1efe8; border-radius: 8px; padding: 1rem;">
      <div style="font-size: 12px; color: #6b6b68;">AUD/INR</div>
      <div style="font-size: 20px; font-weight: 500; margin-top: 4px;">{aud_inr if aud_inr else 'N/A'}</div>
    </div>
    <div style="background: #f1efe8; border-radius: 8px; padding: 1rem;">
      <div style="font-size: 12px; color: #6b6b68;">USD/INR</div>
      <div style="font-size: 20px; font-weight: 500; margin-top: 4px;">{usd_inr if usd_inr else 'N/A'}</div>
    </div>
  </div>

  <div style="font-size: 16px; font-weight: 500; margin: 1.5rem 0 0.75rem;">India vs global benchmark (USD/kg, milled/clean-bean equivalent)</div>
  <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
    <tr style="background: #f1efe8;">
      <td style="padding: 8px 10px; font-weight: 500;">Origin / grade</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Unmilled USD/kg</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Milled-equivalent USD/kg</td>
    </tr>{bench_rows}
  </table>
  <div style="font-size: 12px; color: #888780; margin-top: 6px; line-height: 1.5;">
    No free live feed publishes actual traded Plantation-grade prices, so the "milled-equivalent" column is an estimate: raw unmilled USD/kg divided by an industry outturn ratio. Brazil/PNG share the ICE NY Arabica benchmark since no origin-specific spot price feed is scraped here.
  </div>

  <div style="font-size: 16px; font-weight: 500; margin: 1.5rem 0 0.75rem;">Estimated AU landed cost (calculated on milled-equivalent basis, not a quote)</div>
  <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
    <tr style="background: #f1efe8;">
      <td style="padding: 8px 10px; font-weight: 500;">Grade</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Est. AUD/kg</td>
    </tr>{landed_rows}
  </table>
  <div style="font-size: 12px; color: #888780; margin-top: 6px; line-height: 1.5;">
    Landed cost is based on the milled-equivalent price (Plantation/Cherry-grade clean bean), not the raw parchment/cherry farm-gate price, since that's what actually clears Australian customs. Uses the freight/customs/warehouse/delivery/margin assumptions from the Assumptions tab &mdash; update with your real numbers for accuracy.
  </div>

  <div style="font-size: 16px; font-weight: 500; margin: 1.5rem 0 0.75rem;">Estimated AU landed cost &mdash; Brazil &amp; PNG (calculated, not a quote)</div>
  <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
    <tr style="background: #f1efe8;">
      <td style="padding: 8px 10px; font-weight: 500;">Grade</td>
      <td style="padding: 8px 10px; font-weight: 500; text-align: right;">Est. AUD/kg</td>
    </tr>
    <tr style="border-bottom: 0.5px solid #e3e1d8;">
      <td style="padding: 8px 10px;">Brazil Arabica</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(brazil_png_landed)}</td>
    </tr>
    <tr>
      <td style="padding: 8px 10px;">PNG Arabica</td>
      <td style="padding: 8px 10px; text-align: right;">{fmt(brazil_png_landed)}</td>
    </tr>
  </table>
  <div style="font-size: 12px; color: #888780; margin-top: 6px; line-height: 1.5;">
    Same method as the India table above: ICE NY Arabica benchmark converted to AUD, run through the Assumptions-tab freight/customs/warehouse/delivery/margin. Both rows use the same benchmark since there's no live scraped Brazil-specific or PNG-specific spot price feed.
  </div>

  <div style="font-size: 11px; color: #c9c7bd; margin-top: 2rem; text-align: center;">Generated by coffee_briefing.py via GitHub Actions &middot; {today.isoformat()}</div>

</div>
</body></html>
"""
    return html


def send_telegram(text):
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("  [skip] Telegram not configured (TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID missing)")
        return
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    try:
        resp = requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=15)
        resp.raise_for_status()
        print("  Telegram message sent.")
    except Exception as e:
        print(f"  [warn] Telegram send failed: {e}")


def send_email(subject, body):
    host = os.getenv("SMTP_HOST")
    port = os.getenv("SMTP_PORT")
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    to_addr = os.getenv("EMAIL_TO")
    if not all([host, port, user, password, to_addr]):
        print("  [skip] Email not configured (SMTP_HOST/PORT/USER/PASSWORD/EMAIL_TO missing)")
        return
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = to_addr
    try:
        with smtplib.SMTP_SSL(host, int(port)) as server:
            server.login(user, password)
            server.sendmail(user, [to_addr], msg.as_string())
        print("  Email sent.")
    except Exception as e:
        print(f"  [warn] Email send failed: {e}")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    today = datetime.date.today()
    print(f"=== Coffee briefing for {today.isoformat()} ===")

    print("Fetching Karnataka local prices (Kirehalli)...")
    kirehalli = fetch_kirehalli()

    print("Fetching ICE Arabica benchmark (Trading Economics)...")
    ice_arabica_benchmark = fetch_ice_arabica_benchmark()

    print("Fetching AUD/INR and USD/INR (xe.com)...")
    aud_inr = fetch_rate("AUD", "INR")
    usd_inr = fetch_rate("USD", "INR")

    ice_arabica = kirehalli.get("ice_arabica_cents_lb") or ice_arabica_benchmark
    ice_robusta = kirehalli.get("ice_robusta_usd_tonne")

    print("Connecting to Google Sheets...")
    client = get_gsheet_client()
    ws_log, ws_assump = get_or_create_worksheets(client)
    assumptions = read_assumptions(ws_assump)
    history = read_price_history(ws_log)

    grade_data = {
        ("Arabica", "Parchment (AP)"): kirehalli.get("arabica_parchment"),
        ("Arabica", "Cherry (AC)"): kirehalli.get("arabica_cherry"),
        ("Robusta", "Parchment (RP)"): kirehalli.get("robusta_parchment"),
        ("Robusta", "Cherry (RC)"): kirehalli.get("robusta_cherry"),
    }
    grade_changes = {
        ("Arabica", "Parchment (AP)"): kirehalli.get("arabica_parchment_chg", ""),
        ("Arabica", "Cherry (AC)"): kirehalli.get("arabica_cherry_chg", ""),
        ("Robusta", "Parchment (RP)"): kirehalli.get("robusta_parchment_chg", ""),
        ("Robusta", "Cherry (RC)"): kirehalli.get("robusta_cherry_chg", ""),
    }

    new_rows = []
    range_statuses = {"Arabica": {}, "Robusta": {}}
    usd_kg_by_grade = {}
    flags_by_grade = {}
    milled_by_grade = {}
    landed_by_grade = {}

    for variety, grade in GRADES:
        price_range = grade_data[(variety, grade)]
        low, high = price_range if price_range else (None, None)
        ice_price = ice_arabica if variety == "Arabica" else ice_robusta
        ice_unit = "US cents/lb (ICE Arabica)" if variety == "Arabica" else "USD/tonne (ICE Robusta)"

        if low is None or aud_inr is None:
            new_rows.append([
                today.isoformat(), variety, grade, "N/A", "N/A", "", "Kirehalli (fetch failed)",
                ice_price or "N/A", ice_unit, aud_inr or "N/A", usd_inr or "N/A",
                "N/A", "", "", "", "Fetch failed - fill in manually",
                "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A",
                "N/A", "none", "N/A", "N/A",
            ])
            continue

        mid = (low + high) / 2
        weeks, tr_high, tr_low, status = tracked_stats(history, variety, grade, today, mid)
        landed = india_landed_cost(mid, aud_inr, assumptions, variety, grade)
        usd_kg = (mid / 50) / usd_inr
        bench_usd_kg = global_benchmark_usd_per_kg(variety, ice_price, variety == "Arabica")
        flag_text, flag_color = window_flag(history, variety, grade, today, mid)
        outturn_pct = outturn_pct_for(variety, grade)
        milled_usd_kg = milled_equivalent_usd_kg(usd_kg, variety, grade)

        range_statuses[variety][grade] = status
        usd_kg_by_grade[(variety, grade)] = usd_kg
        flags_by_grade[(variety, grade)] = (flag_text, flag_color)
        milled_by_grade[(variety, grade)] = milled_usd_kg
        landed_by_grade[(variety, grade)] = landed

        new_rows.append([
            today.isoformat(), variety, grade, low, high, grade_changes[(variety, grade)],
            "Kirehalli (Sakleshpur/Chikmagalur)", ice_price, ice_unit, aud_inr, usd_inr,
            round(landed, 2), "", "", "", "",
            mid, round(weeks, 1), tr_high, tr_low, status,
            round(usd_kg, 2), round(bench_usd_kg, 2) if bench_usd_kg else "N/A",
            flag_text, flag_color,
            round(milled_usd_kg, 2), f"{outturn_pct:.0%}",
        ])

    first_new_row = len(ws_log.get_all_values()) + 1
    ws_log.append_rows(new_rows, value_input_option="USER_ENTERED")
    print(f"Appended {len(new_rows)} rows to the Google Sheet.")

    color_flag_cells(ws_log, first_new_row, new_rows)

    arabica_signal, arabica_reason = build_signal("Arabica", range_statuses["Arabica"])
    robusta_signal, robusta_reason = build_signal("Robusta", range_statuses["Robusta"])

    brazil_png_landed = brazil_png_landed_cost_aud_per_kg(ice_arabica, usd_inr, aud_inr, assumptions)
    bench_usd_kg_arabica = global_benchmark_usd_per_kg("Arabica", ice_arabica, True)
    bench_usd_kg_robusta = global_benchmark_usd_per_kg("Robusta", ice_robusta, False)

    lines = [
        f"Coffee briefing - {today.isoformat()}",
        "",
        "Karnataka local prices (Rs/50kg, raw curing-works price - unmilled):",
    ]
    for variety, grade in GRADES:
        low, high = grade_data[(variety, grade)] or (None, None)
        chg = grade_changes[(variety, grade)]
        if low is not None:
            lines.append(f"  {variety} {grade}: {low:,}-{high:,} ({chg or 'n/a'})")
        else:
            lines.append(f"  {variety} {grade}: data unavailable today")
    lines += [
        "",
        f"ICE Arabica (NY): {ice_arabica if ice_arabica else 'N/A'} US cents/lb"
        f" (~${bench_usd_kg_arabica:.2f}/kg)" if bench_usd_kg_arabica else "",
        f"ICE Robusta (London): {ice_robusta if ice_robusta else 'N/A'} USD/tonne"
        f" (~${bench_usd_kg_robusta:.2f}/kg)" if bench_usd_kg_robusta else "",
        f"AUD/INR: {aud_inr if aud_inr else 'N/A'}   USD/INR: {usd_inr if usd_inr else 'N/A'}",
        "",
        f"Arabica signal: {arabica_signal} - {arabica_reason}",
        f"Robusta signal: {robusta_signal} - {robusta_reason}",
        "",
        "3/6-month price flags (green=historically cheap/safer, red=historically expensive/riskier):",
    ]
    any_flag = False
    for variety, grade in GRADES:
        flag_text, flag_color = flags_by_grade.get((variety, grade), ("N/A", "none"))
        if flag_color == "green":
            lines.append(f"  [LOW]  {variety} {grade}: {flag_text}")
            any_flag = True
        elif flag_color == "red":
            lines.append(f"  [HIGH] {variety} {grade}: {flag_text}")
            any_flag = True
    if not any_flag:
        lines.append("  No grade is currently at a 3- or 6-month extreme.")

    lines += [
        "",
        "Market grade - estimated milled/clean-bean-equivalent USD/kg (this is what's actually "
        "comparable to the global benchmark; no free live feed publishes real Plantation-grade "
        "quotes, so this is Parchment/Cherry converted via industry outturn ratios, not a "
        "lab-measured figure):",
    ]
    for variety, grade in GRADES:
        milled = milled_by_grade.get((variety, grade))
        if milled is None:
            continue
        bench = bench_usd_kg_arabica if variety == "Arabica" else bench_usd_kg_robusta
        name = market_grade_name(variety, grade)
        if bench:
            lines.append(f"  {name}: ${milled:.2f}/kg (vs global {variety} benchmark ${bench:.2f}/kg)")
        else:
            lines.append(f"  {name}: ${milled:.2f}/kg")

    lines += [
        "",
        "Estimated AU landed cost (AUD/kg, based on milled-equivalent price - not raw parchment/cherry):",
    ]
    for variety, grade in GRADES:
        landed_val = landed_by_grade.get((variety, grade))
        if landed_val is not None:
            lines.append(f"  {market_grade_name(variety, grade)}: ${landed_val:.2f}/kg")

    lines += [
        "",
        f"Brazil/PNG landed cost estimate (AUD/kg): {round(brazil_png_landed, 2) if brazil_png_landed else 'N/A'}"
        " (same ICE NY benchmark for both - no live origin-specific spot price feed)",
        "",
        "Reminder: Sydney/Newcastle wholesale rates aren't published anywhere - log real supplier"
        " quotes in the Google Sheet's manual columns.",
    ]
    message = "\n".join(l for l in lines if l is not None)
    print("\n" + message)
    # No Telegram/email - dashboard only, per preference. send_telegram()/
    # send_email() are still defined above and still no-op safely if their
    # secrets aren't set, but main() deliberately never calls them.

    # Render the same visual dashboard as the Cowork artifact to a static
    # HTML file and commit it to docs/ - the workflow publishes docs/ to
    # GitHub Pages so it's viewable from any device/browser, no extension
    # or Cowork session required.
    print("Rendering static dashboard HTML...")
    dashboard_html = render_dashboard_html(
        today, grade_data, grade_changes, range_statuses, flags_by_grade,
        milled_by_grade, landed_by_grade, usd_kg_by_grade, ice_arabica,
        ice_robusta, aud_inr, usd_inr, bench_usd_kg_arabica,
        bench_usd_kg_robusta, arabica_signal, arabica_reason,
        robusta_signal, robusta_reason, brazil_png_landed, history,
    )
    os.makedirs("docs", exist_ok=True)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(dashboard_html)
    print("Wrote docs/index.html")

    print("\nDone.")


if __name__ == "__main__":
    main()
