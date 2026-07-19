"""
Builds Coffee_Price_Log.xlsx from scratch (Assumptions + Price Log + Read Me tabs).
Run once to (re)create the workbook. Requires: pip install openpyxl

Usage:
    python3 build_log_local.py
Creates Coffee_Price_Log.xlsx in the current directory.
"""
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "Coffee_Price_Log.xlsx"

wb = openpyxl.Workbook()

# ---------- Sheet 1: Assumptions ----------
ws_a = wb.active
ws_a.title = "Assumptions"
ARIAL = "Arial"
blue = Font(name=ARIAL, color="0000FF")
black = Font(name=ARIAL, color="000000")
header_fill = PatternFill("solid", fgColor="1F4E78")
header_font = Font(name=ARIAL, bold=True, color="FFFFFF")
yellow_fill = PatternFill("solid", fgColor="FFFF00")
note_font = Font(name=ARIAL, italic=True, size=9, color="666666")

ws_a["A1"] = "Landed Cost Assumptions (India -> Australia)"
ws_a["A1"].font = Font(name=ARIAL, bold=True, size=13)
ws_a["A2"] = "Edit the yellow cells below with your real freight/customs/margin numbers."
ws_a["A2"].font = note_font
ws_a.merge_cells("A2:D2")

rows = [
    ("Freight (AUD/kg)", 0.55, "Placeholder - update with your actual freight quote"),
    ("Customs & Clearance (% of value)", 0.05, "Placeholder - update with your actual customs/duty rate"),
    ("Warehouse (AUD/kg)", 0.15, "Placeholder - update with your actual warehousing cost"),
    ("Delivery (AUD/kg)", 0.10, "Placeholder - update with your local delivery cost"),
    ("Desired Margin (% of landed cost)", 0.12, "Placeholder - update with your target margin"),
]
r = 4
ws_a[f"A{r}"] = "Item"; ws_a[f"B{r}"] = "Value"; ws_a[f"C{r}"] = "Note"
for c in ("A", "B", "C"):
    ws_a[f"{c}{r}"].font = header_font
    ws_a[f"{c}{r}"].fill = header_fill
r += 1
assump_rows = {}
for label, val, note in rows:
    ws_a[f"A{r}"] = label
    ws_a[f"A{r}"].font = black
    ws_a[f"B{r}"] = val
    ws_a[f"B{r}"].font = blue
    ws_a[f"B{r}"].fill = yellow_fill
    ws_a[f"C{r}"] = note
    ws_a[f"C{r}"].font = note_font
    assump_rows[label] = r
    r += 1

ws_a.column_dimensions["A"].width = 34
ws_a.column_dimensions["B"].width = 12
ws_a.column_dimensions["C"].width = 55

FREIGHT_ROW = assump_rows["Freight (AUD/kg)"]
CUSTOMS_ROW = assump_rows["Customs & Clearance (% of value)"]
WAREHOUSE_ROW = assump_rows["Warehouse (AUD/kg)"]
DELIVERY_ROW = assump_rows["Delivery (AUD/kg)"]
MARGIN_ROW = assump_rows["Desired Margin (% of landed cost)"]
# Fixed rows if you rebuild: Freight=5, Customs=6, Warehouse=7, Delivery=8, Margin=9

# ---------- Sheet 2: Price Log ----------
ws = wb.create_sheet("Price Log")
headers = [
    "Date", "Variety", "Grade/Type", "Karnataka Price Low (Rs/50kg)", "Karnataka Price High (Rs/50kg)",
    "Change vs prior", "Source", "ICE Futures Price", "ICE Futures Unit", "AUD/INR", "USD/INR",
    "Est. AU Landed Cost (AUD/kg, milled-equivalent basis)", "AU Wholesale Quote - Sydney (AUD/kg, manual)",
    "AU Wholesale Quote - Newcastle (AUD/kg, manual)", "Bean Origin (manual, e.g. Indian/Brazilian)", "Notes",
    "Mid Price (Rs/50kg)", "Weeks Tracked", "Tracked High (Rs/50kg)", "Tracked Low (Rs/50kg)", "Range Status",
    "Price (USD/kg)", "Global Benchmark (USD/kg)", "3/6-Month Flag",
    "Milled/Clean Equivalent (USD/kg)", "Outturn % Used", "Market Grade Name (Indian trade terminology)"
]
for i, h in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=i, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(wrap_text=True, vertical="center")

widths = [11, 10, 16, 14, 14, 12, 12, 14, 13, 9, 9, 24, 18, 20, 22, 26, 12, 10, 13, 13, 26, 14, 18, 30, 26, 14, 32]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Seed row (today's real figures as of 2026-07-16/17 — replace when you run this yourself)
data_rows = [
    (datetime.date(2026,7,16), "Arabica", "Parchment (AP)", 24400, 25000, "No change", "Kirehalli (Sakleshpur/Chikmagalur)", 312.60, "US cents/lb (Sep26 ICE Arabica)", 67.4743, 96.6461, "Global Arabica eased to a 1-week low, off last week's 5.5-month high"),
    (datetime.date(2026,7,16), "Arabica", "Cherry (AC)", 13800, 15400, "No change", "Kirehalli (Sakleshpur/Chikmagalur)", 312.60, "US cents/lb (Sep26 ICE Arabica)", 67.4743, 96.6461, ""),
    (datetime.date(2026,7,16), "Robusta", "Parchment (RP)", 18000, 18500, "+Rs 200", "Kirehalli (Sakleshpur/Chikmagalur)", 3797, "USD/tonne (Sep26 ICE Robusta)", 67.4743, 96.6461, "Local Robusta rising even as ICE Robusta eased slightly"),
    (datetime.date(2026,7,16), "Robusta", "Cherry (RC)", 10200, 11000, "+Rs 100", "Kirehalli (Sakleshpur/Chikmagalur)", 3797, "USD/tonne (Sep26 ICE Robusta)", 67.4743, 96.6461, ""),
]

r = 2
for date, variety, grade, low, high, change, source, ice_price, ice_unit, aud_inr, usd_inr, note in data_rows:
    ws.cell(row=r, column=1, value=date).font = black
    ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=2, value=variety).font = black
    ws.cell(row=r, column=3, value=grade).font = black
    ws.cell(row=r, column=4, value=low).font = blue
    ws.cell(row=r, column=5, value=high).font = blue
    ws.cell(row=r, column=6, value=change).font = black
    ws.cell(row=r, column=7, value=source).font = black
    ws.cell(row=r, column=8, value=ice_price).font = blue
    ws.cell(row=r, column=9, value=ice_unit).font = black
    ws.cell(row=r, column=10, value=aud_inr).font = blue
    ws.cell(row=r, column=11, value=usd_inr).font = blue
    # Landed cost is computed off the MILLED-equivalent price, not the raw
    # parchment/cherry price - what actually gets exported/cleared through
    # customs is milled clean green bean (Plantation/Cherry grade), not raw
    # parchment. Using the raw price here would understate true cost.
    outturn_expr = (
        f'IF(B{r}="Arabica",IF(ISNUMBER(SEARCH("Parchment",C{r})),0.8,0.48),'
        f'IF(ISNUMBER(SEARCH("Parchment",C{r})),0.82,0.5))'
    )
    milled_base = f'((AVERAGE(D{r}:E{r})/50/J{r})/({outturn_expr}))'
    formula = (
        f"={milled_base}*(1+Assumptions!$B${CUSTOMS_ROW})"
        f"+Assumptions!$B${FREIGHT_ROW}+Assumptions!$B${WAREHOUSE_ROW}+Assumptions!$B${DELIVERY_ROW}"
        f"+{milled_base}*Assumptions!$B${MARGIN_ROW}"
    )
    ws.cell(row=r, column=12, value=formula).font = black
    ws.cell(row=r, column=13, value=None).fill = yellow_fill
    ws.cell(row=r, column=14, value=None).fill = yellow_fill
    ws.cell(row=r, column=15, value=None).fill = yellow_fill
    ws.cell(row=r, column=16, value=note).font = note_font

    # Mid price, weeks tracked, tracked high/low, and a readable range status.
    # These are computed from YOUR OWN logged history (not a 3rd-party index), so on
    # day 1 there's no range yet - it becomes a genuine N-week high/low as days accumulate,
    # and a real 52-week high/low once ~1 year of daily entries exist.
    ws.cell(row=r, column=17, value=f"=AVERAGE(D{r}:E{r})").font = black
    ws.cell(row=r, column=18, value=f"=(A{r}-_xlfn.MINIFS($A$2:A{r},$B$2:B{r},B{r},$C$2:C{r},C{r}))/7").font = black
    ws.cell(row=r, column=18).number_format = "0.0"
    ws.cell(row=r, column=19, value=f"=_xlfn.MAXIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r})").font = black
    ws.cell(row=r, column=20, value=f"=_xlfn.MINIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r})").font = black
    range_formula = (
        f'=IF(R{r}<1,"Day 1 - building history",'
        f'IF(Q{r}>=S{r},TEXT(ROUND(R{r},0),"0")&"-wk high (since tracking began)",'
        f'IF(Q{r}<=T{r},TEXT(ROUND(R{r},0),"0")&"-wk low (since tracking began)",'
        f'TEXT(ROUND(R{r},0),"0")&"-wk range, mid-range")))'
    )
    ws.cell(row=r, column=21, value=range_formula).font = black

    # V: local price converted to USD/kg. W: the global ICE benchmark converted to USD/kg
    # (Arabica: US cents/lb -> USD/kg; Robusta: USD/tonne -> USD/kg) for a same-currency comparison.
    ws.cell(row=r, column=22, value=f"=Q{r}/50/K{r}").font = black
    ws.cell(row=r, column=23, value=f'=IF(B{r}="Arabica",H{r}/100/0.45359237,H{r}/1000)').font = black

    # X: 3/6-Month Flag - genuinely computed from YOUR OWN logged history (same
    # self-referential approach as Tracked High/Low). Green = at a 3- or
    # 6-month low (historically cheap, safer to buy); red = at a 3- or
    # 6-month high (historically expensive, riskier to buy). Needs 90+/180+
    # days of real logged history to say anything other than "insufficient
    # history yet" - this activates automatically as days accumulate.
    hist = f"(A{r}-_xlfn.MINIFS($A$2:A{r},$B$2:B{r},B{r},$C$2:C{r},C{r}))"
    six_low = f'_xlfn.MINIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r},$A$2:A{r},">="&(A{r}-180))'
    six_high = f'_xlfn.MAXIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r},$A$2:A{r},">="&(A{r}-180))'
    three_low = f'_xlfn.MINIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r},$A$2:A{r},">="&(A{r}-90))'
    three_high = f'_xlfn.MAXIFS($Q$2:Q{r},$B$2:B{r},B{r},$C$2:C{r},C{r},$A$2:A{r},">="&(A{r}-90))'
    flag_formula = (
        f'=IF(AND({hist}>=180,Q{r}<={six_low}),"6-MONTH LOW - historically cheap, safer to buy",'
        f'IF(AND({hist}>=180,Q{r}>={six_high}),"6-MONTH HIGH - historically expensive, riskier to buy",'
        f'IF(AND({hist}>=90,Q{r}<={three_low}),"3-MONTH LOW - historically cheap, safer to buy",'
        f'IF(AND({hist}>=90,Q{r}>={three_high}),"3-MONTH HIGH - historically expensive, riskier to buy",'
        f'IF({hist}>=90,"within 3/6-month range","N/A - insufficient history yet (need 90+ days)")))))'
    )
    ws.cell(row=r, column=24, value=flag_formula).font = black

    # Y/Z: Milled/clean-bean-equivalent USD/kg + the outturn % used to get there.
    # India's raw Kirehalli price is parchment/cherry (unmilled) weight; the global
    # ICE benchmark is clean green bean. These industry rule-of-thumb outturn ratios
    # (Arabica parchment ~80%, Arabica cherry ~48%, Robusta parchment ~82%, Robusta
    # cherry ~50%) convert unmilled USD/kg to a comparable milled-equivalent figure.
    # Not lab-measured for any specific lot - ask your curer for their actual outturn.
    outturn_expr = (
        f'IF(B{r}="Arabica",IF(ISNUMBER(SEARCH("Parchment",C{r})),0.8,0.48),'
        f'IF(ISNUMBER(SEARCH("Parchment",C{r})),0.82,0.5))'
    )
    ws.cell(row=r, column=25, value=f"=V{r}/{outturn_expr}").font = black
    ws.cell(row=r, column=26, value=(
        f'=IF(B{r}="Arabica",IF(ISNUMBER(SEARCH("Parchment",C{r})),"80%","48%"),'
        f'IF(ISNUMBER(SEARCH("Parchment",C{r})),"82%","50%"))'
    )).font = black

    # Indian trade terminology: washed (parchment) Arabica is sold as "Plantation"
    # once milled; Robusta's washed/milled form stays "Parchment"; natural/dry-
    # processed coffee stays "Cherry" before and after milling. No free live feed
    # publishes actual Plantation-grade quotes (checked directly), so this labels
    # our estimated milled-equivalent figure, not a live market quote.
    ws.cell(row=r, column=27, value=(
        f'=IF(B{r}="Arabica",IF(ISNUMBER(SEARCH("Parchment",C{r})),'
        f'"Arabica Plantation (estimated milled)","Arabica Cherry (estimated milled)"),'
        f'IF(ISNUMBER(SEARCH("Parchment",C{r})),"Robusta Parchment (estimated milled)",'
        f'"Robusta Cherry (estimated milled)"))'
    )).font = black
    r += 1

ws.freeze_panes = "A2"

# Conditional formatting: green when the flag says LOW, red when it says HIGH
from openpyxl.formatting.rule import FormulaRule
flag_range = f"X2:X{r - 1}"
green_fill = PatternFill("solid", fgColor="C6EFCE")
red_fill = PatternFill("solid", fgColor="FFC7CE")
ws.conditional_formatting.add(flag_range, FormulaRule(formula=['ISNUMBER(SEARCH("LOW",X2))'], fill=green_fill))
ws.conditional_formatting.add(flag_range, FormulaRule(formula=['ISNUMBER(SEARCH("HIGH",X2))'], fill=red_fill))

# ---------- Sheet 3: Read Me ----------
ws_l = wb.create_sheet("Read Me")
ws_l["A1"] = "How this workbook works"
ws_l["A1"].font = Font(name=ARIAL, bold=True, size=13)
lines = [
    "",
    "Price Log tab: one row per variety/grade per day.",
    "",
    "Tracked High / Tracked Low / Range Status (columns S-U) are computed from YOUR OWN logged history, not a",
    "third-party index. Day 1 reads 'Day 1 - building history'; it becomes a genuine N-week high/low as days",
    "accumulate, and a real 52-week high/low once ~1 year of daily entries exist.",
    "",
    "Yellow cells are for YOU to fill in:",
    "  - Assumptions tab: your real freight, customs, warehouse, delivery and margin numbers",
    "  - Price Log columns M/N/O: no public wholesale index exists for Sydney/Newcastle green beans by origin.",
    "    Log real quotes after calling suppliers (e.g. Mercanta, InterAmerican Coffee, Genovese Coffee, Complete Coffee).",
    "",
    "Est. AU Landed Cost is a CALCULATED ESTIMATE (Karnataka price -> AUD/kg + freight + customs + warehouse + delivery + margin), not a real quote.",
]
r = 2
for line in lines:
    ws_l.cell(row=r, column=1, value=line).font = Font(name=ARIAL, size=10)
    ws_l.merge_cells(f"A{r}:F{r}")
    r += 1
ws_l.column_dimensions["A"].width = 100

wb.save(OUTPUT_PATH)
print(f"Saved {OUTPUT_PATH}")
