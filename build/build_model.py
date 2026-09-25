"""Builds CPRT_Valuation_Model.xlsx (formula-driven) for the HFAC x Citadel pitch."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

OUT = "CPRT_Valuation_Model.xlsx"

wb = Workbook()
BLUE = Font(color="0000FF")            # hard-coded input
BLACK = Font(color="000000")           # formula
GREEN = Font(color="008000")           # link to other sheet
BOLD = Font(bold=True)
HDR = Font(bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F3864")
SFILL = PatternFill("solid", fgColor="D9E1F2")
IFILL = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="999999")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)

PCT = '0.0%'; NUM = '#,##0;(#,##0)'; NUM1 = '#,##0.0;(#,##0.0)'; USD = '$#,##0.00'; MULT = '0.0"x"'


def hdr(ws, row, text, ncols=8):
    ws.cell(row=row, column=1, value=text).font = HDR
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = HFILL


def sub(ws, row, text, ncols=8):
    ws.cell(row=row, column=1, value=text).font = BOLD
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = SFILL


def put(ws, row, col, val, fmt=None, font=None, fill=None):
    c = ws.cell(row=row, column=col, value=val)
    if fmt: c.number_format = fmt
    if font: c.font = font
    elif isinstance(val, str) and val.startswith("="): c.font = BLACK
    elif isinstance(val, (int, float)): c.font = BLUE
    if fill: c.fill = fill
    return c


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[L(i)].width = x

# =====================================================================
# INPUTS
# =====================================================================
inp = wb.active; inp.title = "Inputs"
widths(inp, [46, 14, 14, 14, 14, 14, 14, 44])
hdr(inp, 1, "COPART (CPRT) - MODEL INPUTS  |  blue = hard-coded input, black = formula, green = link", 8)
put(inp, 2, 1, "All $ in millions except per share. Fiscal year ends July 31.")

sub(inp, 4, "Market data & capitalization", 8)
mkt = {}
rows = [
    ("Valuation date", "2026-09-25", None, "Memo date"),
    ("Share price ($)", 28.04, USD, "NASDAQ close 24-Sep-2026"),
    ("Basic shares outstanding (M)", 925.811, NUM1, "CapIQ, latest (post FY26 buybacks)"),
    ("FY26 diluted weighted shares (M)", 956.86, NUM1, "FY26 10-K / CapIQ"),
    ("Cash & HTM/short-term investments", 4489.8, NUM1, "Balance sheet 31-Jul-2026"),
    ("Total debt (leases; no funded debt)", 88.4, NUM1, "Balance sheet 31-Jul-2026"),
    ("Minority interest", 16.6, NUM1, "Balance sheet 31-Jul-2026"),
    ("Market capitalization", "=B6*B7", NUM, "formula"),
    ("Enterprise value", "=B12-B9+B10+B11", NUM, "formula"),
    ("Target horizon (months)", 12, None, "Competition: 3-12 month horizon"),
]
r = 5
for lab, v, f, src in rows:
    put(inp, r, 1, lab); put(inp, r, 2, v, f); put(inp, r, 8, src).font = Font(italic=True, color="666666")
    mkt[lab] = f"Inputs!$B${r}"; r += 1
PRICE = "Inputs!$B$6"; SHARES = "Inputs!$B$7"; DILSH = "Inputs!$B$8"; CASH0 = "Inputs!$B$9"; DEBT0 = "Inputs!$B$10"; MI0 = "Inputs!$B$11"; MCAP = "Inputs!$B$12"; EV0 = "Inputs!$B$13"

sub(inp, 16, "Scenario selector (drives 'Model' and 'DCF' sheets)", 8)
put(inp, 17, 1, "Active scenario  (1 = Bear, 2 = Base, 3 = Bull)")
put(inp, 17, 2, 2, fill=IFILL)
dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False); inp.add_data_validation(dv); dv.add("B17")
put(inp, 17, 3, '=CHOOSE(B17,"Bear","Base","Bull")')
SCEN = "Inputs!$B$17"

sub(inp, 19, "Operating assumptions by scenario", 8)
for c, t in zip(range(2, 8), ["Bear", "Base", "Bull", "", "", ""]):
    put(inp, 20, c, t).font = BOLD
put(inp, 20, 8, "Rationale / source").font = BOLD
put(inp, 21, 1, "Scenario probability", fmt=None)
for c, v in zip((2, 3, 4), (0.25, 0.50, 0.25)):
    put(inp, 21, c, v, PCT)
put(inp, 21, 5, "=SUM(B21:D21)", PCT)
put(inp, 21, 8, "Must sum to 100%; base weighted to reflect Q4-26 assignment data").font = Font(italic=True, color="666666")

# assumption blocks: name -> {scenario: [FY27..FY31]}
years = ["FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]
assump = [
    ("US unit growth", {1: [-0.03, -0.01, 0.0, 0.0, 0.0], 2: [-0.01, 0.03, 0.025, 0.025, 0.025], 3: [0.01, 0.045, 0.04, 0.04, 0.04]},
     PCT, "Q4-26 US insurance assignments +2.3% ex-one lost customer (Progressive); Q1-27 still down MSD, growth returns by Q4-27 (base)"),
    ("International unit growth", {1: [0.025] * 5, 2: [0.055] * 5, 3: [0.08] * 5},
     PCT, "Q4-26 intl units +10%, intl insurance +11.2%; FY26 intl revenue +8.5%; ex-FX mid/high-single digits"),
    ("Revenue per unit (RPU) growth", {1: [0.01, -0.01, -0.01, 0.0, 0.0], 2: [0.035] * 5, 3: [0.055, 0.055, 0.05, 0.05, 0.05]},
     PCT, "FY26 RPU +5.7%; rising ASP & total-loss frequency (23.3% record CQ2-26), Title Express/long-haul fees. Bear = used-car deflation"),
    ("Core EBIT margin (ex-ACV)", {1: [0.335, 0.325, 0.32, 0.32, 0.32], 2: [0.35, 0.36, 0.365, 0.37, 0.37], 3: [0.36, 0.375, 0.385, 0.39, 0.39]},
     PCT, "FY26 35.4% (FY25 37.7%). Operating leverage on fixed land base reverses when units grow; ~1pt of unit growth = ~$31M EBIT (deep dive)"),
    ("D&A % revenue (ex-ACV)", {1: [0.05] * 5, 2: [0.05] * 5, 3: [0.05] * 5}, PCT, "FY26 $229M = 4.9% of revenue"),
    ("Capex % revenue (ex-ACV)", {1: [0.08] * 5, 2: [0.085] * 5, 3: [0.095] * 5}, PCT, "FY26 7.2% (land buying paused); FY24-25 ~12% incl. land banking; maintenance ~$230M (~5%)"),
    ("Change in NWC (% of revenue change)", {1: [0.03] * 5, 2: [0.03] * 5, 3: [0.03] * 5}, PCT, "Historical NWC ~1% of OCF"),
    ("Effective tax rate", {1: [0.195] * 5, 2: [0.195] * 5, 3: [0.195] * 5}, PCT, "FY26 19.3%, FY25 18.3%"),
    ("Yield on cash", {1: [0.035] * 5, 2: [0.038] * 5, 3: [0.038] * 5}, PCT, "FY26 interest income $182M on ~$4.6B avg cash = 3.9%"),
    ("Share buybacks ($M / yr)", {1: [200, 200, 300, 300, 300], 2: [500, 700, 800, 800, 800], 3: [800, 1000, 1000, 1000, 1000]},
     NUM, "FY26 $1.63B; paused for ACV; ~$2.6B cash left post-deal + $1.25B revolver"),
    ("Assumed avg buyback price ($)", {1: [26, 24, 24, 26, 28], 2: [30, 36, 40, 44, 48], 3: [32, 42, 50, 56, 62]},
     USD, "Assumption consistent with each scenario's price path"),
    ("Gross share issuance (SBC dilution)", {1: [0.003] * 5, 2: [0.003] * 5, 3: [0.003] * 5}, PCT, "FY26 ~$31M issuance / SBC $39M"),
    ("ACV pretax synergies ($M)", {1: [0, 0, 10, 20, 25], 2: [0, 25, 60, 90, 110], 3: [0, 50, 110, 160, 190]},
     NUM, "ACV deep dive: ~$25M needed for FY28 accretion; levers = yard sharing, logistics, buyer cross-sell, floorplan, G&A"),
    ("ACV revenue growth (CY)", {1: [0.06] * 5, 2: [0.12, 0.11, 0.10, 0.09, 0.08], 3: [0.15, 0.14, 0.12, 0.11, 0.10]},
     PCT, "ACV CY25 $760M -> CY26 guide ~$850M (+12%)"),
    ("ACV adj. EBITDA margin (standalone)", {1: [0.085, 0.09, 0.095, 0.10, 0.10], 2: [0.10, 0.12, 0.14, 0.15, 0.16], 3: [0.11, 0.14, 0.16, 0.18, 0.19]},
     PCT, "CY26 guide $73-77M on $850M = 8.8%"),
    ("Target P/E on NTM EPS (12-mo)", {1: [15.0], 2: [24.0], 3: [30.0]}, MULT,
     "Now 17.7x NTM (CapIQ); 2-yr avg 27x; 5-10 yr avg ~30x+. Bear = de-rate to RBA/KAR level"),
    ("Target EV/EBITDA on NTM EBITDA (12-mo)", {1: [10.0], 2: [15.5], 3: [19.0]}, MULT,
     "Now 11.3x NTM; 2-yr avg 18.9x; RBA ~13-15x LTM (levered, 30% lower margin)"),
]
A = {}  # name -> dict(scenario -> list of cell refs (absolute)), plus 'live' -> list of refs
r = 22
for name, d, fmt, note in assump:
    n = len(d[2])
    put(inp, r, 1, name).font = BOLD
    put(inp, r, 8, note).font = Font(italic=True, color="666666")
    if n == 5:
        for i, y in enumerate(years):
            put(inp, r, 2 + i, y).font = Font(italic=True)
        r += 1
        A[name] = {}
        for s, lab in ((1, "Bear"), (2, "Base"), (3, "Bull")):
            put(inp, r, 1, "   " + lab)
            A[name][s] = []
            for i, v in enumerate(d[s]):
                put(inp, r, 2 + i, v, fmt)
                A[name][s].append(f"Inputs!${L(2+i)}${r}")
            r += 1
        put(inp, r, 1, "   Active (per selector)").font = Font(italic=True)
        A[name]["live"] = []
        for i in range(5):
            col = L(2 + i)
            put(inp, r, 2 + i, f"=CHOOSE({SCEN},{col}{r-3},{col}{r-2},{col}{r-1})", fmt)
            A[name]["live"].append(f"Inputs!${col}${r}")
        r += 2
    else:
        A[name] = {}
        for s, c in ((1, 2), (2, 3), (3, 4)):
            put(inp, r, c, d[s][0], fmt); A[name][s] = f"Inputs!${L(c)}${r}"
        put(inp, r, 5, f"=CHOOSE({SCEN},B{r},C{r},D{r})", fmt); A[name]["live"] = f"Inputs!$E${r}"
        put(inp, r, 6, "<- active").font = Font(italic=True)
        r += 1
r += 1
sub(inp, r, "ACV Auctions acquisition (announced 10-Sep-2026)", 8); r += 1
acv_rows = [
    ("Equity purchase price ($M, cash)", 1900, NUM, "$10.50/sh tender; ~35x CY25 adj. EBITDA, 2.7x sales"),
    ("Expected close (fiscal month of FY27, 1-12)", 5, None, "Close targeted by CY-end 2026 -> ~Dec-2026 = month 5 of FY27"),
    ("Months consolidated in FY27", "=12-B{r1}", None, "formula"),
    ("ACV CY2026E revenue ($M)", 850, NUM, "ACV guidance"),
    ("ACV stock-based comp ($M/yr, CY26)", 70, NUM, "Assumption: gap between adj. EBITDA (~$75M) and GAAP op. loss (~-$46M) = SBC + D&A"),
    ("ACV existing D&A ($M/yr, CY26)", 50, NUM, "Assumption (see above)"),
    ("Acquired intangibles as % of price", 0.30, PCT, "Assumption; purchase price allocation not yet published"),
    ("Intangible amortization life (yrs)", 10, None, "Assumption"),
    ("Annual acquired-intangible amortization ($M)", "=B{r7}*B{r8}/B{r9}", NUM, "formula"),
    ("ACV cost growth (SBC & D&A, per yr)", 0.05, PCT, "Assumption"),
]
acv = {}
r0 = r
for i, (lab, v, f, src) in enumerate(acv_rows):
    if isinstance(v, str):
        v = v.format(r1=r0 + 1, r7=r0, r8=r0 + 6, r9=r0 + 7)
    put(inp, r, 1, lab); put(inp, r, 2, v, f); put(inp, r, 8, src).font = Font(italic=True, color="666666")
    acv[lab] = f"Inputs!$B${r}"; r += 1
ACV_PRICE = acv["Equity purchase price ($M, cash)"]; ACV_MONTHS = acv["Months consolidated in FY27"]
ACV_REV0 = acv["ACV CY2026E revenue ($M)"]; ACV_SBC = acv["ACV stock-based comp ($M/yr, CY26)"]; ACV_DA = acv["ACV existing D&A ($M/yr, CY26)"]
ACV_AMORT = acv["Annual acquired-intangible amortization ($M)"]; ACV_CG = acv["ACV cost growth (SBC & D&A, per yr)"]

r += 1
sub(inp, r, "DCF assumptions", 8); r += 1
dcf_in = {}
for lab, v, f, src in [
    ("Risk-free rate", 0.042, PCT, "10-yr UST, Sep-2026 (approx.)"),
    ("Equity risk premium", 0.045, PCT, "Assumption"),
    ("Beta", 1.0, "0.00", "CNBC/StockAnalysis 5-yr beta ~1.03-1.04"),
    ("Cost of equity", "=B{a}+B{c}*B{b}", PCT, "CAPM"),
    ("Pre-tax cost of debt", 0.05, PCT, "Undrawn revolver; immaterial"),
    ("Debt weight", 0.0, PCT, "No funded debt"),
    ("WACC", "=B{d}*(1-B{f})+B{e}*(1-0.195)*B{f}", PCT, "formula"),
    ("Terminal FCF growth", 0.035, PCT, "RPU growth alone has run +3-6%/yr; intl expansion; conservative vs 10-yr revenue CAGR ~13%"),
    ("Mid-year convention (1 = yes)", 1, None, ""),
]:
    if isinstance(v, str):
        v = v.format(a=r - 3, b=r - 2, c=r - 1, d=r - 3, e=r - 2, f=r - 1) if lab in ("Cost of equity",) else v
    put(inp, r, 1, lab); put(inp, r, 2, v, f); put(inp, r, 8, src).font = Font(italic=True, color="666666")
    dcf_in[lab] = f"Inputs!$B${r}"; r += 1
# fix WACC formula refs explicitly
wr = int(dcf_in["WACC"].split("$")[-1])
inp[f"B{wr}"] = f"=B{wr-3}*(1-B{wr-1})+B{wr-2}*(1-0.195)*B{wr-1}"
cer = int(dcf_in["Cost of equity"].split("$")[-1])
inp[f"B{cer}"] = f"=B{cer-3}+B{cer-1}*B{cer-2}"
WACC = dcf_in["WACC"]; TG = dcf_in["Terminal FCF growth"]; MIDYR = dcf_in["Mid-year convention (1 = yes)"]

# =====================================================================
# HISTORICAL
# =====================================================================
h = wb.create_sheet("Historical")
widths(h, [40, 12, 12, 12, 12, 12, 12, 12])
hdr(h, 1, "HISTORICAL FINANCIALS (S&P Capital IQ export; FY ends Jul-31; $M)", 8)
yrs_h = ["FY21A", "FY22A", "FY23A", "FY24A", "FY25A", "FY26A"]
for i, y in enumerate(yrs_h):
    put(h, 3, 2 + i, y).font = BOLD
hist = {
    "Total revenue": [2692.5, 3500.9, 3869.5, 4236.8, 4647.0, 4666.2],
    "  US revenue": [2272.1, 2945.2, 3189.6, 3464.7, 3855.1, 3807.4],
    "  International revenue": [420.4, 555.8, 679.9, 772.1, 791.9, 858.8],
    "  Service revenue": [None, None, None, None, 3968.7, 3969.5],
    "  Vehicle sales": [None, None, None, None, 678.3, 696.7],
    "Gross profit (co. definition)": [1343.1, 1606.2, 1737.0, 1939.2, 2099.6, 2083.9],
    "Operating income (EBIT)": [1136.4, 1375.0, 1486.6, 1573.2, 1752.7, 1652.6],
    "  US EBIT": [1023.6, 1247.6, 1368.1, 1428.0, 1480.9, 1419.0],
    "  International EBIT": [112.9, 127.4, 118.5, 144.0, 215.8, 233.6],
    "D&A": [112.4, 128.8, 152.3, 181.1, 202.8, 229.4],
    "EBITDA": [1248.8, 1503.8, 1638.9, 1754.3, 1955.5, 1882.0],
    "Interest & investment income (net)": [-20.2, -16.7, 65.9, 145.7, 178.9, 181.9],
    "Pre-tax income": [1121.8, 1341.0, 1554.3, 1714.6, 1895.6, 1835.0],
    "Income tax": [185.4, 250.8, 316.6, 352.3, 347.2, 354.6],
    "Net income": [936.5, 1090.1, 1237.7, 1363.0, 1552.4, 1484.3],
    "Diluted EPS ($)": [0.97, 1.13, 1.28, 1.40, 1.59, 1.55],
    "Diluted shares (M)": [961.2, 964.6, 966.6, 974.8, 977.6, 956.9],
    "Cash from operations": [990.9, 1176.7, 1364.2, 1472.6, 1799.8, 1604.5],
    "Capital expenditure": [-463.0, -337.4, -516.6, -511.0, -569.0, -337.4],
    "Share repurchases": [-6.1, -1.9, -4.7, -6.6, -5.3, -1636.3],
    "Cash & ST investments (year-end)": [1048.3, 1384.2, 2364.0, 3422.2, 4789.1, 4489.8],
    "Total debt (leases)": [518.1, 119.5, 120.5, 118.7, 103.7, 88.4],
    "Land (at cost)": [1428.3, 1526.4, 1812.0, 2027.6, 2394.6, None],
}
hr = {}
r = 4
for k, vals in hist.items():
    put(h, r, 1, k).font = BOLD if not k.startswith("  ") else BLACK
    for i, v in enumerate(vals):
        if v is not None:
            put(h, r, 2 + i, v, USD if "EPS" in k else NUM1)
    hr[k] = r; r += 1
r += 1
sub(h, r, "Derived ratios", 8); r += 1
for lab, f, fmt in [
    ("Revenue growth", "=IF({p}=\"\",\"\",{c}/{p}-1)", PCT),
    ("EBIT margin", "={e}/{rev}", PCT),
    ("EBITDA margin", "={eb}/{rev}", PCT),
    ("Free cash flow (CFO - capex)", "={cfo}+{cx}", NUM1),
    ("FCF margin", "=({cfo}+{cx})/{rev}", PCT),
    ("Capex % revenue", "=-{cx}/{rev}", PCT),
    ("International % revenue", "={intl}/{rev}", PCT),
    ("Effective tax rate", "={tax}/{pbt}", PCT),
]:
    put(h, r, 1, lab)
    for i in range(6):
        c = L(2 + i); p = L(1 + i)
        ff = f.format(p=f"{p}{hr['Total revenue']}" if i > 0 else '""', c=f"{c}{hr['Total revenue']}",
                      e=f"{c}{hr['Operating income (EBIT)']}", rev=f"{c}{hr['Total revenue']}", eb=f"{c}{hr['EBITDA']}",
                      cfo=f"{c}{hr['Cash from operations']}", cx=f"{c}{hr['Capital expenditure']}",
                      intl=f"{c}{hr['  International revenue']}", tax=f"{c}{hr['Income tax']}", pbt=f"{c}{hr['Pre-tax income']}")
        if i == 0 and lab == "Revenue growth":
            continue
        put(h, r, 2 + i, ff, fmt)
    hr[lab] = r; r += 1
r += 1
sub(h, r, "Valuation history (CapIQ, quarterly averages of daily multiples)", 8); r += 1
qs = ["CQ4-24", "CQ1-25", "CQ2-25", "CQ3-25", "CQ4-25", "CQ1-26", "CQ2-26", "CQ3-26*"]
widths(h, [40, 12, 12, 12, 12, 12, 12, 12, 12])
for i, q in enumerate(qs):
    put(h, r, 2 + i, q).font = BOLD
r += 1
for lab, vals in [("TEV / NTM EBITDA (avg)", [25.8, 24.5, 23.8, 19.6, 17.0, 15.4, 13.0, 11.9]),
                  ("P / NTM EPS (avg)", [36.1, 34.5, 33.7, 28.4, 24.6, 22.9, 20.0, 18.5]),
                  ("TEV / NTM revenue (avg)", [11.0, 10.5, 10.2, 8.4, 7.3, 6.6, 5.5, 5.0])]:
    put(h, r, 1, lab)
    for i, v in enumerate(vals):
        put(h, r, 2 + i, v, MULT)
    put(h, r, 10, f"=AVERAGE(B{r}:I{r})", MULT); put(h, r, 11, "2-yr avg").font = Font(italic=True)
    hr[lab] = r; r += 1
put(h, r, 1, "* quarter-to-date through 24-Sep-2026. Source: Capital IQ 'Multiples' tab of provided export.").font = Font(italic=True, color="666666")
HIST_PE_AVG = f"Historical!$J${hr['P / NTM EPS (avg)']}"; HIST_EVE_AVG = f"Historical!$J${hr['TEV / NTM EBITDA (avg)']}"

# =====================================================================
# MODEL (scenario-driven) -- generic builder so we can reuse for Scenarios sheet
# =====================================================================
FY26 = {"us_rev": 3807.4, "intl_rev": 858.8, "ebit": 1652.6, "da": 229.4, "int": 181.9, "tax": 354.6, "pbt": 1835.0,
        "ni": 1484.3, "eps": 1.55, "dil": 956.86, "cash": 4489.8, "capex": 337.4}


def build_model(ws, start_row, scen, title, ncols=8):
    """Writes a full 5-year operating model starting at start_row. scen = 1/2/3 or 'live'.
    Returns dict of row numbers for key lines. Columns: A label, B FY26A, C..G FY27E..FY31E."""
    def a(name, i):
        return A[name][scen][i]
    R = {}
    hdr(ws, start_row, title, ncols)
    r = start_row + 1
    for i, y in enumerate(["FY26A"] + years):
        put(ws, r, 2 + i, y).font = BOLD
    r += 1

    def line(label, fy26, fn, fmt=NUM1, bold=False):
        nonlocal r
        c = put(ws, r, 1, label)
        if bold: c.font = BOLD
        if fy26 is not None:
            put(ws, r, 2, fy26, fmt, font=GREEN if isinstance(fy26, str) else BLUE)
        for i in range(5):
            col = L(3 + i); prev = L(2 + i)
            put(ws, r, 3 + i, fn(i, col, prev), fmt)
        R[label] = r; r += 1
        return r - 1

    sub(ws, r, "Core Copart (ex-ACV)", ncols); r += 1
    line("US revenue", FY26["us_rev"], lambda i, c, p: f"={p}{r}*(1+{a('US unit growth', i)})*(1+{a('Revenue per unit (RPU) growth', i)})")
    line("International revenue", FY26["intl_rev"], lambda i, c, p: f"={p}{r}*(1+{a('International unit growth', i)})*(1+{a('Revenue per unit (RPU) growth', i)})")
    line("Core revenue", f"=B{R['US revenue']}+B{R['International revenue']}", lambda i, c, p: f"={c}{R['US revenue']}+{c}{R['International revenue']}", bold=True)
    line("  growth", None, lambda i, c, p: f"={c}{R['Core revenue']}/{p}{R['Core revenue']}-1", PCT)
    line("Core EBIT margin", f"=B{R['Core revenue']+3}/B{R['Core revenue']}", lambda i, c, p: f"={a('Core EBIT margin (ex-ACV)', i)}", PCT)
    line("Core EBIT", FY26["ebit"], lambda i, c, p: f"={c}{R['Core revenue']}*{c}{R['Core EBIT margin']}", bold=True)
    line("Core D&A", FY26["da"], lambda i, c, p: f"={c}{R['Core revenue']}*{a('D&A % revenue (ex-ACV)', i)}")
    line("Core capex", FY26["capex"], lambda i, c, p: f"={c}{R['Core revenue']}*{a('Capex % revenue (ex-ACV)', i)}")

    sub(ws, r, "ACV Auctions (consolidated from close; CY figures used as FY proxy)", ncols); r += 1
    line("Consolidation factor (share of year)", 0, lambda i, c, p: f"={ACV_MONTHS}/12" if i == 0 else "=1", "0.00")
    line("ACV revenue (annualized run-rate)", None, lambda i, c, p: (f"={ACV_REV0}*(1+{a('ACV revenue growth (CY)', i)})" if i == 0 else f"={p}{r}*(1+{a('ACV revenue growth (CY)', i)})"))
    line("ACV revenue consolidated", 0, lambda i, c, p: f"={c}{R['ACV revenue (annualized run-rate)']}*{c}{R['Consolidation factor (share of year)']}")
    line("ACV adj. EBITDA (standalone)", None, lambda i, c, p: f"={c}{R['ACV revenue consolidated']}*{a('ACV adj. EBITDA margin (standalone)', i)}")
    line("  less: ACV SBC", None, lambda i, c, p: f"=-{ACV_SBC}*(1+{ACV_CG})^{i+1}*{c}{R['Consolidation factor (share of year)']}")
    line("  less: ACV existing D&A", None, lambda i, c, p: f"=-{ACV_DA}*(1+{ACV_CG})^{i+1}*{c}{R['Consolidation factor (share of year)']}")
    line("  less: acquired-intangible amortization", None, lambda i, c, p: f"=-{ACV_AMORT}*{c}{R['Consolidation factor (share of year)']}")
    line("  plus: synergies (pre-tax)", None, lambda i, c, p: f"={a('ACV pretax synergies ($M)', i)}")
    line("ACV EBIT contribution (GAAP)", 0, lambda i, c, p: f"=SUM({c}{R['ACV adj. EBITDA (standalone)']}:{c}{R['  plus: synergies (pre-tax)']})", bold=True)
    line("ACV D&A (total)", 0, lambda i, c, p: f"=-{c}{R['  less: ACV existing D&A']}-{c}{R['  less: acquired-intangible amortization']}")
    line("ACV capex", 0, lambda i, c, p: f"={c}{R['ACV revenue consolidated']}*0.02")

    sub(ws, r, "Consolidated P&L", ncols); r += 1
    line("Total revenue", f"=B{R['Core revenue']}", lambda i, c, p: f"={c}{R['Core revenue']}+{c}{R['ACV revenue consolidated']}", bold=True)
    line("  growth", None, lambda i, c, p: f"={c}{R['Total revenue']}/{p}{R['Total revenue']}-1", PCT)
    line("EBIT", f"=B{R['Core EBIT']}", lambda i, c, p: f"={c}{R['Core EBIT']}+{c}{R['ACV EBIT contribution (GAAP)']}", bold=True)
    line("  EBIT margin", f"=B{R['EBIT']}/B{R['Total revenue']}", lambda i, c, p: f"={c}{R['EBIT']}/{c}{R['Total revenue']}", PCT)
    line("D&A", f"=B{R['Core D&A']}", lambda i, c, p: f"={c}{R['Core D&A']}+{c}{R['ACV D&A (total)']}")
    line("EBITDA", f"=B{R['EBIT']}+B{R['D&A']}", lambda i, c, p: f"={c}{R['EBIT']}+{c}{R['D&A']}", bold=True)
    line("Interest income on cash", FY26["int"], lambda i, c, p: f"=AVERAGE({p}{r+11},{c}{r+11})*{a('Yield on cash', i)}")  # placeholder fixed below
    line("Pre-tax income", FY26["pbt"], lambda i, c, p: f"={c}{R['EBIT']}+{c}{R['Interest income on cash']}", bold=True)
    line("Tax", -FY26["tax"], lambda i, c, p: f"=-{c}{R['Pre-tax income']}*{a('Effective tax rate', i)}")
    line("Net income", FY26["ni"], lambda i, c, p: f"={c}{R['Pre-tax income']}+{c}{R['Tax']}", bold=True)
    line("Diluted shares (M, avg)", FY26["dil"], lambda i, c, p: f"=AVERAGE({p}{r+2},{c}{r+2})")  # placeholder fixed below
    line("Diluted EPS ($)", FY26["eps"], lambda i, c, p: f"={c}{R['Net income']}/{c}{R['Diluted shares (M, avg)']}", USD, bold=True)
    line("  EPS growth", None, lambda i, c, p: f"={c}{R['Diluted EPS ($)']}/{p}{R['Diluted EPS ($)']}-1", PCT)
    line("Cash EPS (ex acquired-intangible amort., $)", FY26["eps"], lambda i, c, p: f"=({c}{R['Net income']}-{c}{R['  less: acquired-intangible amortization']}*(1-{a('Effective tax rate', i)}))/{c}{R['Diluted shares (M, avg)']}", USD)

    sub(ws, r, "Cash flow & share count", ncols); r += 1
    line("Unlevered FCF (EBIT*(1-t) + D&A - capex - dNWC)", None,
         lambda i, c, p: f"={c}{R['EBIT']}*(1-{a('Effective tax rate', i)})+{c}{R['D&A']}-{c}{R['Core capex']}-{c}{R['ACV capex']}-({c}{R['Total revenue']}-{p}{R['Total revenue']})*{a('Change in NWC (% of revenue change)', i)}", bold=True)
    line("Levered FCF (NI + D&A - capex - dNWC)", FY26["ni"] + FY26["da"] - FY26["capex"] - 0,
         lambda i, c, p: f"={c}{R['Net income']}+{c}{R['D&A']}-{c}{R['Core capex']}-{c}{R['ACV capex']}-({c}{R['Total revenue']}-{p}{R['Total revenue']})*{a('Change in NWC (% of revenue change)', i)}")
    line("Share buybacks", -1636.3, lambda i, c, p: f"=-{a('Share buybacks ($M / yr)', i)}")
    line("ACV acquisition", 0, lambda i, c, p: f"=-{ACV_PRICE}" if i == 0 else "=0")
    line("Ending cash & investments", FY26["cash"], lambda i, c, p: f"={p}{r}+{c}{R['Levered FCF (NI + D&A - capex - dNWC)']}+{c}{R['Share buybacks']}+{c}{R['ACV acquisition']}", bold=True)
    line("Shares repurchased (M)", None, lambda i, c, p: f"={a('Share buybacks ($M / yr)', i)}/{a('Assumed avg buyback price ($)', i)}")
    line("Ending diluted shares (M)", FY26["dil"] - 30, lambda i, c, p: f"={p}{r}*(1+{a('Gross share issuance (SBC dilution)', i)})-{c}{R['Shares repurchased (M)']}")
    ws.cell(row=R["Ending diluted shares (M)"], column=2).font = BLUE
    put(ws, R["Ending diluted shares (M)"], 9, "FY26 end-point: 925.8M basic + ~1M dilution").font = Font(italic=True, color="666666")
    ws.cell(row=R["Ending diluted shares (M)"], column=2, value=927.0)
    # fix placeholders for interest income & avg shares
    for i in range(5):
        c = L(3 + i); p = L(2 + i)
        ws.cell(row=R["Interest income on cash"], column=3 + i,
                value=f"=({p}{R['Ending cash & investments']}+{c}{R['Share buybacks']}+{c}{R['ACV acquisition']})*{a('Yield on cash', i)}")
        ws.cell(row=R["Diluted shares (M, avg)"], column=3 + i,
                value=f"=AVERAGE({p}{R['Ending diluted shares (M)']},{c}{R['Ending diluted shares (M)']})")
    put(ws, R["Interest income on cash"], 9, "On beginning cash less buybacks & ACV outlay (avoids circularity)").font = Font(italic=True, color="666666")
    return R, r


m = wb.create_sheet("Model")
widths(m, [48, 12, 12, 12, 12, 12, 12, 4, 40])
MR, end = build_model(m, 1, "live", "OPERATING MODEL - ACTIVE SCENARIO (change Inputs!B17)  |  $M, FY ends Jul-31", 7)
m.cell(row=1, column=8, value=f'=CHOOSE({SCEN},"BEAR","BASE","BULL")').font = HDR
m.cell(row=1, column=8).fill = HFILL

# =====================================================================
# SCENARIOS sheet: three full models + 12-month target price
# =====================================================================
s = wb.create_sheet("Scenarios")
widths(s, [48, 12, 12, 12, 12, 12, 12, 4, 40])
SR = {}
r = 1
for sc, nm in ((1, "BEAR"), (2, "BASE"), (3, "BULL")):
    SR[sc], r = build_model(s, r, sc, f"{nm} CASE - full model", 7)
    r += 1

# =====================================================================
# VALUATION sheet (12-month target price, probability-weighted)
# =====================================================================
v = wb.create_sheet("Valuation")
widths(v, [50, 14, 14, 14, 14, 40])
hdr(v, 1, "12-MONTH TARGET PRICE - SCENARIO ANALYSIS (target date ~Sep-2027; NTM = FY28E)", 6)
for c, t in zip((2, 3, 4, 5), ("Bear", "Base", "Bull", "Prob-wtd")):
    put(v, 2, c, t).font = BOLD
put(v, 2, 6, "Notes").font = BOLD
VR = {}
r = 3


def vline(label, fn, fmt=NUM1, wtd=None, note="", bold=False):
    global r
    c = put(v, r, 1, label)
    if bold: c.font = BOLD
    for i, sc in enumerate((1, 2, 3)):
        put(v, r, 2 + i, fn(sc, L(2 + i)), fmt)
    if wtd == "sumproduct":
        put(v, r, 5, f"=SUMPRODUCT(B{r}:D{r},Inputs!$B$21:$D$21)", fmt)
    elif wtd:
        put(v, r, 5, wtd, fmt)
    if note: put(v, r, 6, note).font = Font(italic=True, color="666666")
    VR[label] = r; r += 1


def S(sc, key, col="D"):  # FY28E column = D in model sheets (B=FY26, C=FY27, D=FY28)
    return f"=Scenarios!${col}${SR[sc][key]}"


vline("Probability", lambda sc, c: f"=Inputs!{L(1+sc)}$21", PCT, "=SUM(B3:D3)")
sub(v, r, "Operating outputs", 6); r += 1
vline("FY27E revenue", lambda sc, c: S(sc, "Total revenue", "C"), NUM, "sumproduct", "Consensus FY27E $4,914M (CapIQ)")
vline("FY28E revenue", lambda sc, c: S(sc, "Total revenue"), NUM, "sumproduct", "Consensus FY28E $5,306M (CapIQ)")
vline("FY27E EPS", lambda sc, c: S(sc, "Diluted EPS ($)", "C"), USD, "sumproduct", "Consensus FY27E $1.63")
vline("FY28E EPS (NTM at target date)", lambda sc, c: S(sc, "Diluted EPS ($)"), USD, "sumproduct", "Consensus FY28E $1.77", bold=True)
vline("FY28E cash EPS (ex deal amortization)", lambda sc, c: S(sc, "Cash EPS (ex acquired-intangible amort., $)"), USD, "sumproduct")
vline("FY28E EBITDA", lambda sc, c: S(sc, "EBITDA"), NUM, "sumproduct")
vline("FY27E ending net cash", lambda sc, c: f"{S(sc, 'Ending cash & investments', 'C')}-{DEBT0}-{MI0}", NUM, "sumproduct", "Post-ACV cash, post-buybacks")
vline("FY27E ending diluted shares (M)", lambda sc, c: S(sc, "Ending diluted shares (M)", "C"), NUM1, "sumproduct")
sub(v, r, "Method 1: P/E on NTM EPS", 6); r += 1
vline("Target P/E", lambda sc, c: f"={A['Target P/E on NTM EPS (12-mo)'][sc]}", MULT, None, "Current 17.7x NTM; 2-yr avg 27x")
vline("Implied price (P/E)", lambda sc, c: f"={c}{VR['Target P/E']}*{c}{VR['FY28E EPS (NTM at target date)']}", USD, "sumproduct", bold=True)
sub(v, r, "Method 2: EV/EBITDA on NTM EBITDA", 6); r += 1
vline("Target EV/EBITDA", lambda sc, c: f"={A['Target EV/EBITDA on NTM EBITDA (12-mo)'][sc]}", MULT, None, "Current 11.3x NTM; 2-yr avg 18.9x")
vline("Implied EV", lambda sc, c: f"={c}{VR['Target EV/EBITDA']}*{c}{VR['FY28E EBITDA']}", NUM)
vline("Implied equity value", lambda sc, c: f"={c}{VR['Implied EV']}+{c}{VR['FY27E ending net cash']}", NUM)
vline("Implied price (EV/EBITDA)", lambda sc, c: f"={c}{VR['Implied equity value']}/{c}{VR['FY27E ending diluted shares (M)']}", USD, "sumproduct", bold=True)
sub(v, r, "Method 3: DCF (intrinsic, from DCF sheet - active scenario only)", 6); r += 1
vline("DCF value per share (active scenario)", lambda sc, c: f"=IF({SCEN}={sc},DCF!$B$30,\"n/a - set selector\")", USD, None, "Run each scenario via Inputs!B17 to populate")
sub(v, r, "Blended 12-month target", 6); r += 1
put(v, r, 1, "Weight: P/E method"); put(v, r, 2, 0.5, PCT); put(v, r, 6, "Remaining weight to EV/EBITDA").font = Font(italic=True, color="666666"); WPE = f"$B${r}"; r += 1
vline("12-month target price", lambda sc, c: f"={WPE}*{c}{VR['Implied price (P/E)']}+(1-{WPE})*{c}{VR['Implied price (EV/EBITDA)']}", USD, "sumproduct", bold=True)
vline("Upside / (downside) vs current", lambda sc, c: f"={c}{VR['12-month target price']}/{PRICE}-1", PCT, f"=E{r-1}/{PRICE}-1", bold=True)
vline("Implied NTM P/E at target (check)", lambda sc, c: f"={c}{VR['12-month target price']}/{c}{VR['FY28E EPS (NTM at target date)']}", MULT)
vline("Implied EV/NTM EBITDA at target (check)", lambda sc, c: f"=({c}{VR['12-month target price']}*{c}{VR['FY27E ending diluted shares (M)']}-{c}{VR['FY27E ending net cash']})/{c}{VR['FY28E EBITDA']}", MULT)
r += 1
sub(v, r, "Risk / reward", 6); r += 1
put(v, r, 1, "Current price"); put(v, r, 2, f"={PRICE}", USD, font=GREEN); r += 1
put(v, r, 1, "Probability-weighted target"); put(v, r, 2, f"=E{VR['12-month target price']}", USD); PWT = f"Valuation!$B${r}"; r += 1
put(v, r, 1, "Probability-weighted expected return"); put(v, r, 2, f"=E{VR['Upside / (downside) vs current']}", PCT); PWR = f"Valuation!$B${r}"; r += 1
put(v, r, 1, "Upside (bull) / downside (bear) ratio"); put(v, r, 2, f"=(D{VR['12-month target price']}-{PRICE})/({PRICE}-B{VR['12-month target price']})", '0.0"x"'); r += 1
put(v, r, 1, "Current NTM P/E (FY27E consensus EPS $1.63)"); put(v, r, 2, f"={PRICE}/1.63", MULT); r += 1
put(v, r, 1, "Current EV / FY26A EBITDA"); put(v, r, 2, f"={EV0}/1882", MULT); r += 1
r += 1
sub(v, r, "Sensitivity: 12-month price = Target P/E x FY28E EPS", 6); r += 1
put(v, r, 1, "FY28E EPS ->").font = BOLD
eps_list = [1.45, 1.55, 1.65, 1.75, 1.85, 1.95]
pe_list = [14, 16, 18, 20, 22, 24, 26, 28, 30]
for j, e in enumerate(eps_list):
    put(v, r, 2 + j, e, USD).font = BOLD
r += 1
for pe in pe_list:
    put(v, r, 1, pe, MULT).font = BOLD
    for j in range(len(eps_list)):
        put(v, r, 2 + j, f"=$A{r}*{L(2+j)}${r-1-pe_list.index(pe)}", USD)
    r += 1
put(v, r, 1, "Green = above current price; red = below").font = Font(italic=True, color="666666")
rng = f"B{r-len(pe_list)}:{L(1+len(eps_list))}{r-1}"
v.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=[PRICE], fill=PatternFill("solid", fgColor="C6EFCE")))
v.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=[PRICE], fill=PatternFill("solid", fgColor="FFC7CE")))
widths(v, [50, 14, 14, 14, 14, 14, 14, 14])

# =====================================================================
# DCF sheet
# =====================================================================
d = wb.create_sheet("DCF")
widths(d, [44, 13, 13, 13, 13, 13, 13, 30])
hdr(d, 1, "DCF - ACTIVE SCENARIO (5-yr explicit, Gordon growth terminal)  |  $M", 8)
d.cell(row=1, column=8, value=f'=CHOOSE({SCEN},"BEAR","BASE","BULL")').font = HDR; d.cell(row=1, column=8).fill = HFILL
for i, y in enumerate(years):
    put(d, 3, 2 + i, y).font = BOLD
put(d, 4, 1, "Unlevered FCF")
for i in range(5):
    put(d, 4, 2 + i, f"=Model!{L(3+i)}{MR['Unlevered FCF (EBIT*(1-t) + D&A - capex - dNWC)']}", NUM1, font=GREEN)
put(d, 5, 1, "Discount period (yrs)")
for i in range(5):
    put(d, 5, 2 + i, f"={i+1}-0.5*{MIDYR}", "0.0")
put(d, 6, 1, "Discount factor")
for i in range(5):
    put(d, 6, 2 + i, f"=1/(1+{WACC})^{L(2+i)}5", "0.000")
put(d, 7, 1, "PV of FCF")
for i in range(5):
    put(d, 7, 2 + i, f"={L(2+i)}4*{L(2+i)}6", NUM1)
r = 9
for lab, f, fmt in [
    ("WACC", f"={WACC}", PCT), ("Terminal growth", f"={TG}", PCT),
    ("Terminal value (FY31 FCF x (1+g) / (WACC-g))", "=F4*(1+B10)/(B9-B10)", NUM),
    ("PV of terminal value", "=B11*F6", NUM),
    ("Sum PV of FCF (FY27-31)", "=SUM(B7:F7)", NUM),
    ("Enterprise value", "=B12+B13", NUM),
    ("  % of EV from terminal value", "=B12/B14", PCT),
    ("Plus: cash & investments (31-Jul-26)", f"={CASH0}", NUM),
    ("Less: ACV purchase price (cash out at close)", f"=-{ACV_PRICE}", NUM),
    ("Less: debt & minority interest", f"=-{DEBT0}-{MI0}", NUM),
    ("Equity value", "=SUM(B14,B16:B18)", NUM),
    ("Diluted shares today (M)", f"=Model!B{MR['Ending diluted shares (M)']}", NUM1),
    ("DCF value per share ($)", "=B19/B20", USD),
    ("Upside vs current", f"=B21/{PRICE}-1", PCT),
    ("Implied EV / FY27E EBITDA", f"=B14/Model!C{MR['EBITDA']}", MULT),
    ("Implied terminal EV / FY31 EBITDA", f"=B11/Model!G{MR['EBITDA']}", MULT),
]:
    put(d, r, 1, lab); put(d, r, 2, f, fmt); r += 1
# rows 9..24 used; put per-share result also at B30 for Valuation link
put(d, 30, 1, "DCF value per share (link)").font = BOLD; put(d, 30, 2, "=B21", USD)
d["A21"].font = BOLD; d["B21"].font = BOLD
r = 32
sub(d, r, "Sensitivity: DCF value per share vs WACC (rows) and terminal growth (cols)", 8); r += 1
gs = [0.02, 0.025, 0.03, 0.035, 0.04]; ws_ = [0.075, 0.08, 0.085, 0.09, 0.095, 0.10]
put(d, r, 1, "WACC \\ g").font = BOLD
for j, g in enumerate(gs):
    put(d, r, 2 + j, g, PCT).font = BOLD
r += 1
for w in ws_:
    put(d, r, 1, w, PCT).font = BOLD
    for j, g in enumerate(gs):
        gc = f"{L(2+j)}${r-1-ws_.index(w)}"; wc = f"$A{r}"
        pv = "+".join([f"{L(2+i)}$4/(1+{wc})^{L(2+i)}$5" for i in range(5)])
        tv = f"F$4*(1+{gc})/({wc}-{gc})/(1+{wc})^F$5"
        put(d, r, 2 + j, f"=({pv}+{tv}+$B$16+$B$17+$B$18)/$B$20", USD)
    r += 1
rng = f"B{r-len(ws_)}:{L(1+len(gs))}{r-1}"
d.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=[PRICE], fill=PatternFill("solid", fgColor="C6EFCE")))
d.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=[PRICE], fill=PatternFill("solid", fgColor="FFC7CE")))

# =====================================================================
# COMPS sheet
# =====================================================================
c = wb.create_sheet("Comps")
widths(c, [34, 12, 12, 12, 12, 12, 12, 12, 46])
hdr(c, 1, "COMPARABLE COMPANIES & CPRT HISTORY  (approx. public data as of 24-Sep-2026; verify before submission)", 9)
heads = ["Company", "Ticker", "EV ($B)", "EV/EBITDA (LTM)", "EV/EBITDA (NTM)", "P/E (NTM)", "EBIT margin", "Net debt/EBITDA", "Comment"]
for i, hd in enumerate(heads):
    put(c, 3, 1 + i, hd).font = BOLD
peers = [
    ("RB Global (owns IAA)", "RBA", 19.3, 14.8, 13.0, 17.9, 0.17, 1.5, "Direct salvage competitor via IAA; levered; margins ~half of CPRT; StockAnalysis/Finviz"),
    ("OPENLANE", "KAR", 2.9, 10.0, 8.5, 20.0, 0.10, 0.5, "Digital dealer wholesale (closest ACV comp); Reuters/TickerSpark"),
    ("Copart - current", "CPRT", f"={EV0}/1000", "=B10/1882*1000", 11.3, 17.7, 0.354, -2.3, "CapIQ export; net cash"),
    ("Copart - 2-yr avg (CQ4-24 to CQ3-26)", "CPRT", None, None, f"={HIST_EVE_AVG}", f"={HIST_PE_AVG}", 0.37, None, "CapIQ Multiples tab"),
    ("Copart - CQ4-2024 avg", "CPRT", None, None, 25.8, 36.1, 0.377, None, "Pre-derating level"),
]
r = 4
for p in peers:
    for i, val in enumerate(p):
        fmt = None
        if i == 2: fmt = '0.0'
        elif i in (3, 4, 5): fmt = MULT
        elif i == 6: fmt = PCT
        elif i == 7: fmt = '0.0"x"'
        if val is not None:
            put(c, r, 1 + i, val, fmt)
    r += 1
c["D6"] = f"={EV0}/1882"
put(c, 10, 1, "Peer median (RBA, KAR)").font = BOLD
for col in (4, 5, 6):
    put(c, 10, col, f"=MEDIAN({L(col)}4:{L(col)}5)", MULT)
r = 12
sub(c, r, "Implied CPRT value at peer / historical multiples (on FY27E base-case figures)", 9); r += 1
for i, hd in enumerate(["Basis", "Multiple", "Metric ($M / $)", "Implied EV", "Implied equity", "Per share", "vs current"]):
    put(c, r, 1 + i, hd).font = BOLD
r += 1
base_ebitda27 = f"Scenarios!$C${SR[2]['EBITDA']}"; base_eps27 = f"Scenarios!$C${SR[2]['Diluted EPS ($)']}"
netcash = f"({CASH0}-{ACV_PRICE}-{DEBT0}-{MI0})"
for lab, mult, metric, kind in [
    ("Peer median EV/NTM EBITDA", "=E10", base_ebitda27, "ev"),
    ("RBA EV/NTM EBITDA", "=E4", base_ebitda27, "ev"),
    ("CPRT 2-yr avg EV/NTM EBITDA", "=E7", base_ebitda27, "ev"),
    ("Peer median NTM P/E", "=F10", base_eps27, "pe"),
    ("CPRT 2-yr avg NTM P/E", "=F7", base_eps27, "pe"),
]:
    put(c, r, 1, lab); put(c, r, 2, mult, MULT); put(c, r, 3, f"={metric}", NUM1 if kind == "ev" else USD, font=GREEN)
    if kind == "ev":
        put(c, r, 4, f"=B{r}*C{r}", NUM); put(c, r, 5, f"=D{r}+{netcash}", NUM); put(c, r, 6, f"=E{r}/{SHARES}", USD)
    else:
        put(c, r, 4, ""); put(c, r, 5, ""); put(c, r, 6, f"=B{r}*C{r}", USD)
    put(c, r, 7, f"=F{r}/{PRICE}-1", PCT); r += 1
put(c, r + 1, 1, "Read-across: CPRT trades at a discount to RBA (IAA's parent) on P/E despite ~2x the EBIT margin, zero debt and $4.5B cash. "
    "The market is pricing structural share loss; the data (assignments +2.3% ex one customer) says cyclical.").font = Font(italic=True)

# =====================================================================
# ACV deal sheet
# =====================================================================
a = wb.create_sheet("ACV Deal")
widths(a, [50, 14, 14, 14, 14, 40])
hdr(a, 1, "ACV AUCTIONS ACQUISITION - DEAL MATH & ACCRETION (base case)", 6)
r = 3
for lab, f, fmt, note in [
    ("Purchase price ($M)", f"={ACV_PRICE}", NUM, "Cash tender at $10.50/sh"),
    ("ACV CY25 revenue", 760, NUM, "ACV 10-K"),
    ("ACV CY25 adj. EBITDA", 59, NUM, ""),
    ("ACV CY26E revenue (guide)", f"={ACV_REV0}", NUM, ""),
    ("ACV CY26E adj. EBITDA (guide mid)", 75, NUM, "$73-77M"),
    ("EV / CY25 revenue", "=B3/B4", MULT, ""),
    ("EV / CY25 adj. EBITDA", "=B3/B5", MULT, ""),
    ("EV / CY26E adj. EBITDA", "=B3/B7", MULT, ""),
    ("Foregone interest on cash (pre-tax, $M/yr)", f"=B3*{A['Yield on cash'][2][0]}", NUM1, "Cost of funding the deal from cash"),
    ("Deal amortization (pre-tax, $M/yr)", f"={ACV_AMORT}", NUM1, "Non-cash"),
    ("Pre-tax synergies required for GAAP EPS breakeven (FY28, ex-amort.)", f"=B11-Scenarios!D{SR[2]['ACV adj. EBITDA (standalone)']}-Scenarios!D{SR[2]['  less: ACV SBC']}-Scenarios!D{SR[2]['  less: ACV existing D&A']}", NUM1, "Foregone interest less ACV standalone EBIT (ex deal amort.)"),
    ("Base-case FY28 synergies assumed", f"={A['ACV pretax synergies ($M)'][2][1]}", NUM1, ""),
    ("FY28E GAAP EPS impact of ACV (base, $)", f"=(Scenarios!D{SR[2]['ACV EBIT contribution (GAAP)']}-B11)*(1-0.195)/Scenarios!D{SR[2]['Diluted shares (M, avg)']}", '$0.000', "Includes deal amortization and foregone interest"),
    ("FY28E cash EPS impact of ACV (base, $)", f"=(Scenarios!D{SR[2]['ACV EBIT contribution (GAAP)']}-Scenarios!D{SR[2]['  less: acquired-intangible amortization']}-B11)*(1-0.195)/Scenarios!D{SR[2]['Diluted shares (M, avg)']}", '$0.000', "Ex deal amortization"),
    ("Strategic logic", "Franchise-dealer access CPRT never cracked organically (20 yrs); 829k ACV units onto 21k acres/281 yards; buyer cross-sell; floorplan; vs. buyback at ~18x P/E = ~5.5% earnings yield", None, ""),
]:
    put(a, r, 1, lab); put(a, r, 2, f, fmt); put(a, r, 6, note).font = Font(italic=True, color="666666"); r += 1
a["B17"].alignment = Alignment(wrap_text=True)
r += 1
sub(a, r, "Synergy sensitivity: FY28E GAAP EPS impact vs pre-tax synergies ($M)", 6); r += 1
put(a, r, 1, "Synergies ->").font = BOLD
syn = [0, 25, 50, 75, 100, 150]
for j, sv in enumerate(syn):
    put(a, r, 2 + j, sv, NUM).font = BOLD
r += 1
put(a, r, 1, "FY28E GAAP EPS impact ($)")
for j in range(len(syn)):
    put(a, r, 2 + j, f"=(Scenarios!$D${SR[2]['ACV EBIT contribution (GAAP)']}-Scenarios!$D${SR[2]['  plus: synergies (pre-tax)']}+{L(2+j)}{r-1}-$B$11)*(1-0.195)/Scenarios!$D${SR[2]['Diluted shares (M, avg)']}", '$0.000')
widths(a, [50, 14, 14, 14, 14, 14, 14, 40])

# =====================================================================
# SUMMARY sheet (first)
# =====================================================================
sm = wb.create_sheet("Summary", 0)
widths(sm, [44, 16, 16, 16, 16, 40])
hdr(sm, 1, "COPART, INC. (NASDAQ: CPRT) - LONG  |  HFAC x Citadel 2026 Stock Pitch  |  Valuation model summary", 6)
r = 3
for lab, f, fmt in [
    ("Recommendation", "LONG", None),
    ("Current price (24-Sep-2026)", f"={PRICE}", USD),
    ("12-month probability-weighted target", f"={PWT}", USD),
    ("Expected return", f"={PWR}", PCT),
    ("Bear / Base / Bull target", f'=TEXT(Valuation!B{VR["12-month target price"]},"$0.00")&" / "&TEXT(Valuation!C{VR["12-month target price"]},"$0.00")&" / "&TEXT(Valuation!D{VR["12-month target price"]},"$0.00")', None),
    ("Bear / Base / Bull return", f'=TEXT(Valuation!B{VR["Upside / (downside) vs current"]},"0%")&" / "&TEXT(Valuation!C{VR["Upside / (downside) vs current"]},"0%")&" / "&TEXT(Valuation!D{VR["Upside / (downside) vs current"]},"0%")', None),
    ("Probabilities (Bear/Base/Bull)", '=TEXT(Inputs!B21,"0%")&" / "&TEXT(Inputs!C21,"0%")&" / "&TEXT(Inputs!D21,"0%")', None),
    ("Market cap ($M)", f"={MCAP}", NUM),
    ("Enterprise value ($M)", f"={EV0}", NUM),
    ("Current NTM P/E (consensus FY27E $1.63)", f"={PRICE}/1.63", MULT),
    ("Current EV / FY26A EBITDA", f"={EV0}/1882", MULT),
    ("DCF value per share (active scenario)", "=DCF!B21", USD),
    ("Active scenario", f"=Inputs!C17", None),
]:
    put(sm, r, 1, lab).font = BOLD; put(sm, r, 2, f, fmt); r += 1
r += 1
sub(sm, r, "Sheet guide", 6); r += 1
for t in [
    "Inputs      - all assumptions (blue), scenario selector (B17), probabilities, ACV deal terms, DCF inputs",
    "Historical  - FY21-FY26 financials (Capital IQ export) and CapIQ multiple history",
    "Model       - 5-year operating model for the ACTIVE scenario incl. ACV consolidation, buybacks, cash",
    "Scenarios   - Bear / Base / Bull models side by side (feeds Valuation)",
    "Valuation   - 12-month target price via P/E & EV/EBITDA per scenario, probability-weighted, P/E x EPS sensitivity",
    "DCF         - 5-yr DCF on active scenario, WACC x g sensitivity",
    "Comps       - RBA / KAR / CPRT history and implied values",
    "ACV Deal    - deal multiples, synergy breakeven, EPS accretion/dilution sensitivity",
    "Sources: Copart FY26 press release & 10-K, Capital IQ export (provided), ACV/Copart merger press release (10-Sep-2026), Undiscovered Compounders deep dives (provided), public quotes 24-Sep-2026.",
]:
    put(sm, r, 1, t); r += 1

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B3" if ws.title in ("Model", "Scenarios", "Historical") else None
wb.save(OUT)
print("saved", OUT)
