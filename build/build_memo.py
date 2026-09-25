"""Builds the 2-page CPRT investment memo (PDF) with reportlab + one matplotlib chart."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

OUT = "CPRT_Investment_Memo.pdf"
CHART = "chart_pe.png"

pdfmetrics.registerFont(TTFont("Lib", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("LibB", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("LibI", "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf"))
pdfmetrics.registerFont(TTFont("LibBI", "/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf"))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("Lib", normal="Lib", bold="LibB", italic="LibI", boldItalic="LibBI")

NAVY = colors.HexColor("#1F3864"); GREY = colors.HexColor("#555555"); LIGHT = colors.HexColor("#E8EEF7"); GREEN = colors.HexColor("#1E7B34"); RED = colors.HexColor("#B00020")
FS = 8.0; LEAD = 9.5
body = ParagraphStyle("b", fontName="Lib", fontSize=FS, leading=LEAD, alignment=TA_LEFT, spaceAfter=0)
bul = ParagraphStyle("bul", parent=body, leftIndent=8, bulletIndent=1, spaceAfter=0.6)
h1 = ParagraphStyle("h1", fontName="LibB", fontSize=9.6, leading=11.5, textColor=NAVY, spaceBefore=3.5, spaceAfter=1.5)
small = ParagraphStyle("s", parent=body, fontSize=6.6, leading=7.8, textColor=GREY)
cell = ParagraphStyle("c", parent=body, fontSize=7.5, leading=8.8)
cellb = ParagraphStyle("cb", parent=cell, fontName="LibB")
cellc = ParagraphStyle("cc", parent=cell, alignment=TA_CENTER)
cellcb = ParagraphStyle("ccb", parent=cellc, fontName="LibB")
title = ParagraphStyle("t", fontName="LibB", fontSize=15, leading=17, textColor=colors.white)
subt = ParagraphStyle("st", fontName="LibI", fontSize=8.6, leading=10, textColor=colors.white)
kpi_l = ParagraphStyle("kl", fontName="Lib", fontSize=6.6, leading=7.6, textColor=colors.white, alignment=TA_CENTER)
kpi_v = ParagraphStyle("kv", fontName="LibB", fontSize=9.6, leading=11, textColor=colors.white, alignment=TA_CENTER)

# ---------------- chart: multiple compression vs flat EPS ----------------
q = ["CQ4-24", "CQ1-25", "CQ2-25", "CQ3-25", "CQ4-25", "CQ1-26", "CQ2-26", "CQ3-26"]
pe = [36.1, 34.5, 33.7, 28.4, 24.6, 22.9, 20.0, 18.5]
eps = [1.49, 1.52, 1.56, 1.59, 1.60, 1.58, 1.56, 1.55]  # approx LTM diluted EPS path (FY25 $1.59 -> FY26 $1.55)
q[-1] = "CQ3-26*"
fig, ax = plt.subplots(figsize=(3.9, 2.05), dpi=220)
ax.bar(q, pe, color="#1F3864", width=0.62, label="P / NTM EPS (qtr avg, CapIQ)")
ax.set_ylim(0, 42); ax.set_ylabel("P/E (x)", fontsize=6.5); ax.tick_params(axis="both", labelsize=5.8)
for i, v in enumerate(pe):
    ax.text(i, v + 0.8, f"{v:.0f}x", ha="center", fontsize=5.6, color="#1F3864")
ax2 = ax.twinx(); ax2.plot(q, eps, color="#B00020", marker="o", ms=2.6, lw=1.4, label="LTM diluted EPS ($, approx.)")
ax.text(0.0, -0.30, "*CQ3-26 quarter-to-date (24-Sep-26). Source: S&amp;P Capital IQ", transform=ax.transAxes, fontsize=5.2, color="#555555")
ax2.set_ylim(0, 2.0); ax2.set_ylabel("EPS ($)", fontsize=6.5); ax2.tick_params(axis="y", labelsize=5.8)
ax.set_title("CPRT: -50% in multiple, EPS ~flat (36x -> 18x NTM P/E)", fontsize=7, color="#1F3864", pad=3)
h1_, l1 = ax.get_legend_handles_labels(); h2_, l2 = ax2.get_legend_handles_labels()
ax.legend(h1_ + h2_, l1 + l2, fontsize=5.4, loc="upper right", frameon=False)
for s in ("top",):
    ax.spines[s].set_visible(False); ax2.spines[s].set_visible(False)
plt.setp(ax.get_xticklabels(), rotation=30, ha="right")
fig.tight_layout(pad=0.3); fig.savefig(CHART, dpi=220); plt.close(fig)

# ---------------- document ----------------
doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=0.45 * inch, rightMargin=0.45 * inch, topMargin=0.34 * inch, bottomMargin=0.32 * inch,
                        title="LONG Copart (CPRT) - HFAC x Citadel 2026 Stock Pitch", author="Kushaagra Sood")
W = letter[0] - 0.9 * inch
story = []


def P(t, st=body): return Paragraph(t, st)


def B(t): return Paragraph(t, bul, bulletText="\u2022")


# Header banner
hdr_left = [P("LONG &nbsp;Copart, Inc. (NASDAQ: CPRT)", title),
            P("Total loss, total opportunity: the market is pricing a broken moat; the data shows a cyclical trough plus one customer", subt)]
kpis = [("Price (9/24/26)", "$28.04"), ("12-mo target (base)", "$41.00"), ("Upside (base)", "+46%"), ("Prob-wtd target", "$40 (+43%)"),
        ("Bear / Bull", "$21 / $58"), ("Mkt cap / EV", "$26.0B / $21.6B"), ("NTM P/E vs 2-yr avg", "17x vs 27x")]
kw = W * 0.60 - 8
kp = Table([[P(v, kpi_v) for _, v in kpis], [P(l, kpi_l) for l, _ in kpis]], colWidths=[kw * 0.135, kw * 0.14, kw * 0.12, kw * 0.16, kw * 0.14, kw * 0.165, kw * 0.14])
kp.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOTTOMPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                        ("LINEAFTER", (0, 0), (-2, -1), 0.4, colors.HexColor("#6F86B5"))]))
banner = Table([[hdr_left, kp]], colWidths=[W * 0.40, W * 0.60])
banner.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
story.append(banner)
story.append(P("Valuation date: 25-Sep-2026 &nbsp;|&nbsp; Horizon: 12 months (catalyst path 3-12 months) &nbsp;|&nbsp; Author: Kushaagra Sood &nbsp;|&nbsp; Model: CPRT_Valuation_Model.xlsx (scenarios, DCF, comps, ACV accretion)", small))

# Executive summary
story.append(P("Executive summary", h1))
story.append(P(
    "Copart is one of two players (with RB Global's IAA) in the US salvage-vehicle auction duopoly: insurers consign totaled cars, Copart tows, stores them on "
    "<b>281 yards / 21,000+ acres (>90% owned)</b> and auctions them online to a global buyer base (300k+ paying members). It earns ~85% of revenue from fees, converts 35% of revenue to EBIT and 27% to FCF, "
    "carries <b>no debt and $4.5B of cash</b>, and compounded revenue at 12% and EPS at 10% p.a. over FY21-26 before FY26's first EPS dip (-2.5%). The stock is -43% from its 52-week high and now trades at <b>17x NTM EPS and 11x EV/EBITDA</b>, in line with "
    "levered, half-the-margin RBA and below its 2-yr average of 27x / 19x. The bear narrative - structural share loss to IAA, broken margins, value-destructive M&amp;A - conflates three things: "
    "<b>(1) a single customer</b> (Progressive) moving low-value volume to IAA after Copart refused to cut fees, <b>(2) the soft point of the auto-insurance cycle</b>, and <b>(3) a $1.9B acquisition</b> (ACV) that is small (1.5x FCF) and not needed for the thesis. "
    "Q4 FY26 (10-Sep) gave the first hard datapoint: <b>US insurance assignments were +2.3% ex that one customer</b> while reported units were -7.5%. As Progressive's shift laps (Q1-Q2 FY27) and assignments convert to units (2-3 month lag), "
    "we expect US insurance units to return to growth by Q4 FY27, operating leverage to reverse (~$31M EBIT per point of unit growth), and the multiple to re-rate toward 24x. Base case <b>$41 (+46%)</b>; probability-weighted <b>$40 (+43%)</b>; DCF $36 without any re-rating."))

# Why the stock is down / variant view table
story.append(P("Where we differ from the market", h1))
rows = [[P("Market believes", cellb), P("We believe", cellb), P("Evidence", cellb)],
        [P("US insurance units -8% in FY26 = structural share loss to IAA; the duopoly is cracking", cell),
         P("It is one customer plus the cycle. Copart <i>chose</i> to lose low-ASP Progressive volume rather than cut fees; the rest of the book is growing", cell),
         P("Q4: domestic insurance assignments <b>+2.3% ex one customer</b> (reported -5%); IAA now serves its 'largest insurance partner' in all 50 states, i.e. the shift is largely done; RBA's take rate fell 21.4% to 20.0% in 3 quarters (IAA bought the volume with price) while Copart RPU rose +5.7% in FY26", cell)],
        [P("Margins are broken (operating margin 36.5% -> 35.4%; Q4 US operating margin 38.6% -> 33.6%)", cell),
         P("Deleverage on a fixed land base plus deliberate investment (Title Express, long-haul delivery, whole-car/wholesale capacity, G&amp;A +7%) - both reverse/pay back as units recover", cell),
         P("Precedent: after the only prior two weak years (FY13-14) management cut G&amp;A 15.5% in FY15; every 1pt of unit growth ~ $31M EBIT on a fixed cost base; international EBIT +8% at a 27% margin in FY26", cell)],
        [P("Total-loss tailwind is over; fewer accidents mean fewer salvage cars", cell),
         P("Frequency dips are cyclical; severity and total-loss frequency (TLF) keep rising as repair complexity grows - the structural volume driver is intact", cell),
         P("US auto premiums <b>-5.1% y/y (Aug-26)</b> after two years of double-digit hikes; GEICO claim frequency +2-7% across coverages (CQ1-26); TLF on a decade-long uptrend (CCC: 15.6% in 2015 to 20%+); insurers' salvage ASP at an all-time high (seasonally adj.) - more insured cars, more claims, more totals", cell)],
        [P("ACV ($1.9B, ~25x fwd EBITDA) shows capital indiscipline; buybacks halted", cell),
         P("Small, cash-funded, GAAP-dilutive by ~$0.10 (FY28E) but gives franchise-dealer access Copart never cracked in 20 years; thesis does not need it to work", cell),
         P("Deal = 7% of market cap; pro-forma cash still ~$2.6B + $1.25B revolver; management said buybacks paused <i>for</i> M&amp;A - resumption at ~17x P/E is itself a catalyst", cell)]]
t = Table(rows, colWidths=[W * 0.27, W * 0.33, W * 0.40])
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")),
                       ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                       ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
story.append(t)

# Thesis pillars + catalysts (two columns)
left = [P("Thesis", h1),
        B("<b>The moat is physical and un-replicable.</b> Permitted, drained, fenced storage acreage near population centers is the bottleneck in salvage; Copart controls 21,000+ acres (>19k in the US, >90% owned) carried at ~$127k/acre historical cost vs. ~$147k paid on recent purchases, in zoning that no new entrant can recreate. "
          "Two-sided liquidity (300k+ paying buyer members, exporters in Europe, Africa, LatAm, the Middle East) lifts insurers' recovery rates, which is why 27 of the top 30 carriers use Copart and split volume between only two vendors."),
        B("<b>Insurance cycle is turning in Copart's favor.</b> After two years of premium hikes, carriers earned record margins in 2025 and are now cutting rates (premiums -5.1% y/y in Aug-26; Berkshire warns GEICO's peers are pressing) and growing policies; claim frequency is rising again (+2-7% at GEICO) while the decade-long TLF uptrend continues. "
          "The carriers regaining share (GEICO, State Farm, Allstate) are the bulk of Copart's mix; their assignments are already growing (+2.3%) and convert to units with a 2-3 month lag."),
        B("<b>Progressive is a bounded, lapping headwind.</b> The shift equals ~9pts of US insurance assignments (Q4), most of it already migrated (IAA: 'all 50 states'). Copart kept pricing discipline (RPU +5.7% in FY26), so lost units were the lowest-fee ones. "
          "The comp turns easy from Q2 FY27 and mechanically flips positive by Q4 FY27."),
        B("<b>Operating leverage cuts both ways.</b> Land is paid for; incremental units cost only towing/handling and no customer-acquisition spend. Our base case takes core EBIT margin from 35.0% (FY27) back to 36.0% (FY28) - still below FY25's 37.7%."),
        B("<b>International is a second engine.</b> Units +10% and insurance units +11.2% in Q4; segment now 18% of revenue and 14% of EBIT at a 27% margin as UK/Canada/Germany convert insurers to consignment and Adair targets Europe expansion."),
        B("<b>Fortress balance sheet funds both offense and defense.</b> $4.5B cash, $1.27B FCF, zero debt: ACV is funded from cash, ~$2.6B remains, and the FY26 pace of buybacks ($1.6B) retired 4% of shares.")]
right = [P("Catalysts (3-12 months)", h1),
         B("<b>Nov-2026 (Q1 FY27 print):</b> Progressive impact peaks; management repeats the 'ex-one-customer' assignment metric it disclosed for the first time in Q4 - expect positive underlying growth vs. negative headline."),
         B("<b>Dec-2026:</b> ACV tender closes; synergy framework and PPA disclosed; overhang removed; buyback authorization resumes."),
         B("<b>Feb-2027 (Q2 FY27):</b> US insurance units improve to down low-single-digits as the shift laps; RPU growth keeps revenue flat-to-up; cost per unit stabilizes."),
         B("<b>May/Sep-2027 (Q3-Q4 FY27):</b> US insurance units return to growth; gross margin inflects; Street FY28 EPS ($1.77) is revised up and the multiple re-rates."),
         B("<b>Anytime:</b> CAT/hurricane season volume; Europe entry announcements; any indication that a top carrier is consolidating volume back to two vendors."),
         P("Why now", h1),
         P("Sentiment trough: -43% from the 52-week high ($49), 52-week low $26.81 (Jul-26), HSBC downgrade (16-Sep), Street mean PT ~$40 (range $25-55) with estimates being cut. Short-dated pain (Q1 FY27 will be ugly) is visible and priced; the inflection is 2-3 quarters out - "
           "inside the competition horizon and before consensus models it. Buyers at 17x get a 35%-margin, net-cash compounder at the multiple of a levered auctioneer.")]
two = Table([[left, right]], colWidths=[W * 0.55, W * 0.45])
two.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (0, 0), (0, 0), 6),
                         ("LEFTPADDING", (1, 0), (1, 0), 6), ("RIGHTPADDING", (1, 0), (1, 0), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                         ("LINEBEFORE", (1, 0), (1, 0), 0.4, colors.HexColor("#B7C4DC"))]))
story.append(two)

# ---------------- PAGE 2: Valuation ----------------
story.append(P("Valuation: re-rating on recovering earnings, with a DCF floor above the current price", h1))
vrows = [[P("", cell), P("Bear (25%)", cellcb), P("Base (50%)", cellcb), P("Bull (25%)", cellcb), P("Comment", cellb)],
         [P("US units FY27E / FY28E", cell), P("-3% / -1%", cellc), P("-1% / +3%", cellc), P("+1% / +4.5%", cellc), P("Q4 assignments +2.3% ex-PGR", cell)],
         [P("Intl units; RPU growth", cell), P("+2.5%; +1% then -1%", cellc), P("+5.5%; +3.5%", cellc), P("+8%; +5.5%", cellc), P("FY26: intl units +10%, RPU +5.7%", cell)],
         [P("Core EBIT margin FY27E / FY28E", cell), P("33.5% / 32.5%", cellc), P("35.0% / 36.0%", cellc), P("36.0% / 37.5%", cellc), P("FY26 35.4%; FY25 37.7% (CapIQ)", cell)],
         [P("ACV synergies FY28E (pre-tax)", cell), P("$0", cellc), P("$25M", cellc), P("$50M", cellc), P("Consolidated from Dec-26", cell)],
         [P("FY28E EPS (NTM at target date)", cellb), P("$1.31", cellcb), P("$1.72", cellcb), P("$1.97", cellcb), P("Consensus $1.77 (pre-ACV)", cell)],
         [P("Target NTM P/E  |  EV/EBITDA", cell), P("15x | 10x", cellc), P("24x | 15.5x", cellc), P("30x | 19x", cellc), P("Now 17x | 11x; 2-yr avg 27x | 19x", cell)],
         [P("12-month target (50/50 blend)", cellb), P("<font color='#B00020'>$21 (-25%)</font>", cellcb), P("<font color='#1E7B34'>$41 (+46%)</font>", cellcb), P("<font color='#1E7B34'>$58 (+105%)</font>", cellcb),
          P("<b>Prob-wtd $40 (+43%); up/down 4.2x</b>", cell)]]
vt = Table(vrows, colWidths=[W * 0.16, W * 0.09, W * 0.09, W * 0.095, W * 0.155])
vt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2),
                        ("BACKGROUND", (0, 7), (-1, 7), colors.HexColor("#F3F6FB"))]))
img = Image(CHART, width=W * 0.40, height=W * 0.40 * 2.05 / 3.9)
valblock = Table([[vt, img]], colWidths=[W * 0.60, W * 0.40])
valblock.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
story.append(valblock)
story.append(Spacer(0, 2))
story.append(P(
    "<b>Cross-checks.</b> (i) <b>DCF</b> (base case, WACC 8.7% = 4.2% rf + 1.0 beta x 4.5% ERP, terminal growth 3.5%, capex 8.5% of revenue vs. 7.2% in FY26, ACV cash out): <b>$36/share (+29%)</b> with 80% of EV in terminal value; "
    "the WACC/g grid spans $25-52 and only the corner with WACC >= 9.5% <i>and</i> terminal growth <= 3% falls below today's price. Implied terminal EV/EBITDA of 12.6x is below where the stock trades today. "
    "(ii) <b>Comps</b>: RBA (IAA parent) trades at 17.9x NTM P/E and ~13x NTM EBITDA with a 17% EBIT margin and 1.5x leverage; KAR at ~20x P/E. Copart at 17x with a 35% margin and net cash is the anomaly, not the peers. "
    "(iii) <b>History</b>: Copart traded at 30-40x NTM EPS for most of the last decade; our base 24x asks for a partial reversion only. "
    "(iv) <b>What is priced in</b>: at 17x, the market implies near-zero EPS growth; consensus already sits at FY27E $1.63 (+5%) and FY28E $1.77 (+9%). Our base FY28E $1.72 is <i>below</i> consensus because we include ACV's GAAP drag and lower interest income - the upside is multiple, not heroic estimates. "
    "<b>Sensitivity</b>: each 1x of P/E is ~$1.70; each $0.10 of FY28 EPS is ~$2.40 at 24x. Fully modeled in the accompanying Excel (Inputs!B17 toggles scenarios)."))

# Risks
story.append(P("Key risks and mitigants", h1))
rrows = [[P("Risk", cellb), P("Assessment / mitigant", cellb), P("Bear case impact", cellb)],
         [P("<b>Another top carrier defects</b> (GEICO/State Farm/Allstate concentration; IAA buying share with fee cuts)", cell),
          P("Progressive was unique: a low-ASP, direct-writer mix that Copart priced to lose. Carriers split volume for redundancy and CAT capacity - Copart absorbed $56M of Helene/Milton costs to prove it. Monitor: any 'customer' language on calls", cell),
          P("Captured in bear (US units -3%/-1%, margin to 32.5%): $21, -25%", cellc)],
         [P("<b>Cycle does not turn / used-car ASP deflation</b> lowers RPU and fee-grid revenue", cell),
          P("TLF and severity are structural (electronics content); ASP fell 2022-23 and Copart still grew EPS 13%. Bear case assumes RPU -1% for two years", cell),
          P("Included in bear", cellc)],
         [P("<b>ACV integration / more M&amp;A</b> (Adair: 'I would take on some debt'; CCC interest reported)", cell),
          P("ACV is 7% of market cap and modeled dilutive; a >$5B levered deal would be a thesis-breaker and is a monitoring item. Board and 20-year record of disciplined tuck-ins", cell),
          P("Not in target; -$0.10 EPS already modeled", cellc)],
         [P("<b>Autonomy/ADAS reduces accidents</b> long-term", cell),
          P("Outside 12-month horizon; ADAS raises repair cost and TLF near-term (more totals per accident) and equips only ~7% of the fleet; average US vehicle age ~12.8 years", cell), P("Negligible in horizon", cellc)],
         [P("<b>Execution/communication</b>: CEO re-transition, sparse disclosure, Q1 FY27 optics", cell),
          P("Q1 will show US insurance units down mid-single digits - known and modeled (FY27E EPS $1.51 < consensus). Position sizing: build through Q1 print", cell), P("Timing risk only", cellc)]]
rt = Table(rrows, colWidths=[W * 0.30, W * 0.50, W * 0.20])
rt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2)]))
story.append(rt)

story.append(P("What would make us wrong (thesis breakers we will monitor each quarter)", h1))
story.append(P("(1) US insurance assignments ex-Progressive turn negative for two consecutive quarters; (2) RPU growth falls below ASP growth (fee pricing power eroding); "
               "(3) a second named customer loss; (4) a >$5B debt-funded acquisition; (5) core EBIT margin below 33% while units are growing (cost inflation, not deleverage). Any of these moves us to the bear case and we exit."))

story.append(P("Appendix: key financials ($M, FY ends Jul-31; historicals per S&amp;P Capital IQ standardized)", h1))
frows = [["", "FY23A", "FY24A", "FY25A", "FY26A", "FY27E base", "FY28E base"],
         ["Revenue (core)", "3,870", "4,237", "4,647", "4,666", "4,839", "5,183"],
         ["  growth", "10.5%", "9.5%", "9.7%", "0.4%", "3.7%", "7.1%"],
         ["Core EBIT / margin", "1,487 / 38.4%", "1,573 / 37.1%", "1,753 / 37.7%", "1,653 / 35.4%", "1,694 / 35.0%", "1,866 / 36.0%"],
         ["ACV EBIT contribution (GAAP, incl. amort.)", "-", "-", "-", "-", "(51)", "(37)"],
         ["Diluted EPS (GAAP)", "$1.28", "$1.40", "$1.59", "$1.55", "$1.51", "$1.72"],
         ["FCF (CFO - capex) / Unlevered FCF (E)", "848", "962", "1,231", "1,267", "1,184", "1,356"],
         ["Cash &amp; investments (year-end) / Debt", "2,364 / 0", "3,422 / 0", "4,789 / 0", "4,490 / 0", "3,337 / 0", "4,074 / 0"],
         ["Diluted shares (M, avg) / buybacks ($M)", "967 / 5", "975 / 7", "978 / 5", "957 / 1,636", "920 / 500", "905 / 700"]]
ft = Table([[P(c, cellb if j == 0 else cellc) if i == 0 else P(c, cell if j == 0 else cellc) for j, c in enumerate(r)] for i, r in enumerate(frows)],
           colWidths=[W * 0.28] + [W * 0.12] * 6)
ft.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 0.8), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8)]))
story.append(ft)
story.append(Spacer(0, 1.5))
# sensitivity + ACV math side by side
sens_pe = [16, 18, 20, 22, 24, 26]; sens_eps = [1.55, 1.65, 1.72, 1.80, 1.90]
sc_ = ParagraphStyle("sc", parent=cellc, fontSize=6.9, leading=8)
srows = [[P("P/E \\ EPS", cellb)] + [P(f"${e:.2f}", cellcb) for e in sens_eps]]
for pe_ in sens_pe:
    row = [P(f"{pe_}x", cellcb)]
    for e in sens_eps:
        px = pe_ * e; up = px / 28.04 - 1
        row.append(P(f"<font color='{'#1E7B34' if up > 0 else '#B00020'}'>{px:.0f} ({up:+.0%})</font>", sc_))
    srows.append(row)
st_ = Table(srows, colWidths=[W * 0.47 * 0.19] + [W * 0.47 * 0.162] * 5)
st_.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("BACKGROUND", (0, 0), (0, -1), LIGHT), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")),
                         ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 0.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.5),
                         ("BOX", (3, 5), (3, 5), 1.0, NAVY)]))
acv_rows = [[P("ACV acquisition ($M)", cellb), P("", cellb)],
            [P("Cash price ($10.50/sh); % mkt cap; x FY26 FCF", cell), P("1,900; 7%; 1.5x", cellc)],
            [P("EV / CY26E revenue | adj. EBITDA ($850M | $75M guide)", cell), P("2.2x | ~25x", cellc)],
            [P("Foregone interest | deal amortization (pre-tax, /yr)", cell), P("~72 | ~57 (30% PPA, 10 yrs)", cellc)],
            [P("FY28E GAAP | cash EPS impact (base, $25M synergies)", cell), P("-$0.10 | -$0.05", cellc)],
            [P("Pre-tax synergies for FY28 cash-EPS breakeven", cell), P("~78 (~9% of ACV sales)", cellc)],
            [P("829k dealer units/yr onto 281 yards; franchise-dealer channel; logistics/floorplan cross-sell", cell), P("Option value; not in target", cellc)]]
at_ = Table(acv_rows, colWidths=[W * 0.52 * 0.64, W * 0.52 * 0.36])
at_.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("SPAN", (0, 0), (-1, 0)), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C4DC")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                         ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 0.8), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8)]))
bottom = Table([[[P("12-mo price ($) sensitivity: P/E x FY28E EPS (base boxed)", cellb), st_], [at_]]], colWidths=[W * 0.48, W * 0.52])
bottom.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (0, 0), 6), ("RIGHTPADDING", (1, 0), (1, 0), 0),
                            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
story.append(bottom)
story.append(Spacer(0, 2))
story.append(P("Sources: Copart FY26 earnings release &amp; Q4 FY26 call (10-Sep-2026); Copart 10-K FY25; S&amp;P Capital IQ (financials, consensus, multiples, 24-Sep-2026); Copart/ACV merger announcement &amp; SC TO-T (Sep-2026); RB Global Q2-26 call; "
               "CCC Crash Course / industry TLF data; GEICO/Berkshire 2025 annual report; Undiscovered Compounders Copart deep dive, Q4 FY26 update and ACV analysis (Aug-Sep 2026); public price data (24-Sep-2026). "
               "FY27E/FY28E figures are the author's base case (see model); consensus figures are CapIQ mean estimates.", small))

doc.build(story)
print("built", OUT)
