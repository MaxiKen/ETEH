"""Capstone edition-3 slide copy: Module P.3, Capstone 1 (SuperMart retail margin review). Figures from w5_fx, i.e. from
the reference analysis behind Capstone1_SuperMart_Stores_SOLUTION.xlsx."""
from slidekit import GREEN, NAVY, ORANGE, RED
from v3kit import T
from w5_fx import (ALL_HI, ALL_LO, ATTACH, BASK, BD25, BLANK, CAP, CAP_M, CLEAN, CTRL, CUST25, DEEP, DEPTH24, DEPTH25, DIRTY,
                   DLOG, DM, DUPS, FURN, FURN_DEEP, G25, K30, LG, LINES, LOSS_ST, LOSSVAL, M24, M25, MIX, MIXE, N_CAP, N_DEEP,
                   N_DROPPED, N_LOSSSKU, N_ON_TGT, N_UNSOLD, NSTORES, NZONES, O25, ORD25, ORDERS_ALL, P24, P25, QD, QM, RATE,
                   S24, S25, SKUS, SOLD25, SUBCATS, SUBL, TAIL30, TARGET, WHATIF, I, N, Nm, P, pts)

FM = '₦#,##0,,"m"'
QS = ["Q1", "Q2", "Q3", "Q4"]
_lg = LG.sort_values("att")
_short = [n for n in _lg.Store_Name]
_lines = O25.Lines.clip(upper=5).value_counts().sort_index()
DASH = dict(kind="dashboard", full=True, h=4.95, banner="SuperMart FY2025: sales up, margin down",
            kpis=[("Revenue", Nm(S25), f"{P(G25)} vs FY2024"), ("Margin", P(M25), pts(DM, 1).replace(" pts", " points"), RED),
                  ("Orders", I(ORD25), f"basket {Nm(BASK['mean_val'], 2)}"), ("Target", P(S25 / TARGET), f"{N_ON_TGT} of {NSTORES} stores on target"),
                  ("Loss-making stores", str(len(LOSS_ST)), Nm(LOSS_ST.P25.sum()), RED)],
            charts=[dict(kind="column", cats=QS, series=[("2024", [QM[(2024, q)] for q in QS]), ("2025", [QM[(2025, q)] for q in QS])],
                         fmt="0%", colors=["B8C2CC", NAVY], title="Margin by quarter"),
                    dict(kind="bar", cats=_short, series=[("Attainment", _lg.att.tolist())], fmt="0%", legend=False,
                         point_colors=[RED if a < 1 else NAVY for a in _lg.att], title="Target attainment by store"),
                    dict(kind="column", cats=list(BD25), series=[("Profit", [v[1] for v in BD25.values()])], fmt=FM, legend=False,
                         point_colors=[GREEN if v[1] > 0 else RED for v in BD25.values()], title="Profit by discount band")],
            widths=[1, 1.15, 1])

M_P3 = dict(
    id="P.3", title="Capstone 1: SuperMart Retail Analytics",
    subtitle="A complete retail build on the store ledger: brief, data, cleaning, model, baskets, margin erosion, slow "
             "movers, dashboard, memo and verification.",
    objectives=["Turn the board's question into success criteria.", "Inventory the sources by grain and clean the returns.",
                "Model the measures and analyse baskets.", "Explain the margin fall with discount bands and rate/mix.",
                "Build the dashboard, write the memo and verify the numbers."],
    topics=[
        T("P.3.1", "Capstone 1 Brief and Success Criteria",
          defn="Capstone 1 answers the board's question: FY2025 sales grew, so why did margin fall, which stores and products "
               "caused it, and what should change in FY2026?",
          simple="Success means reconciled numbers, a measured cause and a costed recommendation.",
          grid=dict(label="The headline numbers", widths=[1.6, 1.5, 1.5], size=12.5, rows=[
              ["Measure", "FY2024", "FY2025"], ["Sales", N(S24), N(S25)], ["Profit", N(P24), N(P25)],
              ["Margin", P(M24, 2), P(M25, 2)], ["Change", "", f"Sales {P(G25)}; margin {pts(DM)}"]]),
          chart=dict(kind="combo", cats=["FY2024", "FY2025"], series=[("Sales", [S24, S25]), ("Margin", [M24, M25])], fmt=FM,
                     y2_fmt="0.0%", title="Sales up, margin down", h=2.5, gap=80),
          order=["grid", "chart"], split=0.5),

        T("P.3.2", "SuperMart Data Inventory",
          defn="A data inventory lists each source with its grain (what one row means), period and size.",
          simple="Knowing the grain stops you calling sales lines 'orders'.",
          grid=dict(label="Capstone 1 inventory", widths=[1.7, 1.4, 3.0], size=12, rows=[
              ["Source", "Grain", "Scale"],
              ["POS ledger", "One sales line", f"{I(LINES)} lines; {I(ORDERS_ALL)} orders; 2022-2025"],
              ["Products (from ledger)", "One SKU", f"{I(SKUS)} SKUs; 3 categories; {SUBCATS} sub-categories"],
              ["Store directory", "One store", f"{NSTORES} stores; {NZONES} zones; FY2025 targets"],
              ["Store returns (dirty)", "One sales line", f"{I(len(DIRTY))} rows › {I(CLEAN)} after cleaning"]]),
          watch=f"FY2025 has {I(CLEAN)} lines but only {I(ORD25)} orders: count orders with DISTINCTCOUNT(Order_ID)."),

        T("P.3.3", "Cleaning the POS Extract",
          defn=f"The dirty store returns carry {len(DLOG)} kinds of defect. Fix each with a named step and log the row count.",
          grid=dict(label="Part of the defect log: SuperMart_Sales_2025_DIRTY.csv", widths=[2.4, 0.8, 2.6], size=11.5, rows=[
              ["Defect", "Rows", "Fix"],
              ["Mixed date formats", I(DLOG["Mixed date formats"]), "Change Type with Locale (en-GB)"],
              ["Sales typed as '₦1,234,560'", I(DLOG["Numbers stored as text"]), "Remove ₦ and commas, then number"],
              ["Discount typed as '20%'", I(DLOG["Percent as text"]), "Convert to 0.2"],
              ["Sales × 100 keying errors", I(DLOG["Keying outliers"]), "Correct from the receipt"],
              ["Duplicates and spacer rows", f"{DUPS} + {BLANK}", f"Remove: {I(len(DIRTY))} › {I(CLEAN)}"]]),
          watch=f"Remove duplicates on all columns, not Order_ID alone: one order has several lines. After cleaning, orders "
                f"({I(CTRL['Distinct orders'])}) and customers ({I(CUST25)}) match the controls."),

        T("P.3.4", "Modelling Retail Sales",
          defn="The model links sales lines to the store table on Store_ID and uses explicit DAX measures.",
          fx=[("SUM ( FactSales[Profit] )", N(P25), "Profit"),
              ("DIVIDE ( [Profit], [Revenue] )", P(M25, 2), "Margin %"),
              ("DIVIDE ( SUMX ( FactSales, FactSales[Discount] * FactSales[Sales] ), [Revenue] )", P(DEPTH25), "Discount depth"),
              ("DIVIDE ( [Revenue], SUM ( DimStore[Annual_Target_2025] ) )", P(S25 / TARGET), "Target attainment")],
          fx_head=["DAX measure", "FY2025", "Meaning"], fx_widths=[4.4, 1.2, 1.4], stack=True,
          watch="Discount depth is sales-weighted. A plain AVERAGE(Discount) treats a ₦500 line like a ₦5m line."),

        T("P.3.5", "Basket Analysis",
          defn="Basket analysis looks at whole orders: how big a typical order is and which categories are bought together.",
          simple="Add the lines up to order level first, then analyse the orders.",
          grid=dict(label=f"FY2025 baskets ({I(ORD25)} orders)", widths=[2.6, 1.3], size=12, rows=[
              ["Measure", "Value"], ["Lines per order: mean / most common", f"{BASK['mean_lines']:.2f} / {BASK['mode_lines']}"],
              ["Basket value: mean", N(BASK["mean_val"])], ["Basket value: median", N(BASK["med_val"])],
              ["Orders with 2+ categories", P(BASK["multi"])],
              ["Technology baskets with Office Supplies", P(ATTACH["Technology"][1]["Office Supplies"])]]),
          chart=dict(kind="column", cats=["1", "2", "3", "4", "5+"], series=[("Orders", _lines.tolist())], fmt="#,##0", labels=True,
                     legend=False, title="Most orders have one line", x_title="Lines per order", h=2.6),
          order=["grid", "chart"], split=0.52,
          watch=f"The mean basket is {BASK['mean_val'] / BASK['med_val']:.1f}× the median: a few large furniture and technology "
                "orders pull it up."),

        T("P.3.6", "Margin Erosion Analysis",
          defn="Explain a margin fall by discount depth and by splitting it into a rate effect (margins fell inside categories) "
               "and a mix effect (sales moved between categories).",
          grid=dict(label="FY2025 by discount band", widths=[1.2, 1.2, 1.1, 1.3], size=12, rows=[
              ["Band", "Share of sales", "Margin", "Profit"]] + [
              [k, P(v[0] / S25), P(v[1] / v[0]), Nm(v[1])] for k, v in BD25.items()]),
          chart=dict(kind="column", cats=QS, series=[("2024", [QM[(2024, q)] for q in QS]), ("2025", [QM[(2025, q)] for q in QS])],
                     fmt="0%", labels=[3], label_fmt="0.0%", colors=["B8C2CC", NAVY], h=2.5,
                     title=f"Q4 margin fell to {P(QM[(2025, 'Q4')])} (discount depth {P(QD[(2025, 'Q4')])})"),
          order=["grid", "chart"], split=0.5,
          watch=f"The {I(N_DEEP)} lines discounted above 20% lost {Nm(-DEEP)}. Rate effect {pts(RATE)}, mix effect {pts(MIXE)}: "
                "the fall is pricing, not product mix.",
          more=[dict(title="Rate and Mix Effects by Category", pill="RATE AND MIX",
                     chart=dict(kind="waterfall", full=True, h=4.4, start=M24 * 100,
                                cats=["FY2024"] + [f"{c} rate" for c in MIX.index] + [f"{c} mix" for c in MIX.index] + ["FY2025"],
                                steps=[float(MIX.loc[c].rate) for c in MIX.index] + [float(MIX.loc[c].mix) for c in MIX.index],
                                fmt="0.00", label=f"FROM {P(M24)} TO {P(M25)}, IN MARGIN POINTS",
                                title=f"Margin bridge FY2024 › FY2025: rate {pts(RATE)}, mix {pts(MIXE)}",
                                note="Rate = share 2025 × (margin 2025 − margin 2024); mix = (share 2025 − share 2024) × margin 2024. "
                                     "The six bars add up to the total change."),
                     notes="Walk the bridge left to right. Office Supplies and Furniture margins fell inside the category; the mix "
                           "moved towards Office Supplies, which helped. Pricing, not product mix, explains the fall.")]),

        T("P.3.7", "Stock and Slow-Mover Analysis",
          defn="Slow movers are products that sell little or nothing. Compare the product list with FY2025 sales using an anti "
               "join: items in one list with no match in the other.",
          grid=dict(label="SKU actions", widths=[3.0, 1.1], size=12.5, rows=[
              ["Measure", "Value"], ["SKUs in the range (2022-2025)", I(SKUS)], ["No sale in FY2025", I(N_UNSOLD)],
              ["Sold in FY2024 but not FY2025 (anti join)", I(N_DROPPED)], ["Loss-making SKUs in FY2025", I(N_LOSSSKU)],
              ["Their combined loss", Nm(-LOSSVAL)],
              [f"Revenue share of the bottom 30% ({I(K30)} SKUs)", P(TAIL30)]]),
          chart=dict(kind="bar", cats=SUBL.Sub_Category.head(5).tolist(), series=[("Profit", SUBL.Profit.head(5).tolist())], fmt=FM,
                     labels=True, label_fmt='₦#,##0.0,,"m"', legend=False, colors=[RED],
                     title="The five weakest sub-categories, FY2025 profit", h=2.6),
          order=["grid", "chart"], split=0.5),

        T("P.3.8", "Building the Retail Dashboard",
          defn="The Capstone 1 dashboard fits one screen: KPI cards, a store ranking, margin by quarter and profit by discount "
               "band, with zone, category and segment slicers on every visual.",
          grid=dict(label="Stores that lost money in FY2025", widths=[2.4, 1.4, 1.3], size=12, rows=[
              ["Store", "FY2025 profit", "Attainment"]] + [
              [r.Store_Name, N(r.P25).replace("₦-", "−₦"), P(r.S25 / r.T25)] for r in LOSS_ST.sort_values("P25").itertuples()]),
          watch="Hitting the sales target is not the same as making money: some loss-makers were on target.",
          more=[dict(title="The Capstone 1 Dashboard", pill="DASHBOARD", chart=DASH,
                     notes="One screen: five KPI cards and three visuals, each answering a question in the brief. The red figures "
                           "are the problems the memo addresses.")]),

        T("P.3.9", "SuperMart Insight Memo",
          defn="The insight memo fits one page: a headline, three supporting numbers, and the recommendation with an owner and "
               "an expected effect shown as a range.",
          grid=dict(label=f"Cap Furniture discounts at {P(CAP, 0)}: extra FY profit", widths=[2.4, 1.4], size=12.5, rows=[
              ["Volume lost on re-priced lines", "Profit gain"]] + [[P(k, 0), Nm(v)] for k, v in WHATIF.items()]),
          points=[f"Furniture margin fell to {P(FURN.mg25)}; {I(N_CAP)} of its lines were discounted above {P(CAP, 0)}",
                  f"Those lines lost {Nm(-FURN_DEEP.Profit.sum())}",
                  f"Chain-wide cap (phase 2): {Nm(ALL_LO)}-{Nm(ALL_HI)}"],
          points_label="SUPPORTING NUMBERS", order=["grid", "points"], split=0.48,
          watch="Show the effect as a range with its assumption (volume lost), never a single confident number."),

        T("P.3.10", "SuperMart Verification and Handover",
          defn="Before handover, prove the headline numbers: sales and profit reconcile to the ledger, row counts are "
               "documented, slicers are tested and the file opens on a second machine.",
          grid=dict(label="Verification checks", widths=[2.4, 2.4, 0.8], size=12, rows=[
              ["Check", "Evidence", "Status"], ["Revenue = POS ledger", N(S25), "Pass"], ["Profit = ledger", N(P25), "Pass"],
              ["Row counts documented", f"{I(CLEAN)} FY2025 lines; {I(LINES)} total", "Pass"],
              ["Slicers tested", "3 slicers × every visual", "Pass"],
              ["Second machine", "Charts and ₦ symbol intact", "Pass"]]),
          watch=f"The what-if margin with the cap ({P(CAP_M)} at 25% volume loss) is an estimate: label it so in the handover."),
    ],
    takeaways=["Start from the board's question and the headline numbers.", "Inventory by grain; clean to the control totals.",
               "Measures: margin = total profit ÷ total sales; depth is sales-weighted.",
               "Deep discounts and Q4 pricing explain the margin fall, not mix.",
               "Recommend with an owner and a range; verify before handover."],
)

MODULES = [M_P3]
