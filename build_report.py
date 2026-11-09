"""Build Kirana_Store_Financial_Report.xlsx from the kirana SQL project's database.

Usage:  python build_report.py
Every number on the analysis sheets is a live Excel formula that traces back
to the Data and Assumptions sheets.
"""
import sqlite3
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

DB = Path("/home/claude/store_db/store.db")
OUT = Path("/home/claude/finance_report/Kirana_Store_Financial_Report.xlsx")

# ------------------------------------------------------------------ data from SQL
con = sqlite3.connect(DB)
q = lambda s: pd.read_sql_query(s, con)
months = q("SELECT DISTINCT month FROM v_sales ORDER BY month").month.tolist()
labels = [pd.Timestamp(m + "-01").strftime("%b-%y") for m in months]
days = [pd.Timestamp(m + "-01").days_in_month for m in months]
sales = q("SELECT month, SUM(revenue) rev, SUM(cost) cogs, COUNT(DISTINCT order_id) bills FROM v_sales GROUP BY month").set_index("month")
purch = q("""SELECT strftime('%Y-%m', po.received_date) month, SUM(pi.quantity * pi.unit_cost) p
             FROM purchase_orders po JOIN purchase_items pi USING (po_id)
             WHERE po.status = 'Received' GROUP BY month""").set_index("month").p
inv = q("""SELECT strftime('%Y-%m', snapshot_date) month, SUM(closing_stock * cost_price) v
           FROM inventory_daily i JOIN products p USING (product_id)
           WHERE snapshot_date IN (SELECT MAX(snapshot_date) FROM inventory_daily GROUP BY strftime('%Y-%m', snapshot_date))
           GROUP BY month""").set_index("month").v
khata = q("""SELECT o_month month, SUM(CASE WHEN pm = 'Credit (khata)' THEN r END) / SUM(r) k FROM
             (SELECT strftime('%Y-%m', o.order_date) o_month, o.payment_mode pm, v.revenue r
              FROM orders o JOIN v_sales v USING (order_id)) GROUP BY o_month""").set_index("month").k
exp = q("SELECT strftime('%Y-%m', expense_date) month, category, SUM(amount) a FROM expenses GROUP BY 1, 2") \
    .pivot(index="category", columns="month", values="a").fillna(0)
cat = q("SELECT category, month, SUM(revenue) r FROM v_sales GROUP BY 1, 2").pivot(index="category", columns="month", values="r").fillna(0)
cat = cat.loc[cat.sum(axis=1).sort_values(ascending=False).index]
open_inv = q("""SELECT SUM(opening_stock * cost_price) v FROM inventory_daily i JOIN products p USING (product_id)
                WHERE snapshot_date = (SELECT MIN(snapshot_date) FROM inventory_daily)""").v[0]

OPEX = ["Rent", "Salaries", "Electricity", "Transport", "Fuel", "Packaging", "Maintenance", "Wastage & Expiry",
        "Internet & Phone", "Software & Fees", "Licences & Taxes", "Marketing", "Miscellaneous"]
OPEX_BUDGET = [30000, 32000, 7000, 6000, 4500, 2500, 3000, 1500, 1597, 649, 1200, 1000, 2000]
FIXED = {"Rent", "Salaries", "Internet & Phone", "Software & Fees", "Licences & Taxes"}

# ------------------------------------------------------------------ styles
F = "Arial"
NAVY, LIGHT, GREYF = "1F3A5F", "EEF2F6", "F4F4F4"
f_in, f_calc, f_link = Font(name=F, color="0000FF", size=10), Font(name=F, size=10), Font(name=F, color="008000", size=10)
f_bold, f_head = Font(name=F, bold=True, size=10), Font(name=F, bold=True, color="FFFFFF", size=10)
f_sec = Font(name=F, bold=True, size=11, color=NAVY)
fill_head, fill_tot, fill_yel = PatternFill("solid", fgColor=NAVY), PatternFill("solid", fgColor=LIGHT), PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="D0D7DE")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
top_line = Border(top=Side(style="thin", color="000000"))
RS = '"₹"#,##0;("₹"#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
MC = [L(c) for c in range(3, 15)]          # month columns C..N
FY = "O"


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(name=F, bold=True, size=15, color=NAVY)
    ws["A2"] = sub
    ws["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    ws.sheet_view.showGridLines = False


def month_header(ws, row, first="Line item", fy=True, link=True):
    cells = [first] + [""] + [f"=Data!{c}$4" if link else lab for c, lab in zip(MC, labels)] + (["FY total"] if fy else [])
    for i, v in enumerate(cells, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font, c.fill, c.border = f_head, fill_head, box
        c.alignment = Alignment(horizontal="center" if i > 2 else "left")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 3
    for c in MC + [FY]:
        ws.column_dimensions[c].width = 11.5


def row_formula(ws, r, label, fn, fmt=RS, font=f_calc, fy="sum", bold=False, total=False):
    ws.cell(row=r, column=1, value=label).font = f_bold if bold else Font(name=F, size=10)
    for i, c in enumerate(MC):
        cell = ws[f"{c}{r}"]
        cell.value = fn(c, i)
        cell.number_format, cell.font = fmt, Font(name=F, size=10, bold=bold, color=font.color)
    if fy == "sum":
        ws[f"{FY}{r}"] = f"=SUM(C{r}:N{r})"
    elif fy:
        ws[f"{FY}{r}"] = fy
    if fy:
        ws[f"{FY}{r}"].number_format, ws[f"{FY}{r}"].font = fmt, Font(name=F, size=10, bold=True)
    if total:
        for col in ["A"] + MC + [FY]:
            ws[f"{col}{r}"].fill, ws[f"{col}{r}"].border = fill_tot, top_line


def section(ws, r, text):
    ws.cell(row=r, column=1, value=text).font = f_sec


wb = Workbook()

# ================================================================== Guide
g = wb.active
g.title = "Guide"
title(g, "Kirana Store: Monthly Financial Report (FY Nov 2025 - Oct 2026)",
      "Gupta General Store, Laxmi Nagar, Delhi (hypothetical). Built from the SQL database in my kirana project. Data is simulated.")
guide = [
    ("What this workbook does", None),
    ("Dashboard", "Pick a month in the yellow cell: key figures, charts and how the month compares with budget."),
    ("P&L", "Monthly profit and loss: revenue, cost of goods, gross profit, every operating expense, operating profit."),
    ("Budget vs Actual", "Revenue, gross profit, expenses and profit against budget, with favourable / (adverse) variances."),
    ("Cash Flow", "Where cash came from and went each month, including credit (khata) sales collected a month later."),
    ("Ratios", "Margins, inventory turnover, days of stock, average bill, break-even sales and margin of safety."),
    ("What-If", "Change price, volume, supplier costs, rent, salaries or wastage and see the effect on annual profit."),
    ("Assumptions", "Opening balances, budget and what-if levers. The only sheet you need to edit."),
    ("Data", "Monthly figures pulled from the SQL database (store.db) with queries; inputs only."),
    ("", ""),
    ("Colour code", None),
    ("Blue text", "Hard-coded inputs (Data and Assumptions sheets)."),
    ("Black text", "Formulas calculated on the same sheet."),
    ("Green text", "Values linked from another sheet."),
    ("Yellow cell", "Key assumption or selector you are meant to change."),
    ("", ""),
    ("Notes", None),
    ("Accounting basis", "GST and income tax are excluded. Equipment is capitalised and depreciated straight-line (life on Assumptions)."),
    ("Negative numbers", "Shown in brackets, e.g. (₹5,000). A dash means zero."),
]
for i, (a, b) in enumerate(guide, 4):
    if b is None:
        g.cell(row=i, column=1, value=a).font = f_sec
        continue
    ca = g.cell(row=i, column=1, value=a)
    ca.font = Font(name=F, bold=True, size=10, color={"Blue text": "0000FF", "Green text": "008000"}.get(a, "000000"))
    if a == "Yellow cell":
        ca.fill = fill_yel
    cb = g.cell(row=i, column=2, value=b)
    cb.font, cb.alignment = Font(name=F, size=10), Alignment(wrap_text=True, vertical="top")
g.column_dimensions["A"].width = 22
g.column_dimensions["B"].width = 110

# ================================================================== Data
d = wb.create_sheet("Data")
title(d, "Data: monthly figures from the SQL database",
      "Source: store.db (kirana SQL project), queries on v_sales, purchase_orders, inventory_daily, orders and expenses. Simulated data. Rs.")
month_header(d, 4, fy=False, link=False)
d.cell(row=4, column=15, value="Source").font = f_bold
data_rows = [
    (5, "Revenue", [sales.rev[m] for m in months], RS, "v_sales: SUM(revenue)"),
    (6, "Cost of goods sold", [sales.cogs[m] for m in months], RS, "v_sales: SUM(cost)"),
    (7, "Bills (count)", [sales.bills[m] for m in months], "#,##0", "COUNT(DISTINCT order_id)"),
    (8, "Stock purchases received", [purch.get(m, 0) for m in months], RS, "purchase_orders x purchase_items"),
    (9, "Closing inventory (at cost)", [inv[m] for m in months], RS, "inventory_daily, last day of month"),
    (10, "Credit (khata) share of sales", [khata[m] for m in months], PCT, "orders.payment_mode"),
    (11, "Days in month", days, "0", "Calendar"),
]
for r, lab, vals, fmt, src in data_rows:
    d.cell(row=r, column=1, value=lab).font = Font(name=F, size=10)
    for c, v in zip(MC, vals):
        d[f"{c}{r}"] = round(float(v), 4) if fmt == PCT else round(float(v))
        d[f"{c}{r}"].font, d[f"{c}{r}"].number_format = f_in, fmt
    d.cell(row=r, column=15, value=src).font = Font(name=F, size=8, italic=True, color="595959")
section(d, 13, "Operating expenses")
for i, name in enumerate(OPEX):
    r = 14 + i
    d.cell(row=r, column=1, value=name).font = Font(name=F, size=10)
    for c, m in zip(MC, months):
        d[f"{c}{r}"] = round(float(exp.loc[name, m]) if name in exp.index else 0)
        d[f"{c}{r}"].font, d[f"{c}{r}"].number_format = f_in, RS
EQ = 14 + len(OPEX)                                     # equipment row (27)
d.cell(row=EQ, column=1, value="Equipment purchase (capital spend)").font = Font(name=F, size=10)
for c, m in zip(MC, months):
    d[f"{c}{EQ}"] = round(float(exp.loc["Equipment", m]))
    d[f"{c}{EQ}"].font, d[f"{c}{EQ}"].number_format = f_in, RS
d.cell(row=14, column=15, value="expenses table").font = Font(name=F, size=8, italic=True, color="595959")
CAT0 = EQ + 3                                           # 30
section(d, CAT0 - 1, "Sales by product category")
d.cell(row=CAT0 - 1, column=15, value="FY total").font = f_bold
for i, name in enumerate(cat.index):
    r = CAT0 + i
    d.cell(row=r, column=1, value=name).font = Font(name=F, size=10)
    for c, m in zip(MC, months):
        d[f"{c}{r}"] = round(float(cat.loc[name, m]))
        d[f"{c}{r}"].font, d[f"{c}{r}"].number_format = f_in, RS
    d[f"O{r}"] = f"=SUM(C{r}:N{r})"
    d[f"O{r}"].number_format, d[f"O{r}"].font = RS, f_calc
CAT1 = CAT0 + len(cat) - 1
d.column_dimensions["O"].width = 30
d.freeze_panes = "C5"

# ================================================================== Assumptions
a = wb.create_sheet("Assumptions")
title(a, "Assumptions", "Blue = inputs. Yellow = the levers you are meant to change. Everything else in the workbook updates from here.")
a.column_dimensions["A"].width = 40
a.column_dimensions["B"].width = 16
a.column_dimensions["C"].width = 70


def inp(r, label, val, fmt, note, key=False):
    a.cell(row=r, column=1, value=label).font = Font(name=F, size=10)
    c = a.cell(row=r, column=2, value=val)
    c.font, c.number_format, c.border = f_in, fmt, box
    if key:
        c.fill = fill_yel
    a.cell(row=r, column=3, value=note).font = Font(name=F, size=9, italic=True, color="595959")


section(a, 4, "Opening balances (1 Nov 2025)")
inp(5, "Opening cash in hand and bank", 200000, RS, "Assumption: owner's estimate of cash at the start of the year.", True)
inp(6, "Khata (credit) owed by customers", 50000, RS, "Assumption: credit sales from October 2025 still to be collected.")
inp(7, "Opening inventory at cost", round(float(open_inv)), RS, "From the database: opening stock x cost price on 1 Nov 2025.")
inp(8, "Equipment useful life (months)", 60, "0", "Assumption: visi-cooler bought in June 2026 depreciated over 5 years.")
section(a, 10, "Budget")
inp(11, "Sales budget for November 2025", 750000, RS, "Assumption: owner's target for the first month.", True)
inp(12, "Monthly sales growth in budget", 0.01, PCT, "Assumption: 1% a month.", True)
inp(13, "Budget gross margin", 0.14, PCT, "Assumption: typical kirana gross margin of 12-15%.", True)
section(a, 15, "Monthly operating-expense budget")
for i, (n, b) in enumerate(zip(OPEX, OPEX_BUDGET)):
    inp(16 + i, n, b, RS, "")
OB_TOT = 16 + len(OPEX)                                 # 29
a.cell(row=OB_TOT, column=1, value="Total monthly opex budget").font = f_bold
a[f"B{OB_TOT}"] = f"=SUM(B16:B{OB_TOT - 1})"
a[f"B{OB_TOT}"].font, a[f"B{OB_TOT}"].number_format, a[f"B{OB_TOT}"].fill = f_bold, RS, fill_tot
W0 = OB_TOT + 2                                         # 31
section(a, W0, "What-if levers (used on the What-If sheet)")
levers = [("Selling price change", 0.0, PCT, "e.g. 1% = all prices up by 1%."),
          ("Sales volume change", 0.0, PCT, "e.g. 5% = 5% more units sold."),
          ("Supplier cost change", 0.0, PCT, "e.g. -1% = buy stock 1% cheaper."),
          ("Rent change (Rs per month)", 0, RS, "e.g. 3000 = rent goes up by Rs 3,000 a month."),
          ("Salary change (Rs per month)", 0, RS, "e.g. -9000 = one fewer delivery staff."),
          ("Wastage reduction", 0.0, PCT, "e.g. 50% = half as much stock thrown away.")]
for i, (n, v, fmt, note) in enumerate(levers):
    inp(W0 + 1 + i, n, v, fmt, note, True)
LV = {k: f"Assumptions!$B${W0 + 1 + i}" for i, k in enumerate(["price", "vol", "cost", "rent", "sal", "waste"])}

# ================================================================== P&L
p = wb.create_sheet("P&L")
title(p, "Profit and loss by month", "Rs. Links in green come from the Data sheet. Brackets = negative.")
month_header(p, 4)
row_formula(p, 5, "Revenue", lambda c, i: f"=Data!{c}5", font=f_link, bold=True)
row_formula(p, 6, "Cost of goods sold", lambda c, i: f"=Data!{c}6", font=f_link)
row_formula(p, 7, "Gross profit", lambda c, i: f"={c}5-{c}6", bold=True, total=True)
row_formula(p, 8, "Gross margin %", lambda c, i: f"=IF({c}5=0,0,{c}7/{c}5)", fmt=PCT, fy="=IF(O5=0,0,O7/O5)")
section(p, 10, "Operating expenses")
for i, n in enumerate(OPEX):
    row_formula(p, 11 + i, n, lambda c, j, i=i: f"=Data!{c}{14 + i}", font=f_link)
DEP = 11 + len(OPEX)                                    # 24
row_formula(p, DEP, "Depreciation on equipment", lambda c, i: f"=SUM(Data!$C${EQ}:{c}{EQ})/Assumptions!$B$8")
TOX = DEP + 1                                           # 25
row_formula(p, TOX, "Total operating expenses", lambda c, i: f"=SUM({c}11:{c}{DEP})", bold=True, total=True)
OP = TOX + 1                                            # 26
row_formula(p, OP, "Operating profit", lambda c, i: f"={c}7-{c}{TOX}", bold=True, total=True)
row_formula(p, OP + 1, "Operating margin %", lambda c, i: f"=IF({c}5=0,0,{c}{OP}/{c}5)", fmt=PCT, fy=f"=IF(O5=0,0,O{OP}/O5)")
p.freeze_panes = "C5"
p.conditional_formatting.add(f"C{OP}:O{OP}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
PL = {n: 11 + i for i, n in enumerate(OPEX)}

# ================================================================== Budget vs Actual
b = wb.create_sheet("Budget vs Actual")
title(b, "Budget vs actual", "Variance: revenue and profit = actual - budget; costs = budget - actual. Positive = favourable, (brackets) = adverse.")
month_header(b, 4)
section(b, 5, "Revenue")
row_formula(b, 6, "Budget", lambda c, i: "=Assumptions!$B$11" if i == 0 else f"={MC[i - 1]}6*(1+Assumptions!$B$12)")
row_formula(b, 7, "Actual", lambda c, i: f"='P&L'!{c}5", font=f_link)
row_formula(b, 8, "Variance", lambda c, i: f"={c}7-{c}6", bold=True)
row_formula(b, 9, "Variance %", lambda c, i: f"=IF({c}6=0,0,{c}8/{c}6)", fmt=PCT, fy="=IF(O6=0,0,O8/O6)")
section(b, 11, "Gross profit")
row_formula(b, 12, "Budget", lambda c, i: f"={c}6*Assumptions!$B$13")
row_formula(b, 13, "Actual", lambda c, i: f"='P&L'!{c}7", font=f_link)
row_formula(b, 14, "Variance", lambda c, i: f"={c}13-{c}12", bold=True)
section(b, 16, "Operating expenses (before depreciation)")
row_formula(b, 17, "Budget", lambda c, i: f"=Assumptions!$B${OB_TOT}", font=f_link)
row_formula(b, 18, "Actual", lambda c, i: f"='P&L'!{c}{TOX}-'P&L'!{c}{DEP}", font=f_link)
row_formula(b, 19, "Variance", lambda c, i: f"={c}17-{c}18", bold=True)
section(b, 21, "Operating profit (before depreciation)")
row_formula(b, 22, "Budget", lambda c, i: f"={c}12-{c}17")
row_formula(b, 23, "Actual", lambda c, i: f"={c}13-{c}18")
row_formula(b, 24, "Variance", lambda c, i: f"={c}23-{c}22", bold=True, total=True)
for r in (8, 9, 14, 19, 24):
    rng = f"C{r}:O{r}"
    b.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, color="C00000", bold=True)))
    b.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=F, color="006100", bold=True)))
section(b, 26, "Expenses by category: full year")
hdr = ["Category", "", "FY budget", "FY actual", "Variance", "Variance %", "Status"]
for i, h in enumerate(hdr, 1):
    c = b.cell(row=27, column=i, value=h)
    c.font, c.fill, c.border = f_head, fill_head, box
for i, n in enumerate(OPEX):
    r = 28 + i
    b.cell(row=r, column=1, value=n).font = Font(name=F, size=10)
    b[f"C{r}"] = f"=Assumptions!B{16 + i}*12"
    b[f"D{r}"] = f"='P&L'!O{PL[n]}"
    b[f"E{r}"] = f"=C{r}-D{r}"
    b[f"F{r}"] = f"=IF(C{r}=0,0,E{r}/C{r})"
    b[f"G{r}"] = f'=IF(E{r}>=0,"Within budget",IF(F{r}<-0.1,"Over by >10%","Slightly over"))'
    for col, fmt, fnt in (("C", RS, f_link), ("D", RS, f_link), ("E", RS, f_calc), ("F", PCT, f_calc), ("G", "@", f_calc)):
        b[f"{col}{r}"].number_format, b[f"{col}{r}"].font, b[f"{col}{r}"].border = fmt, fnt, box
CT = 28 + len(OPEX)
b.cell(row=CT, column=1, value="Total").font = f_bold
for col in "CDE":
    b[f"{col}{CT}"] = f"=SUM({col}28:{col}{CT - 1})"
    b[f"{col}{CT}"].number_format, b[f"{col}{CT}"].font, b[f"{col}{CT}"].fill = RS, f_bold, fill_tot
b[f"F{CT}"] = f"=IF(C{CT}=0,0,E{CT}/C{CT})"
b[f"F{CT}"].number_format, b[f"F{CT}"].font = PCT, f_bold
b.conditional_formatting.add(f"G28:G{CT - 1}", FormulaRule(formula=['G28="Over by >10%"'], fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=F, color="9C0006", bold=True)))
b.conditional_formatting.add(f"G28:G{CT - 1}", FormulaRule(formula=['G28="Slightly over"'], fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=F, color="9C5700")))
b.conditional_formatting.add(f"G28:G{CT - 1}", FormulaRule(formula=['G28="Within budget"'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=F, color="006100")))
b.column_dimensions["G"].width = 15
b.freeze_panes = "C5"

# ================================================================== Cash Flow
cf = wb.create_sheet("Cash Flow")
title(cf, "Cash flow by month", "Cash sales are collected the same month; khata (credit) sales are collected the following month. Suppliers are paid on delivery.")
month_header(cf, 4)
row_formula(cf, 5, "Opening cash", lambda c, i: "=Assumptions!$B$5" if i == 0 else f"={MC[i - 1]}17", bold=True, fy="=C5")
section(cf, 6, "Cash in")
row_formula(cf, 7, "Cash and UPI sales", lambda c, i: f"='P&L'!{c}5*(1-Data!{c}10)", font=f_link)
row_formula(cf, 8, "Khata collected (last month's credit)", lambda c, i: "=Assumptions!$B$6" if i == 0 else f"='P&L'!{MC[i - 1]}5*Data!{MC[i - 1]}10", font=f_link)
row_formula(cf, 9, "Total cash in", lambda c, i: f"={c}7+{c}8", bold=True, total=True)
section(cf, 10, "Cash out")
row_formula(cf, 11, "Paid to suppliers for stock", lambda c, i: f"=Data!{c}8", font=f_link)
row_formula(cf, 12, "Operating expenses paid", lambda c, i: f"='P&L'!{c}{TOX}-'P&L'!{c}{DEP}", font=f_link)
row_formula(cf, 13, "Equipment bought", lambda c, i: f"=Data!{c}{EQ}", font=f_link)
row_formula(cf, 14, "Total cash out", lambda c, i: f"=SUM({c}11:{c}13)", bold=True, total=True)
row_formula(cf, 16, "Net cash flow", lambda c, i: f"={c}9-{c}14", bold=True)
row_formula(cf, 17, "Closing cash", lambda c, i: f"={c}5+{c}16", bold=True, total=True, fy="=N17")
row_formula(cf, 19, "Khata owed by customers at month end", lambda c, i: f"='P&L'!{c}5*Data!{c}10", fy="=N19")
row_formula(cf, 20, "Inventory at cost at month end", lambda c, i: f"=Data!{c}9", font=f_link, fy="=N20")
cf["A22"] = "Check: closing cash = opening cash + total cash in - total cash out"
cf["A22"].font = Font(name=F, size=9, italic=True, color="595959")
cf["O22"] = "=IF(ABS(O5+O9-O14-O17)<1,\"OK\",\"Check\")"
cf["O22"].font = f_bold
cf.conditional_formatting.add("C16:O17", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
cf.freeze_panes = "C5"

# ================================================================== Ratios
ra = wb.create_sheet("Ratios")
title(ra, "Financial ratios by month", "Fixed costs = rent, salaries, internet, software, licences and depreciation. Everything else is treated as variable.")
month_header(ra, 4)
fixed_terms = lambda c: "+".join(f"'P&L'!{c}{PL[n]}" for n in OPEX if n in FIXED) + f"+'P&L'!{c}{DEP}"
row_formula(ra, 5, "Gross margin %", lambda c, i: f"='P&L'!{c}8", fmt=PCT, font=f_link, fy=f"='P&L'!O8")
row_formula(ra, 6, "Operating margin %", lambda c, i: f"='P&L'!{c}{OP + 1}", fmt=PCT, font=f_link, fy=f"='P&L'!O{OP + 1}")
row_formula(ra, 7, "Operating expenses as % of sales", lambda c, i: f"=IF('P&L'!{c}5=0,0,'P&L'!{c}{TOX}/'P&L'!{c}5)", fmt=PCT, fy=f"=IF('P&L'!O5=0,0,'P&L'!O{TOX}/'P&L'!O5)")
row_formula(ra, 8, "Average bill value", lambda c, i: f"=IF(Data!{c}7=0,0,'P&L'!{c}5/Data!{c}7)", fy="=IF(SUM(Data!C7:N7)=0,0,'P&L'!O5/SUM(Data!C7:N7))")
row_formula(ra, 9, "Bills per day", lambda c, i: f"=Data!{c}7/Data!{c}11", fmt=NUM1, fy="=SUM(Data!C7:N7)/SUM(Data!C11:N11)")
row_formula(ra, 10, "Average inventory at cost", lambda c, i: f"=(Assumptions!$B$7+Data!{c}9)/2" if i == 0 else f"=(Data!{MC[i - 1]}9+Data!{c}9)/2", fy="=AVERAGE(C10:N10)")
row_formula(ra, 11, "Inventory turnover (times a year)", lambda c, i: f"=IF({c}10=0,0,'P&L'!{c}6*12/{c}10)", fmt=NUM1, fy="=IF(O10=0,0,'P&L'!O6/O10)")
row_formula(ra, 12, "Days of stock on hand", lambda c, i: f"=IF('P&L'!{c}6=0,0,{c}10/('P&L'!{c}6/Data!{c}11))", fmt=NUM1, fy="=IF('P&L'!O6=0,0,O10/('P&L'!O6/SUM(Data!C11:N11)))")
row_formula(ra, 13, "Fixed costs", lambda c, i: "=" + fixed_terms(c), fy="=SUM(C13:N13)")
row_formula(ra, 14, "Contribution margin %", lambda c, i: f"=IF('P&L'!{c}5=0,0,('P&L'!{c}7-('P&L'!{c}{TOX}-{c}13))/'P&L'!{c}5)", fmt=PCT,
            fy=f"=IF('P&L'!O5=0,0,('P&L'!O7-('P&L'!O{TOX}-O13))/'P&L'!O5)")
row_formula(ra, 15, "Break-even sales", lambda c, i: f"=IF({c}14<=0,0,{c}13/{c}14)", bold=True, fy="=IF(O14<=0,0,O13/O14)")
row_formula(ra, 16, "Margin of safety %", lambda c, i: f"=IF('P&L'!{c}5=0,0,('P&L'!{c}5-{c}15)/'P&L'!{c}5)", fmt=PCT, fy="=IF('P&L'!O5=0,0,('P&L'!O5-O15)/'P&L'!O5)")
row_formula(ra, 17, "Closing cash", lambda c, i: f"='Cash Flow'!{c}17", font=f_link, fy="='Cash Flow'!O17")
ra.conditional_formatting.add("C16:O16", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
ra.freeze_panes = "C5"

# ================================================================== What-If
wi = wb.create_sheet("What-If")
title(wi, "What-if: what moves annual profit?", "Change the yellow levers on the Assumptions sheet. Base = actual full year from the P&L.")
for i, h in enumerate(["Line item (full year)", "", "Base", "Scenario", "Change"], 1):
    c = wi.cell(row=4, column=i, value=h)
    c.font, c.fill, c.border = f_head, fill_head, box
var_opex = f"('P&L'!O{PL['Packaging']}+'P&L'!O{PL['Transport']}+'P&L'!O{PL['Fuel']})"
rows = [
    ("Revenue", "='P&L'!O5", f"=C5*(1+{LV['price']})*(1+{LV['vol']})"),
    ("Cost of goods sold", "='P&L'!O6", f"=C6*(1+{LV['vol']})*(1+{LV['cost']})"),
    ("Gross profit", "=C5-C6", "=D5-D6"),
    ("Operating expenses", f"='P&L'!O{TOX}",
     f"=C8+12*{LV['rent']}+12*{LV['sal']}-'P&L'!O{PL['Wastage & Expiry']}*{LV['waste']}+{var_opex}*{LV['vol']}"),
    ("Operating profit", "=C7-C8", "=D7-D8"),
    ("Operating margin %", "=IF(C5=0,0,C9/C5)", "=IF(D5=0,0,D9/D5)"),
]
for i, (lab, base, scen) in enumerate(rows):
    r = 5 + i
    wi.cell(row=r, column=1, value=lab).font = f_bold if lab in ("Gross profit", "Operating profit") else Font(name=F, size=10)
    wi[f"C{r}"], wi[f"D{r}"] = base, scen
    wi[f"E{r}"] = f"=D{r}-C{r}"
    fmt = PCT if "%" in lab else RS
    for col in "CDE":
        wi[f"{col}{r}"].number_format, wi[f"{col}{r}"].border = fmt, box
        wi[f"{col}{r}"].font = f_link if (col == "C" and base.startswith("='P&L'")) else Font(name=F, size=10, bold=lab == "Operating profit")
wi["A11"] = "Change in operating profit"
wi["A11"].font = f_sec
wi["C11"] = "=IF(C9=0,0,E9/ABS(C9))"
wi["C11"].number_format, wi["C11"].font, wi["C11"].fill = PCT, Font(name=F, bold=True, size=12), fill_yel
wi.conditional_formatting.add("E5:E10", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
wi.conditional_formatting.add("E5:E10", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=F, bold=True, color="006100")))
section(wi, 13, "Sensitivity: annual operating profit if selling prices change (everything else as base)")
for i, h in enumerate(["Price change", "", "Revenue", "Operating profit", "vs base"], 1):
    c = wi.cell(row=14, column=i, value=h)
    c.font, c.fill, c.border = f_head, fill_head, box
for i, pc in enumerate([-0.02, -0.01, 0, 0.01, 0.02, 0.03]):
    r = 15 + i
    wi[f"A{r}"] = pc
    wi[f"A{r}"].number_format, wi[f"A{r}"].font, wi[f"A{r}"].border = '+0%;-0%;0%', f_in, box
    wi[f"C{r}"] = f"=$C$5*(1+A{r})"
    wi[f"D{r}"] = f"=C{r}-$C$6-$C$8"
    wi[f"E{r}"] = f"=D{r}-$C$9"
    for col in "CDE":
        wi[f"{col}{r}"].number_format, wi[f"{col}{r}"].font, wi[f"{col}{r}"].border = RS, f_calc, box
wi.conditional_formatting.add("D15:E20", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, color="C00000")))
wi["A22"] = "Reading it: a kirana works on thin margins, so a 1% change in selling prices moves annual profit far more than 1%."
wi["A22"].font = Font(name=F, size=9, italic=True, color="595959")
wi.column_dimensions["A"].width = 34
wi.column_dimensions["B"].width = 3
for col in "CDE":
    wi.column_dimensions[col].width = 16

# ================================================================== Dashboard
db = wb.create_sheet("Dashboard", 1)
title(db, "Dashboard", "Pick a month in the yellow cell. Figures below update for that month; charts show the full year.")
db["A4"] = "Month"
db["A4"].font = Font(name=F, bold=True, size=11)
db["B4"] = labels[-1]
db["B4"].font, db["B4"].fill, db["B4"].border = Font(name=F, bold=True, size=12, color="0000FF"), fill_yel, box
dv = DataValidation(type="list", formula1="=Data!$C$4:$N$4", allow_blank=False)
db.add_data_validation(dv)
dv.add("B4")
db["C4"] = "=MATCH($B$4,Data!$C$4:$N$4,0)"
db["C4"].font = Font(name=F, size=8, color="808080")
db["D4"] = "<- column number of the selected month (helper)"
db["D4"].font = Font(name=F, size=8, italic=True, color="808080")
pick = lambda sheet, r: f"=INDEX('{sheet}'!$C${r}:$N${r},$C$4)"
for i, h in enumerate(["Selected month", "Value", "Full year"], 1):
    c = db.cell(row=6, column=i, value=h)
    c.font, c.fill, c.border = f_head, fill_head, box
tiles = [
    ("Revenue", pick("P&L", 5), "='P&L'!O5", RS),
    ("Revenue vs budget", pick("Budget vs Actual", 9), "='Budget vs Actual'!O9", PCT),
    ("Gross margin", pick("P&L", 8), "='P&L'!O8", PCT),
    ("Operating expenses", pick("P&L", TOX), f"='P&L'!O{TOX}", RS),
    ("Operating profit", pick("P&L", OP), f"='P&L'!O{OP}", RS),
    ("Operating margin", pick("P&L", OP + 1), f"='P&L'!O{OP + 1}", PCT),
    ("Closing cash", pick("Cash Flow", 17), "='Cash Flow'!O17", RS),
    ("Average bill value", pick("Ratios", 8), "=Ratios!O8", RS),
    ("Days of stock on hand", pick("Ratios", 12), "=Ratios!O12", NUM1),
    ("Break-even sales", pick("Ratios", 15), "=Ratios!O15", RS),
    ("Margin of safety", pick("Ratios", 16), "=Ratios!O16", PCT),
]
for i, (lab, f1, f2, fmt) in enumerate(tiles):
    r = 7 + i
    db.cell(row=r, column=1, value=lab).font = Font(name=F, size=10)
    for col, f in (("B", f1), ("C", f2)):
        db[f"{col}{r}"] = f
        db[f"{col}{r}"].number_format, db[f"{col}{r}"].font, db[f"{col}{r}"].border = fmt, Font(name=F, bold=True, size=11, color="008000"), box
    db[f"A{r}"].border = box
for r in (8, 11, 12, 17):
    db.conditional_formatting.add(f"B{r}:C{r}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
db.column_dimensions["A"].width = 26
db.column_dimensions["B"].width = 16
db.column_dimensions["C"].width = 16
db["A19"] = "Biggest insight"
db["A19"].font = f_sec
db["A20"] = '="A 1% price rise would add "&TEXT(\'What-If\'!E18,"₹#,##0")&" to annual operating profit ("&TEXT(\'What-If\'!E18/\'What-If\'!C9,"0%")&")."'
db["A20"].font = Font(name=F, size=10)

# charts
c1 = BarChart(); c1.type = "col"; c1.title = "Revenue vs budget"; c1.height, c1.width = 7.5, 16
c1.add_data(Reference(b, min_col=1, max_col=14, min_row=7, max_row=7), from_rows=True, titles_from_data=True)
c1.set_categories(Reference(b, min_col=3, max_col=14, min_row=4))
l1 = LineChart(); l1.add_data(Reference(b, min_col=1, max_col=14, min_row=6, max_row=6), from_rows=True, titles_from_data=True)
c1.y_axis.numFmt = "#,##0"; c1.legend.position = "b"
for ser in l1.series: ser.smooth = False
c1 += l1
db.add_chart(c1, "E3")
c2 = BarChart(); c2.type = "col"; c2.title = "Operating profit by month"; c2.height, c2.width = 7.5, 16
c2.add_data(Reference(p, min_col=3, max_col=14, min_row=OP, max_row=OP), from_rows=True, titles_from_data=False)
c2.set_categories(Reference(p, min_col=3, max_col=14, min_row=4))
c2.legend = None; c2.y_axis.numFmt = "#,##0"
db.add_chart(c2, "N3")
c3 = LineChart(); c3.title = "Closing cash"; c3.height, c3.width = 7.5, 16
c3.add_data(Reference(cf, min_col=3, max_col=14, min_row=17, max_row=17), from_rows=True, titles_from_data=False)
c3.set_categories(Reference(cf, min_col=3, max_col=14, min_row=4))
c3.legend = None; c3.y_axis.numFmt = "#,##0"
for ser in c3.series: ser.smooth = False
db.add_chart(c3, "E19")
c4 = BarChart(); c4.type = "bar"; c4.title = "Where operating costs go (full year)"; c4.height, c4.width = 7.5, 16
c4.add_data(Reference(p, min_col=15, min_row=11, max_row=DEP))
c4.set_categories(Reference(p, min_col=1, min_row=11, max_row=DEP))
c4.legend = None; c4.x_axis.scaling.orientation = "maxMin"; c4.y_axis.numFmt = "#,##0"
db.add_chart(c4, "N19")
c5 = BarChart(); c5.type = "bar"; c5.title = "Sales by category (full year)"; c5.height, c5.width = 8, 16
c5.add_data(Reference(d, min_col=15, min_row=CAT0, max_row=CAT1))
c5.set_categories(Reference(d, min_col=1, min_row=CAT0, max_row=CAT1))
c5.legend = None; c5.x_axis.scaling.orientation = "maxMin"; c5.y_axis.numFmt = "#,##0"
db.add_chart(c5, "E35")

order = ["Guide", "Dashboard", "P&L", "Budget vs Actual", "Cash Flow", "Ratios", "What-If", "Assumptions", "Data"]
wb._sheets = [wb[n] for n in order]
wb.active = 1
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {"Dashboard": NAVY, "Assumptions": "C9A227", "What-If": "2A7F75"}.get(ws.title)
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved", OUT)
