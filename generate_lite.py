# -*- coding: utf-8 -*-
"""生成免费精简版《ecom-recon-lite.xlsx》：单区100行对账引擎，无看板/差异清单/图表（Pro独占）。"""
import random, datetime as dt
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

random.seed(20260910)
OUT = "/Users/kk/workspaces/side-hustle/recon-pro/lite/ecom-recon-lite.xlsx"
MAX_ROW = 101

def D2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
PASTE_FILL = PatternFill("solid", fgColor="FFF7E6")
THIN = Border(*[Side(style="thin", color="D9D9D9")] * 4)

def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN

orders, funds = [], []
plats = ["抖店", "淘宝", "拼多多"]
for i in range(12):
    p = random.choice(plats)
    paid = D2(round(random.uniform(29, 399), 2))
    ship = random.choice([0, 0, 8])
    refund = 0 if random.random() > 0.15 else D2(paid * 0.8)
    comm = D2(paid * 0.05)
    promo = D2(paid * 0.03)
    oid = f"DEMO{20260900 + i}"
    orders.append((p, oid, paid, ship, refund))
    if i != 2:  # 埋2处演示差错: i=2未结算 / i=7金额差异
        arrive = D2(paid + ship - comm - promo - refund)
        arrive = D2(arrive - 10) if i == 7 else arrive
        funds.append((p, oid, paid + ship, comm, promo, refund, arrive))

wb = Workbook()
ws = wb.active; ws.title = "说明"
ws.sheet_view.showGridLines = False
ws["B2"] = "电商对账 Lite（免费版）"; ws["B2"].font = Font(bold=True, size=18, color="1F4E79")
notes = [
    "用法：把订单粘进【订单数据】，结算账单粘进【资金账单】，【自动对账】出结果。",
    "列格式：订单=平台/订单号/实付/运费/退款；账单=平台/订单号/订单金额/佣金/推广费/退款扣回/实际到账。",
    "原理：应收(实付+运费) = 到账+佣金+推广+扣回，差>0.01元标红。",
    "自带12单演示数据：1笔⚠未结算(DEMO202609002)、1笔❌金额差异(DEMO202609007)。",
    "免费版限制：100行、无利润看板/差异清单/图表/教程。",
    "Pro版(600行多平台+真实利润看板+差异清单+图文教程)：见仓库 README。",
]
r = 4
for n in notes:
    ws.cell(row=r, column=2, value=n).font = Font(size=12)
    r += 1
ws.column_dimensions["B"].width = 100

ws = wb.create_sheet("订单数据")
for i, h in enumerate(["平台", "订单号", "买家实付(元)", "运费(元)", "退款(元)"], 1):
    ws.cell(row=1, column=i, value=h)
style_header(ws, 5)
for r, (p, oid, paid, ship, refund) in enumerate(orders, 2):
    for c, v in enumerate((p, oid, paid, ship, refund), 1):
        cell = ws.cell(row=r, column=c, value=v)
        if c > 2: cell.number_format = "#,##0.00"
for r in range(2, MAX_ROW + 1):
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = PASTE_FILL
        ws.cell(row=r, column=c).border = THIN
for col, w in zip("ABCDE", [10, 20, 14, 10, 10]):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("资金账单")
for i, h in enumerate(["平台", "订单号", "订单金额(元)", "佣金(元)", "推广费(元)", "退款扣回(元)", "实际到账(元)"], 1):
    ws.cell(row=1, column=i, value=h)
style_header(ws, 7)
for r, (p, oid, amt, comm, promo, deduct, arrive) in enumerate(funds, 2):
    for c, v in enumerate((p, oid, amt, comm, promo, deduct, arrive), 1):
        cell = ws.cell(row=r, column=c, value=v)
        if c > 2: cell.number_format = "#,##0.00"
for r in range(2, MAX_ROW + 1):
    for c in range(1, 8):
        ws.cell(row=r, column=c).fill = PASTE_FILL
        ws.cell(row=r, column=c).border = THIN
for col, w in zip("ABCDEFG", [10, 20, 14, 10, 11, 12, 12]):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("自动对账")
for i, h in enumerate(["平台", "订单号", "应收(实付+运费)", "资金侧核对额", "结算笔数", "金额差异", "退款差异", "结果"], 1):
    ws.cell(row=1, column=i, value=h)
style_header(ws, 8)
ZJ, OD = "资金账单", "订单数据"
def fs(r, col):
    return f'SUMIFS({ZJ}!${col}$2:${col}${MAX_ROW},{ZJ}!$B$2:$B${MAX_ROW},$B{r})'
for r in range(2, MAX_ROW + 1):
    ws.cell(row=r, column=1, value=f'=IF({OD}!B{r}="","",{OD}!A{r})')
    ws.cell(row=r, column=2, value=f'=IF({OD}!B{r}="","",{OD}!B{r})')
    ws.cell(row=r, column=3, value=f'=IF($B{r}="","",ROUND({OD}!C{r}+{OD}!D{r},2))').number_format = "#,##0.00"
    ws.cell(row=r, column=4, value=f'=IF($B{r}="","",ROUND({fs(r,"G")}+{fs(r,"D")}+{fs(r,"E")}+{fs(r,"F")},2))').number_format = "#,##0.00"
    ws.cell(row=r, column=5, value=f'=IF($B{r}="","",COUNTIF({ZJ}!$B$2:$B${MAX_ROW},$B{r}))')
    ws.cell(row=r, column=6, value=f'=IF($B{r}="","",ROUND($C{r}-$D{r},2))').number_format = "#,##0.00"
    ws.cell(row=r, column=7, value=f'=IF($B{r}="","",ROUND({OD}!E{r}-{fs(r,"F")},2))').number_format = "#,##0.00"
    ws.cell(row=r, column=8, value=(
        f'=IF($B{r}="","",IF($E{r}=0,"⚠未结算",IF($E{r}>1,"❌重复结算",'
        f'IF(AND(ABS($G{r})>0.01,ABS($F{r})<=0.01),"❌退款未同步",'
        f'IF(OR(ABS($F{r})>0.01,ABS($G{r})>0.01),"❌金额差异","✓")))))'))
rng = f"H2:H{MAX_ROW}"
ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT($H2,1)="⚠"'],
    fill=PatternFill("solid", fgColor="FFE699"), font=Font(color="9C6500", bold=True)))
ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT($H2,1)="❌"'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$H2="✓"'],
    fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100")))
for col, w in zip("ABCDEFGH", [10, 20, 16, 14, 10, 11, 11, 13]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

wb.save(OUT)
print("lite saved:", OUT, "| orders:", len(orders), "funds:", len(funds))
