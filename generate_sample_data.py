"""
見本用の架空データを作成するスクリプト。
3店舗のカフェが、それぞれ違う形式で売上をExcelに記録している、という想定。
（実在の企業・データとは一切関係ありません）
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

random.seed(42)
OUT = Path("input")
OUT.mkdir(exist_ok=True)

PRODUCTS = {  # 商品名: 単価
    "ブレンドコーヒー": 450,
    "カフェラテ": 520,
    "チーズケーキ": 480,
    "サンドイッチ": 650,
    "コーヒー豆200g": 1500,
}
MONTHS = [(2026, 7), (2026, 8), (2026, 9)]


def days_in(y, m):
    d = date(y, m, 1)
    while d.month == m:
        yield d
        d += timedelta(days=1)


def daily_rows(y, m, scale):
    for d in days_in(y, m):
        for name, price in PRODUCTS.items():
            if random.random() < 0.8:
                qty = max(1, int(random.gauss(12, 5) * scale))
                yield d, name, qty, price


# 新宿店：比較的きれい。ただし商品名の前後に空白が混じる
wb = Workbook(); ws = wb.active; ws.title = "売上"
ws.append(["日付", "商品名", "数量", "単価", "売上"])
for y, m in MONTHS:
    for d, name, qty, price in daily_rows(y, m, 1.3):
        n = f" {name}" if random.random() < 0.1 else name
        ws.append([d, n, qty, price, qty * price])
wb.save(OUT / "新宿店_売上_2026Q3.xlsx")

# 吉祥寺店：タイトル行あり、日付は「7月1日」の文字列、金額は「1,350円」の文字列、
# 数字に全角が混じる、空行あり、別名の列見出し
wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
ws.append(["吉祥寺店 売上記録（7月〜9月）"])
ws.append([])
ws.append(["売上日", "品名", "個数", "金額"])
z = str.maketrans("0123456789", "０１２３４５６７８９")
for y, m in MONTHS:
    for d, name, qty, price in daily_rows(y, m, 0.9):
        q = str(qty).translate(z) if random.random() < 0.15 else qty
        ws.append([f"{d.month}月{d.day}日", name, q, f"{qty * price:,}円"])
        if random.random() < 0.02:
            ws.append([])
wb.save(OUT / "吉祥寺店_売上.xlsx")

# 立川店：月ごとに別シート、英語の列見出し、半角カナの商品名、売上列なし、末尾に合計行
HALF = {"ブレンドコーヒー": "ﾌﾞﾚﾝﾄﾞｺｰﾋｰ", "カフェラテ": "ｶﾌｪﾗﾃ"}
wb = Workbook(); wb.remove(wb.active)
for y, m in MONTHS:
    ws = wb.create_sheet(f"{m}月")
    ws.append(["Date", "Item", "Qty", "UnitPrice"])
    total = 0
    for d, name, qty, price in daily_rows(y, m, 1.0):
        n = HALF.get(name, name) if random.random() < 0.3 else name
        ws.append([d.strftime("%Y-%m-%d"), n, qty, price])
        total += qty
    ws.append(["合計", None, total, None])
wb.save(OUT / "tachikawa_sales.xlsx")

print("input/ に3店舗分の架空データを作成しました")
