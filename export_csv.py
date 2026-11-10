"""Write refills.csv (sorted by date/time) from drink_refills.xlsx for the dashboard."""
import csv
from openpyxl import load_workbook

ws = load_workbook("drink_refills.xlsx")["Refills"]
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    d, t, loc, drinks, price = r[:5]
    rows.append([d.strftime("%Y-%m-%d"), t.strftime("%H:%M"), loc, int(drinks), f"{float(price):.2f}"])
rows.sort(key=lambda x: (x[0], x[1]))
with open("refills.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["date", "time", "location", "drinks", "price"])
    w.writerows(rows)
print(f"{len(rows)} rows -> refills.csv")
