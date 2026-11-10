"""Append a refill to drink_refills.xlsx, stamped with the current Mountain Time.

Usage: python3 add_refill.py <drinks> "<location>" <price> ["YYYY-MM-DD HH:MM"]
The optional last argument overrides the timestamp (Mountain Time).
"""
import sys
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

FILE = "drink_refills.xlsx"
TZ = ZoneInfo("America/Denver")

drinks, location, price = int(sys.argv[1]), sys.argv[2], float(sys.argv[3].lstrip("$"))
when = (datetime.strptime(sys.argv[4], "%Y-%m-%d %H:%M") if len(sys.argv) > 4
        else datetime.now(TZ).replace(tzinfo=None))

wb = load_workbook(FILE)
ws = wb["Refills"]
row = 2
while ws[f"A{row}"].value is not None:
    row += 1
ws[f"A{row}"] = when.date()
ws[f"B{row}"] = when.time().replace(second=0, microsecond=0)
ws[f"C{row}"] = location
ws[f"D{row}"] = drinks
ws[f"E{row}"] = price
ws[f"F{row}"] = f"=IF(D{row}>0,E{row}/D{row},\"\")"
ws[f"A{row}"].number_format = "mm/dd/yyyy"
ws[f"B{row}"].number_format = "h:mm AM/PM"
ws[f"E{row}"].number_format = ws[f"F{row}"].number_format = "$#,##0.00"
wb.save(FILE)
print(f"Row {row}: {when:%m/%d/%Y %I:%M %p} MT | {location} | {drinks} drinks | ${price:.2f}")

import subprocess
subprocess.run([sys.executable, "export_csv.py"], check=True)
