import openpyxl
from pathlib import Path

wb_path = r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\Skoda DSR (1).xlsx'
wb = openpyxl.load_workbook(wb_path, read_only=True)
sheet = wb.active
headers = [cell.value for cell in sheet[1]]
print("ZOHO HEADERS:")
for i, h in enumerate(headers):
    print(f"{i}: {h}")

master_path = r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\SKODA_MASTER_DSR.xlsx'
if Path(master_path).exists():
    wb_m = openpyxl.load_workbook(master_path, read_only=True)
    sheet_m = wb_m.active
    master_headers = [cell.value for cell in sheet_m[1]]
    print("\nMASTER HEADERS:")
    for i, h in enumerate(master_headers):
        print(f"{i}: {h}")
