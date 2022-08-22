from openpyxl import load_workbook
wb = load_workbook(filename = './one.xlsx')
sheet_ranges = wb['Plan1']
print(sheet_ranges)
