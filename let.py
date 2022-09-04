import openpyxl
f = openpyxl.load_workbook('./Treino.xlsx') 
sheet = f.active
print(sheet)

