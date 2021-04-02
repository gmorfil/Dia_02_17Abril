import openpyxl
miarchivo=openpyxl.Workbook()
mihoja=miarchivo.active
mihoja["A1"]=5
mihoja["A2"]=7
miarchivo.save('hoja_excel.xlsx')
miarchivo=openpyxl.load_workbook('hoja_excel.xlsx')
print(mihoja['A1'].value)


