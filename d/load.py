import openpyxl
path = "/home/karthikeyan/vscode/py/d/5-15 TO 7-8(nifty).xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active