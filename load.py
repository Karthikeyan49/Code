import openpyxl
path = "/home/karthikeyan/vscode/py/5-15 TO 7-8(close)(nifty bank).xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active