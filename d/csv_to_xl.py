
import pandas as pd
df_new = pd.read_csv('file1.csv')
 
# saving xlsx file
GFG = pd.ExcelWriter('/home/karthikeyan/vscode/py/d/5-15 TO 7-8(nifty).xlsx')
df_new.to_excel(GFG, index=False)
 
GFG.save()
