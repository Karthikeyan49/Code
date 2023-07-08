
import pandas as pd
df_new = pd.read_csv('file1.csv')
 
# saving xlsx file
GFG = pd.ExcelWriter('5-15 TO 7-8(close)(nifty bank).xlsx')
df_new.to_excel(GFG, index=False)
 
GFG.save()
