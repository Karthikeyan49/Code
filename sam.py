from yahoo_fin import stock_info
import yfinance as yf
# df = yf.download("^NSEI", start="2023-05-15", end="2023-07-08", interval="5m")
df = yf.download("^NSEBANK", start="2023-05-15", end="2023-07-08", interval="5m")
# nifty=yf.download("^NSEI",start="2023-05-09",end="2023-06-09", interval="15m")
# print(stock_info.get_live_price("^NSEI"))
# print(type(df))
# print(df.index.tolist)
# n=input("enter time : ")
# for i in range(df):
#     df['Datatime']==n
# file_name = 'MarksData.xlsx'
# df.index.tolist
#  saving the excel
# df.to_excel(file_name)
df.to_csv('file1.csv')
