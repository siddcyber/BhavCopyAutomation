from bs4 import BeautifulSoup
import requests
import pandas as pd

header = {
    'user-agent': 'Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.188 Safari/537.36 CrKey/1.54.250320',
    'accept-language': 'en-IN,en;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-GB;q=0.6,en-US;q=0.5',
    'accept-encoding': 'gzip, deflate, br',
    'content-type': 'application/json; charset=utf-8'
}
holidays = requests.get(' https://www.nseindia.com/api/holiday-master?type=trading', headers=header).json()
df = pd.DataFrame.from_dict(holidays, orient='index')
df = df.transpose()
print(df.keys())
df = df['CM']
df2 = pd.DataFrame()
for row in range(len(df)):
    df2 = df2.append(df.loc[row], ignore_index=True)
print(df2.to_string())
print(df2['tradingDate'].to_string())
# print(df.keys())
# print(df.index.values)
# print(df.loc[0])
# print(df.iloc[:,:1])
