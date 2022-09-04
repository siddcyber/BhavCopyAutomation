import requests
import pandas as pd


class holidaylist:
    def __init__(self):
        self.header = {
            'user-agent': 'Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.188 Safari/537.36 CrKey/1.54.250320',
            'accept-language': 'en-IN,en;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-GB;q=0.6,en-US;q=0.5',
            'accept-encoding': 'gzip, deflate, br',
            'content-type': 'application/json; charset=utf-8'
        }
        self.holidays = requests.get(' https://www.nseindia.com/api/holiday-master?type=trading', headers=self.header).json()
        self.df = pd.DataFrame.from_dict(self.holidays, orient='index')
        self.df = self.df.transpose()
        self.df = self.df['CM']
        self.df2 = pd.DataFrame()
        for row in range(len(self.df)):
            self.df2 = self.df2.append(self.df.loc[row], ignore_index=True)
        self.dflist = self.df2['tradingDate'].tolist()
