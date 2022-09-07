# install packages: pywin32, pandas, requests, future, and winshell

import os
from datetime import date, datetime, timedelta
from zipfile import ZipFile
import time
import requests
import pandas as pd
from dateutil.parser import parse

# function to get and parse holiday list from nse website directly
def holidaylist():
        header = {
            'user-agent': 'Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.188 Safari/537.36 CrKey/1.54.250320',
            'accept-language': 'en-IN,en;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-GB;q=0.6,en-US;q=0.5',
            'accept-encoding': 'gzip, deflate, br',
            'content-type': 'application/json; charset=utf-8'
        }
        holidays = requests.get(' https://www.nseindia.com/api/holiday-master?type=trading', headers=header).json()
        df = pd.DataFrame.from_dict(holidays, orient='index')
        df = df.transpose()
        df = df['CM']
        df2 = pd.DataFrame()
        for row in range(len(df)):
            df2 = df2.append(df.loc[row], ignore_index=True)
        dflist = df2['tradingDate'].tolist()
        for i in range(len(dflist)):
            dflist[i] = parse(dflist[i])
            dflist[i] = dflist[i].date()
        return dflist

# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6

#  Function to check if today is a holiday
def checkholiday(day):
    dflist = holidaylist()
    holiday = False
    if day in dflist:
        holiday = True
    #  if today is a holiday return true
    return holiday

today = date.today()  # today's date

if today.weekday() == 6:
    today = today - timedelta(days=2)  # Friday's date for file naming

elif today.weekday() == 0:
    today = today - timedelta(days=3)  # Friday's date for file naming

elif checkholiday(today):
    today = today - timedelta(days=2)  # Holiday's date for file naming

else:
    today = today - timedelta(days=1)  # Yesterday's date for file naming

#  Name of the zip file to be downloaded
zipname = str("cm" + today.strftime("%d%b%Y").upper() + 'bhav.csv.zip')

# Url of the File to be Downloaded
URL2 = "https://archives.nseindia.com/content/historical/EQUITIES/" + \
       str(today.strftime("%Y") + "/" + today.strftime("%b").upper()) + '/' + zipname

while True:
    # main function to repeat until the file is downloaded, unzipped and delete the zip file
    try:
        # read the bhavcopy_location file for the save location of the NSE Bhavcopy
        with open((os.getcwd() + r"\BhavCopy_location.txt"), 'r') as file:
            path = str(file.readline())
        response = requests.get(URL2)  # download the data behind the URL
        open(zipname, "wb").write(response.content)  # Open the response into a new file
        # extract zip file to specified location
        with ZipFile(zipname, 'r') as zip_file:
            zip_file.extractall(path=path)
        os.remove(zipname)  # removes the downloaded zip file
        break
    except requests.exceptions.ConnectionError:
        print("preSleepConnectionErrorCatchSuccess")
        time.sleep(10)
        print("postSleepConnectionErrorCatchSuccess")

    except(FileNotFoundError, NameError):
        print(r"preSleepNameError/FileNotFoundErrorCatchSuccess")
        os.startfile(os.getcwd() + r"\BhavCopyAutoSettings.exe")
        time.sleep(60)
        print(r"postSleepNameError/FileNotFoundErrorCatchSuccess")
