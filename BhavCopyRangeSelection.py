# install packages: pywin32, pandas, requests, future, and winshell

import os
import requests
from datetime import date, datetime, timedelta
from zipfile import ZipFile, BadZipfile
import time


# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6
def extractAndDelete(filename, url):
    while True:
        # main function to repeat until the file is downloaded, unzipped and delete the zip file
        try:
            # read the bhavcopy_location file for the save location of the NSE Bhavcopy
            with open((os.getcwd() + r"\BhavCopy_location.txt"), 'r') as file:
                path = str(file.readline())
            response = requests.get(URL2)  # download the data behind the URL
            open(zipname, "wb").write(response.content)  # Open the response into a new file
            # extract zip file to specified location
            print('file extraction start')
            with ZipFile(zipname, 'r') as zip_file:
                zip_file.extractall(path=path)
                print('file extracted')
                print('sleeping for 0.1 second')
                time.sleep(0.1)
                print('complete')
            print('deleting file')
            os.remove(zipname)  # removes the downloaded zip file
            print('file deleted')
            break
        except requests.exceptions.ConnectionError:
            print("preSleepConnectionErrorCatchSuccess")
            time.sleep(10)
            print("postSleepConnectionErrorCatchSuccess")
        # except(FileNotFoundError, NameError):
        #     print(r"preSleepNameError/FileNotFoundErrorCatchSuccess")
        #     os.startfile(os.getcwd() + r"\BhavCopyAutoSettings.exe")
        #     time.sleep(60)
        #     print(r"postSleepNameError/FileNotFoundErrorCatchSuccess")


def checkWeekend(day):
    weekend = False
    if day.weekday() == 6 or day.weekday() == 5:
        weekend = True
    return weekend


startDate = date(2022, 1, 1)
endDate = date(2022, 8, 1)
today = date.today()  # today's date
# deltaDate = endDate-startDate # returns deltatime
deltaDate = today - startDate  # returns deltatime

if today.weekday() == 6:
    today = today - timedelta(days=2)  # Friday's date for file naming
elif today.weekday() == 0:
    today = today - timedelta(days=3)  # Friday's date for file naming
else:
    today = today - timedelta(days=1)  # Yesterday's date for file naming
print('start')
for i in range(deltaDate.days + 1):
    currentDay = startDate + timedelta(days=i)
    if checkWeekend(currentDay):
        pass
    else:
        #  Name of the zip file to be downloaded
        zipname = str("cm" + currentDay.strftime("%d%b%Y").upper() + 'bhav.csv.zip')
        # Url of the File to be Downloaded
        URL2 = "https://www1.nseindia.com/content/historical/EQUITIES/" + \
               str(currentDay.strftime("%Y") + "/" + currentDay.strftime("%b").upper()) + '/' + zipname
        try:
            print('start function')
            extractAndDelete(zipname, URL2)
        except BadZipfile:
            print('BadZipfile')
            os.remove(zipname)
            continue
print('done')