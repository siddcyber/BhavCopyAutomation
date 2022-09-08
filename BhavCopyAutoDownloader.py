# install packages: pywin32, pandas, requests, future, and winshell
# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6


import os
from datetime import date, timedelta
from zipfile import ZipFile, BadZipfile
import time
import requests

today = date.today()  # today's date

if today.weekday() == 6:
    today = today - timedelta(days=2)  # Friday's date for file naming

elif today.weekday() == 0:
    today = today - timedelta(days=3)  # Friday's date for file naming

else:
    today = today - timedelta(days=1)  # Yesterday's date for file naming

#  Name of the zip file to be downloaded
zipname = str("cm" + today.strftime("%d%b%Y").upper() + 'bhav.csv.zip')

# Url of the File to be Downloaded
URL2 = "https://www1.nseindia.com/content/historical/EQUITIES/" + \
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
    except BadZipfile:
        print('badzip')
        os.remove(zipname)
        print('remove successful')
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
