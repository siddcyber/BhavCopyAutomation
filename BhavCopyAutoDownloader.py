# install pywin32, future, nsepy and winshell

import os
import requests
from datetime import date, datetime, timedelta
from zipfile import ZipFile
import time
# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6

# nsepy method
# from nsepy.history import get_price_list
# yesterday = today - timedelta(days=1)  # yesterday's date
# nameDateYesterday = nameDateToday - timedelta(days=1)  # Yesterday's date for file naming
# prises = get_price_list(dt=yesterday)  # gets latest Bhav NSE copy list
# # saving of file in CSV format with name
# prises.to_csv(os.path.join(path, 'cm' + nameDateYesterday.strftime("%d%b%Y") + 'bhav.csv'))

#  Add try except statements to check if internet works if not retry after 1 minute
#  and convert code into class(if possible)
with open((os.getcwd() + "\BhavCopy_location.txt"), 'r') as file:
    path = str(file.readline())

today = date.today()  # today's date
nameDateToday = datetime.today()  # today's date for file naming

if today.weekday() == 6:
    nameDateYesterday = nameDateToday - timedelta(days=2)  # Friday's date for file naming

elif today.weekday() == 0:
    nameDateYesterday = nameDateToday - timedelta(days=3)  # Friday's date for file naming

else:
    nameDateYesterday = nameDateToday - timedelta(days=1)  # Yesterday's date for file naming

#  Name of the zip file to be downloaded
zipname = str("cm" + nameDateYesterday.strftime("%d%b%Y").upper() + 'bhav.csv.zip')

# URL = "https://archives.nseindia.com/content/historical/EQUITIES/" + \
#       str(nameDateYesterday.strftime("%Y") + "/" + nameDateYesterday.strftime("%b").upper() + "/cm" +
#           nameDateYesterday.strftime("%d%b%Y").upper()) + 'bhav.csv.zip'

URL2 = "https://archives.nseindia.com/content/historical/EQUITIES/" + \
      str(nameDateYesterday.strftime("%Y") + "/" + nameDateYesterday.strftime("%b").upper()) + '/' + zipname
# Method one Simple for loop
## add value while false if file still not downloaded
# no_of_attempts = 5 # set number of attempts
# for i in range (no_of_attempts):
#     try:
#         response = requests.get(URL2)  # download the data behind the URL
#         open(zipname, "wb").write(response.content)  # Open the response into a new file
#         # extract zip file to specified location
#         with ZipFile(zipname, 'r') as zip_file:
#             zip_file.extractall(path=path)
#         os.remove(zipname)  # removes the downloaded zip file
#         print("itworks")
#     except (requests.exceptions.ConnectionError, FileNotFoundError):
#         print("finally the error")
#         time.sleep(3)
#         continue

# method 2 effective but efficient? check add total no of iterations to go through
def doSomething():
    try:
        response = requests.get(URL2)  # download the data behind the URL
        open(zipname, "wb").write(response.content)  # Open the response into a new file
        # extract zip file to specified location
        with ZipFile(zipname, 'r') as zip_file:
            zip_file.extractall(path=path)
        os.remove(zipname)  # removes the downloaded zip file
        print("itworks")
    except (requests.exceptions.ConnectionError, FileNotFoundError):
        print("finally the error")
        #  retry the try part after some seconds
        time.sleep(1)
        # Try again
        doSomething()
doSomething()
