# install packages: pywin32, pandas, requests, future, and winshell

import os
import requests
from datetime import date, datetime, timedelta
from zipfile import ZipFile
import time
start = time.time()
# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6

# read the bhavcopy_location file for the save location of the NSE Bhavcopy
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

# Url of the File to be Downloaded
URL2 = "https://archives.nseindia.com/content/historical/EQUITIES/" + \
       str(nameDateYesterday.strftime("%Y") + "/" + nameDateYesterday.strftime("%b").upper()) + '/' + zipname

# # main function to repeat until the file is downloaded, unzipped and delete the zip file
# recursion method takes 0.63 to 0.17 seconds
# def main():
#     try:
#         response = requests.get(URL2)  # download the data behind the URL
#         open(zipname, "wb").write(response.content)  # Open the response into a new file
#         # extract zip file to specified location
#         with ZipFile(zipname, 'r') as zip_file:
#             zip_file.extractall(path=path)
#         os.remove(zipname)  # removes the downloaded zip file
#         print("success")
#     except (requests.exceptions.ConnectionError, FileNotFoundError):
#         #  retry the try part after some seconds
#         time.sleep(60)
#         # Try again(recursion)
#         main()
# main()

# while true method takes 0.7 to 0.19 seconds
# while True:
#     try:
#         response = requests.get(URL2)  # download the data behind the URL
#         open(zipname, "wb").write(response.content)  # Open the response into a new file
#         # extract zip file to specified location
#         with ZipFile(zipname, 'r') as zip_file:
#             zip_file.extractall(path=path)
#         os.remove(zipname)  # removes the downloaded zip file
#         print("success")
#         break
#     except (requests.exceptions.ConnectionError, FileNotFoundError):
#         print("error, trying again in 10s")
#         time.sleep(1)
end = time.time()
print("The time of execution of above program is :", end-start)
