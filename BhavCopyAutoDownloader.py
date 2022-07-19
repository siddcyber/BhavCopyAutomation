# install pywin32, future, nsepy and winshell

import os
from datetime import date, datetime, timedelta
import requests

# file opens after 5 minutes of startup, asks user to edit file (provide auto edit option, ask first time for auto edit)

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
    # saving of file in CSV format with name

elif today.weekday() == 0:
    nameDateYesterday = nameDateToday - timedelta(days=3)  # Friday's date for file naming
    # saving of file in CSV format with name

else:
    nameDateYesterday = nameDateToday - timedelta(days=1)  # Yesterday's date for file naming
    # saving of file in CSV format with name


URL = "https://archives.nseindia.com/content/historical/EQUITIES/" + \
      str(nameDateYesterday.strftime("%Y")+"/" + nameDateYesterday.strftime("%b").upper() + "/cm" +
          nameDateYesterday.strftime("%d%b%Y").upper()) + 'bhav.csv.zip'


# 2. download the data behind the URL
response = requests.get(URL)

# 3. Open the response into a new file called instagram.ico
open(str("cm" + nameDateYesterday.strftime("%d%b%Y").upper() + 'bhav.csv.zip'), "wb").write(response.content)

# extract zip file to specified location

#  delete the zip file
