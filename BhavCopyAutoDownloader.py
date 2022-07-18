import os
from datetime import date, datetime, timedelta
from nsepy.history import get_price_list

# file opens after 5 minutes of startup, asks user to edit file (provide auto edit option, ask first time for auto edit)

# monday    -> friday   ->0
# tuesday   -> monday   ->1
# wednesday -> tuesday  ->2
# Thursday  -> wednesday->3
# friday    -> thursday ->4
# saturday  -> friday   ->5
# sunday    -> friday   ->6

#  Add try except statements to check if internet works if not retry after 1 minute
#  and convert code into class(if possible)
with open((os.getcwd() + "\BhavCopy_location.txt"), 'r') as file:
    path = str(file.readline())

today = date.today()  # today's date
nameDateToday = datetime.today()  # today's date for file naming

if today.weekday() == 6:
    yesterday = today - timedelta(days=2)
    nameDateYesterday = nameDateToday - timedelta(days=2)  # Friday's date for file naming
    prises = get_price_list(dt=yesterday)  # gets latest Bhav NSE copy list
    # saving of file in CSV format with name
    prises.to_csv(os.path.join(path, 'cm' + nameDateYesterday.strftime("%d%b%Y") + 'bhav.csv'))

elif today.weekday() == 0:
    yesterday = today - timedelta(days=3)
    nameDateYesterday = nameDateToday - timedelta(days=3)  # Friday's date for file naming
    prises = get_price_list(dt=yesterday)  # gets latest Bhav NSE copy list
    # saving of file in CSV format with name
    prises.to_csv(os.path.join(path, 'cm' + nameDateYesterday.strftime("%d%b%Y") + 'bhav.csv'))

else:
    yesterday = today - timedelta(days=1)  # yesterday's date
    nameDateYesterday = nameDateToday - timedelta(days=1)  # Yesterday's date for file naming
    prises = get_price_list(dt=yesterday)  # gets latest Bhav NSE copy list
    # saving of file in CSV format with name
    prises.to_csv(os.path.join(path, 'cm' + nameDateYesterday.strftime("%d%b%Y") + 'bhav.csv'))
