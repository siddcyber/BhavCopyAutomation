# install packages: pywin32, pandas, requests, future, tkcalendar and winshell

import os
import requests
from datetime import date, datetime, timedelta
from zipfile import ZipFile, BadZipfile
import time
from tkcalendar import DateEntry
from tkinter import filedialog, Label, Radiobutton, BooleanVar, Tk, Button, ttk, Scrollbar, LabelFrame
import os.path
import pandas as pd


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
            response = requests.get(url)  # download the data behind the URL
            open(filename, "wb").write(response.content)  # Open the response into a new file
            # extract zip file to specified location
            print('file extraction start')
            with ZipFile(filename, 'r') as zip_file:
                zip_file.extractall(path=path)
                print('file extracted')
                print('sleeping for 0.1 second')
                time.sleep(0.1)
                print('complete')
            print('deleting file')
            os.remove(filename)  # removes the downloaded zip file
            print('file deleted')
            break
        except requests.exceptions.ConnectionError:
            print("preSleepConnectionErrorCatchSuccess")
            time.sleep(10)
            print("postSleepConnectionErrorCatchSuccess")


def checkWeekend(day):
    weekend = False
    if day.weekday() == 6 or day.weekday() == 5:
        weekend = True
    return weekend


def downloadRange(start, end):
    deltaDate = end - start  # returns deltatime
    print('start')
    for i in range(deltaDate.days + 1):
        currentDay = start + timedelta(days=i)
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


# function to change properties of button on hover
class ChangeOnHover:
    def __init__(self, button, colorOnHover, colorOnLeave):
        # adjusting background of the widget background on entering widget
        button.bind("<Enter>", func=lambda e: button.config(background=colorOnHover))
        # background color on leaving widget
        button.bind("<Leave>", func=lambda e: button.config(background=colorOnLeave))


class WidgetConfig:
    def __init__(self, widget, row, col):
        if row == 'row':
            widget.rowconfigure(index=3, weight=1)
        else:
            pass
        if col == 'col':
            widget.columnconfigure(index=0, weight=1)
        else:
            pass


# function to check if BhavCopy_location.txt file is located or  not
def CheckOrignalFilePath():
    import winshell
    if os.path.exists(os.getcwd() + '\\' + 'BhavCopy_location.txt'):
        #  if file is present then reads file
        with open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'r') as file:
            orignalPath = str(file.readline())
    elif not os.path.exists(os.getcwd() + '\\' + 'BhavCopy_location.txt'):
        #  if file not present creates empty file and assigns orignal path as desktop
        with open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'w+') as file:
            orignalPath = winshell.desktop()
            file.write(str(orignalPath))  # append the file with new location
    file.close()
    return orignalPath


def BrowseFile():
    # opens file browser
    filename = filedialog.askdirectory(initialdir=str(CheckOrignalFilePath()), title="select the folder")
    labelFilename.configure(text=filename)  # reconfigures the Filename label with new location in GUI
    #   if folder not present create folder
    newPath = open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'w')  # opens txt file in write mode
    newPath.truncate(0)  # erase all the contents of the file
    newPath.write(str(labelFilename.cget('text')))  # append the file with new location
    newPath.close()


def confrirmAndDownload():
    #     start progressbar and progresbar status
    startDate = cal.get_date()
    # deltaDate = endDate-startDate # returns deltatime
    if not endDateCheck.get():
        endDate = calEnd.get_date()
    else:
        endDate = date.today()
        if endDate.weekday() == 6:
            endDate = endDate - timedelta(days=2)  # Friday's date for file naming
        elif endDate.weekday() == 0:
            endDate = endDate - timedelta(days=3)  # Friday's date for file naming
        else:
            endDate = endDate - timedelta(days=1)  # Yesterday's date for file naming
    rangeFrameLabel.configure(
        text="range selected from " + startDate.strftime("%d %b %Y") + " to " + endDate.strftime("%d %b %Y"))
    downloadRange(startDate, endDate)


window = Tk()
endDateCheck = BooleanVar()
endDateCheck.set(False)
window.geometry('700x360')
window.config(background="White")

window.title("BhavCopy Range Downloader")
window.iconbitmap('BhavCopyLogo.ico')


heading = Label(window, text="BhavCopy Range Downloader", relief='raised', background="White")
sampleHeading = Label(window, text="\nSample Downloaded NSE Bhav Copy:", background="White")

browserFrame = LabelFrame(window,background = 'White')
labelFilename = Label(browserFrame,background = 'White')
labelFilename.configure(text=CheckOrignalFilePath())
fileBrowserButton = Button(browserFrame, text='File Browser', command=BrowseFile, background="White")

# frame with all the widgets relating to range downloading
rangeFrame = LabelFrame(window, background='White')
rangeFrameLabel = Label(rangeFrame, text='Select the range to be Downloaded', background='White')
today = date.today()
cal = DateEntry(rangeFrame, selectmode='day', year=today.year, month=today.month, day=today.day)
calEnd = DateEntry(rangeFrame, selectmode='day', year=today.year, month=today.month, day=today.day)
rangeFrameStartDate = Label(rangeFrame, text="select the start date: ", background='White')
rangeFrameEndDate = Label(rangeFrame, text="select the End date:", background='White')
rangeFrameButton = Button(rangeFrame, text="Confirm\nAnd\nDownload", command=confrirmAndDownload, background='White')
rangeFrameRadioEnd = Radiobutton(rangeFrame, text="End Date",background = 'White',variable =endDateCheck,value=False)
rangeFrameRadioToday =Radiobutton(rangeFrame, text="Today",background = 'White',variable =endDateCheck,value=True)

# selecting file as a sample Downloaded Bhav Copy (first 5 entries)
treeframe = LabelFrame(window)
Dataset = pd.read_csv('Sample_BhavCopy.csv').head()
tv1 = ttk.Treeview(treeframe)
tv1.place(relheight=1, relwidth=1)
treescrolly = Scrollbar(treeframe, orient="vertical", command=tv1.yview)
treescrollx = Scrollbar(treeframe, orient="horizontal", command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")
tv1["column"] = list(Dataset.columns)
tv1["show"] = "headings"
for columns in tv1["column"]:
    tv1.heading(columns, text=columns)
df_rows = Dataset.to_numpy().tolist()
for row in df_rows:
    tv1.insert("", "end", values=row)

exitButton = Button(window, text='Exit', command=window.destroy, background="White")

# Color changes on hover from white to light gray
ChangeOnHover(fileBrowserButton, "Light Gray", "White")
ChangeOnHover(exitButton, "Light Gray", "White")
ChangeOnHover(rangeFrameButton, "Light Gray", "White")
# widget row/column Config
WidgetConfig(window, "row", "col")
WidgetConfig(browserFrame, '', 'col')
WidgetConfig(rangeFrame, 'row', 'col')

# change to grid
heading.grid(row=0, column=0, sticky='NSEW')
sampleHeading.grid(row=2, column=0, sticky='NSEW')
treeframe.grid(row=3, column=0, sticky='EW', ipady=60)

browserFrame.grid(row=4, column=0, sticky='EW')
labelFilename.grid(row=0, column=0, sticky='W')
fileBrowserButton.grid(row=0, column=1, sticky='E')

rangeFrame.grid(row=5,column=0,sticky='NSEW')
rangeFrameLabel.grid(row=0,column=0,columnspan=8,sticky='NSEW')
rangeFrameStartDate.grid(row=1,column=0,sticky='W')
cal.grid(row=1,column=1,sticky='EW')
rangeFrameEndDate.grid(row=1, column=2, sticky='NSEW')
rangeFrameRadioEnd.grid(row=1, column=3, sticky='NSEW')
calEnd.grid(row=1, column=4, sticky='EW')
rangeFrameRadioToday.grid(row=1, column=5, sticky='E')
rangeFrameButton.grid(row=1, column=6, sticky='E')

exitButton.grid(row=7, column=0, sticky='NSEW')

window.mainloop()
