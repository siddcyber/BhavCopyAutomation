# install pywin32, future, nsepy and winshell

from tkinter import filedialog, Label, Checkbutton, IntVar, Tk, Button, ttk, Scrollbar, LabelFrame
import os.path
import pandas as pd


#  function to create shortcut in the specified path for the BhavCopyAutoDownloader.exe
def createShortcut(path, target='', wDir='', icon=''):
    from win32com.client import Dispatch
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    if icon == '':
        pass
    else:
        shortcut.IconLocation = icon
    shortcut.save()


# function to change properties of button on hover
def changeOnHover(button, colorOnHover, colorOnLeave):
    # adjusting background of the widget background on entering widget
    button.bind("<Enter>", func=lambda e: button.config(background=colorOnHover))
    # background color on leaving widget
    button.bind("<Leave>", func=lambda e: button.config(background=colorOnLeave))


# function to check if BhavCopy_location.txt file is located or  not
def CheckOrignalFilePath():
    import winshell
    if os.path.exists(os.getcwd() + '\\' + 'BhavCopy_location.txt'):
        #  if file is present then reads file
        with open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'r') as file:
            orignalPath = str(file.readline())
    elif not os.path.exists(os.getcwd() + '\\' + 'BhavCopy_location.txt'):
        #  if file not present creates empty file and assigns orignal path as desktop
#   if folder not present create folder
        with open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'w+') as file:
            orignalPath = winshell.desktop()
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


# function to add application to startup
def StartupFunctionAndExit():
    import getpass
    # check if checkbox is clicked or not
    if startupCheck.get() == 1:  # checkbox clicked
        USER_NAME = getpass.getuser()
        startupPath = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start ' \
                      r'Menu\Programs\Startup\BhavCopyAutoDownloader.lnk' % USER_NAME
        file_path = str(os.getcwd()) + "\BhavCopyAutoDownloader.exe"
        createShortcut(startupPath, file_path, str(os.getcwd()), '')
        window.destroy()
    elif startupCheck.get() == 0:  # checkbox not clicked
        USER_NAME = getpass.getuser()
        startupPath = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start ' \
                      r'Menu\Programs\Startup\BhavCopyAutoDownloader.lnk' % USER_NAME
        try:
            os.remove(startupPath)  # removes shortcut if unselected
        except FileNotFoundError:
            pass  # handles exception in case file was no found(deleted/never made)
        window.destroy()


#  function to check if shortcut file is present in the startup folder or not
def CheckStartupFile():
    import getpass
    USER_NAME = getpass.getuser()
    startupPath = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start ' \
                  r'Menu\Programs\Startup\BhavCopyAutoDownloader.lnk' % USER_NAME
    if os.path.exists(startupPath):
        filePresent = True
    else:
        filePresent = False
    return filePresent


window = Tk()
startupCheck = IntVar()
window.geometry('700x330')
window.config(background="White")

window.title("BhavCopy Download Automation Settings GUI")
window.iconbitmap('BhavCopyLogo.ico')

window.columnconfigure(index=0, weight=1)
# window.rowconfigure(index, weight)

heading = Label(window, text="BhavCopy Download Automation Settings", relief='raised', background="White")
sampleHeading = Label(window, text="\nSample Downloaded NSE Bhav Copy:", background="White")
treeframe = LabelFrame(window)
browserFrame = LabelFrame(window)
browserFrame.columnconfigure(index=0, weight=1)

labelFilename = Label(browserFrame)
labelFilename.configure(text=CheckOrignalFilePath())
fileBrowserButton = Button(browserFrame, text='File Browser', command=BrowseFile, background="White")

startupTrue = Checkbutton(window, text='Download On System Startup', variable=startupCheck, onvalue=1, offvalue=0,
                          background="White")

if CheckStartupFile():
    startupTrue.select()
else:
    pass

# selecting file as a sample Downloaded Bhav Copy (first 5 entries)
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

exitButton = Button(window, text='Save and Exit', command=StartupFunctionAndExit, background="White")

# Color changes on hover from white to light gray
changeOnHover(fileBrowserButton, "Light Gray", "White")
changeOnHover(exitButton, "Light Gray", "White")

# change to grid
heading.grid(row=0, column=0, sticky='NSEW')
sampleHeading.grid(row=2, column=0)
treeframe.grid(row=3, column=0, sticky='EW', ipady=60)

browserFrame.grid(row=4, column=0, sticky='EW')
labelFilename.grid(row=0, column=0, sticky='W')
fileBrowserButton.grid(row=0, column=1, sticky='E')

startupTrue.grid(row=5, column=0, sticky='NSEW')
exitButton.grid(row=6, column=0, sticky='NSEW')

window.mainloop()
