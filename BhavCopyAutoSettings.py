# install pywin32, future, nsepy and winshell

from tkinter import filedialog, Label, Checkbutton, IntVar, Tk, Button, ttk, Scrollbar, LabelFrame
import os.path

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
    file.close()
    return orignalPath


#  function to open the file browser to assign directory and replace/append the BhavCopy_location.txt file
#  with new location
def BrowseFile():
    # opens file browser
    filename = filedialog.askdirectory(initialdir=str(CheckOrignalFilePath()), title="select the folder")
    labelFilename.configure(text=filename)  # reconfigures the Filename label with new location in GUI

    newPath = open(os.getcwd() + '\\' + 'BhavCopy_location.txt', 'w')   # opens txt file in write mode
    newPath.truncate(0)                                                 # erase all the contents of the file
    newPath.write(str(labelFilename.cget('text')))                      # append the file with new location
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


def ShowExcel():
        import pandas as pd
        Dataset = pd.read_csv('Sample_BhavCopy.csv').head()
        treeframe = LabelFrame(window)
        treeframe.grid(row=1, column=0, sticky='EW', columnspan=100, ipady=60)
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
        window.update()


window = Tk()
startupCheck = IntVar()

window.title("BhavCopy Download Automation Settings GUI")
window.iconbitmap('BhavCopyLogo.ico')

heading = Label(window, text="BhavCopy Download Automation Settings")
labelFilename = Label(window)
labelFilename.configure(text=CheckOrignalFilePath())
fileBrowserButton = Button(window, text='File Browser', command=BrowseFile)
startupTrue = Checkbutton(window, text='Download On System Startup', variable=startupCheck, onvalue=1, offvalue=0)
if CheckStartupFile():
    startupTrue.select()
else:
    pass

exitButton = Button(window, text='Save and Exit', command=StartupFunctionAndExit)
# change to grid
heading.grid(row=0, column=0, columnspan=10)
labelFilename.grid(row=5, column=0, columnspan=7)
fileBrowserButton.grid(row=5, column=8, columnspan=2)
startupTrue.grid(row=6, column=0, columnspan=10)
exitButton.grid(row=7, column=3, columnspan=3)

window.mainloop()
