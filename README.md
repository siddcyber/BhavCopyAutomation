# BhavCopyAutomation

Automation tool for Downloading latest NSE Bhav Copy with GUI Settings

Just use exe files Executables folder in case you do not want to go through the code and open BhavCopyAutoDownloader.exe to start downloading the Latest NSE BhavCopy

1. Packages used: Tkinter, pandas, requests, zipfile, os, pywin32, winshell, Time


2. The BhavCopyAutoDownloader file is the main code to download the Latest NSE BhavCopy from official website.


3. The BhavCopyAutoSettings file is the GUI used to specify the file download location and if download should start on startup or not.


4. The source code is in python


5. The executable files BhavCopyAutoSettings.exe, BhavCopyAutoDownloader.exe as well as BhavCopy_location.txt, BhavCopyLogo.ico and
   Sample_BhavCopy.csv should be in same folder for the program to work


6. In every run the program will check if download location is present or not and opens settings for the program to set the download destination and/or set the program on startup and save and Exit.The program will restart after 60-80 seconds.


7. In case of any connection error the program will retry after every 10 seconds until the file is downloaded.


8. Do not delete the BhavCopy_location.txt it contains the download location of the file.


9. If you Do not need the python scripts then delete and use only the exe files as they are the executable version of the python script


10. If you want the program to start with startup then click on "download on system startup" in the BhavCopyAutoSettings.exe
 

update:
1. new function to search the internet and skip past the holidays in the year implemented.
2. the previous method was inefficient, a change in website and adding badZipFile exception handling fixed the problem with holidays 