@echo off
echo Now we will try moving a file from Google Drive to our test directory
::Powershell.exe
pause
copy "C:\Users\talop\My Drive (thalia.edwards2203@gmail.com)\application-form-responses.gsheet" C:\Users\talop\Documents\intro-to-batch\testDir
pause
cd /d C:\Users\talop\Documents\intro-to-batch\testDir
dir
pause

::Outline of functionality:
::1. Grab file from google drive
::2. Extracts info and resume document from sheet
::3. 