@echo off
echo Hello Mrs.Simpson, press enter to collect your applications.
pause
echo Collecting applications...
python pythonProject\readGoogleSheets\ReadWriteGoogleSheets.py
echo Done! Your applications have been collected, press enter to exit.
pause

::Outline of functionality:
::1. Run python script