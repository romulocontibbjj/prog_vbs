@ ECHO OFF
ECHO Copying ActiveGUI.ocx ...
copy ActiveGUI.ocx %SYSTEMROOT%\system32\
regsvr32 %SYSTEMROOT%\system32\ActiveGUI.ocx

ECHO Copying ASIFce1.ocx
copy AS-IFce1.ocx %SYSTEMROOT%\system32\
regsvr32 %SYSTEMROOT%\system32\AS-IFce1.ocx

ECHO Copying ADockedForm.ocx
copy DockedForm.ocx %SYSTEMROOT%\system32\
regsvr32 %SYSTEMROOT%\system32\DockedForm.ocx


