CLS 
@ ECHO OFF
title EWS Bluezone Scripts Power Pad Installer

REM ===================================================================================================
REM This is a Windows command script that will install a BlueZone application on the user's desktop.
REM The BlueZone script configuration is et up, will display the power pad you've chosen.
REM ===================================================================================================

REM ====================================================================================================
REM Created 11/02/2019 by Ilse Ferris
REM Reason for update: Desktop and documents folders are being synced soon with OneDrive, which 
REM		changes their path location. The installer will now install the files to the OneDrive-synced desktop 
REM		if it exists, otherwise to the normal desktop. The installer will now install the .zmd files to 
REM		%userprofile%\desktop, which is not synced to OneDrive. The end result is that the installer should
REM		function correctly if the user regardless of whether the user has fully completed the migration to 
REM		OneDrive or if they have not.
REM ===================================================================================================

REM ===================================================================================================
REM MENU - SELECT AN OPTION
REM ===================================================================================================

:START

ECHO.
ECHO This installer will copy a version of BlueZone application configured to support the BlueZone scripts on your desktop.
ECHO Close any open BlueZone sessions before installing a power pad. Choose which power pad you wish to install.
ECHO.
ECHO Select 1. Install the BlueZone Scripts Power Pad
ECHO Select 2. Install the Specialty Power Pad
ECHO Select 3. CANCEL INSTALLATION
ECHO.

SET CHOICE=
SET /p CHOICE=Select a number and press enter... 
if not '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO OPTION_ONE
IF '%CHOICE%'=='2' GOTO OPTION_TWO
IF '%CHOICE%'=='3' GOTO END 
ECHO "%choice%" is not valid please try again
ECHO.
GOTO START

REM ===================================================================================================
REM OPTION ONE - Install the BlueZone Scripts Power Pad
REM ===================================================================================================
:OPTION_ONE
@ECHO OFF
Taskkill /IM bzmd.exe 2> NUL 

REM ---Deleting the .bzs script files that may have been installed previously 
RD /S/Q "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts"
RD /S/Q "%userprofile%\Documents\BlueZone\Scripts"
RD /S/Q "C:\Bluezone_HSR_Scripts" 

REM Removing Hennepin session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
    DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
) ELSE (
    DEL /Q "%userprofile%\Desktop\Hennepin.zmd"
)

REM Removing Hennepin-Specialty session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
    DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"
) ELSE (
    DEL /Q "%userprofile%\Desktop\Hennepin-Speciality.zmd"
)

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\Desktop\Hennepin.zmd"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A the new Hennepin icon has been added to your desktop.
GOTO END

REM ===================================================================================================
REM OPTION TWO - Install the Specialty Power Pad
REM ===================================================================================================
:OPTION_TWO
@ECHO OFF
Taskkill /IM bzmd.exe 2> NUL 

REM ---Deleting the .bzs script files that may have been installed previously 
RD /S/Q "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts"
RD /S/Q "%userprofile%\Documents\BlueZone\Scripts"
RD /S/Q "C:\Bluezone_HSR_Scripts"
    
REM Removing Hennepin session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
    DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
) ELSE (
    DEL /Q "%userprofile%\Desktop\Hennepin.zmd"
)

REM Removing Hennepin-Specialty session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
    DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"
) ELSE (
    DEL /Q "%userprofile%\Desktop\Hennepin-Speciality.zmd"
)

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\Desktop\Hennepin-Speciality.zmd"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A the new Hennepin-Specialty icon has been added to your desktop.
GOTO END

:END
ECHO.
PAUSE 