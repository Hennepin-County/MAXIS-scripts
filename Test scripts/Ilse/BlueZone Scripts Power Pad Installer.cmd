CLS 

@ ECHO OFF
title Bluezone Scripts Power Pad Installer

REM ===================================================================================================
REM This is a Windows command script that will install a BlueZone configuration on the user's desktop
REM The BlueZone configurations are set up to display the appropriate power pad
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
ECHO This installer will copy a version of BlueZone configured to support the BlueZone script tool on your desktop.
ECHO Close any open BlueZone sessions before installing a power pad.
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
Taskkill /F /IM bzmd.exe
ECHO.

REM ---Deleting the .bzs script files that may have been installed previously 
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts" DEL "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts"
REM IF EXISTS "%userprofile%\Documents\BlueZone\Scripts" DEL "%userprofile%\Documents\BlueZone\Scripts"
REM 
REM IF EXISTS "C:\Bluezone_HSR_Scripts" DEL "C:\Bluezone_HSR_Scripts"

REM REM ---Deleting any previous version of BlueZone with BlueZone Scripts from the Desktop
REM REM Removing Hennepin .zmd
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
REM IF EXISTS "%userprofile%\Desktop\Hennepin.zmd" DEL "%userprofile%\Desktop\Hennepin.zmd"
REM 
REM REM Removing Specialty .zmd
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Specialty.zmd" DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Specialty.zmd"
REM IF EXISTS "%userprofile%\Desktop\Hennepin-Specialty.zmd" DEL "%userprofile%\Desktop\Hennepin-Specialty.zmd"

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
ECHO A new BlueZone icon has been added to your desktop.
GOTO END

REM ===================================================================================================
REM OPTION TWO - Install the Specialty Power Pad
REM ===================================================================================================
:OPTION_TWO
@ECHO OFF
Taskkill /F /IM bzmd.exe
ECHO.

REM ---Deleting the .bzs script files that may have been installed previously 
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts" DEL "%userprofile%\OneDrive - Hennepin County\My Documents\Documents\BlueZone\Scripts"
REM IF EXISTS "%userprofile%\Documents\BlueZone\Scripts" DEL "%userprofile%\Documents\BlueZone\Scripts"
REM 
REM IF EXISTS "C:\Bluezone_HSR_Scripts" DEL "C:\Bluezone_HSR_Scripts"
REM 
REM REM Deleting any previous version of BlueZone with BlueZone Scripts from the Desktop
REM REM Removing General version 
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
REM IF EXISTS "%userprofile%\Desktop\Hennepin.zmd" DEL "%userprofile%\Desktop\Hennepin.zmd"
REM 
REM REM Removing Specialty version 
REM IF EXISTS "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd" DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"
REM IF EXISTS "%userprofile%\Desktop\Hennepin-Speciality.zmd" DEL "%userprofile%\Desktop\Hennepin-Speciality.zmd"

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
ECHO A new BlueZone icon has been added to your desktop.
GOTO 

:END
ECHO.
PAUSE 