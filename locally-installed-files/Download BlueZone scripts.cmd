CLS
@ECHO OFF
Title ES BlueZone Scripts Power Pad Installer

REM ===================================================================================================
REM This is a Windows command script that will install a BlueZone application on the user's desktop.
REM The BlueZone script configuration is et up, will display the power pad you've chosen.
REM ===================================================================================================

REM ====================================================================================================
REM Updated 12/16/2020 by Ilse Ferris
REM Reason for update: Added support for confirming if BlueZone is downloaded, if connection to the T drive exists, and file access exists. Suppresses error messages. 
REM ===================================================================================================

REM ====================================================================================================
REM Updated 08/06/2020 by Ilse Ferris
REM Reason for update: Removed ELSE statements. Some users do not have OneDrive/C Drives fully linked.
REM ===================================================================================================

REM ====================================================================================================
REM Updated 08/04/2020 by Ilse Ferris
REM Reason for update: Commented out coding to remove 'old scripts' or .bzs files. All users should no  
REM		longer have these files on their C drives or network drives. 
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
REM VERIFY BLUEZONE IS INSTALLED AND COMPUTER IS CONNECTED TO THE T:\ DRIVE
REM ===================================================================================================
:START
IF NOT EXIST "C:\Program Files (x86)\BlueZone\7.1\bzmd.exe" GOTO BLUEZONE_NOT_INSTALLED
REM IF NOT EXIST "T:\" GOTO NOT_CONNECTED_TO_T_DRIVE
IF NOT EXIST "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\" GOTO NOT_CONNECTED_TO_T_DRIVE

REM IF NOT EXIST "T:\Eligibility Support\Scripts\Hennepin.zmd" GOTO ES_ACCESS_ERROR
IF NOT EXIST "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Eligibility Support\Scripts\Hennepin.zmd" GOTO ES_ACCESS_ERROR

REM ===================================================================================================
REM VERIFY BLUEZONE IS INSTALLED AND COMPUTER IS CONNECTED TO THE T:\ DRIVE
REM ===================================================================================================

CLS
ECHO Which Economic Supports power pad would you like to install?
ECHO.
ECHO Select 1. Install the BlueZone Scripts Power Pad
ECHO Select 2. Install the Specialty Power Pad
ECHO Select 3. Cancel installation
ECHO.
SET CHOICE=
SET /p CHOICE=Select a number and press enter...
IF NOT '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO CHOICE_ES_BLUEZONE_SCRIPTS_POWER_PAD
IF '%CHOICE%'=='2' GOTO CHOICE_ES_SPECIALTY_POWER_PAD
IF '%CHOICE%'=='3' GOTO CHOICE_CANCEL
ECHO "%choice%" is not a valid choice, try again
ECHO.
GOTO CHOICE_BUSINESS_AREA_ES 
ECHO.

REM ===================================================================================================
REM CHOICE_ES_BLUEZONE_SCRIPTS_POWER_PAD - Install the ES BlueZone Scripts Power Pad
REM ===================================================================================================
:CHOICE_ES_BLUEZONE_SCRIPTS_POWER_PAD
CLS

Taskkill /IM bzmd.exe 2> NUL 

REM ------------------------------------Removing Hennepin session from desktop. "2> NUL" coding suppresses error messages. 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" 2> NUL
DEL /Q "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 

DEL /Q "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd" 2> NUL

REM -----------------------------------If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"
IF EXIST "%userprofile%\Desktop\Hennepin.zmd" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\Desktop\Hennepin.zmd"

REM CLS
ECHO.
ECHO Installation complete. 
ECHO.
ECHO If your Hennepin BlueZone session is not on your desktop, please contact the BZST at: HSPH.EWS.BlueZoneScripts@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_ES_SPECIALTY_POWER_PAD - Install the ES Specialty Power Pad
REM ===================================================================================================
:CHOICE_ES_SPECIALTY_POWER_PAD 
CLS

Taskkill /IM bzmd.exe 2> NUL 

REM ------------------------------------Removing Hennepin session from desktop. "2> NUL" coding suppresses error messages. 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" 2> NUL
DEL /Q "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 

DEL /Q "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd" 2> NUL

REM ------------------------------------If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"
IF EXIST "%userprofile%\Desktop\Hennepin-Speciality.zmd" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\Desktop\Hennepin-Speciality.zmd"

REM CLS
ECHO.
ECHO Installation complete. 
ECHO.
ECHO If your Hennepin Specialty BlueZone session is not on your desktop, please contact the BZST at: HSPH.EWS.BlueZoneScripts@Hennepin.us.
GOTO END

REM ========================================================================
REM CHOICE_CANCEL - CANCEL INSTALLATION
REM ========================================================================
:CHOICE_CANCEL
CLS
ECHO.
ECHO Installation canceled.
GOTO END

REM ========================================================================
REM BLUEZONE_NOT_INSTALLED
REM ========================================================================
:BLUEZONE_NOT_INSTALLED
ECHO.
ECHO ***ERROR***
ECHO.
ECHO BlueZone must be installed for the power pad installer to work properly.
ECHO.
ECHO Exit the installer, install BlueZone (in the Software Center), then run this power pad installer again.
ECHO.
GOTO END

REM ========================================================================
REM ES_ACCESS_ERROR
REM ========================================================================
:ES_ACCESS_ERROR
ECHO.
ECHO ***ERROR***
ECHO.
ECHO Installation canceled. 
ECHO.
ECHO Access was denied to T:\Eligibility Support\Scripts.
ECHO.
ECHO Contact your supervisor or the Help Desk to resolve the issue.
ECHO.
GOTO END

REM ========================================================================
REM NOT_CONNECTED_TO_T_DRIVE
REM ========================================================================
:NOT_CONNECTED_TO_T_DRIVE
ECHO.
ECHO ***ERROR***
ECHO.
ECHO You are not connected to the T: drive, so the installer cannot run.
ECHO.
ECHO Reconnect to the T: drive then press any key in the installer to try again.
ECHO.
ECHO Contact the Help Desk if you need assistance reconnecting to the T: drive.
ECHO.
PAUSE
CLS
GOTO START

REM ========================================================================
REM END
REM ========================================================================
:END
ECHO.
PAUSE