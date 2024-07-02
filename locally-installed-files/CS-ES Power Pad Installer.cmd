CLS
@ECHO OFF
title BlueZone Scripts Power Pad Installer

REM ===================================================================================================
REM Created 11/4/2020 by Kyle Nelson Version 0.0
REM 	This is a Windows command script that installs a BlueZone configuration file on the user's
REM 	computer. The BlueZone configurations are set up to display the appropriate power pad. This
REM 	combined CS-ES power pad installer was created so that it can be added to the Software Center.
REM 	Currently, users access a power pad installer from a link on a corresponding SharePoint page,
REM	but these links do not function in Edge. Edge is being used more and more and will become the
REM	default browser, so putting this in the Software Center should fix this problem.
REM Updated 1/26/2021 by Kyle Nelson and Ilse Ferris Version 1.0
REM 	Updated to clean up and simplify some of the coding and to improve the text displayed in the
REM	Command Prompt. Added contact email address at the end message for Child Support users.
REM Updated 3/15/2021 by Kyle Nelson 
REM	Updated - if user is not connected to T: drive, script would previously loop indefinitely
REM	(each time user pressed a key) until they were connected to the T: drive or X'd out of the
REM	script. Per recommendation of John Parizek, who is packaging the script for the Software
REM	Center, I removed the infinite loop because X'ing out of the script will cause the Software
REM	Center to say that installation of the script failed. Now if the user is not connected to the
REM	T: drive, the script will display a message indicating as such and then when they press any
REM	key the script will end. The user will need to reconnect to the T: drive and run the installer
REM	again.
REM Updated 3/26/2021 by Kyle Nelson
REM	Updated - removed an extraneous k in the location the script checks to verify connection to 
REM	the T:\ drive.
REM Updated 06/12/2024 by Ilse Ferris 
REM     Updated BlueZone application path from 7.1 to 10.1. 
REM ===================================================================================================

REM ===================================================================================================
REM VERIFY BLUEZONE IS INSTALLED AND COMPUTER IS CONNECTED TO THE T:\ DRIVE
REM ===================================================================================================
:START
IF NOT EXIST "C:\Program Files (x86)\BlueZone\10.1\bzmd.exe" GOTO BLUEZONE_NOT_INSTALLED
IF NOT EXIST "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\" GOTO NOT_CONNECTED_TO_T_DRIVE

REM ===================================================================================================
REM SELECT A BUSINESS AREA (ES OR CS)
REM ===================================================================================================
ECHO.
ECHO This installer will create a shortcut to BlueZone on your desktop.
ECHO.
ECHO The shortcut will be configured to support BlueZone scripts for your area.
ECHO.
ECHO Close any open BlueZone sessions before proceeding.
ECHO.
ECHO Which area do you work in?
ECHO.
ECHO Select 1. Economic Supports
ECHO Select 2. Child Support
ECHO Select 3. Cancel installation
ECHO.
SET CHOICE=
SET /p CHOICE=Select a number and press enter...
IF NOT '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO CHOICE_BUSINESS_AREA_ES
IF '%CHOICE%'=='2' GOTO CHOICE_BUSINESS_AREA_CS
IF '%CHOICE%'=='3' GOTO CHOICE_CANCEL
ECHO "%choice%" is not a valid choice, try again
ECHO.
GOTO START

REM ===================================================================================================
REM CHOICE_BUSINESS_AREA_ES 
REM ===================================================================================================
:CHOICE_BUSINESS_AREA_ES
IF NOT EXIST "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Eligibility Support\Scripts\Hennepin.zmd" GOTO ES_ACCESS_ERROR

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

REM ------------------------------------Removing Hennepin session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd"  2> NUL 
IF EXIST "%userprofile%\Desktop\Hennepin.zmd" DEL /Q "%userprofile%\Desktop\Hennepin.zmd"  2> NUL 

REM -----------------------------------Removing Hennepin-Specialty session from desktop 
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd"  2> NUL 
IF EXIST "%userprofile%\Desktop\Hennepin-Speciality.zmd" DEL /Q "%userprofile%\Desktop\Hennepin-Speciality.zmd" 2> NUL 

REM -----------------------------------If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" 2> NUL 
IF EXIST "%userprofile%\Desktop\Hennepin.zmd" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin.zmd" "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 

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
Taskkill /IM bzmd.exe 2> NUL 

REM ------------------------------------Removing Hennepin session from desktop 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin.zmd" 2> NUL 
DEL /Q "%userprofile%\Desktop\Hennepin.zmd" 2> NUL 

REM -----------------------------------Removing Hennepin-Specialty session from desktop 
DEL /Q "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd" 2> NUL 
DEL /Q "%userprofile%\Desktop\Hennepin-Speciality.zmd" 2> NUL 

REM ------------------------------------If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\Hennepin-Speciality.zmd" 2> NUL 
IF EXIST "%userprofile%\Desktop\Hennepin-Speciality.zmd" COPY /Y "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Hennepin-Speciality.zmd" "%userprofile%\Desktop\Hennepin-Speciality.zmd" 2> NUL 

REM CLS
ECHO.
ECHO Installation complete. 
ECHO.
ECHO If your Hennepin Specialty BlueZone session is not on your desktop, please contact the BZST at: HSPH.EWS.BlueZoneScripts@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_BUSINESS_AREA_CS
REM ===================================================================================================
:CHOICE_BUSINESS_AREA_CS
IF NOT EXIST "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\BlueZone Scripts.zmd" GOTO CS_ACCESS_ERROR

CLS
ECHO.
ECHO Which Child Support power pad would you like to install?
ECHO.
ECHO Select 1. Clerical Power Pad
ECHO Select 2. Scripts Power Pad
ECHO Select 3. Cancel installation
ECHO.
SET CHOICE=
SET /p CHOICE=Select a number and press enter...
IF NOT '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO CHOICE_CS_CLERICAL_POWER_PAD
IF '%CHOICE%'=='2' GOTO CHOICE_CS_SCRIPTS_POWER_PAD
IF '%CHOICE%'=='3' GOTO CHOICE_CANCEL
ECHO "%choice%" is not a valid choice, try again
ECHO.
GOTO CHOICE_BUSINESS_AREA_CS
ECHO.

REM ===================================================================================================
REM CHOICE_CS_CLERICAL_POWER_PAD
REM ===================================================================================================
:CHOICE_CS_CLERICAL_POWER_PAD
CLS
ECHO.
ECHO CUSTOM SETTINGS 1: BlueZone will be configured with some optional settings that you may find helpful. 
ECHO The Enter key will perform the Enter function.
ECHO.
ECHO CUSTOM SETTINGS 2: BlueZone will be configured with some optional settings that you may find helpful. 
ECHO The Enter key will perform the New Line function.
ECHO.
ECHO DEFAULT SETTINGS: BlueZone will be configured with default settings. 
ECHO The Enter key will perform the New Line function.
ECHO.
ECHO STATE ICON INFO
ECHO BlueZone will be configured with default settings without access to the power pad or scripts.
ECHO.
ECHO SETTINGS CHOICES
ECHO Select 1. Install the Clerical Power Pad (custom settings 1)
ECHO Select 2. Install the Clerical Power Pad (custom settings 2)
ECHO Select 3. Install the Clerical Power Pad (default settings)
ECHO Select 4. Install the STATE icon
ECHO Select 5. Cancel installation
ECHO.
SET CHOICE=
SET /p CHOICE=Select a number and press enter...
IF NOT '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_1
IF '%CHOICE%'=='2' GOTO CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_2
IF '%CHOICE%'=='3' GOTO CHOICE_CS_CLERICAL_DEFAULT_SETTINGS
IF '%CHOICE%'=='4' GOTO CHOICE_CS_STATE_ICON
IF '%CHOICE%'=='5' GOTO CHOICE_CANCEL
ECHO "%choice%" is not a valid choice, try again
ECHO.
GOTO CHOICE_CS_CLERICAL_POWER_PAD
ECHO.

REM ===================================================================================================
REM CHOICE_CS_SCRIPTS_POWER_PAD
REM ===================================================================================================
:CHOICE_CS_SCRIPTS_POWER_PAD
CLS
ECHO.
ECHO CUSTOM SETTINGS 1 INFO
ECHO BlueZone will be configured with a number of optional settings that you may find helpful.
ECHO The Enter key will perform the Enter function.
ECHO.
ECHO CUSTOM SETTINGS 2 INFO
ECHO BlueZone will be configured with a number of optional settings that you may find helpful.
ECHO The Enter key will perform the New Line function.
ECHO.
ECHO DEFAULT SETTINGS INFO
ECHO BlueZone will be configured with default settings. 
ECHO The Enter key will perform the New Line function.
ECHO.
ECHO STATE ICON INFO
ECHO BlueZone will be configured with default settings without access to the power pad or scripts.
ECHO.
ECHO SETTINGS CHOICES
ECHO Select 1. Install the Scripts Power Pad (custom settings 1)
ECHO Select 2. Install the Scripts Power Pad (custom settings 2)
ECHO Select 3. Install the Scripts Power Pad (default settings)
ECHO Select 4. Install the STATE icon
ECHO Select 5. Cancel installation
ECHO.
SET CHOICE=
SET /p CHOICE=Select a number and press enter...
IF NOT '%choice%'=='' set choice=%choice:~0,1%
IF '%CHOICE%'=='1' GOTO CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_1
IF '%CHOICE%'=='2' GOTO CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_2
IF '%CHOICE%'=='3' GOTO CHOICE_CS_SCRIPTS_DEFAULT_SETTINGS
IF '%CHOICE%'=='4' GOTO CHOICE_CS_STATE_ICON
IF '%CHOICE%'=='5' GOTO CHOICE_CANCEL
ECHO "%choice%" is not a valid choice, try again
ECHO.
GOTO CHOICE_CS_SCRIPTS_POWER_PAD
ECHO.


REM ===================================================================================================
REM CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_1 (DESKTOP ICON NAME: CLERICAL+)
REM ===================================================================================================
:CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_1
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical+.zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical+ (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\Clerical Scripts+.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical+.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical+ (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\Clerical Scripts+.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts+.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical+.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical+ (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts+.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical+.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical+ (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical Scripts+.lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Clerical Scripts+.zmd" "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts+.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts+.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts+.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts+.lnk" "%userprofile%\Desktop\Clerical Scripts+.lnk"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Clerical Scripts+".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_2 (DESKTOP ICON NAME: CLERICAL-)
REM ===================================================================================================
:CHOICE_CS_CLERICAL_CUSTOM_SETTINGS_2
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical-.zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical- (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\Clerical Scripts-.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical-.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical- (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\Clerical Scripts-.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts-.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical-.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical- (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts-.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical-.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical- (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical Scripts-.lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Clerical Scripts-.zmd" "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts-.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts-.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts-.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts-.lnk" "%userprofile%\Desktop\Clerical Scripts-.lnk"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Clerical Scripts-".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_CLERICAL_DEFAULT_SETTINGS (DESKTOP ICON NAME: CLERICAL)
REM ===================================================================================================
:CHOICE_CS_CLERICAL_DEFAULT_SETTINGS 
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical.zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Clerical (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\Clerical Scripts.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Clerical (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\Clerical Scripts.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical.lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Clerical Scripts.lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Clerical Scripts.zmd" "%userprofile%\CS_BlueZone_Scripts\Clerical Scripts.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Clerical Scripts.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Clerical Scripts.lnk" "%userprofile%\Desktop\Clerical Scripts.lnk"
)

CLS
ECHO.
ECHO Installation successful! 
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Clerical Scripts".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_1 (DESKTOP ICON NAME: SCRIPTS+)
REM ===================================================================================================
:CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_1 
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts +.zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts+ (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts+.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts +.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts+ (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts+.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts+.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts+.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts+ (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts+.lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts+ (7.1.5).lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\BlueZone Scripts+.zmd" "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts+.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts+.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts+.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts+.lnk" "%userprofile%\Desktop\Scripts+.lnk"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Scripts+".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_2 (DESKTOP ICON NAME: SCRIPTS-)
REM ===================================================================================================
:CHOICE_CS_SCRIPTS_CUSTOM_SETTINGS_2 
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts -.zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts- (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts-.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts -.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts- (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts-.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts-.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts-.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts- (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts-.lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts- (7.1.5).lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\BlueZone Scripts-.zmd" "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts-.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts-.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts-.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts-.lnk" "%userprofile%\Desktop\Scripts-.lnk"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Scripts-".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_SCRIPTS_DEFAULT_SETTINGS (DESKTOP ICON NAME: SCRIPTS)
REM ===================================================================================================
:CHOICE_CS_SCRIPTS_DEFAULT_SETTINGS
Taskkill /F /IM bzmd.exe 2> NUL 
ECHO.

REM Delete any previous custom settings 1 installations of the .zmd file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts .zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Documents\BlueZone\BlueZone Scripts.zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts .zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts (7.1.5).zmd" 2> NUL 
DEL "%userprofile%\Documents\BlueZone\BlueZone Scripts.zmd" 2> NUL 
DEL "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts.zmd" 2> NUL 

REM Delete any previous custom settings 1 installations of the .lnk file (named with a '+')
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts.lnk" 2> NUL 
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts (7.1.5).lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts.lnk" 2> NUL 
DEL "%userprofile%\desktop\Scripts (7.1.5).lnk" 2> NUL 

REM Create the CS_BlueZone_Scripts folder if it doesn't exist
IF NOT EXIST "%userprofile%\CS_BlueZone_Scripts" MKDIR "%userprofile%\CS_BlueZone_Scripts"

REM Install the configuration file to the CS_BlueZone_Scripts folder
COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\BlueZone Scripts.zmd" "%userprofile%\CS_BlueZone_Scripts\BlueZone Scripts.zmd"

REM If the user's desktop is synced to OneDrive, install the shortcut file to the OneDrive-synced desktop, otherwise install the shortcut file on the normal Desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts.lnk" "%userprofile%\OneDrive - Hennepin County\Desktop\Scripts.lnk"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\Shortcut Files\Scripts.lnk" "%userprofile%\Desktop\Scripts.lnk"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "Scripts".
ECHO.
ECHO To access scripts on the power pad, open BlueZone via the new icon.
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.
GOTO END

REM ===================================================================================================
REM CHOICE_CS_STATE_ICON (DESKTOP ICON NAME: STATE)
REM ===================================================================================================
:CHOICE_CS_STATE_ICON 
Taskkill /F /IM bzmd.exe 2> NUL 

REM Delete any previous installations of the STATE.zmd file
DEL "%userprofile%\OneDrive - Hennepin County\Desktop\STATE.zmd"  2> NUL 
DEL "%userprofile%\desktop\STATE.zmd" 2> NUL 

REM If the user's desktop is synced to OneDrive, install the STATE.zmd file to the OneDrive-synced desktop
IF EXIST "%userprofile%\OneDrive - Hennepin County\Desktop" (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\STATE.zmd" "%userprofile%\OneDrive - Hennepin County\Desktop\STATE.zmd"
) ELSE (
	COPY /Y "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Child Support\EA_CSD_Common\BlueZone Scripts\BlueZone Configuration Files\STATE.zmd" "%userprofile%\Desktop\STATE.zmd"
)

CLS
ECHO.
ECHO Installation successful!
ECHO.
ECHO A new BlueZone icon has been added to your desktop named "STATE".
ECHO.
ECHO Questions? Problems? Contact the Child Support BlueZone script writer at HSPH.CS.InfoTeam@Hennepin.us.

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
REM CS_ACCESS_ERROR
REM ========================================================================
:CS_ACCESS_ERROR
ECHO.
ECHO ***ERROR***
ECHO.
ECHO Installation canceled. 
ECHO.
ECHO Access was denied to T:\Child Support\EA_CSD_Common\BlueZone\BlueZone Scripts.
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
ECHO Unable to connect to the T: drive, so the installer cannot run.
ECHO.
ECHO Reconnect to the T: drive and try running the installer again.
ECHO. 
ECHO Contact the Help Desk if you need assistance with connecting to the T: drive.
ECHO.
GOTO END

REM ========================================================================
REM END
REM ========================================================================
:END
ECHO.
PAUSE