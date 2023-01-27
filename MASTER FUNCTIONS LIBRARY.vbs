'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script "library" contains functions and variables that the other BlueZone scripts use very commonly. The other BlueZone scripts contain a few lines of code that run
'this script and get the functions. This saves time in writing and copy/pasting the same functions in many different places. Only add functions to this script if they've
'been tested in other scripts first. This document is actively used by live scripts, so it needs to be functionally complete at all times.
'
'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'CHANGELOG BLOCK ===========================================================================================================
actual_script_name = name_of_script
name_of_script = "Functions Library"
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/02/2022", "Retired Script - UTILITIES - LOST APPLYMN. This has been replaced by UTILITIES - APPLICATION INQUIRY. Power pad also updated to reflect this change.", "Ilse Ferris, Hennepin County")
Call changelog_update("03/15/2022", "DHS is reporting the MAXIS Background slowdown has been resolved. Cases should be coming through BGTX in a typical amount of time.##~## ##~##If you have any issues with MAXIS Background going forward, please contack MAXIS Help Desk.", "Casey Love, Hennepin County")
Call changelog_update("03/15/2022", "************** ATTENTION *************##~## ##~##There appears to be slowdown in MAXIS Background Transaction (BGTX) and cases are getting stuck for an extended amount of time.##~## ##~##This is not an issue with any script but could become apparant during a script run where the script seems 'stuck' or could error out. As the issue is with MAXIS - we cannot handle for this with any script updates or fixes.##~## ##~##DHS is repporting they are aware of the problem.##~## ##~##If possible, ensure your case is through background before running any scripts to reduce the impact on your script runs.", "Casey Love, Hennepin County")
Call changelog_update("09/24/2021", "************** ATTENTION *************##~## ##~##DHS HAS UPDATED THE LAYOUT OF THE ADDR PANEL##~## ##~##Effective in the Footer Month of 10/21 there are additional fields on ADDR and some of the information has moved.##~## ##~##This may cause some issues in the BlueZone Scripts as it may read incorrect information.##~## ##~##We are working on updating these script files as quickly as we can. However, we only had a few days notice on this change in MAXIS and these updates take time. (We have many scripts that read from the ADDR panel!)##~## ##~##Please send an email to HSPH.EWS.BlueZoneScripts@hennepin.us if you notice anything with a script running wrong or providing incorrect ADDR information. We will continue working on our list and you will see changes throughout the day.##~## ##~##THANK YOU!##~##", "Casey Love, Hennepin County")
Call changelog_update("09/16/2021", "!!! NEW SCRIPT !!! ##~## ##~##UC VERIFICATION REQUEST##~## ##~##This script will email unemployment liasons to request UC verification.##~##", "MiKayla Handley, Hennepin County")
Call changelog_update("06/08/2021", "##~##UPDATE TO THE POWER PAD##~## ##~##We have updated the Power Pad (the buttons at the top of your MAXIS session that allows you to access the scripts). ##~## ##~##This new layout allows for more direct access to some of our most used scripts and some support scripts to reach the right people quickly.##~##As with any new functionality, there may be some things that need tweaking, so please let us know if you see anything out of place or in need of update.##~## ##~##If you do not see the new buttons, they will come up the next time you reopen or start a new session of MAXIS/MMIS.", "Casey Love, Hennepin County")
Call changelog_update("04/14/2021", "##~## ##~##NOTICES - PA VERIF##~## ##~##This script has been retired as it was not in compliance with the procedures at Hennepin County for providing public assistance verifications. The tool will be re-evaluated for in-person supports in the near future. If you have a resident or community partner who is requesting public assistance verifications, please refer to the HSR Manual references: Data Privacy, Client Requests, Verifying Public Housing, and Public Assistance Verification.", "Ilse Ferris, Hennepin County")
Call changelog_update("11/25/2020", "!!! NEW SCRIPT !!! ##~## ##~##VA VERIFICATION REQUEST##~## ##~##This script will email veterans services to request VA income.##~##", "MiKayla Handley, Hennepin County")
Call changelog_update("09/23/2020", "!!! NEW SCRIPT !!! ##~## ##~##Job Change Reported##~## ##~##This script will replace 'New Job Reported', it handles the functionality of a new job reported and adds functionality to handle for changes to existing jobs and stop work reports.##~##This script is not intended for budgeting with received verifications, it is used for the initial report when no verif have been received.##~## ##~##Expect 'New Job Reported' to be retired soon.##~##", "Casey Love, Hennepin County")
Call changelog_update("06/12/2020", "There have been some changes to ALL of the MAXIS scripts Categories Menus. You may have noticed the changes, but the scripts should be mostly in the same place and work the same way. ##~## What has changed?##~## ##~## 1. The script button names are all the same length now (so much nicer looking). ##~## 2. The menus now change size based on the number of scripts (again nicer looking). ##~## 3. There is a '?' button in front of most script names. This button will open the script instructions. ##~## ##~## Let us know what you think, or email us if anything is missing or out of place.", "Casey Love, Hennepin County")
Call changelog_update("10/01/2019", "Remember, the old NOTES - CAF script retires today. ##~## ##~## The new NOTES - CAF is enhanced to meet documentation requirements for processing a CAF. ##~## ##~## Please join us on Thursday for a Live Skype Demo. ##~## ##~## Check today's Hot Topic for additional information.", "Casey Love, Hennepin County")
Call changelog_update("09/24/2019", "Join us today on SKYPE for a LIVE DEMO of the new CAF script. ##~## ##~## Today 09-24-2019 at 10:00AM. ##~## ##~## Find the link to join the Skype session on the BlueZone Sharepoint Page or in this week's Hot Topics!##~## ##~## We hope to see you there!", "Casey Love, Hennepin County")
Call changelog_update("09/19/2019", "Join us today on SKYPE for a LIVE DEMO of the new CAF script. ##~## ##~## Today 09-19-2019 at 2:00PM. ##~## ##~## Find the link to join the Skype session on the BlueZone Shorepoint Page or in this week's Hot Topics!##~## ##~## We hope to see you there!", "Casey Love, Hennepin County")
Call changelog_update("09/17/2019", "The NEW CAF Script is available. Try it now! Find it in the NOTES menu. The script can be called using the button '**NEW** CAF'. ##~## ##~## Details can be found in HOT TOPICS from today (9/17/19).##~## ##~## The old version of the script is still available but will be taken out soon.", "Casey Love, Hennepin County")
Call changelog_update("09/10/2019", "CHANGES to the CAF Script. One of our most used scripts is getting a revamp. Updated functionality and format. Check out this week's HOT TOPICS for a sneak peak or ask around, we have testers working on it in the regions right now!", "Casey Love, Hennepin County")
call changelog_update("08/28/2019", "Have you seen a script ask at the end if you need to send an error report? Want to know what that is about? Read this week's HOT TOPICS for details about In-Script Error Reporting.", "Casey Love, Hennepin County")
call changelog_update("06/25/2019", "We want to hear from YOU! Please respond to our Survey, link and details can be found in Hot Topics.", "Casey Love, Hennepin County")
call changelog_update("06/21/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
name_of_script = actual_script_name
'END CHANGELOG BLOCK =======================================================================================================

'LOADING LIST OF SCRIPTS FROM GITHUB REPOSITORY===========================================================================
IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
	IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
		script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
	Else											'Everyone else should use the release branch.
		script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
	End if

	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_list_URL
	req.open "GET", script_list_URL, FALSE							'Attempts to open the script_list_URL
	req.send													'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
		Execute req.responseText								'Executes the script code
	ELSE														'Error message
		critical_error_msgbox = MsgBox ("Something has gone wrong. The script list code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                        "Script list URL: " & script_list_URL & vbNewLine & vbNewLine &_
                                        "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                        vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
        StopScript
	END IF
ELSE
	script_list_URL = "C:\MAXIS-scripts\COMPLETE LIST OF SCRIPTS.vbs"
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
Dim checked, unchecked, cancel, OK, blank, t_drive, STATS_counter, STATS_manualtime, STATS_denomination, script_run_lowdown, testing_run, MAXIS_case_number		'Declares this for Option Explicit users

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs
blank = ""
t_drive = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team"

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

'Global function to actually RUN'
If IsArray(tester_array) = True Then Call confirm_tester_information

'Time arrays which can be used to fill an editbox with the convert_array_to_droplist_items function
time_array_15_min = array("7:00 AM", "7:15 AM", "7:30 AM", "7:45 AM", "8:00 AM", "8:15 AM", "8:30 AM", "8:45 AM", "9:00 AM", "9:15 AM", "9:30 AM", "9:45 AM", "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "1:00 PM", "1:15 PM", "1:30 PM", "1:45 PM", "2:00 PM", "2:15 PM", "2:30 PM", "2:45 PM", "3:00 PM", "3:15 PM", "3:30 PM", "3:45 PM", "4:00 PM", "4:15 PM", "4:30 PM", "4:45 PM", "5:00 PM", "5:15 PM", "5:30 PM", "5:45 PM", "6:00 PM")
time_array_30_min = array("7:00 AM", "7:30 AM", "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM")

'Array of all the upcoming holidays
HOLIDAYS_ARRAY = Array(#11/11/22#, #11/24/22#, #11/25/22#, #12/26/22#, #01/2/23#, #1/16/23#, #2/20/23#, #5/29/23#, #6/19/23#, #7/4/23#, #9/4/23#, #11/10/23#, #11/23/23#, #11/24/23#, #12/25/23#)

'Determines CM and CM+1 month and year using the two rightmost chars of both the month and year. Adds a "0" to all months, which will only pull over if it's a single-digit-month
Dim CM_mo, CM_yr, CM_plus_1_mo, CM_plus_1_yr, CM_plus_2_mo, CM_plus_2_yr
'var equals...  the right part of...    the specific part...    of either today or next month... just the right 2 chars!
CM_mo =         right("0" &             DatePart("m",           date                             ), 2)
CM_yr =         right(                  DatePart("yyyy",        date                             ), 2)

CM_plus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", 1, date)            ), 2)
CM_plus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", 1, date)            ), 2)

CM_plus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", 2, date)            ), 2)
CM_plus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", 2, date)            ), 2)

If worker_county_code   = "" then worker_county_code = "MULTICOUNTY"
IF PRISM_script <> true then county_name = ""		'VKC NOTE 08/12/2016: ADDED IF...THEN CONDITION BECAUSE PRISM IS STILL USING THIS VARIABLE IN ALL SCRIPTS.vbs. IT WILL BE REMOVED AND THIS CAN BE RESTORED.

If ButtonPressed <> "" then ButtonPressed = ""		'Defines ButtonPressed if not previously defined, allowing scripts the benefit of not having to declare ButtonPressed all the time

'All 10-day cutoff dates are provided in POLI/TEMP TE19.132
IF CM_mo = "01" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #01/19/2023#
ELSEIF CM_mo = "02" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #02/16/2023#
ELSEIF CM_mo = "03" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #03/21/2023#
ELSEIF CM_mo = "04" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #04/20/2023#
ELSEIF CM_mo = "05" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #05/19/2023#
ELSEIF CM_mo = "06" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #06/20/2023#
ELSEIF CM_mo = "07" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #07/20/2023#
ELSEIF CM_mo = "08" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #08/21/2023#
ELSEIF CM_mo = "09" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #09/20/2023#
ELSEIF CM_mo = "10" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #10/19/2023#
ELSEIF CM_mo = "11" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #11/20/2023#
ELSEIF CM_mo = "12" AND CM_yr = "23" THEN
    ten_day_cutoff_date = #12/20/2023#
ELSEIF CM_mo = "11" AND CM_yr = "22" THEN
    ten_day_cutoff_date = #11/18/2022#
ELSEIF CM_mo = "12" AND CM_yr = "22" THEN
    ten_day_cutoff_date = #12/21/2022#                                'last month of current year
ELSE
	MsgBox "You have entered a date (" & CM_mo & "/" & CM_yr & ") not supported by this function. Please contact a scripts administrator to determine if the script requires updating.", vbInformation + vbSystemModal, "NOTICE"
END IF

'preloading boolean variables for tabbing dialogs
pass_one = False
pass_two = False
pass_three = False
pass_four = False
pass_five = False
pass_six = False
pass_seven = False
pass_eight = False
pass_nine = False
pass_ten = False

show_one = True
show_two = True
show_three = True
show_four = True
show_five = True
show_six = True
show_seven = True
show_eight = True
show_nine = True
show_ten = True

tab_button = False

'list of all the counties for drop downs in dialogs
'Should work for both ComboBox and DropListBox
county_list = "01 - Aitkin"
county_list = county_list+chr(9)+"02 - Anoka"
county_list = county_list+chr(9)+"03 - Becker"
county_list = county_list+chr(9)+"04 - Beltrami"
county_list = county_list+chr(9)+"05 - Benton"
' county_list = county_list+chr(9)+"06 - Big Stone"
county_list = county_list+chr(9)+"07 - Blue Earth"
county_list = county_list+chr(9)+"08 - Brown"
county_list = county_list+chr(9)+"09 - Carlton"
county_list = county_list+chr(9)+"10 - Carver"
county_list = county_list+chr(9)+"11 - Cass"
' county_list = county_list+chr(9)+"12 - Chippewa"
county_list = county_list+chr(9)+"13 - Chisago"
county_list = county_list+chr(9)+"14 - Clay"
' county_list = county_list+chr(9)+"15 - Clearwater"
' county_list = county_list+chr(9)+"16 - Cook"
' county_list = county_list+chr(9)+"17 - Cottonwood"
county_list = county_list+chr(9)+"18 - Crow Wing"
county_list = county_list+chr(9)+"19 - Dakota"
county_list = county_list+chr(9)+"20 - Dodge"
county_list = county_list+chr(9)+"21 - Douglas"
county_list = county_list+chr(9)+"22 - Faribault"
county_list = county_list+chr(9)+"23 - Fillmore"
county_list = county_list+chr(9)+"24 - Freeborn"
county_list = county_list+chr(9)+"25 - Goodhue"
' county_list = county_list+chr(9)+"26 - Grant"
county_list = county_list+chr(9)+"27 - Hennepin"
county_list = county_list+chr(9)+"28 - Houston"
county_list = county_list+chr(9)+"29 - Hubbard"
county_list = county_list+chr(9)+"30 - Isanti"
county_list = county_list+chr(9)+"31 - Itasca"
' county_list = county_list+chr(9)+"32 - Jackson"
county_list = county_list+chr(9)+"33 - Kanabec"
county_list = county_list+chr(9)+"34 - Kandiyohi"
' county_list = county_list+chr(9)+"35 - Kittson"
' county_list = county_list+chr(9)+"36 - Koochiching"
' county_list = county_list+chr(9)+"37 - Lac Qui Parle"
' county_list = county_list+chr(9)+"38 - Lake"
' county_list = county_list+chr(9)+"39 - Lake Of Woods"
county_list = county_list+chr(9)+"40 - Le Sueur"
' county_list = county_list+chr(9)+"41 - Lincoln"
county_list = county_list+chr(9)+"42 - Lyon"
county_list = county_list+chr(9)+"43 - Mcleod"
' county_list = county_list+chr(9)+"44 - Mahnomen"
' county_list = county_list+chr(9)+"45 - Marshall"
county_list = county_list+chr(9)+"46 - Martin"
county_list = county_list+chr(9)+"47 - Meeker"
county_list = county_list+chr(9)+"48 - Mille Lacs"
county_list = county_list+chr(9)+"49 - Morrison"
county_list = county_list+chr(9)+"50 - Mower"
' county_list = county_list+chr(9)+"51 - Murray"
county_list = county_list+chr(9)+"52 - Nicollet"
county_list = county_list+chr(9)+"53 - Nobles"
' county_list = county_list+chr(9)+"54 - Norman"
county_list = county_list+chr(9)+"55 - Olmsted"
county_list = county_list+chr(9)+"56 - Otter Tail"
county_list = county_list+chr(9)+"57 - Pennington"
county_list = county_list+chr(9)+"58 - Pine"
' county_list = county_list+chr(9)+"59 - Pipestone"
county_list = county_list+chr(9)+"60 - Polk"
' county_list = county_list+chr(9)+"61 - Pope"
county_list = county_list+chr(9)+"62 - Ramsey"
' county_list = county_list+chr(9)+"63 - Red Lake"
county_list = county_list+chr(9)+"64 - Redwood"
county_list = county_list+chr(9)+"65 - Renville"
county_list = county_list+chr(9)+"66 - Rice"
' county_list = county_list+chr(9)+"67 - Rock"
county_list = county_list+chr(9)+"68 - Roseau"
county_list = county_list+chr(9)+"69 - St. Louis"
county_list = county_list+chr(9)+"70 - Scott"
county_list = county_list+chr(9)+"71 - Sherburne"
county_list = county_list+chr(9)+"72 - Sibley"
county_list = county_list+chr(9)+"73 - Stearns"
county_list = county_list+chr(9)+"74 - Steele"
' county_list = county_list+chr(9)+"75 - Stevens"
' county_list = county_list+chr(9)+"76 - Swift"
county_list = county_list+chr(9)+"77 - Todd"
' county_list = county_list+chr(9)+"78 - Traverse"
county_list = county_list+chr(9)+"79 - Wabasha"
' county_list = county_list+chr(9)+"80 - Wadena"
county_list = county_list+chr(9)+"81 - Waseca"
county_list = county_list+chr(9)+"82 - Washington"
' county_list = county_list+chr(9)+"83 - Watonwan"
' county_list = county_list+chr(9)+"84 - Wilkin"
county_list = county_list+chr(9)+"85 - Winona"
county_list = county_list+chr(9)+"86 - Wright"
' county_list = county_list+chr(9)+"87 - Yellow Medicine"
county_list = county_list+chr(9)+"89 - Out-of-State"

state_list = "NB - MN Newborn"
state_list = state_list+chr(9)+"FC - Foreign Country"
state_list = state_list+chr(9)+"UN - Unknown"
state_list = state_list+chr(9)+"AL - Alabama"
state_list = state_list+chr(9)+"AK - Alaska"
state_list = state_list+chr(9)+"AZ - Arizona"
state_list = state_list+chr(9)+"AR - Arkansas"
state_list = state_list+chr(9)+"CA - California"
state_list = state_list+chr(9)+"CO - Colorado"
state_list = state_list+chr(9)+"CT - Connecticut"
state_list = state_list+chr(9)+"DE - Delaware"
state_list = state_list+chr(9)+"DC - District Of Columbia"
state_list = state_list+chr(9)+"FL - Florida"
state_list = state_list+chr(9)+"GA - Georgia"
state_list = state_list+chr(9)+"HI - Hawaii"
state_list = state_list+chr(9)+"ID - Idaho"
state_list = state_list+chr(9)+"IL - Illnois"
state_list = state_list+chr(9)+"IN - Indiana"
state_list = state_list+chr(9)+"IA - Iowa"
state_list = state_list+chr(9)+"KS - Kansas"
state_list = state_list+chr(9)+"KY - Kentucky"
state_list = state_list+chr(9)+"LA - Louisiana"
state_list = state_list+chr(9)+"MA - Massachusetts"
state_list = state_list+chr(9)+"MD - Maryland"
state_list = state_list+chr(9)+"ME - Maine"
state_list = state_list+chr(9)+"MI - Michigan"
state_list = state_list+chr(9)+"MN - Minnesota"
state_list = state_list+chr(9)+"MS - Mississippi"
state_list = state_list+chr(9)+"MO - Missouri"
state_list = state_list+chr(9)+"MT - Montana"
state_list = state_list+chr(9)+"NE - Nebraska"
state_list = state_list+chr(9)+"NV - Nevada"
state_list = state_list+chr(9)+"NH - New Hampshire"
state_list = state_list+chr(9)+"NJ - New Jersey"
state_list = state_list+chr(9)+"NM - New Mexico"
state_list = state_list+chr(9)+"NY - New York"
state_list = state_list+chr(9)+"NC - North Carolina"
state_list = state_list+chr(9)+"ND - North Dakota"
state_list = state_list+chr(9)+"OH - Ohio"
state_list = state_list+chr(9)+"OK - Oklahoma"
state_list = state_list+chr(9)+"OR - Oregon"
state_list = state_list+chr(9)+"PA - Pennsylvania"
state_list = state_list+chr(9)+"RI - Rhode Island"
state_list = state_list+chr(9)+"SC - South Carolina"
state_list = state_list+chr(9)+"SD - South Dakota"
state_list = state_list+chr(9)+"TN - Tennessee"
state_list = state_list+chr(9)+"TX - Texas"
state_list = state_list+chr(9)+"UT - Utah"
state_list = state_list+chr(9)+"VT - Vermont"
state_list = state_list+chr(9)+"VA - Virginia"
state_list = state_list+chr(9)+"WA - Washington"
state_list = state_list+chr(9)+"WV - West Virginia"
state_list = state_list+chr(9)+"WI - Wisconsin"
state_list = state_list+chr(9)+"WY - Wyoming"
state_list = state_list+chr(9)+"PR - Puerto Rico"
state_list = state_list+chr(9)+"VI - Virgin Islands"

UNEA_type_list = "01 - RSDI, Disa"
UNEA_type_list = UNEA_type_list+chr(9)+"02 - RSDI, No Disa"
UNEA_type_list = UNEA_type_list+chr(9)+"03 - SSI"
UNEA_type_list = UNEA_type_list+chr(9)+"06 - Non-MN PA"
UNEA_type_list = UNEA_type_list+chr(9)+"11 - VA Disability"
UNEA_type_list = UNEA_type_list+chr(9)+"12 - VA Pension"
UNEA_type_list = UNEA_type_list+chr(9)+"13 - VA Other"
UNEA_type_list = UNEA_type_list+chr(9)+"38 - VA Aid & Attendance"
UNEA_type_list = UNEA_type_list+chr(9)+"14 - Unemployment Insurance"
UNEA_type_list = UNEA_type_list+chr(9)+"15 - Worker's Comp"
UNEA_type_list = UNEA_type_list+chr(9)+"16 - Railroad Retirement"
UNEA_type_list = UNEA_type_list+chr(9)+"17 - Other Retirement"
UNEA_type_list = UNEA_type_list+chr(9)+"18 - Military Enrirlement"
UNEA_type_list = UNEA_type_list+chr(9)+"19 - FC Child req FS"
UNEA_type_list = UNEA_type_list+chr(9)+"20 - FC Child not req FS"
UNEA_type_list = UNEA_type_list+chr(9)+"21 - FC Adult req FS"
UNEA_type_list = UNEA_type_list+chr(9)+"22 - FC Adult not req FS"
UNEA_type_list = UNEA_type_list+chr(9)+"23 - Dividends"
UNEA_type_list = UNEA_type_list+chr(9)+"24 - Interest"
UNEA_type_list = UNEA_type_list+chr(9)+"25 - Cnt gifts/prizes"
UNEA_type_list = UNEA_type_list+chr(9)+"26 - Strike Benefits"
UNEA_type_list = UNEA_type_list+chr(9)+"27 - Contract for Deed"
UNEA_type_list = UNEA_type_list+chr(9)+"28 - Illegal Income"
UNEA_type_list = UNEA_type_list+chr(9)+"29 - Other Countable"
UNEA_type_list = UNEA_type_list+chr(9)+"30 - Infrequent"
UNEA_type_list = UNEA_type_list+chr(9)+"31 - Other - FS Only"
UNEA_type_list = UNEA_type_list+chr(9)+"08 - Direct Child Support"
UNEA_type_list = UNEA_type_list+chr(9)+"35 - Direct Spousal Support"
UNEA_type_list = UNEA_type_list+chr(9)+"36 - Disbursed Child Support"
UNEA_type_list = UNEA_type_list+chr(9)+"37 - Disbursed Spousal Support"
UNEA_type_list = UNEA_type_list+chr(9)+"39 - Disbursed CS Arrears"
UNEA_type_list = UNEA_type_list+chr(9)+"40 - Disbursed Spsl Sup Arrears"
UNEA_type_list = UNEA_type_list+chr(9)+"43 - Disbursed Excess CS"
UNEA_type_list = UNEA_type_list+chr(9)+"44 - MSA - Excess Income for SSI"
UNEA_type_list = UNEA_type_list+chr(9)+"47 - Tribal Income"
UNEA_type_list = UNEA_type_list+chr(9)+"48 - Trust Income"
UNEA_type_list = UNEA_type_list+chr(9)+"49 - Non-Recurring"

ACCT_type_list = "SV - Savings"
ACCT_type_list = ACCT_type_list+chr(9)+"CK - Checking"
ACCT_type_list = ACCT_type_list+chr(9)+"CE - Certificate of Deposit"
ACCT_type_list = ACCT_type_list+chr(9)+"MM - Money Market"
ACCT_type_list = ACCT_type_list+chr(9)+"DC - Debit Card"
ACCT_type_list = ACCT_type_list+chr(9)+"KO - Keogh Account"
ACCT_type_list = ACCT_type_list+chr(9)+"FT - Fed Thrift Savings Plan"
ACCT_type_list = ACCT_type_list+chr(9)+"SL - State & Local Govt"
ACCT_type_list = ACCT_type_list+chr(9)+"RA - Employee Ret Annuities"
ACCT_type_list = ACCT_type_list+chr(9)+"NP - Non-Profit Emmployee Ret"
ACCT_type_list = ACCT_type_list+chr(9)+"IR - Indiv Ret Acct"
ACCT_type_list = ACCT_type_list+chr(9)+"RH - Roth IRA"
ACCT_type_list = ACCT_type_list+chr(9)+"FR - Ret Plan for Employers"
ACCT_type_list = ACCT_type_list+chr(9)+"CT - Corp Ret Trust"
ACCT_type_list = ACCT_type_list+chr(9)+"RT - Other Ret Fund"
ACCT_type_list = ACCT_type_list+chr(9)+"QT - Qualified Tuition (529)"
ACCT_type_list = ACCT_type_list+chr(9)+"CA - Coverdell SV (530)"
ACCT_type_list = ACCT_type_list+chr(9)+"OE - Other Educational"
ACCT_type_list = ACCT_type_list+chr(9)+"OT - Other"

SECU_type_list = "LI - Life Insurance"
SECU_type_list = SECU_type_list+chr(9)+"ST - Stocks"
SECU_type_list = SECU_type_list+chr(9)+"BO - Bonds"
SECU_type_list = SECU_type_list+chr(9)+"CD - Ctrct for Deed"
SECU_type_list = SECU_type_list+chr(9)+"MO - Mortgage Note"
SECU_type_list = SECU_type_list+chr(9)+"AN - Annuity"
SECU_type_list = SECU_type_list+chr(9)+"OT - Other"

CARS_type_list = "1 - Car"
CARS_type_list = CARS_type_list+chr(9)+"2 - Truck"
CARS_type_list = CARS_type_list+chr(9)+"3 - Van"
CARS_type_list = CARS_type_list+chr(9)+"4 - Camper"
CARS_type_list = CARS_type_list+chr(9)+"5 - Motorcycle"
CARS_type_list = CARS_type_list+chr(9)+"6 - Trailer"
CARS_type_list = CARS_type_list+chr(9)+"7 - Other"

REST_type_list = "1 - House"
REST_type_list = REST_type_list+chr(9)+"2 - Land"
REST_type_list = REST_type_list+chr(9)+"3 - Buildings"
REST_type_list = REST_type_list+chr(9)+"4 - Mobile Home"
REST_type_list = REST_type_list+chr(9)+"5 - Life Estate"
REST_type_list = REST_type_list+chr(9)+"6 - Other"

JOBS_type_list = "J - WIOA"
JOBS_type_list = JOBS_type_list+chr(9)+"W - Wages (Incl Tips)"
JOBS_type_list = JOBS_type_list+chr(9)+"E - EITC"
JOBS_type_list = JOBS_type_list+chr(9)+"G - Experience Works"
JOBS_type_list = JOBS_type_list+chr(9)+"F - Federal Work Study"
JOBS_type_list = JOBS_type_list+chr(9)+"S - State Work Study"
JOBS_type_list = JOBS_type_list+chr(9)+"O - Other"
JOBS_type_list = JOBS_type_list+chr(9)+"C - Contract Income"
JOBS_type_list = JOBS_type_list+chr(9)+"T - Training Program"
JOBS_type_list = JOBS_type_list+chr(9)+"P - Service Program"
JOBS_type_list = JOBS_type_list+chr(9)+"R - Rehab Program"

function remove_dash_from_droplist(list_to_alter)
	list_to_alter = replace(list_to_alter, " - ", " ")
end function


'Preloading worker_signature, as a constant to be used in scripts---------------------------------------------------------------------------------------------------------

Const ForReading = 1			'These Constants are used for enumeration in file manipulation functions.
Const ForWriting = 2			'See the 'OpenAsTextStream' method
Const ForAppending = 8			'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method

Set wshshell = CreateObject("WScript.Shell")						'creating the wscript method to interact with the system
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"	'defining the my documents folder for use in saving script details/variables between script runs

'Now it determines the signature
With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
	If .FileExists(user_myDocs_folder & "workersig.txt") Then												'If the workersig.txt file exists...
		Set get_worker_sig = CreateObject("Scripting.FileSystemObject")										'Create another FSO
		Set worker_sig_command = get_worker_sig.OpenTextFile(user_myDocs_folder & "workersig.txt")			'Open the text file
		worker_sig = worker_sig_command.ReadAll																'Read the text file
		IF worker_sig <> "" THEN worker_signature = worker_sig												'If it isn't blank then the worker_signature should be populated with the contents of the file
		worker_sig_command.Close																			'Closes the file
	END IF
END WITH

'Now we check MyDocs to see if there is script saved work files that need to be removed.
'This is necessary for correct records managment and data security.
'This runs every time we open the Func Lib and happens in the background, ensuring no interruption in other script functions
Set objFolder = objFSO.GetFolder(user_myDocs_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder
For Each objFile in colFiles																'looping through each file
	delete_this_file = False																'Default to NOT delete the file
	this_file_name = objFile.Name															'Grabing the file name
	this_file_type = objFile.Type															'Grabing the file type
	this_file_created_date = objFile.DateCreated											'Reading the date created
	this_file_path = objFile.Path															'Grabing the path for the file

	If InStr(this_file_name, "caf-answers-") <> 0 Then delete_this_file = True				'We want to delete files that say 'caf-answers-' as this is how the UTILITIES - Complete Phone CAF script creates the save your work doc
	If InStr(this_file_name, "caf-variables-") <> 0 Then delete_this_file = True				'We want to delete files that say 'caf-answers-' as this is how the UTILITIES - Complete Phone CAF script creates the save your work doc
	If InStr(this_file_name, "interview-answers-") <> 0 Then delete_this_file = True		'We want to delete files that say 'interview-answers-' as this is how the NOTES - Interview script creates the save your work doc
	If this_file_type <> "Text Document" then delete_this_file = False						'We do NOT want to delete files that are NOT TXT file types
	If DateDiff("d", this_file_created_date, date) < 8 Then delete_this_file = False		'We do NOT want to delete files that are 7 days old or less - we may need to reference the saved work in these files.

	If delete_this_file = True Then objFSO.DeleteFile(this_file_path)						'If we have determined that we need to delete the file - here we delete it
Next

'----------------------------------------------------------------------------------------------------Email addresses for the teams
IF current_worker_number =	"X127F3P" 	THEN email_address = "HSPH.ES.MA.EPD.Adult@hennepin.us"
IF current_worker_number =	"X127F3K" 	THEN email_address = "HSPH.ES.MA.EPD.FAM@hennepin.us"
IF current_worker_number =	"X127F3F"	THEN email_address = "HSPH.ES.MA.EPD.ADS@hennepin.us"
IF current_worker_number =	"X127EA0" 	THEN email_address = "hsph.es.ea.team@hennepin.us"
IF current_worker_number =	"X127EAK" 	THEN email_address = "hsph.es.ea.team@hennepin.us"
IF current_worker_number =	"X127EM3" 	THEN email_address = "hsph.es.extendicare@hennepin.us"
IF current_worker_number =	"X127EM4" 	THEN email_address = "hsph.es.extendicare@hennepin.us"
IF current_worker_number =	"X127FG6" 	THEN email_address = "hsph.es.goldenliving@hennepin.us"
IF current_worker_number =	"X127FG7" 	THEN email_address = "hsph.es.goldenliving@hennepin.us"
IF current_worker_number =	"X127LE1" 	THEN email_address = "hsph.es.littleearth@hennepin.us"
IF current_worker_number =	"X127NP0"	THEN email_address = "hsph.es.northpoint@hennepin.us"
IF current_worker_number =	"X127NPC" 	THEN email_address = "hsph.es.northpoint@hennepin.us"
IF current_worker_number =	"X127FF4" 	THEN email_address = "hsph.es.northridge@hennepin.us"
IF current_worker_number =	"X127FF5" 	THEN email_address = "hsph.es.northridge@hennepin.us"
IF current_worker_number =	"X127ED8" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EAJ" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EN1" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EN2" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EN3" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EN4"  	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127ED6" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127ED7" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127FE6" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF current_worker_number =	"X127EJ9" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127ER6" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EE2" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EE3" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EE4"	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EE5" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EG5" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF current_worker_number =	"X127EQ1" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EQ2" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EQ5" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EK8" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EQ4" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127FH9" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EG6" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF current_worker_number =	"X127EE1"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127FB2"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127EG7" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127ED9"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127EE0" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127EH4" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127EH5" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127F3D" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127FH8" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF current_worker_number =	"X127EQ8" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127EE6" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127EE7" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127ER1" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127EG8" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127FH2" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF current_worker_number =	"X127EF8" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF current_worker_number =	"X127EF9" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF current_worker_number =	"X127EG9" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF current_worker_number =	"X127EG0" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF current_worker_number =	"X127EP8" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF current_worker_number =	"X127EP6" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF current_worker_number =	"X127EP7"	THEN email_address = "hsph.es.team.170@hennepin.us"
IF current_worker_number =	"X127EG4"	THEN email_address = "hsph.es.team.170@hennepin.us"
IF current_worker_number =	"X127FG8" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF current_worker_number =	"X127EH1" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF current_worker_number =	"X127EH7" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF current_worker_number =	"X127EH2"  	THEN email_address = "hsph.es.team.251@hennepin.us"
IF current_worker_number =	"X127EH3" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF current_worker_number =	"X127FH4" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF current_worker_number =	"X127EH8" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EQ3" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EJ2" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EJ3" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127FH1" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127FG4" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127F3C" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127F3G"	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127F3L" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EJ1" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EH9" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EM2" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF current_worker_number =	"X127EJ6" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127FE5" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EK3" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EK1" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EK2" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EJ7" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EJ8" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EJ5" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF current_worker_number =	"X127EL8" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL9" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127FE1" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL2" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL3" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL4" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL5" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL6" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EL7" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF current_worker_number =	"X127EH6" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF current_worker_number =	"X127EM1" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF current_worker_number =	"X127FI7" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF current_worker_number =	"X127EM7" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127FI2" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127FG3" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127EM8" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127EM9" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127EJ4" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF current_worker_number =	"X127EK4" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EK5" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127FH5" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EN7"	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EK6" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EK9" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EN6" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF current_worker_number =	"X127EP3" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127EP4" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127EP5" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127EP9" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127F3U"	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127F3V" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF current_worker_number =	"X127EF7" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF current_worker_number =	"X127EN5" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF current_worker_number =	"X127EF5" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF current_worker_number =	"X127EK7" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF current_worker_number =	"X127EF6" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF current_worker_number =	"X127EQ9" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF current_worker_number =	"X127ER2" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF current_worker_number =	"X127ER3" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF current_worker_number =	"X127ER4" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF current_worker_number =	"X127ER5" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF current_worker_number =	"X127FF6" 	THEN email_address = "hsph.es.team.262@hennepin.us"
IF current_worker_number =	"X127FF7" 	THEN email_address = "hsph.es.team.262@hennepin.us"
IF current_worker_number =	"X127FF8" 	THEN email_address = "hsph.es.team.300@hennepin.us"
IF current_worker_number =	"X127FF9" 	THEN email_address = "hsph.es.team.300@hennepin.us"
IF current_worker_number =	"X127FF3" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF current_worker_number =	"X127EX3" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF current_worker_number =	"X127ES7" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF current_worker_number =	"X127EX2" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF current_worker_number =	"X127EX1"	THEN email_address = "hsph.es.team.410@hennepin.us"
IF current_worker_number =	"X127ET9"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EU4"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EW2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EW3"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EU1"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EU3"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127BV2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127EU2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127FH7"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF current_worker_number =	"X127FA1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127FA4" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127BV1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127FA2" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127F3R" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127F3Y" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127FA3"	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127FJ1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF current_worker_number =	"X127ER8" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127F3B" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127ES1" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127ES3" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127FB6" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127F3H" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127F4E" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127FB4" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127F3A" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127FB5" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127FB3" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127ER9" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127ES2" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127F3M" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF current_worker_number =	"X127EY8"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF current_worker_number =	"X127EY9"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF current_worker_number =	"X127EZ1"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF current_worker_number =	"X127EX7" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127EY1" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127FJ5" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127F3Q" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127EX9" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127F3T" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127EX8" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127F3Z" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF current_worker_number =	"X127EU5" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127EU6" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127EY2" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127F3W" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127EU8" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127F3X" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127EU7" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127F3S" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127EU9" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127FJ3" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127FJ4" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF current_worker_number =	"X127ES4" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127ES8" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127EM6" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127ES5" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127ES6" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127ES9" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF current_worker_number =	"X127EV1" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF current_worker_number =	"X127EV5" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF current_worker_number =	"X127EV2" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF current_worker_number =	"X127EV4" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF current_worker_number =	"X127EV3" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF current_worker_number =	"X127ET2"  	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127ET3" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127FJ2" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127ET1" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EM5" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EZ2" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EZ9" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EZ4" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EZ3" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EW9"	THEN email_address = "hsph.es.team.462@hennepin.us"
IF current_worker_number =	"X127EZ5" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127EZ8" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127FH6" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127EZ6" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127EZ7" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127EZ0" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF current_worker_number =	"X127FA5" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127FA6" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127FA7" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127FA8" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127FB1" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127FA9" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF current_worker_number =	"X127ET4" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127ET6" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127ET8" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127F4C" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127F4F" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127F4D" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127ET7"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127ET5" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127BV3"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127FB9"  	THEN email_address = "hsph.es.team.467@hennepin.us"
IF current_worker_number =	"X127FC1" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF current_worker_number =	"X127FC2"  	THEN email_address = "hsph.es.team.467@hennepin.us"
IF current_worker_number =	"X127EL1"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF current_worker_number =	"X127FB8" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF current_worker_number =	"X127FB7" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF current_worker_number =	"X127FD4" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FD5" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FD8" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FD6" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FD9" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FD7" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127EDD" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF current_worker_number =	"X127FG1" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X127EW6" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X1274EC" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X127FG2" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X127EW4" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X127EW5" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF current_worker_number =	"X127FE7" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF current_worker_number =	"X127FE8" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF current_worker_number =	"X127FE9" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF current_worker_number =	"X127EX4" 	THEN email_address = "hsph.es.team.601@hennepin.us"
IF current_worker_number =	"X127EX5" 	THEN email_address = "hsph.es.team.601@hennepin.us"
IF current_worker_number =	"X127FF1"	THEN email_address = "hsph.es.team.601@hennepin.us"
IF current_worker_number =	"X127FF2"	THEN email_address = "hsph.es.team.601@hennepin.us"
IF current_worker_number =	"X127EN8" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127EN9" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127FH3" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127F3E" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127F3J" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127F3N" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127FI6" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF current_worker_number =	"X127F4A" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF current_worker_number =	"X127F4B" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF current_worker_number =	"X127FI1"	THEN email_address = "hsph.es.team.603@hennepin.us"
IF current_worker_number =	"X127FI3" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF current_worker_number =	"X127EQ6" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF current_worker_number =	"X127EQ7" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF current_worker_number =	"X127EP1" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF current_worker_number =	"X127EP2" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF current_worker_number =	"X127FE2" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF current_worker_number =	"X127FE3" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF current_worker_number =	"X127FG5" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF current_worker_number =	"X127FG9" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF current_worker_number =	"X127EW7"	THEN email_address = "hsph.es.team.ebenezer@hennepin.us"
IF current_worker_number =	"X127EW8"	THEN email_address = "hsph.es.team.ebenezer@hennepin.us"
IF current_worker_number =	"X127ER7" 	THEN email_address = "hsph.es.team.mhc@hennepin.us"
IF current_worker_number =	"X127SH1" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"
IF current_worker_number =	"X127AN1" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"
IF current_worker_number =	"X127EHD" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"

'=========================================================================================================================================================================== CLASSES USED BY SCRIPTS
'A class for each script item
class script

	public script_name             	'The familiar name of the script
	public file_name               	'The actual file name
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
    public SIR_instructions_URL    	'The instructions URL in SIR
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run
	public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property

end class
'=========================================================================================================================================================================== END OF CLASSES

























'BELOW ARE THE ACTUAL FUNCTIONS--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function ABAWD_FSET_exemption_finder()
'--- This function screens for ABAWD/FSET exemptions for SNAP.
'===== Keywords: MAXIS, ABAWD, FSET, exemption, SNAP
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    '>>>>>Checking for privileged<<<<<
    row = 1
    col = 1
    EMSearch "PRIVILEGED", row, col
    IF row <> 0 THEN script_end_procedure("This case appears to be privileged. The script cannot access it.")

    DO
    	CALL HH_member_custom_dialog(HH_member_array)
    	IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
    LOOP UNTIL uBound(HH_member_array) <> -1

    'Building a placeholder array for EATS group comparison
    placeholder_HH_array = ""
    person_count = 0
    FOR EACH person IN HH_member_array
    	placeholder_HH_array = placeholder_HH_array & person & ","
    NEXT

    CALL check_for_MAXIS(False)

    closing_message = ""

    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen cl_age, 2, 8, 76
    		IF cl_age = "  " THEN cl_age = 0
    		cl_age = cl_age * 1
    		IF cl_age < 18 OR cl_age >= 50 THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have exemption. Age = " & cl_age & "."
    	END IF
    NEXT

    CALL navigate_to_MAXIS_screen("STAT", "DISA")
    FOR EACH person IN HH_member_array
    	disa_status = false
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_DISA, 1, 2, 78
    		IF num_of_DISA <> "0" THEN
    			EMReadScreen disa_end_dt, 10, 6, 69
    			disa_end_dt = replace(disa_end_dt, " ", "/")
    			EMReadScreen cert_end_dt, 10, 7, 69
    			cert_end_dt = replace(cert_end_dt, " ", "/")
    			IF IsDate(disa_end_dt) = True THEN
    				IF DateDiff("D", date, disa_end_dt) > 0 THEN
    					closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA end date = " & disa_end_dt & "."
    					disa_status = True
    				END IF
    			ELSE
    				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    					closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA has no end date."
    					disa_status = True
    				END IF
    			END IF
    			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
    				IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
    			ELSE
    				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
    					EMReadScreen cert_begin_dt, 8, 7, 47
    					IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA certification has no end date."
    				END IF
    			END IF
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>>> EATS GROUP
    FOR EACH person IN HH_member_array
    	CALL navigate_to_MAXIS_screen("STAT", "EATS")
    	eats_group_members = ""
    	memb_found = True
    	EMReadScreen all_eat_together, 1, 4, 72
    	IF all_eat_together = "_" THEN
    		eats_group_members = "01" & ","
    	ELSEIF all_eat_together = "Y" THEN
    		eats_row = 5
    		DO
    			EMReadScreen eats_person, 2, eats_row, 3
    			eats_person = replace(eats_person, " ", "")
    			IF eats_person <> "" THEN
    				eats_group_members = eats_group_members & eats_person & ","
    				eats_row = eats_row + 1
    			END IF
    		LOOP UNTIL eats_person = ""
    	ELSEIF all_eat_together = "N" THEN
    		eats_row = 13
    		DO
    			EMReadScreen eats_group, 38, eats_row, 39
    			find_memb01 = InStr(eats_group, person)
    			IF find_memb01 = 0 THEN
    				eats_row = eats_row + 1
    				IF eats_row = 18 THEN
    					memb_found = False
    					EXIT DO
    				END IF
    			END IF
    		LOOP UNTIL find_memb01 <> 0
    		eats_col = 39
    		DO
    			EMReadScreen eats_group, 2, eats_row, eats_col
    			IF eats_group <> "__" THEN
    				eats_group_members = eats_group_members & eats_group & ","
    				eats_col = eats_col + 4
    			END IF
    		LOOP UNTIL eats_group = "__"
    	END IF

    	IF memb_found = True THEN
    		IF placeholder_HH_array <> eats_group_members THEN script_end_procedure("You are asking the script to verify ABAWD and SNAP E&T exemptions for a household that does not match the EATS group. The script cannot support this request. It will now end." & vbCr & vbCr & "Please re-run the script selecting only the individuals in the EATS group.")
    		eats_group_members = trim(eats_group_members)
    		eats_group_members = split(eats_group_members, ",")

    		IF all_eat_together <> "_" THEN
    			CALL write_value_and_transmit("MEMB", 20, 71)
    			FOR EACH eats_pers IN eats_group_members
    				IF eats_pers <> "" AND person <> eats_pers THEN
    					CALL write_value_and_transmit(eats_pers, 20, 76)
    					EMReadScreen cl_age, 2, 8, 76
    					IF cl_age = "  " THEN cl_age = 0
    						cl_age = cl_age * 1
    						IF cl_age =< 17 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": May have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
    						END IF
    				END IF
    			NEXT
    		END IF

    		CALL write_value_and_transmit("DISA", 20, 71)
    		FOR EACH disa_pers IN eats_group_members
    			disa_status = false
    			IF disa_pers <> "" AND disa_pers <> person THEN
    				CALL write_value_and_transmit(disa_pers, 20, 76)
    				EMReadScreen num_of_DISA, 1, 2, 78
    				IF num_of_DISA <> "0" THEN
    					EMReadScreen disa_end_dt, 10, 6, 69
    					disa_end_dt = replace(disa_end_dt, " ", "/")
    					EMReadScreen cert_end_dt, 10, 7, 69
    					cert_end_dt = replace(cert_end_dt, " ", "/")
    					IF IsDate(disa_end_dt) = True THEN
    						IF DateDiff("D", date, disa_end_dt) > 0 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": MAY have an exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = TRUE
    						END IF
    					ELSEIF IsDate(disa_end_dt) = False THEN
    						IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    							closing_message = closing_message & vbCr & "* M" & person & " : MAY have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = true
    						END IF
    					END IF
    					IF IsDate(cert_end_dt) = True AND disa_status = False THEN
    						IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
    					ELSE
    						IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN
    							EMReadScreen cert_begin_dt, 8, 7, 47
    							IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
    						END IF
    					END IF
    				END IF
    			END IF
    		NEXT
    	END IF
    NEXT

    '>>>>>>>>>>>>>>EARNED INCOME
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		prosp_inc = 0
    		prosp_hrs = 0
    		prospective_hours = 0

    		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
    		EMWritescreen person, 20, 76
    		EMWritescreen "01", 20, 79				'ensures that we start at 1st job
    		transmit
    		EMReadScreen num_of_JOBS, 1, 2, 78
    		IF num_of_JOBS <> "0" THEN
    			DO
    			 	EMReadScreen jobs_end_dt, 8, 9, 49
    				EMReadScreen cont_end_dt, 8, 9, 73
    				IF jobs_end_dt = "__ __ __" THEN
    					CALL write_value_and_transmit("X", 19, 38)     'Entering the PIC
    					EMReadScreen prosp_monthly, 8, 18, 56
    					prosp_monthly = trim(prosp_monthly)
    					IF prosp_monthly = "" THEN prosp_monthly = 0
    					prosp_inc = prosp_inc + prosp_monthly
    					EMReadScreen prosp_hrs, 8, 16, 50
    					IF prosp_hrs = "        " THEN prosp_hrs = 0
    					prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
    					EMReadScreen pay_freq, 1, 5, 64
    					IF pay_freq = "1" THEN
    						prosp_hrs = prosp_hrs
    					ELSEIF pay_freq = "2" THEN
    						prosp_hrs = (2 * prosp_hrs)
    					ELSEIF pay_freq = "3" THEN
    						prosp_hrs = (2.15 * prosp_hrs)
    					ELSEIF pay_freq = "4" THEN
    						prosp_hrs = (4.3 * prosp_hrs)
    					END IF
                        transmit		'to exit PIC
    					prospective_hours = prospective_hours + prosp_hrs
    				ELSE
    					jobs_end_dt = replace(jobs_end_dt, " ", "/")
    					IF DateDiff("D", date, jobs_end_dt) > 0 THEN
    						'Going into the PIC for a job with an end date in the future
    						CALL write_value_and_transmit("X", 19, 38)        'Entering the PIC
    						EMReadScreen prosp_monthly, 8, 18, 56
    						prosp_monthly = trim(prosp_monthly)
    						IF prosp_monthly = "" THEN prosp_monthly = 0
    						prosp_inc = prosp_inc + prosp_monthly
    						EMReadScreen prosp_hrs, 8, 16, 50
    						IF prosp_hrs = "        " THEN prosp_hrs = 0
    						prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
    						EMReadScreen pay_freq, 1, 5, 64
    						IF pay_freq = "1" THEN
    							prosp_hrs = prosp_hrs
    						ELSEIF pay_freq = "2" THEN
    							prosp_hrs = (2 * prosp_hrs)
    						ELSEIF pay_freq = "3" THEN
    							prosp_hrs = (2.15 * prosp_hrs)
    						ELSEIF pay_freq = "4" THEN
    							prosp_hrs = (4.3 * prosp_hrs)
    						END IF
                            transmit		'to exit PIC
    						'added seperate incremental variable to account for multiple jobs
    						prospective_hours = prospective_hours + prosp_hrs
    					END IF
    				END IF

    				EMReadScreen JOBS_panel_current, 1, 2, 73
    				'looping until all the jobs panels are calculated
    				If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
    			Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
    		END IF

    		EMWriteScreen "BUSI", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_BUSI, 1, 2, 78
    		IF num_of_BUSI <> "0" THEN
    			DO
    				EMReadScreen busi_end_dt, 8, 5, 72
    				busi_end_dt = replace(busi_end_dt, " ", "/")
    				IF IsDate(busi_end_dt) = True THEN
    					IF DateDiff("D", date, busi_end_dt) > 0 THEN
    						EMReadScreen busi_inc, 8, 10, 69
    						busi_inc = trim(busi_inc)
    						EMReadScreen busi_hrs, 3, 13, 74
    						busi_hrs = trim(busi_hrs)
    						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
    						prosp_inc = prosp_inc + busi_inc
    						prosp_hrs = prosp_hrs + busi_hrs
    						prospective_hours = prospective_hours + busi_hrs
    					END IF
    				ELSE
    					IF busi_end_dt = "__/__/__" THEN
    						EMReadScreen busi_inc, 8, 10, 69
    						busi_inc = trim(busi_inc)
    						EMReadScreen busi_hrs, 3, 13, 74
    						busi_hrs = trim(busi_hrs)
    						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
    						prosp_inc = prosp_inc + busi_inc
    						prosp_hrs = prosp_hrs + busi_hrs
    						prospective_hours = prospective_hours + busi_hrs
    					END IF
    				END IF
    				transmit
    				EMReadScreen enter_a_valid, 13, 24, 2
    			LOOP UNTIL enter_a_valid = "ENTER A VALID"
    		END IF

    		EMWriteScreen "RBIC", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_RBIC, 1, 2, 78
    		IF num_of_RBIC <> "0" THEN closing_message = closing_message & vbCr & "* M" & person & ": Has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
    		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working 30 hours/wk (regardless of wage level) or earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
    		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption and SNAP E&T exemptions."
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>>>UNEA
    CALL navigate_to_MAXIS_screen("STAT", "UNEA")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_UNEA, 1, 2, 78
    		IF num_of_UNEA <> "0" THEN
    			DO
    				EMReadScreen unea_type, 2, 5, 37
    				EMReadScreen unea_end_dt, 8, 7, 68
    				unea_end_dt = replace(unea_end_dt, " ", "/")
    				IF IsDate(unea_end_dt) = True THEN
    					IF DateDiff("D", date, unea_end_dt) > 0 THEN
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
    					END IF
    				ELSE
    					IF unea_end_dt = "__/__/__" THEN
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
    					END IF
    				END IF
    				transmit
    				EMReadScreen enter_a_valid, 13, 24, 2
    			LOOP UNTIL enter_a_valid = "ENTER A VALID"
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>PBEN
    CALL navigate_to_MAXIS_screen("STAT", "PBEN")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		EMWriteScreen "PBEN", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_PBEN, 1, 2, 78
    		IF num_of_PBEN <> "0" THEN
    			pben_row = 8
    			DO
    			    IF pben_type = "12" THEN		'UI pending'
    					EMReadScreen pben_disp, 1, pben_row, 77
    					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
    						closing_message = closing_message & vbCr & "* M" & person & ": Appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption."
    						EXIT DO
    					END IF
    				ELSE
    					pben_row = pben_row + 1
    				END IF
    			LOOP UNTIL pben_row = 14
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>PREG
    CALL navigate_to_MAXIS_screen("STAT", "PREG")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_PREG, 1, 2, 78
            EMReadScreen preg_due_dt, 8, 10, 53
            preg_due_dt = replace(preg_due_dt, " ", "/")
    		EMReadScreen preg_end_dt, 8, 12, 53

    		IF num_of_PREG <> "0" THen
                If preg_due_dt <> "__/__/__" Then
                    If DateDiff("d", date, preg_due_dt) > 0 AND preg_end_dt = "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active pregnancy. Please review for ABAWD exemption."
                    If DateDiff("d", date, preg_due_dt) < 0 Then closing_message = closing_message & vbCr & "* M" & person & ": Appears to have an overdue pregnancy, person may meet a minor child exemption. Contact client."
                End If
            End If
        END IF
    NEXT

    '>>>>>>>>>>PROG
    CALL navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash1_status, 4, 6, 74
    EMReadScreen cash2_status, 4, 7, 74
    IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN closing_message = closing_message & vbCr & "* Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."

    '>>>>>>>>>>ADDR
    CALL navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_code, 1, 10, 43
    EmReadscreen addr_line_01, 16, 6, 43

    IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN closing_message = closing_message & vbCr & "* Client is claiming homelessness. If client has barriers to employment, they could meet the 'Unfit for Employment' exemption. Exemption began 05/2018."

    '>>>>>>>>>SCHL/STIN/STEC
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
            CALL navigate_to_MAXIS_screen("STAT", "SCHL")
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_SCHL, 1, 2, 78
    		IF num_of_SCHL = "1" THEN
    			EMReadScreen school_status, 1, 6, 40
    			IF school_status <> "N" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
    		ELSE
    			EMWriteScreen "STIN", 20, 71
    			CALL write_value_and_transmit(person, 20, 76)
    			EMReadScreen num_of_STIN, 1, 2, 78
    			IF num_of_STIN = "1" THEN
    				STIN_row = 8
    				DO
    					EMReadScreen cov_thru, 5, STIN_row, 67
    					IF cov_thru <> "__ __" THEN
    						cov_thru = replace(cov_thru, " ", "/01/")
    						cov_thru = DateAdd("M", 1, cov_thru)
    						cov_thru = DateAdd("D", -1, cov_thru)
    						IF DateDiff("D", date, cov_thru) > 0 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
    							EXIT DO
    						ELSE
    							STIN_row = STIN_row + 1
    							IF STIN_row = 18 THEN
    								PF20
    								STIN_row = 8
    								EMReadScreen last_page, 21, 24, 2
    								IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
    							END IF
    						END IF
    					ELSE
    						EXIT DO
    					END IF
    				LOOP
    			ELSE
    				EMWriteScreen "STEC", 20, 71
    				CALL write_value_and_transmit(person, 20, 76)
    				EMReadScreen num_of_STEC, 1, 2, 78
    				IF num_of_STEC = "1" THEN
    					STEC_row = 8
    					DO
    						EMReadScreen stec_thru, 5, STEC_row, 48
    						IF stec_thru <> "__ __" THEN
    							stec_thru = replace(stec_thru, " ", "/01/")
    							stec_thru = DateAdd("M", 1, stec_thru)
    							stec_thru = DateAdd("D", -1, stec_thru)
    							IF DateDiff("D", date, stec_thru) > 0 THEN
    								closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
    								EXIT DO
    							ELSE
    								STEC_row = STEC_row + 1
    								IF STEC_row = 17 THEN
    									PF20
    									STEC_row = 8
    									EMReadScreen last_page, 21, 24, 2
    									IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
    								END IF
    							END IF
    						ELSE
    							EXIT DO
    						END IF
    					LOOP
    				END IF
    			END IF
    		END IF
    	END IF
    	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
    NEXT

    household_persons = ""
    pers_count = 0

    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		IF pers_count = uBound(HH_member_array) THEN
    			IF pers_count = 0 THEN
    				household_persons = household_persons & person
    			ELSE
    				household_persons = household_persons & "and " & person
    			END IF
    		ELSE
    			household_persons = household_persons & person & ", "
    			pers_count = pers_count + 1
    		END IF
    	END IF
    NEXT

    IF closing_message = "" THEN
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "It appears there are NO missed exemptions for ABAWD or SNAP E&T in MAXIS for this case. The script has checked ADDR, EATS, MEMB, DISA, JOBS, BUSI, RBIC, UNEA, PREG, PROG, PBEN, SCHL, STIN, and STEC for member(s) " & household_persons & "." & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    ELSE
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "The script has checked for ABAWD and SNAP E&T exemptions coded in MAXIS for member(s) " & household_persons & "." & vbCr & closing_message & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    END IF

    'Displaying the results...now with added MsgBox bling.
    'vbSystemModal will keep the results in the foreground.
    MsgBox closing_message, vbInformation + vbSystemModal, "ABAWD/FSET Exemption Check -- Results"

    STATS_counter = STATS_counter - 1		'Removing one instance from the STATS Counter as it started with one at the beginning
End Function

function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ access_type:indicates if we need to read or write to the ADDR panel, only options are 'READ' or 'WRITE' - case does not matter - the function wil ucase it
'~~~~~ notes_on_address: string - information read from the panel and used typically in an editbox
'~~~~~ resi_line_one: string - information from the first line of the residence street address
'~~~~~ resi_line_two: string - information from the second line of the residence street address
'~~~~~ resi_street_full: string - information from the first and second line of the residence street address
'~~~~~ resi_city: string - information from the city line of the residence address
'~~~~~ resi_state: string - information from the state line of the residence address - this is formatted as MN - Minnesota
'~~~~~ resi_zip: string - information of the zip code from the residence address
'~~~~~ resi_county: string - information of the county of the residence address - this is formatted as 27 Hennepin
'~~~~~ addr_verif: string - information of the verification listed on ADDR - formattes as SF - Shlter Form
'~~~~~ addr_homeless: string of 'Yes' or 'No' from the homeless yes/no code on ADDR
'~~~~~ addr_reservation: string of 'Yes' or 'No' from the reservation yes/no code on ADDR
'~~~~~ addr_living_sit: string - information of the living situation - formatted as 01 - Own Home, or as Blank
'~~~~~ reservation_name: string - information of the reservation name - formatted as LL - Leech Lake
'~~~~~ mail_line_one: string - information from first line of street address of the mailing address
'~~~~~ mail_line_two: string - information from the second line of street address of the mailing address
'~~~~~ mail_street_full: string - information of both the first and second line of the street address of the mailing address
'~~~~~ mail_city: string - information of the city listed on the mailing address
'~~~~~ mail_state: string - information of the state listed on the mailing address - this is formatted as MN - Minnesota
'~~~~~ mail_zip: string - information of the zip code listed on the mailing address
'~~~~~ addr_eff_date: string - the date listed as the effective date on the top of the panel. Formatted as MM/DD/YY - script may read as a date
'~~~~~ addr_future_date: string - the date listed at the top if there is a future date of a change in a future month. formatted as MM/DD/YY - script may read as a date
'~~~~~ phone_one: string - information listed on the first phone line - formatted as xxx-xxx-xxxx
'~~~~~ phone_two: string - information listed on the second phone line - formatted as xxx-xxx-xxxx
'~~~~~ phone_three: string - information listed on the third phone line - formatted as xxx-xxx-xxxx
'~~~~~ type_one: string - information listed for the type of the first phone  - formatted as C - Cell
'~~~~~ type_two: string - information listed for the type of the second phone  - formatted as C - Cell
'~~~~~ type_three: string - information listed for the type of the third phone  - formatted as C - Cell
'~~~~~ text_yn_one: string - information listed as the authorization on the first phone for texting - formatted as Y or N
'~~~~~ text_yn_two: string - information listed as the authorization on the second phone for texting - formatted as Y or N
'~~~~~ text_yn_three: string - information listed as the authorization on the third phone for texting - formatted as Y or N
'~~~~~ addr_email: string - information listed on the email line. any trailing '_' will be trimmed off but NOT any in the middle
'~~~~~ verif_received: string - StRING - used as entry on WRITE only
'~~~~~ original_information: string that saves the original detail known on the panel - this is used to see if there is an update to any information
'~~~~~ update_attempted: boolean - output after 'WRITE' to indicate that an update of panel information was attempted.
'===== Keywords: MAXIS, update, ADDR
	access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")							'going to ADDR
		EMReadScreen curr_addr_footer_month, 2, 20, 55							'reading the footer month and year on the panel because there is footer months pecific differences
		EMReadScreen curr_addr_footer_year, 2, 20, 58
		the_footer_month_date = curr_addr_footer_month &  "/1/" & curr_addr_footer_year		'making a date out of the footer month
		the_footer_month_date = DateAdd("d", 0, the_footer_month_date)
		new_version = True														'defaulting to using the new verion of the panel with text authorization and email
		If DateDiff("d", the_footer_month_date, #10/1/2021#) > 0 Then new_version = False	'if the footer months is BEFORE 10/1/2021, then we need to read the old version

        EMReadScreen line_one, 22, 6, 43										'reading the information from the top half of the ADDR panel
        EMReadScreen line_two, 22, 7, 43
        EMReadScreen city_line, 15, 8, 43
        EMReadScreen state_line, 2, 8, 66
        EMReadScreen zip_line, 7, 9, 43
        EMReadScreen county_line, 2, 9, 66
        EMReadScreen verif_line, 2, 9, 74
        EMReadScreen homeless_line, 1, 10, 43
        EMReadScreen reservation_line, 1, 10, 74
        EMReadScreen living_sit_line, 2, 11, 43
		EMReadScreen reservation_name, 2, 11, 74

        resi_line_one = replace(line_one, "_", "")								'formatting the residence address information to remove the '_'
        resi_line_two = replace(line_two, "_", "")
		resi_street_full = trim(resi_line_one & " " & resi_line_two)
        resi_city = replace(city_line, "_", "")
        resi_zip = replace(zip_line, "_", "")

        If county_line = "01" Then addr_county = "01 Aitkin"					'Adding the county name to the county string
        If county_line = "02" Then addr_county = "02 Anoka"
        If county_line = "03" Then addr_county = "03 Becker"
        If county_line = "04" Then addr_county = "04 Beltrami"
        If county_line = "05" Then addr_county = "05 Benton"
        If county_line = "06" Then addr_county = "06 Big Stone"
        If county_line = "07" Then addr_county = "07 Blue Earth"
        If county_line = "08" Then addr_county = "08 Brown"
        If county_line = "09" Then addr_county = "09 Carlton"
        If county_line = "10" Then addr_county = "10 Carver"
        If county_line = "11" Then addr_county = "11 Cass"
        If county_line = "12" Then addr_county = "12 Chippewa"
        If county_line = "13" Then addr_county = "13 Chisago"
        If county_line = "14" Then addr_county = "14 Clay"
        If county_line = "15" Then addr_county = "15 Clearwater"
        If county_line = "16" Then addr_county = "16 Cook"
        If county_line = "17" Then addr_county = "17 Cottonwood"
        If county_line = "18" Then addr_county = "18 Crow Wing"
        If county_line = "19" Then addr_county = "19 Dakota"
        If county_line = "20" Then addr_county = "20 Dodge"
        If county_line = "21" Then addr_county = "21 Douglas"
        If county_line = "22" Then addr_county = "22 Faribault"
        If county_line = "23" Then addr_county = "23 Fillmore"
        If county_line = "24" Then addr_county = "24 Freeborn"
        If county_line = "25" Then addr_county = "25 Goodhue"
        If county_line = "26" Then addr_county = "26 Grant"
        If county_line = "27" Then addr_county = "27 Hennepin"
        If county_line = "28" Then addr_county = "28 Houston"
        If county_line = "29" Then addr_county = "29 Hubbard"
        If county_line = "30" Then addr_county = "30 Isanti"
        If county_line = "31" Then addr_county = "31 Itasca"
        If county_line = "32" Then addr_county = "32 Jackson"
        If county_line = "33" Then addr_county = "33 Kanabec"
        If county_line = "34" Then addr_county = "34 Kandiyohi"
        If county_line = "35" Then addr_county = "35 Kittson"
        If county_line = "36" Then addr_county = "36 Koochiching"
        If county_line = "37" Then addr_county = "37 Lac Qui Parle"
        If county_line = "38" Then addr_county = "38 Lake"
        If county_line = "39" Then addr_county = "39 Lake Of Woods"
        If county_line = "40" Then addr_county = "40 Le Sueur"
        If county_line = "41" Then addr_county = "41 Lincoln"
        If county_line = "42" Then addr_county = "42 Lyon"
        If county_line = "43" Then addr_county = "43 Mcleod"
        If county_line = "44" Then addr_county = "44 Mahnomen"
        If county_line = "45" Then addr_county = "45 Marshall"
        If county_line = "46" Then addr_county = "46 Martin"
        If county_line = "47" Then addr_county = "47 Meeker"
        If county_line = "48" Then addr_county = "48 Mille Lacs"
        If county_line = "49" Then addr_county = "49 Morrison"
        If county_line = "50" Then addr_county = "50 Mower"
        If county_line = "51" Then addr_county = "51 Murray"
        If county_line = "52" Then addr_county = "52 Nicollet"
        If county_line = "53" Then addr_county = "53 Nobles"
        If county_line = "54" Then addr_county = "54 Norman"
        If county_line = "55" Then addr_county = "55 Olmsted"
        If county_line = "56" Then addr_county = "56 Otter Tail"
        If county_line = "57" Then addr_county = "57 Pennington"
        If county_line = "58" Then addr_county = "58 Pine"
        If county_line = "59" Then addr_county = "59 Pipestone"
        If county_line = "60" Then addr_county = "60 Polk"
        If county_line = "61" Then addr_county = "61 Pope"
        If county_line = "62" Then addr_county = "62 Ramsey"
        If county_line = "63" Then addr_county = "63 Red Lake"
        If county_line = "64" Then addr_county = "64 Redwood"
        If county_line = "65" Then addr_county = "65 Renville"
        If county_line = "66" Then addr_county = "66 Rice"
        If county_line = "67" Then addr_county = "67 Rock"
        If county_line = "68" Then addr_county = "68 Roseau"
        If county_line = "69" Then addr_county = "69 St. Louis"
        If county_line = "70" Then addr_county = "70 Scott"
        If county_line = "71" Then addr_county = "71 Sherburne"
        If county_line = "72" Then addr_county = "72 Sibley"
        If county_line = "73" Then addr_county = "73 Stearns"
        If county_line = "74" Then addr_county = "74 Steele"
        If county_line = "75" Then addr_county = "75 Stevens"
        If county_line = "76" Then addr_county = "76 Swift"
        If county_line = "77" Then addr_county = "77 Todd"
        If county_line = "78" Then addr_county = "78 Traverse"
        If county_line = "79" Then addr_county = "79 Wabasha"
        If county_line = "80" Then addr_county = "80 Wadena"
        If county_line = "81" Then addr_county = "81 Waseca"
        If county_line = "82" Then addr_county = "82 Washington"
        If county_line = "83" Then addr_county = "83 Watonwan"
        If county_line = "84" Then addr_county = "84 Wilkin"
        If county_line = "85" Then addr_county = "85 Winona"
        If county_line = "86" Then addr_county = "86 Wright"
        If county_line = "87" Then addr_county = "87 Yellow Medicine"
        If county_line = "89" Then addr_county = "89 Out-of-State"
        resi_county = addr_county

		Call get_state_name_from_state_code(state_line, resi_state, TRUE)		'This function makes the state code to be the state name written out - including the code

        If homeless_line = "Y" Then addr_homeless = "Yes"						'formatting the Y and N to 'Yes' or 'No'
        If homeless_line = "N" Then addr_homeless = "No"
        If reservation_line = "Y" Then addr_reservation = "Yes"
        If reservation_line = "N" Then addr_reservation = "No"

		If reservation_name = "__" Then reservation_name = ""					'Filling in the detail of the reservation name
		If reservation_name = "BD" Then reservation_name = "BD - Bois Forte - Deer Creek"
		If reservation_name = "BN" Then reservation_name = "BN - Bois Forte - Nett Lake"
		If reservation_name = "BV" Then reservation_name = "BV - Bois Forte - Vermillion Lk"
		If reservation_name = "FL" Then reservation_name = "FL - Fond du Lac"
		If reservation_name = "GP" Then reservation_name = "GP - Grand Portage"
		If reservation_name = "LL" Then reservation_name = "LL - Leach Lake"
		If reservation_name = "LS" Then reservation_name = "LS - Lower Sioux"
		If reservation_name = "ML" Then reservation_name = "ML - Mille Lacs"
		If reservation_name = "PL" Then reservation_name = "PL - Prairie Island Community"
		If reservation_name = "RL" Then reservation_name = "RL - Red Lake"
		If reservation_name = "SM" Then reservation_name = "SM - Shakopee Mdewakanton"
		If reservation_name = "US" Then reservation_name = "US - Upper Sioux"
		If reservation_name = "WE" Then reservation_name = "WE - White Earth"

        If verif_line = "SF" Then addr_verif = "SF - Shelter Form"				'filling in the detail of the verification listed
        If verif_line = "CO" Then addr_verif = "CO - Coltrl Stmt"
        If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
        If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
        If verif_line = "CD" Then addr_verif = "CD - Contrct for Deed"
        If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
        If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
        If verif_line = "OT" Then addr_verif = "OT - Other Document"
        If verif_line = "NO" Then addr_verif = "NO - No Verif"
        If verif_line = "?_" Then addr_verif = "? - Delayed"
        If verif_line = "__" Then addr_verif = "Blank"


        If living_sit_line = "__" Then living_situation = "Blank"				'Adding detail to the living situation code
        If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roommate"
        If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
        If living_sit_line = "03" Then living_situation = "03 -  servc prvdr- foster/group home"
        If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
        If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
        If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
        If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
        If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
        If living_sit_line = "09" Then living_situation = "09 - Declined"
        If living_sit_line = "10" Then living_situation = "10 - Unknown"
        addr_living_sit = living_situation

        EMReadScreen addr_eff_date, 8, 4, 43									'reading the dates on the panel
        EMReadScreen addr_future_date, 8, 4, 66

		If new_version = False Then												'reading the bottom half of the panel based on if we are in the new or old version
	        EMReadScreen mail_line_one, 22, 13, 43
	        EMReadScreen mail_line_two, 22, 14, 43
	        EMReadScreen mail_city_line, 15, 15, 43
	        EMReadScreen mail_state_line, 2, 16, 43
	        EMReadScreen mail_zip_line, 7, 16, 52

			EMReadScreen phone_one, 14, 17, 45
			EMReadScreen phone_two, 14, 18, 45
			EMReadScreen phone_three, 14, 19, 45

			EMReadScreen type_one, 1, 17, 67
			EMReadScreen type_two, 1, 18, 67
			EMReadScreen type_three, 1, 19, 67
		End If

		If new_version = True Then
			EMReadScreen mail_line_one, 22, 12, 49
			EMReadScreen mail_line_two, 22, 13, 49
			EMReadScreen mail_city_line, 15, 14, 49
			EMReadScreen mail_state_line, 2, 15, 49
			EMReadScreen mail_zip_line, 7, 15, 58

			EMReadScreen phone_one, 14, 16, 39
			EMReadScreen phone_two, 14, 17, 39
			EMReadScreen phone_three, 14, 18, 39

			EMReadScreen type_one, 1, 16, 61
			EMReadScreen type_two, 1, 17, 61
			EMReadScreen type_three, 1, 18, 61

			EMReadScreen text_yn_one, 1, 16, 76
			EMReadScreen text_yn_two, 1, 17, 76
			EMReadScreen text_yn_three, 1, 18, 76


			EMReadScreen addr_email, 50, 19, 31
		End If

        addr_eff_date = replace(addr_eff_date, " ", "/")						'formatting the information from the second half
        addr_future_date = trim(addr_future_date)
        addr_future_date = replace(addr_future_date, " ", "/")
        mail_line_one = replace(mail_line_one, "_", "")
        mail_line_two = replace(mail_line_two, "_", "")
		mail_street_full = trim(mail_line_one & " " & mail_line_two)
        mail_city = replace(mail_city_line, "_", "")
        mail_state = replace(mail_state_line, "_", "")
        mail_zip = replace(mail_zip_line, "_", "")

		If text_yn_one = "_" Then text_yn_one = ""								'Changing blanks to nulls
		If text_yn_two = "_" Then text_yn_two = ""
		If text_yn_three = "_" Then text_yn_three = ""
		'here we remove the '_' from the front and back of the string - we cannot use replace here because '_' may be a part of the email address.
		Do
			If right(addr_email, 1) = "_" Then addr_email = left(addr_email, len(addr_email) - 1)
			If left(addr_email, 1) = "_" Then addr_email = right(addr_email, len(addr_email) - 1)
		Loop until right(addr_email, 1) <> "_" AND  left(addr_email, 1) <> "_"

        notes_on_address = "Address effective on " & addr_eff_date & "."
        ' If mail_line_one <> "" Then
        '     If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        '     If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        ' End If
        If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."
		notes_on_address = notes_on_address & "; "

        phone_one = replace(phone_one, " ) ", "-")								'formatting phone numbers
        phone_one = replace(phone_one, " ", "-")
        If phone_one = "___-___-____" Then phone_one = ""

        phone_two = replace(phone_two, " ) ", "-")
        phone_two = replace(phone_two, " ", "-")
        If phone_two = "___-___-____" Then phone_two = ""

        phone_three = replace(phone_three, " ) ", "-")
        phone_three = replace(phone_three, " ", "-")
        If phone_three = "___-___-____" Then phone_three = ""

        If type_one = "H" Then type_one = "H - Home"
        If type_one = "W" Then type_one = "W - Work"
        If type_one = "C" Then type_one = "C - Cell"
        If type_one = "M" Then type_one = "M - Message"
        If type_one = "T" Then type_one = "T - TTY/TDD"
        If type_one = "_" Then type_one = ""

        If type_two = "H" Then type_two = "H - Home"
        If type_two = "W" Then type_two = "W - Work"
        If type_two = "C" Then type_two = "C - Cell"
        If type_two = "M" Then type_two = "M - Message"
        If type_two = "T" Then type_two = "T - TTY/TDD"
        If type_two = "_" Then type_two = ""

        If type_three = "H" Then type_three = "H - Home"
        If type_three = "W" Then type_three = "W - Work"
        If type_three = "C" Then type_three = "C - Cell"
        If type_three = "M" Then type_three = "M - Message"
        If type_three = "T" Then type_three = "T - TTY/TDD"
        If type_three = "_" Then type_three = ""

		'here we save the information we gathered to start with so that we can compare it and know if it changed
		original_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							   addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							   mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&_
							   text_yn_one&"|"&text_yn_two&"|"&text_yn_three&"|"&addr_email&"|"&addr_verif
		original_information = UCASE(original_information)
    End If

    If access_type = "WRITE" Then
		' verif_received 'add functionality to change how this is updated based on if we have verif or not.

		update_attempted = False
		resi_line_one = trim(resi_line_one)
		resi_line_two = trim(resi_line_two)
		resi_street_full = trim(resi_street_full)
		resi_city = trim(resi_city)
		resi_state = trim(resi_state)
		resi_zip = trim(resi_zip)
		resi_county = trim(resi_county)
		addr_verif = trim(addr_verif)
		addr_homeless = trim(addr_homeless)
		addr_reservation = trim(addr_reservation)
		addr_living_sit = trim(addr_living_sit)
		mail_line_one = trim(mail_line_one)
		mail_line_two = trim(mail_line_two)
		mail_street_full = trim(mail_street_full)
		mail_city = trim(mail_city)
		mail_state = trim(mail_state)
		mail_zip = trim(mail_zip)
		addr_eff_date = trim(addr_eff_date)
		addr_future_date = trim(addr_future_date)
		phone_one = trim(phone_one)
		phone_two = trim(phone_two)
		phone_three = trim(phone_three)
		type_one = trim(type_one)
		type_two = trim(type_two)
		type_three = trim(type_three)
		verif_received = trim(verif_received)

		current_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							  addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							  mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&_
							  text_yn_one&"|"&text_yn_two&"|"&text_yn_three&"|"&addr_email&"|"&addr_verif


		current_information = UCase(current_information)
		' MsgBox "THIS" & vbCR & "ORIGINAL" & vbCr & original_information & vbCr & vbCr & "CURRENT" & vbCr & current_information
		If current_information <> original_information Then						'If the information in the beginning and the information inthe end do not match - we need to update
			update_attempted = True

			Call navigate_to_MAXIS_screen("STAT", "ADDR")							'going to ADDR
			EMReadScreen curr_addr_footer_month, 2, 20, 55							'reading the footer month and year on the panel because there is footer months pecific differences
			EMReadScreen curr_addr_footer_year, 2, 20, 58
			the_footer_month_date = curr_addr_footer_month &  "/1/" & curr_addr_footer_year		'making a date out of the footer month
			the_footer_month_date = DateAdd("d", 0, the_footer_month_date)
			new_version = True														'defaulting to using the new verion of the panel with text authorization and email
			If DateDiff("d", the_footer_month_date, #10/1/2021#) > 0 Then new_version = False	'if the footer months is BEFORE 10/1/2021, then we need to read the old version

	        PF9																	'Put it in edit mode

			Call clear_line_of_text(6, 43) 	'residence addr 1
			Call clear_line_of_text(7, 43) 	'residence addr 2
			Call clear_line_of_text(8, 43) 	'residence city
			Call clear_line_of_text(8, 66) 	'residence state
			Call clear_line_of_text(9, 43) 	'residence zip

			If new_version = True then
				Call clear_line_of_text(12, 49) 	'mail addr 1
				Call clear_line_of_text(13, 49) 	'mail addr 2
				Call clear_line_of_text(14, 49) 	'mail city
				Call clear_line_of_text(15, 49) 	'mail state
				Call clear_line_of_text(15, 58) 	'mail zip

				Call clear_line_of_text(16, 39)		'phone information'
				Call clear_line_of_text(16, 45)
				Call clear_line_of_text(16, 49)
				Call clear_line_of_text(16, 61)
				Call clear_line_of_text(16, 76)

				Call clear_line_of_text(17, 39)
				Call clear_line_of_text(17, 45)
				Call clear_line_of_text(17, 49)
				Call clear_line_of_text(17, 61)
				Call clear_line_of_text(17, 76)

				Call clear_line_of_text(18, 39)
				Call clear_line_of_text(18, 45)
				Call clear_line_of_text(18, 49)
				Call clear_line_of_text(18, 61)
				Call clear_line_of_text(18, 76)

				Call clear_line_of_text(19, 31)
			End If
			If new_version = False then
				Call clear_line_of_text(13, 43) 	'mail addr 1
				Call clear_line_of_text(14, 43) 	'mail addr 2
				Call clear_line_of_text(15, 43) 	'mail city
				Call clear_line_of_text(16, 43) 	'mail state
				Call clear_line_of_text(16, 52) 	'mail zip

				Call clear_line_of_text(17, 45)		'phone information'
				Call clear_line_of_text(17, 15)
				Call clear_line_of_text(17, 55)
				Call clear_line_of_text(17, 67)

				Call clear_line_of_text(18, 45)
				Call clear_line_of_text(18, 15)
				Call clear_line_of_text(18, 55)
				Call clear_line_of_text(18, 67)

				Call clear_line_of_text(19, 45)
				Call clear_line_of_text(19, 15)
				Call clear_line_of_text(19, 55)
				Call clear_line_of_text(19, 67)
			End If

			'Now we write all the information
	        Call create_mainframe_friendly_date(addr_eff_date, 4, 43, "YY")

			resi_street_full = trim(trim(resi_line_one) & " " & trim(resi_line_two))
			If len(resi_line_one) > 22 or len(resi_line_two) > 22 Then
				'This functionality will write the address word by word so that if it needs to word wrap to the second line, it can move the words to the next line
	            resi_words = split(resi_street_full, " ")
	            write_resi_line_one = ""
	            write_resi_line_two = ""
	            For each word in resi_words
	                If write_resi_line_one = "" Then
	                    write_resi_line_one = word
	                ElseIf len(write_resi_line_one & " " & word) =< 22 Then
	                    write_resi_line_one = write_resi_line_one & " " & word
	                Else
	                    If write_resi_line_two = "" Then
	                        write_resi_line_two = word
	                    Else
	                        write_resi_line_two = write_resi_line_two & " " & word
	                    End If
	                End If
	            Next
	        Else
	            write_resi_line_one = resi_line_one
				write_resi_line_two = resi_line_two
	        End If
	        EMWriteScreen write_resi_line_one, 6, 43
	        EMWriteScreen write_resi_line_two, 7, 43
	        EMWriteScreen resi_city, 8, 43
	        EMWriteScreen left(resi_county, 2), 9, 66
	        EMWriteScreen left(resi_state, 2), 8, 66
	        EMWriteScreen resi_zip, 9, 43
			If addr_living_sit <> "Blank" AND addr_living_sit <> "Select" AND len(addr_living_sit) >=2 Then EMWriteScreen left(addr_living_sit, 2), 11, 43
			EMWriteScreen left(addr_homeless, 1), 10, 43
			EMWriteScreen left(addr_reservation, 1), 10, 74
			IF addr_reservation = "No" THEN Call clear_line_of_text(11, 74)

	        EMWriteScreen left(addr_verif, 2), 9, 74

			mail_street_full = trim(trim(mail_line_one) & " " & trim(mail_line_two))
			If len(mail_line_one) > 22 or len(mail_line_two) > 22 Then
				'This functionality will write the address word by word so that if it needs to word wrap to the second line, it can move the words to the next line
	            mail_words = split(mail_street_full, " ")
	            write_mail_line_one = ""
	            write_mail_line_two = ""
	            For each word in mail_words
	                If write_mail_line_one = "" Then
	                    write_mail_line_one = word
	                ElseIf len(write_mail_line_one & " " & word) =< 22 Then
	                    write_mail_line_one = write_mail_line_one & " " & word
	                Else
	                    If write_mail_line_two = "" Then
	                        write_mail_line_two = word
	                    Else
	                        write_mail_line_two = write_mail_line_two & " " & word
	                    End If
	                End If
	            Next
	        Else
	            write_mail_line_one = mail_line_one
				write_mail_line_two = mail_line_two
	        End If


			If new_version = False then
				EMWriteScreen write_mail_line_one, 13, 43
				EMWriteScreen write_mail_line_two, 14, 43
				EMWriteScreen mail_city, 15, 43
				If write_mail_line_one <> "" Then EMWriteScreen left(mail_state, 2), 16, 43
				EMWriteScreen mail_zip, 16, 52

		        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
		        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
		        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

		        EMWriteScreen phone_one_left, 17, 45
		        EMWriteScreen phone_one_mid, 17, 51
		        EMWriteScreen phone_one_right, 17, 55
		        If type_one <> "Select One..." Then EMWriteScreen type_one, 17, 67

		        EMWriteScreen phone_two_left, 18, 45
		        EMWriteScreen phone_two_mid, 18, 51
		        EMWriteScreen phone_two_right, 18, 55
		        If type_two <> "Select One..." Then EMWriteScreen type_two, 18, 67

		        EMWriteScreen phone_three_left, 19, 45
		        EMWriteScreen phone_three_mid, 19, 51
		        EMWriteScreen phone_three_right, 19, 55
		        If type_three <> "Select One..." Then EMWriteScreen type_three, 19, 67
			End If

			If new_version = True then
				EMWriteScreen write_mail_line_one, 12, 49
		        EMWriteScreen write_mail_line_two, 13, 49
		        EMWriteScreen mail_city, 14, 49
		        If write_mail_line_one <> "" Then EMWriteScreen left(mail_state, 2), 15, 49
		        EMWriteScreen mail_zip, 15, 58

		        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
		        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
		        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

		        EMWriteScreen phone_one_left, 16, 39
		        EMWriteScreen phone_one_mid, 16, 45
		        EMWriteScreen phone_one_right, 16, 49
		        If type_one <> "Select One..." Then EMWriteScreen type_one, 16, 61
				If phone_one <> "" Then EMWriteScreen text_yn_one, 16, 76

		        EMWriteScreen phone_two_left, 17, 39
		        EMWriteScreen phone_two_mid, 17, 45
		        EMWriteScreen phone_two_right, 17, 49
		        If type_two <> "Select One..." Then EMWriteScreen type_two, 17, 61
				If phone_two <> "" Then EMWriteScreen text_yn_two, 17, 76

		        EMWriteScreen phone_three_left, 18, 39
		        EMWriteScreen phone_three_mid, 18, 45
		        EMWriteScreen phone_three_right, 18, 49
		        If type_three <> "Select One..." Then EMWriteScreen type_three, 18, 61
				If phone_three <> "" Then EMWriteScreen text_yn_three, 18, 76

				EMSetCursor 19, 31
				EMSendKey "<eraseEOF>"
				EMWriteScreen addr_email, 19, 31
			End If

	        save_attempt = 1
	        Do
	            transmit
				' MsgBox "Pause - " & save_attempt
	            EMReadScreen resi_standard_note, 33, 24, 2
	            If resi_standard_note = "RESIDENCE ADDRESS IS STANDARDIZED" Then transmit
	            EMReadScreen mail_standard_note, 31, 24, 2
	            If mail_standard_note = "MAILING ADDRESS IS STANDARDIZED" Then transmit

				EMReadScreen warn_msg, 60, 24, 2
				warn_msg = trim(warn_msg)
				If warn_msg = "ENTER A VALID COMMAND OR PF-KEY" Then Exit Do
	            row = 0
	            col = 0
	            EMSearch "Warning:", row, col

	            If row <> 0 Then
	                Do
	                    EMReadScreen warning_note, 55, row, col
	                    warning_note = trim(warning_note)
	                    warning_message = warning_message & "; " & warning_note
	                Loop until warning_note = ""
	            End If

	            save_attempt = save_attempt + 1
	        Loop until save_attempt = 20
		End IF
    End If
end function

function access_HEST_panel(access_type, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ access_type:indicates if we need to read or write to the HEST panel, only options are 'READ' or 'WRITE' - case does not matter - the function wil ucase it
'~~~~~ all_persons_paying: string - information with reference numbers for all of the members listed on the panel as currently paying
'~~~~~ choice_date: information from the panel of the date the selection was made - formatted as mm/dd/yy - script may read this as a date - could be null
'~~~~~ actual_initial_exp: string - information listed as the initial expense
'~~~~~ retro_heat_ac_yn: string - the y/n selection for retro heat/ac - will be either 'Y', 'N', ''
'~~~~~ retro_heat_ac_units: string - two digit entry of the number of units responsible for this expense for retro heat/ac
'~~~~~ retro_heat_ac_amt: number - the expense amount of retro heat/ac listed on the HEST panel
'~~~~~ retro_electric_yn: string - the y/n selection for retro electric - will be either 'Y', 'N', ''
'~~~~~ retro_electric_units: string - two digit entry of the number of units responsible for this expense for retro electric
'~~~~~ retro_electric_amt: number - the expense amount of retro electric listed on the HEST panel
'~~~~~ retro_phone_yn: string - the y/n selection for retro phone - will be either 'Y', 'N', ''
'~~~~~ retro_phone_units: string - two digit entry of the number of units responsible for this expense for retro phone
'~~~~~ retro_phone_amt: number - the expense amount of retro phone listed on the HEST panel
'~~~~~ prosp_heat_ac_yn: string - the y/n selection for prosp heat/ac - will be either 'Y', 'N', ''
'~~~~~ prosp_heat_ac_units: string - two digit entry of the number of units responsible for this expense for prosp heat/ac
'~~~~~ prosp_heat_ac_amt: number - the expense amount of prosp heat/ac listed on the HEST panel
'~~~~~ prosp_electric_yn: string - the y/n selection for prosp electric - will be either 'Y', 'N', ''
'~~~~~ prosp_electric_units: string - two digit entry of the number of units responsible for this expense for prosp electric
'~~~~~ prosp_electric_amt: number - the expense amount of prosp electric listed on the HEST panel
'~~~~~ prosp_phone_yn: string - the y/n selection for prosp phone - will be either 'Y', 'N', ''
'~~~~~ prosp_phone_units: string - two digit entry of the number of units responsible for this expense for prosp phone
'~~~~~ prosp_phone_amt: number - the expense amount of prosp phone listed on the HEST panel
'~~~~~ total_utility_expense: number - the amount that will be budgeted for SNAP with SUA
'===== Keywords: MAXIS, edit, HEST
    access_type = UCase(access_type)
	Call navigate_to_MAXIS_screen("STAT", "HEST")
    If access_type = "READ" Then
        hest_col = 40
        Do
            EMReadScreen pers_paying, 2, 6, hest_col
            If pers_paying <> "__" Then
                all_persons_paying = all_persons_paying & ", " & pers_paying
            Else
                exit do
            End If
            hest_col = hest_col + 3
        Loop until hest_col = 70
        If left(all_persons_paying, 1) = "," Then all_persons_paying = right(all_persons_paying, len(all_persons_paying) - 2)

        EMReadScreen choice_date, 8, 7, 40
        EMReadScreen actual_initial_exp, 8, 8, 61

        EMReadScreen retro_heat_ac_yn, 1, 13, 34
        EMReadScreen retro_heat_ac_units, 2, 13, 42
        EMReadScreen retro_heat_ac_amt, 6, 13, 49
        EMReadScreen retro_electric_yn, 1, 14, 34
        EMReadScreen retro_electric_units, 2, 14, 42
        EMReadScreen retro_electric_amt, 6, 14, 49
        EMReadScreen retro_phone_yn, 1, 15, 34
        EMReadScreen retro_phone_units, 2, 15, 42
        EMReadScreen retro_phone_amt, 6, 15, 49

        EMReadScreen prosp_heat_ac_yn, 1, 13, 60
        EMReadScreen prosp_heat_ac_units, 2, 13, 68
        EMReadScreen prosp_heat_ac_amt, 6, 13, 75
        EMReadScreen prosp_electric_yn, 1, 14, 60
        EMReadScreen prosp_electric_units, 2, 14, 68
        EMReadScreen prosp_electric_amt, 6, 14, 75
        EMReadScreen prosp_phone_yn, 1, 15, 60
        EMReadScreen prosp_phone_units, 2, 15, 68
        EMReadScreen prosp_phone_amt, 6, 15, 75

        choice_date = replace(choice_date, " ", "/")
        If choice_date = "__/__/__" Then choice_date = ""
        actual_initial_exp = trim(actual_initial_exp)
        actual_initial_exp = replace(actual_initial_exp, "_", "")

        retro_heat_ac_yn = replace(retro_heat_ac_yn, "_", "")
        retro_heat_ac_units = replace(retro_heat_ac_units, "_", "")
        retro_heat_ac_amt = trim(retro_heat_ac_amt)
		If retro_heat_ac_amt = "" Then retro_heat_ac_amt = 0
		retro_heat_ac_amt = retro_heat_ac_amt * 1
        retro_electric_yn = replace(retro_electric_yn, "_", "")
        retro_electric_units = replace(retro_electric_units, "_", "")
        retro_electric_amt = trim(retro_electric_amt)
		If retro_electric_amt = "" Then retro_electric_amt = 0
		retro_electric_amt = retro_electric_amt * 1
        retro_phone_yn = replace(retro_phone_yn, "_", "")
        retro_phone_units = replace(retro_phone_units, "_", "")
        retro_phone_amt = trim(retro_phone_amt)
		If retro_phone_amt = "" Then retro_phone_amt = 0
		retro_phone_amt = retro_phone_amt * 1

        prosp_heat_ac_yn = replace(prosp_heat_ac_yn, "_", "")
        prosp_heat_ac_units = replace(prosp_heat_ac_units, "_", "")
        prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
        If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
		prosp_heat_ac_amt = prosp_heat_ac_amt * 1
        prosp_electric_yn = replace(prosp_electric_yn, "_", "")
        prosp_electric_units = replace(prosp_electric_units, "_", "")
        prosp_electric_amt = trim(prosp_electric_amt)
        If prosp_electric_amt = "" Then prosp_electric_amt = 0
		prosp_electric_amt = prosp_electric_amt * 1
        prosp_phone_yn = replace(prosp_phone_yn, "_", "")
        prosp_phone_units = replace(prosp_phone_units, "_", "")
        prosp_phone_amt = trim(prosp_phone_amt)
        If prosp_phone_amt = "" Then prosp_phone_amt = 0
		prosp_phone_amt = prosp_phone_amt * 1

        total_utility_expense = 0
        If prosp_heat_ac_yn = "Y" Then
            total_utility_expense =  prosp_heat_ac_amt
        ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt + prosp_phone_amt
        ElseIf prosp_electric_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt
        Elseif prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_phone_amt
        End If

    End If

	If access_type = "WRITE" Then
		EMReadScreen hest_version, 1, 2, 73
		If hest_version = "1" Then PF9
		If hest_version = "0" Then
			EMWriteScreen "NN", 20, 79
			transmit
		End If

		all_persons_paying = trim(all_persons_paying)
		If all_persons_paying <> "" Then
			If InStr(all_persons_paying, ",") = 0 Then
				persons_array = array(all_persons_paying)
			Else
				persons_array = split(all_persons_paying, ",")
			End If

			hest_col = 40
			for each pers_paying in persons_array
				EMWriteScreen pers_paying, 6, hest_col
				hest_col = hest_col + 3
			Next

			If IsDate(choice_date) = True Then Call create_mainframe_friendly_date(choice_date, 7, 40, "YY")
	        EMWriteScreen actual_initial_exp, 8, 61

			EMWriteScreen retro_heat_ac_yn, 13, 34
	        If retro_heat_ac_yn = "Y" Then EMWriteScreen "01", 13, 42
	        EMWriteScreen retro_electric_yn, 14, 34
	        If retro_electric_yn = "Y" Then EMWriteScreen "01", 14, 42
	        EMWriteScreen retro_phone_yn, 15, 34
	        If retro_phone_yn = "Y" Then EMWriteScreen "01", 15, 42

	        EMWriteScreen prosp_heat_ac_yn, 13, 60
	        If prosp_heat_ac_yn = "Y" Then EMWriteScreen "01", 13, 68
	        EMWriteScreen prosp_electric_yn, 14, 60
	        If prosp_electric_yn = "Y" Then EMWriteScreen "01", 14, 68
	        EMWriteScreen prosp_phone_yn, 15, 60
	        If prosp_phone_yn = "Y" Then EMWriteScreen "01", 15, 68

			transmit

			EMReadScreen retro_heat_ac_amt, 6, 13, 49
			EMReadScreen retro_electric_amt, 6, 14, 49
			EMReadScreen retro_phone_amt, 6, 15, 49

			EMReadScreen prosp_heat_ac_amt, 6, 13, 75
			EMReadScreen prosp_electric_amt, 6, 14, 75
			EMReadScreen prosp_phone_amt, 6, 15, 75

			retro_heat_ac_amt = trim(retro_heat_ac_amt)
			If retro_heat_ac_amt = "" Then retro_heat_ac_amt = 0
			retro_heat_ac_amt = retro_heat_ac_amt * 1
			retro_electric_amt = trim(retro_electric_amt)
			If retro_electric_amt = "" Then retro_electric_amt = 0
			retro_electric_amt = retro_electric_amt * 1
			retro_phone_amt = trim(retro_phone_amt)
			If retro_phone_amt = "" Then retro_phone_amt = 0
			retro_phone_amt = retro_phone_amt * 1

			prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
			If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
			prosp_heat_ac_amt = prosp_heat_ac_amt * 1
			prosp_electric_amt = trim(prosp_electric_amt)
			If prosp_electric_amt = "" Then prosp_electric_amt = 0
			prosp_electric_amt = prosp_electric_amt * 1
			prosp_phone_amt = trim(prosp_phone_amt)
			If prosp_phone_amt = "" Then prosp_phone_amt = 0
			prosp_phone_amt = prosp_phone_amt * 1
		End If
	End If
end function

function access_SHEL_panel(access_type, shel_ref_number, hud_sub_yn, shared_yn, paid_to, rent_retro_amt, rent_retro_verif, rent_prosp_amt, rent_prosp_verif, lot_rent_retro_amt, lot_rent_retro_verif, lot_rent_prosp_amt, lot_rent_prosp_verif, mortgage_retro_amt, mortgage_retro_verif, mortgage_prosp_amt, mortgage_prosp_verif, insurance_retro_amt, insurance_retro_verif, insurance_prosp_amt, insurance_prosp_verif, tax_retro_amt, tax_retro_verif, tax_prosp_amt, tax_prosp_verif, room_retro_amt, room_retro_verif, room_prosp_amt, room_prosp_verif, garage_retro_amt, garage_retro_verif, garage_prosp_amt, garage_prosp_verif, subsidy_retro_amt, subsidy_retro_verif, subsidy_prosp_amt, subsidy_prosp_verif, original_information)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ access_type:indicates if we need to read or write to the HEST panel, only options are 'READ' or 'WRITE' - case does not matter - the function wil ucase it
'~~~~~ shel_ref_number: String - the 2 digit reference number of the member the panels is for
'~~~~~ hud_sub_yn: string - the y/n information about if the rent is HUD Subsidized - options: 'Y', 'N', ''
'~~~~~ shared_yn: string - the y/n information about if the expense is shared - options: 'Y', 'N', ''
'~~~~~ paid_to: string - information about who the expense is paid to
'~~~~~ rent_retro_amt: number - amount entered of the retro rent
'~~~~~ rent_retro_verif: string - verif entered for the retro rent - formatted as XX - VERIF NAME
'~~~~~ rent_prosp_amt: number - amount entered of the prosp rent
'~~~~~ rent_prosp_verif: string - verif entered for the prosp rent
'~~~~~ lot_rent_retro_amt: number - amount entered of the retro lot rent
'~~~~~ lot_rent_retro_verif: string - verif entered for the retro lot rent
'~~~~~ lot_rent_prosp_amt: number - amount entered of the prosp lot rent
'~~~~~ lot_rent_prosp_verif: string - verif entered for the prosp lot rent
'~~~~~ mortgage_retro_amt: number - amount entered of the retro mortgage
'~~~~~ mortgage_retro_verif: string - verif entered for the retro mortgage
'~~~~~ mortgage_prosp_amt: number - amount entered of the prosp mortgage
'~~~~~ mortgage_prosp_verif: string - verif entered for the prosp mortgage
'~~~~~ insurance_retro_amt: number - amount entered of the retro insurance
'~~~~~ insurance_retro_verif: string - verif entered for the retro insurance
'~~~~~ insurance_prosp_amt: number - amount entered of the prosp insurance
'~~~~~ insurance_prosp_verif: string - verif entered for the prosp insurance
'~~~~~ tax_retro_amt: number - amount entered of the retro tax
'~~~~~ tax_retro_verif: string - verif entered for the retro tax
'~~~~~ tax_prosp_amt: number - amount entered of the prosp tax
'~~~~~ tax_prosp_verif: string - verif entered for the prosp tax
'~~~~~ room_retro_amt: number - amount entered of the retro room
'~~~~~ room_retro_verif: string - verif entered for the retro room
'~~~~~ room_prosp_amt: number - amount entered of the prosp room
'~~~~~ room_prosp_verif: string - verif entered for the prosp room
'~~~~~ garage_retro_amt: number - amount entered of the retro garage
'~~~~~ garage_retro_verif: string - verif entered for the retro garage
'~~~~~ garage_prosp_amt: number - amount entered of the prosp garage
'~~~~~ garage_prosp_verif: string - verif entered for the prsop garage
'~~~~~ subsidy_retro_amt: number - amount entered of the retro subsidy
'~~~~~ subsidy_retro_verif: string - verif entered for the retro subsidy
'~~~~~ subsidy_prosp_amt: number - amount entered of the prosp subsidy
'~~~~~ subsidy_prosp_verif: string - verif entered for the prosp subsidy
'~~~~~ original_information: string that saves the original detail known on the panel - this is used to see if there is an update to any information
'===== Keywords: MAXIS, edit, update, SHEL
	Call navigate_to_MAXIS_screen("STAT", "SHEL")
	EMWriteScreen shel_ref_number, 20, 76
	transmit

	access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen hud_sub_yn,            1, 6, 46
        EMReadScreen shared_yn,             1, 6, 64
        EMReadScreen paid_to,               25, 7, 50

		if hud_sub_yn = "_" Then hud_sub_yn = ""
		if shared_yn = "_" Then shared_yn = ""
        paid_to = replace(paid_to, "_", "")

        EMReadScreen rent_retro_amt,        8, 11, 37
        EMReadScreen rent_retro_verif,      2, 11, 48
        EMReadScreen rent_prosp_amt,        8, 11, 56
        EMReadScreen rent_prosp_verif,      2, 11, 67

        rent_retro_amt = replace(rent_retro_amt, "_", "")
        rent_retro_amt = trim(rent_retro_amt)
		If rent_retro_amt = "" Then rent_retro_amt = 0
		rent_retro_amt = rent_retro_amt * 1
        If rent_retro_verif = "SF" Then rent_retro_verif = "SF - Shelter Form"
        If rent_retro_verif = "LE" Then rent_retro_verif = "LE - Lease"
        If rent_retro_verif = "RE" Then rent_retro_verif = "RE - Rent Receipt"
        If rent_retro_verif = "OT" Then rent_retro_verif = "OT - Other Doc"
        If rent_retro_verif = "NC" Then rent_retro_verif = "NC - Chg, Neg Impact"
        If rent_retro_verif = "PC" Then rent_retro_verif = "PC - Chg, Pos Impact"
        If rent_retro_verif = "NO" Then rent_retro_verif = "NO - No Verif"
		If rent_retro_verif = "?_" Then rent_retro_verif = "? - Delayed Verif"
        If rent_retro_verif = "__" Then rent_retro_verif = ""
        rent_prosp_amt = replace(rent_prosp_amt, "_", "")
        rent_prosp_amt = trim(rent_prosp_amt)
		If rent_prosp_amt = "" Then rent_prosp_amt = 0
		rent_prosp_amt = rent_prosp_amt * 1
        If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
        If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
        If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
        If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Doc"
        If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg, Neg Impact"
        If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg, Pos Impact"
        If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Verif"
		If rent_prosp_verif = "?_" Then rent_prosp_verif = "? - Delayed Verif"
        If rent_prosp_verif = "__" Then rent_prosp_verif = ""

        EMReadScreen lot_rent_retro_amt,    8, 12, 37
        EMReadScreen lot_rent_retro_verif,  2, 12, 48
        EMReadScreen lot_rent_prosp_amt,    8, 12, 56
        EMReadScreen lot_rent_prosp_verif,  2, 12, 67

        lot_rent_retro_amt = replace(lot_rent_retro_amt, "_", "")
        lot_rent_retro_amt = trim(lot_rent_retro_amt)
		If lot_rent_retro_amt = "" Then lot_rent_retro_amt = 0
		lot_rent_retro_amt = lot_rent_retro_amt * 1
        If lot_rent_retro_verif = "LE" Then lot_rent_retro_verif = "LE - Lease"
        If lot_rent_retro_verif = "RE" Then lot_rent_retro_verif = "RE - Rent Receipt"
        If lot_rent_retro_verif = "BI" Then lot_rent_retro_verif = "BI - Billing Stmt"
        If lot_rent_retro_verif = "OT" Then lot_rent_retro_verif = "OT - Other Doc"
        If lot_rent_retro_verif = "NC" Then lot_rent_retro_verif = "NC - Chg, Neg Impact"
        If lot_rent_retro_verif = "PC" Then lot_rent_retro_verif = "PC - Chg, Pos Impact"
        If lot_rent_retro_verif = "NO" Then lot_rent_retro_verif = "NO - No Verif"
		If lot_rent_retro_verif = "?_" Then lot_rent_retro_verif = "? - Delayed Verif"
        If lot_rent_retro_verif = "__" Then lot_rent_retro_verif = ""
        lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
        lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
		If lot_rent_prosp_amt = "" Then lot_rent_prosp_amt = 0
		lot_rent_prosp_amt = lot_rent_prosp_amt * 1
        If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
        If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
        If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
        If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Doc"
        If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg, Neg Impact"
        If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg, Pos Impact"
        If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Verif"
		If lot_rent_prosp_verif = "?_" Then lot_rent_prosp_verif = "? - Delayed Verif"
        If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""

        EMReadScreen mortgage_retro_amt,    8, 13, 37
        EMReadScreen mortgage_retro_verif,  2, 13, 48
        EMReadScreen mortgage_prosp_amt,    8, 13, 56
        EMReadScreen mortgage_prosp_verif,  2, 13, 67

        mortgage_retro_amt = replace(mortgage_retro_amt, "_", "")
        mortgage_retro_amt = trim(mortgage_retro_amt)
		If mortgage_retro_amt = "" Then mortgage_retro_amt = 0
		mortgage_retro_amt = mortgage_retro_amt * 1
        If mortgage_retro_verif = "MO" Then mortgage_retro_verif = "MO - Mortgage Pmt Book"
        If mortgage_retro_verif = "CD" Then mortgage_retro_verif = "CD - Ctrct fro Deed"
        If mortgage_retro_verif = "OT" Then mortgage_retro_verif = "OT - Other Doc"
        If mortgage_retro_verif = "NC" Then mortgage_retro_verif = "NC - Chg, Neg Impact"
        If mortgage_retro_verif = "PC" Then mortgage_retro_verif = "PC - Chg, Pos Impact"
        If mortgage_retro_verif = "NO" Then mortgage_retro_verif = "NO - No Verif"
		If mortgage_retro_verif = "?_" Then mortgage_retro_verif = "? - Delayed Verif"
        If mortgage_retro_verif = "__" Then mortgage_retro_verif = ""
        mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
        mortgage_prosp_amt = trim(mortgage_prosp_amt)
		If mortgage_prosp_amt = "" Then mortgage_prosp_amt = 0
		mortgage_prosp_amt = mortgage_prosp_amt * 1
        If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
        If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
        If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Doc"
        If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg, Neg Impact"
        If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg, Pos Impact"
        If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Verif"
		If mortgage_prosp_verif = "?_" Then mortgage_prosp_verif = "? - Delayed Verif"
        If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""

        EMReadScreen insurance_retro_amt,   8, 14, 37
        EMReadScreen insurance_retro_verif, 2, 14, 48
        EMReadScreen insurance_prosp_amt,   8, 14, 56
        EMReadScreen insurance_prosp_verif, 2, 14, 67

        insurance_retro_amt = replace(insurance_retro_amt, "_", "")
        insurance_retro_amt = trim(insurance_retro_amt)
		If insurance_retro_amt = "" Then insurance_retro_amt = 0
		insurance_retro_amt = insurance_retro_amt * 1
        If insurance_retro_verif = "BI" Then insurance_retro_verif = "BI - Billing Stmt"
        If insurance_retro_verif = "OT" Then insurance_retro_verif = "OT - Other Doc"
        If insurance_retro_verif = "NC" Then insurance_retro_verif = "NC - Chg, Neg Impact"
        If insurance_retro_verif = "PC" Then insurance_retro_verif = "PC - Chg, Pos Impact"
        If insurance_retro_verif = "NO" Then insurance_retro_verif = "NO - No Verif"
		If insurance_retro_verif = "?_" Then insurance_retro_verif = "? - Delayed Verif"
        If insurance_retro_verif = "__" Then insurance_retro_verif = ""
        insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
        insurance_prosp_amt = trim(insurance_prosp_amt)
		If insurance_prosp_amt = "" Then insurance_prosp_amt = 0
		insurance_prosp_amt = insurance_prosp_amt * 1
        If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
        If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Doc"
        If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg, Neg Impact"
        If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg, Pos Impact"
        If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Verif"
		If insurance_prosp_verif = "?_" Then insurance_prosp_verif = "? - Delayed Verif"
        If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""

        EMReadScreen tax_retro_amt,         8, 15, 37
        EMReadScreen tax_retro_verif,       2, 15, 48
        EMReadScreen tax_prosp_amt,         8, 15, 56
        EMReadScreen tax_prosp_verif,       2, 15, 67

        tax_retro_amt = replace(tax_retro_amt, "_", "")
        tax_retro_amt = trim(tax_retro_amt)
		If tax_retro_amt = "" Then tax_retro_amt = 0
		tax_retro_amt = tax_retro_amt * 1
        If tax_retro_verif = "TX" Then tax_retro_verif = "TX - Prop Tax Stmt"
        If tax_retro_verif = "OT" Then tax_retro_verif = "OT - Other Doc"
        If tax_retro_verif = "NC" Then tax_retro_verif = "NC - Chg, Neg Impact"
        If tax_retro_verif = "PC" Then tax_retro_verif = "PC - Chg, Pos Impact"
        If tax_retro_verif = "NO" Then tax_retro_verif = "NO - No Verif"
		If tax_retro_verif = "?_" Then tax_retro_verif = "? - Delayed Verif"
        If tax_retro_verif = "__" Then tax_retro_verif = ""
        tax_prosp_amt = replace(tax_prosp_amt, "_", "")
        tax_prosp_amt = trim(tax_prosp_amt)
		If tax_prosp_amt = "" Then tax_prosp_amt = 0
		tax_prosp_amt = tax_prosp_amt * 1
        If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
        If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Doc"
        If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg, Neg Impact"
        If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg, Pos Impact"
        If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Verif"
		If tax_prosp_verif = "?_" Then tax_prosp_verif = "? - Delayed Verif"
        If tax_prosp_verif = "__" Then tax_prosp_verif = ""

        EMReadScreen room_retro_amt,        8, 16, 37
        EMReadScreen room_retro_verif,      2, 16, 48
        EMReadScreen room_prosp_amt,        8, 16, 56
        EMReadScreen room_prosp_verif,      2, 16, 67

        room_retro_amt = replace(room_retro_amt, "_", "")
        room_retro_amt = trim(room_retro_amt)
		If room_retro_amt = "" Then room_retro_amt = 0
		room_retro_amt = room_retro_amt * 1
        If room_retro_verif = "SF" Then room_retro_verif = "SF - Shelter Form"
        If room_retro_verif = "LE" Then room_retro_verif = "LE - Lease"
        If room_retro_verif = "RE" Then room_retro_verif = "RE - Rent Receipt"
        If room_retro_verif = "OT" Then room_retro_verif = "OT - Other Doc"
        If room_retro_verif = "NC" Then room_retro_verif = "NC - Chg, Neg Impact"
        If room_retro_verif = "PC" Then room_retro_verif = "PC - Chg, Pos Impact"
        If room_retro_verif = "NO" Then room_retro_verif = "NO - No Verif"
		If room_retro_verif = "?_" Then room_retro_verif = "? - Delayed Verif"
        If room_retro_verif = "__" Then room_retro_verif = ""
        room_prosp_amt = replace(room_prosp_amt, "_", "")
        room_prosp_amt = trim(room_prosp_amt)
		If room_prosp_amt = "" Then room_prosp_amt = 0
		room_prosp_amt = room_prosp_amt * 1
        If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
        If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
        If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
        If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Doc"
        If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg, Neg Impact"
        If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg, Pos Impact"
        If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Verif"
		If room_prosp_verif = "?_" Then room_prosp_verif = "? - Delayed Verif"
        If room_prosp_verif = "__" Then room_prosp_verif = ""

        EMReadScreen garage_retro_amt,      8, 17, 37
        EMReadScreen garage_retro_verif,    2, 17, 48
        EMReadScreen garage_prosp_amt,      8, 17, 56
        EMReadScreen garage_prosp_verif,    2, 17, 67

        garage_retro_amt = replace(garage_retro_amt, "_", "")
        garage_retro_amt = trim(garage_retro_amt)
		If garage_retro_amt = "" Then garage_retro_amt = 0
		garage_retro_amt = garage_retro_amt * 1
        If garage_retro_verif = "SF" Then garage_retro_verif = "SF - Shelter Form"
        If garage_retro_verif = "LE" Then garage_retro_verif = "LE - Lease"
        If garage_retro_verif = "RE" Then garage_retro_verif = "RE - Rent Receipt"
        If garage_retro_verif = "OT" Then garage_retro_verif = "OT - Other Doc"
        If garage_retro_verif = "NC" Then garage_retro_verif = "NC - Chg, Neg Impact"
        If garage_retro_verif = "PC" Then garage_retro_verif = "PC - Chg, Pos Impact"
        If garage_retro_verif = "NO" Then garage_retro_verif = "NO - No Verif"
		If garage_retro_verif = "?_" Then garage_retro_verif = "? - Delayed Verif"
        If garage_retro_verif = "__" Then garage_retro_verif = ""
        garage_prosp_amt = replace(garage_prosp_amt, "_", "")
        garage_prosp_amt = trim(garage_prosp_amt)
		If garage_prosp_amt = "" Then garage_prosp_amt = 0
		garage_prosp_amt = garage_prosp_amt * 1
        If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
        If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
        If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
        If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Doc"
        If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg, Neg Impact"
        If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg, Pos Impact"
        If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Verif"
		If garage_prosp_verif = "?_" Then garage_prosp_verif = "? - Delayed Verif"
        If garage_prosp_verif = "__" Then garage_prosp_verif = ""

        EMReadScreen subsidy_retro_amt,     8, 18, 37
        EMReadScreen subsidy_retro_verif,   2, 18, 48
        EMReadScreen subsidy_prosp_amt,     8, 18, 56
        EMReadScreen subsidy_prosp_verif,   2, 18, 67

        subsidy_retro_amt = replace(subsidy_retro_amt, "_", "")
        subsidy_retro_amt = trim(subsidy_retro_amt)
		If subsidy_retro_amt = "" Then subsidy_retro_amt = 0
		subsidy_retro_amt = subsidy_retro_amt * 1
        If subsidy_retro_verif = "SF" Then subsidy_retro_verif = "SF - Shelter Form"
        If subsidy_retro_verif = "LE" Then subsidy_retro_verif = "LE - Lease"
        If subsidy_retro_verif = "OT" Then subsidy_retro_verif = "OT - Other Doc"
        If subsidy_retro_verif = "NO" Then subsidy_retro_verif = "NO - No Verif"
		If subsidy_retro_verif = "?_" Then subsidy_retro_verif = "? - Delayed Verif"
        If subsidy_retro_verif = "__" Then subsidy_retro_verif = ""
        subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
        subsidy_prosp_amt = trim(subsidy_prosp_amt)
		If subsidy_prosp_amt = "" Then subsidy_prosp_amt = 0
		subsidy_prosp_amt = subsidy_prosp_amt * 1
        If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
        If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
        If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Doc"
        If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Verif"
		If subsidy_prosp_verif = "?_" Then subsidy_prosp_verif = "? - Delayed Verif"
        If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""

		original_information = hud_sub_yn&"|"&shared_yn&"|"&paid_to&"|"&rent_retro_amt&"|"&rent_retro_verif&"|"&rent_prosp_amt&"|"&rent_prosp_verif&"|"&lot_rent_retro_amt&"|"&lot_rent_retro_verif&"|"&lot_rent_prosp_amt&"|"&_
							   lot_rent_prosp_verif&"|"&mortgage_retro_amt&"|"&mortgage_retro_verif&"|"&mortgage_prosp_amt&"|"&mortgage_prosp_verif&"|"&insurance_retro_amt&"|"&insurance_retro_verif&"|"&insurance_prosp_amt&"|"&_
							   insurance_prosp_verif&"|"&tax_retro_amt&"|"&tax_retro_verif&"|"&tax_prosp_amt&"|"&tax_prosp_verif&"|"&room_retro_amt&"|"&room_retro_verif&"|"&room_prosp_amt&"|"&room_prosp_verif&"|"&garage_retro_amt&"|"&_
							   garage_retro_verif&"|"&garage_prosp_amt&"|"&garage_prosp_verif&"|"&subsidy_retro_amt&"|"&subsidy_retro_verif&"|"&subsidy_prosp_amt&"|"&subsidy_prosp_verif
    End If

	If access_type = "WRITE" Then
		hud_sub_yn = trim(hud_sub_yn)
		shared_yn = trim(shared_yn)
		paid_to = trim(paid_to)
		rent_retro_amt = trim(rent_retro_amt)
		rent_retro_verif = trim(rent_retro_verif)
		rent_prosp_amt = trim(rent_prosp_amt)
		rent_prosp_verif = trim(rent_prosp_verif)
		lot_rent_retro_amt = trim(lot_rent_retro_amt)
		lot_rent_retro_verif = trim(lot_rent_retro_verif)
		lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
		lot_rent_prosp_verif = trim(lot_rent_prosp_verif)
		mortgage_retro_amt = trim(mortgage_retro_amt)
		mortgage_retro_verif = trim(mortgage_retro_verif)
		mortgage_prosp_amt = trim(mortgage_prosp_amt)
		mortgage_prosp_verif = trim(mortgage_prosp_verif)
		insurance_retro_amt = trim(insurance_retro_amt)
		insurance_retro_verif = trim(insurance_retro_verif)
		insurance_prosp_amt = trim(insurance_prosp_amt)
		insurance_prosp_verif = trim(insurance_prosp_verif)
		tax_retro_amt = trim(tax_retro_amt)
		tax_retro_verif = trim(tax_retro_verif)
		tax_prosp_amt = trim(tax_prosp_amt)
		tax_prosp_verif = trim(tax_prosp_verif)
		room_retro_amt = trim(room_retro_amt)
		room_retro_verif = trim(room_retro_verif)
		room_prosp_amt = trim(room_prosp_amt)
		room_prosp_verif = trim(room_prosp_verif)
		garage_retro_amt = trim(garage_retro_amt)
		garage_retro_verif = trim(garage_retro_verif)
		garage_prosp_amt = trim(garage_prosp_amt)
		garage_prosp_verif = trim(garage_prosp_verif)
		subsidy_retro_amt = trim(subsidy_retro_amt)
		subsidy_retro_verif = trim(subsidy_retro_verif)
		subsidy_prosp_amt = trim(subsidy_prosp_amt)
		subsidy_prosp_verif = trim(subsidy_prosp_verif)

		current_shel_details = hud_sub_yn&"|"&shared_yn&"|"&paid_to&"|"&rent_retro_amt&"|"&rent_retro_verif&"|"&rent_prosp_amt&"|"&rent_prosp_verif&"|"&lot_rent_retro_amt&"|"&lot_rent_retro_verif&"|"&lot_rent_prosp_amt&"|"&_
							   lot_rent_prosp_verif&"|"&mortgage_retro_amt&"|"&mortgage_retro_verif&"|"&mortgage_prosp_amt&"|"&mortgage_prosp_verif&"|"&insurance_retro_amt&"|"&insurance_retro_verif&"|"&insurance_prosp_amt&"|"&_
							   insurance_prosp_verif&"|"&tax_retro_amt&"|"&tax_retro_verif&"|"&tax_prosp_amt&"|"&tax_prosp_verif&"|"&room_retro_amt&"|"&room_retro_verif&"|"&room_prosp_amt&"|"&room_prosp_verif&"|"&garage_retro_amt&"|"&_
							   garage_retro_verif&"|"&garage_prosp_amt&"|"&garage_prosp_verif&"|"&subsidy_retro_amt&"|"&subsidy_retro_verif&"|"&subsidy_prosp_amt&"|"&subsidy_prosp_verif



		If current_shel_details <> original_information Then
			EMReadScreen shel_version, 1, 2, 73
			If shel_version = "1" Then PF9
			If shel_version = "0" Then
				EMWriteScreen "NN", 20, 79
				transmit
			End If

			rent_retro_verif = 		rent_retro_verif & "  "
			rent_prosp_verif = 		rent_prosp_verif & "  "
			lot_rent_retro_verif = 	lot_rent_retro_verif & "  "
			lot_rent_prosp_verif = 	lot_rent_prosp_verif & "  "
			mortgage_retro_verif = 	mortgage_retro_verif & "  "
			mortgage_prosp_verif = 	mortgage_prosp_verif & "  "
			insurance_retro_verif = insurance_retro_verif & "  "
			insurance_prosp_verif = insurance_prosp_verif & "  "
			tax_retro_verif = 		tax_retro_verif & "  "
			tax_prosp_verif = 		tax_prosp_verif & "  "
			room_retro_verif = 		room_retro_verif & "  "
			room_prosp_verif = 		room_prosp_verif & "  "
			garage_retro_verif = 	garage_retro_verif & "  "
			garage_prosp_verif = 	garage_prosp_verif & "  "
			subsidy_retro_verif = 	subsidy_retro_verif & "  "
			subsidy_prosp_verif = 	subsidy_prosp_verif & "  "

			If hud_sub_yn = "" Then
				EMSetCursor 6, 46
				EMSendKey "<eraseEOF>"
			Else
				EMWriteScreen hud_sub_yn, 6, 46
			End If
			If shared_yn = "" Then
				EMSetCursor 6, 64
				EMSendKey "<eraseEOF>"
			Else
	        	EMWriteScreen shared_yn, 6, 64
			End If
			If paid_to = "" Then
				EMSetCursor 7, 50
				EMSendKey "<eraseEOF>"
			Else
	        	EMWriteScreen paid_to, 7, 50
			End If

			EMWriteScreen right("        " & rent_retro_amt, 8), 		11, 37
	        EMWriteScreen left(rent_retro_verif, 2),      				11, 48
	        EMWriteScreen right("        " & rent_prosp_amt, 8),    	11, 56
	        EMWriteScreen left(rent_prosp_verif, 2),      				11, 67

			EMWriteScreen right("        " & lot_rent_retro_amt, 8),    12, 37
	        EMWriteScreen left(lot_rent_retro_verif, 2),  				12, 48
	        EMWriteScreen right("        " & lot_rent_prosp_amt, 8),    12, 56
	        EMWriteScreen left(lot_rent_prosp_verif, 2),  				12, 67

			EMWriteScreen right("        " & mortgage_retro_amt, 8),    13, 37
	        EMWriteScreen left(mortgage_retro_verif, 2),  				13, 48
	        EMWriteScreen right("        " & mortgage_prosp_amt, 8),    13, 56
	        EMWriteScreen left(mortgage_prosp_verif, 2),  				13, 67

			EMWriteScreen right("        " & insurance_retro_amt, 8),   14, 37
	        EMWriteScreen left(insurance_retro_verif, 2), 				14, 48
	        EMWriteScreen right("        " & insurance_prosp_amt, 8),   14, 56
	        EMWriteScreen left(insurance_prosp_verif, 2), 				14, 67

			EMWriteScreen right("        " & tax_retro_amt, 8),         15, 37
	        EMWriteScreen left(tax_retro_verif, 2),       				15, 48
	        EMWriteScreen right("        " & tax_prosp_amt, 8),         15, 56
	        EMWriteScreen left(tax_prosp_verif, 2),       				15, 67

			EMWriteScreen right("        " & room_retro_amt, 8),        16, 37
	        EMWriteScreen left(room_retro_verif, 2),      				16, 48
	        EMWriteScreen right("        " & room_prosp_amt, 8),        16, 56
	        EMWriteScreen left(room_prosp_verif, 2),      				16, 67

			EMWriteScreen right("        " & garage_retro_amt, 8),      17, 37
			EMWriteScreen left(garage_retro_verif, 2),    				17, 48
			EMWriteScreen right("        " & garage_prosp_amt, 8),      17, 56
			EMWriteScreen left(garage_prosp_verif, 2),    				17, 67

			EMWriteScreen right("        " & subsidy_retro_amt, 8),     18, 37
			EMWriteScreen left(subsidy_retro_verif, 2),   				18, 48
			EMWriteScreen right("        " & subsidy_prosp_amt, 8),     18, 56
			EMWriteScreen left(subsidy_prosp_verif, 2),   				18, 67

		End If
	End If
end function

function add_ACCI_to_variable(ACCI_variable)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ ACCI_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, ACCI
  EMReadScreen ACCI_date, 8, 6, 73
  ACCI_date = replace(ACCI_date, " ", "/")
  If datediff("yyyy", ACCI_date, now) < 5 then
    EMReadScreen ACCI_type, 2, 6, 47
    If ACCI_type = "01" then ACCI_type = "Auto"
    If ACCI_type = "02" then ACCI_type = "Workers Comp"
    If ACCI_type = "03" then ACCI_type = "Homeowners"
    If ACCI_type = "04" then ACCI_type = "No Fault"
    If ACCI_type = "05" then ACCI_type = "Other Tort"
    If ACCI_type = "06" then ACCI_type = "Product Liab"
    If ACCI_type = "07" then ACCI_type = "Med Malprac"
    If ACCI_type = "08" then ACCI_type = "Legal Malprac"
    If ACCI_type = "09" then ACCI_type = "Diving Tort"
    If ACCI_type = "10" then ACCI_type = "Motorcycle"
    If ACCI_type = "11" then ACCI_type = "MTC or Other Bus Tort"
    If ACCI_type = "12" then ACCI_type = "Pedestrian"
    If ACCI_type = "13" then ACCI_type = "Other"
    ACCI_variable = ACCI_variable & ACCI_type & " on " & ACCI_date & ".; "
  End if
end function

function add_ACCT_to_variable(ACCT_variable)
'--- This function adds STAT/ACCT data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ ACCT_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, ACCT
  EMReadScreen ACCT_amt, 8, 10, 46
  ACCT_amt = trim(ACCT_amt)
  ACCT_amt = "$" & ACCT_amt
  EMReadScreen ACCT_type, 2, 6, 44
  EMReadScreen ACCT_location, 20, 8, 44
  ACCT_location = replace(ACCT_location, "_", "")
  ACCT_location = split(ACCT_location)
  For each ACCT_part in ACCT_location
    If ACCT_part <> "" then
      first_letter = ucase(left(ACCT_part, 1))
      other_letters = LCase(right(ACCT_part, len(ACCT_part) -1))
      If len(ACCT_part) > 3 then
        new_ACCT_location = new_ACCT_location & first_letter & other_letters & " "
      Else
        new_ACCT_location = new_ACCT_location & ACCT_part & " "
      End if
    End if
  Next
  EMReadScreen ACCT_ver, 1, 10, 63
  If ACCT_ver = "N" then
    ACCT_ver = ", no proof provided"
  Else
    ACCT_ver = ""
  End if
  ACCT_variable = ACCT_variable & ACCT_type & " at " & new_ACCT_location & "(" & ACCT_amt & ")" & ACCT_ver & ".; "
  new_ACCT_location = ""
end function

function add_BUSI_to_variable(variable_name_for_BUSI)
'--- This function adds STAT/BUSI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ BUSI_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, BUSI

	'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
	EMReadScreen BUSI_footer_month, 5, 20, 55
	BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")

	'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
	If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		EMWriteScreen "X", 7, 26
		EMSendKey "<enter>"
		EMWaitReady 0, 0
		If cash_check = 1 then
			EMReadScreen BUSI_ver, 1, 9, 73
		ElseIf HC_check = 1 then
			EMReadScreen BUSI_ver, 1, 12, 73
			If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
		ElseIf SNAP_check = 1 then
			EMReadScreen BUSI_ver, 1, 11, 73
		End if
		EMSendKey "<PF3>"
		EMWaitReady 0, 0
		If SNAP_check = 1 then
			EMReadScreen BUSI_amt, 8, 11, 68
			BUSI_amt = trim(BUSI_amt)
		ElseIf cash_check = 1 then
			EMReadScreen BUSI_amt, 8, 9, 54
			BUSI_amt = trim(BUSI_amt)
		ElseIf HC_check = 1 then
			EMWriteScreen "X", 17, 29
			EMSendKey "<enter>"
			EMWaitReady 0, 0
			EMReadScreen BUSI_amt, 8, 15, 54
			If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
			BUSI_amt = trim(BUSI_amt)
			EMSendKey "<PF3>"
			EMWaitReady 0, 0
		End if
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI"
		EMReadScreen BUSI_income_end_date, 8, 5, 71
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
		If IsDate(BUSI_income_end_date) = True then
			variable_name_for_BUSI = variable_name_for_BUSI & " (ended " & BUSI_income_end_date & ")"
		Else
			If BUSI_amt <> "" then variable_name_for_BUSI = variable_name_for_BUSI & ", ($" & BUSI_amt & "/monthly)"
		End if
		If BUSI_ver = "N" or BUSI_ver = "?" then
			variable_name_for_BUSI = variable_name_for_BUSI & ", no proof provided.; "
		Else
			variable_name_for_BUSI = variable_name_for_BUSI & ".; "
		End if
	Else		'------------This was updated 01/07/2015.
		'Checks the current footer month. If this is the future, it will know later on to read the HC pop-up
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		If datediff("d", date, BUSI_footer_month) > 0 then
			pull_future_HC = TRUE
		Else
			pull_future_HC = FALSE
		End if

		'Converting BUSI type code to a human-readable string
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"

		'Reading and converting BUSI Self employment method into human-readable
		EMReadScreen BUSI_method, 2, 16, 53
		IF BUSI_method = "01" THEN BUSI_method = "50% Gross Income"
		IF BUSI_method = "02" THEN BUSI_method = "Tax Forms"

		'Going to the Gross Income Calculation pop-up
		EMWriteScreen "X", 6, 26
		transmit

		'Getting the verification codes for each type. Only does income, expenses are not included at this time.
		EMReadScreen BUSI_cash_ver, 1, 9, 73
		EMReadScreen BUSI_IVE_ver, 1, 10, 73
		EMReadScreen BUSI_SNAP_ver, 1, 11, 73
		EMReadScreen BUSI_HCA_ver, 1, 12, 73
		EMReadScreen BUSI_HCB_ver, 1, 13, 73

		'Converts each ver type to human readable
		If BUSI_cash_ver = "1" then BUSI_cash_ver = "tax returns provided"
		If BUSI_cash_ver = "2" then BUSI_cash_ver = "receipts provided"
		If BUSI_cash_ver = "3" then BUSI_cash_ver = "client ledger provided"
		If BUSI_cash_ver = "6" then BUSI_cash_ver = "other doc provided"
		If BUSI_cash_ver = "N" then BUSI_cash_ver = "no proof provided"
		If BUSI_cash_ver = "?" then BUSI_cash_ver = "no proof provided"
		If BUSI_IVE_ver = "1" then BUSI_IVE_ver = "tax returns provided"
		If BUSI_IVE_ver = "2" then BUSI_IVE_ver = "receipts provided"
		If BUSI_IVE_ver = "3" then BUSI_IVE_ver = "client ledger provided"
		If BUSI_IVE_ver = "6" then BUSI_IVE_ver = "other doc provided"
		If BUSI_IVE_ver = "N" then BUSI_IVE_ver = "no proof provided"
		If BUSI_IVE_ver = "?" then BUSI_IVE_ver = "no proof provided"
		If BUSI_SNAP_ver = "1" then BUSI_SNAP_ver = "tax returns provided"
		If BUSI_SNAP_ver = "2" then BUSI_SNAP_ver = "receipts provided"
		If BUSI_SNAP_ver = "3" then BUSI_SNAP_ver = "client ledger provided"
		If BUSI_SNAP_ver = "6" then BUSI_SNAP_ver = "other doc provided"
		If BUSI_SNAP_ver = "N" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_SNAP_ver = "?" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_HCA_ver = "1" then BUSI_HCA_ver = "tax returns provided"
		If BUSI_HCA_ver = "2" then BUSI_HCA_ver = "receipts provided"
		If BUSI_HCA_ver = "3" then BUSI_HCA_ver = "client ledger provided"
		If BUSI_HCA_ver = "6" then BUSI_HCA_ver = "other doc provided"
		If BUSI_HCA_ver = "N" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCA_ver = "?" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCB_ver = "1" then BUSI_HCB_ver = "tax returns provided"
		If BUSI_HCB_ver = "2" then BUSI_HCB_ver = "receipts provided"
		If BUSI_HCB_ver = "3" then BUSI_HCB_ver = "client ledger provided"
		If BUSI_HCB_ver = "6" then BUSI_HCB_ver = "other doc provided"
		If BUSI_HCB_ver = "N" then BUSI_HCB_ver = "no proof provided"
		If BUSI_HCB_ver = "?" then BUSI_HCB_ver = "no proof provided"

		'Back to the main screen
		PF3

		'Reading each income amount, trimming them to clean out unneeded spaces.
		EMReadScreen BUSI_cash_retro_amt, 8, 8, 55
		BUSI_cash_retro_amt = trim(BUSI_cash_retro_amt)
		EMReadScreen BUSI_cash_pro_amt, 8, 8, 69
		BUSI_cash_pro_amt = trim(BUSI_cash_pro_amt)
		EMReadScreen BUSI_IVE_amt, 8, 9, 69
		BUSI_IVE_amt = trim(BUSI_IVE_amt)
		EMReadScreen BUSI_SNAP_retro_amt, 8, 10, 55
		BUSI_SNAP_retro_amt = trim(BUSI_SNAP_retro_amt)
		EMReadScreen BUSI_SNAP_pro_amt, 8, 10, 69
		BUSI_SNAP_pro_amt = trim(BUSI_SNAP_pro_amt)

		'Pulls prospective amounts for HC, either from prosp side or from HC inc est.
		If pull_future_HC = False then
			EMReadScreen BUSI_HCA_amt, 8, 11, 69
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 12, 69
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
		Else
			EMWriteScreen "x", 17, 27
			transmit
			EMReadScreen BUSI_HCA_amt, 8, 15, 54
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 16, 54
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
			PF3
		End if

		'Reads end date logic (in case it ended), converts to an actual date
		EMReadScreen BUSI_income_end_date, 8, 5, 72
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")

		'Entering the variable details based on above
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI:; "
		If IsDate(BUSI_income_end_date) = True then	variable_name_for_BUSI = variable_name_for_BUSI & "- Income ended " & BUSI_income_end_date & ".; "
		If BUSI_cash_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH retro: $" & BUSI_cash_retro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_cash_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH pro: $" & BUSI_cash_pro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_IVE_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- IV-E: $" & BUSI_IVE_amt & " budgeted, " & BUSI_IVE_ver & "; "
		If BUSI_SNAP_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP retro: $" & BUSI_SNAP_retro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_SNAP_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP pro: $" & BUSI_SNAP_pro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_HCA_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method A: $" & BUSI_HCA_amt & " budgeted, " & BUSI_HCA_ver & "; "
		If BUSI_HCB_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method B: $" & BUSI_HCB_amt & " budgeted, " & BUSI_HCB_ver & "; "
		'Checks to see if pre 01/15 or post 02/15 then decides what to put in case note based on what was found/needed on the self employment method.
		If IsDate(BUSI_income_end_date) = false then
			IF BUSI_method <> "__" or BUSI_method = "" THEN
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: " & BUSI_method & "; "
			Else
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: None; "
			END IF
		End if
	End if
end function

function add_CARS_to_variable(CARS_variable)
'--- This function adds STAT/CARS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ CARS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, CARS
  EMReadScreen CARS_year, 4, 8, 31
  EMReadScreen CARS_make, 15, 8, 43
  CARS_make = replace(CARS_make, "_", "")
  EMReadScreen CARS_model, 15, 8, 66
  CARS_model = replace(CARS_model, "_", "")
  CARS_type = CARS_year & " " & CARS_make & " " & CARS_model
  CARS_type = split(CARS_type)
  For each CARS_part in CARS_type
    If len(CARS_part) > 1 then
      first_letter = ucase(left(CARS_part, 1))
      other_letters = LCase(right(CARS_part, len(CARS_part) -1))
      new_CARS_type = new_CARS_type & first_letter & other_letters & " "
    End if
  Next
  EMReadScreen CARS_amt, 8, 9, 45
  CARS_amt = trim(CARS_amt)
  CARS_amt = "$" & CARS_amt
  CARS_variable = CARS_variable & trim(new_CARS_type) & ", (" & CARS_amt & "); "
  new_CARS_type = ""
end function

function add_JOBS_to_variable(variable_name_for_JOBS)
'--- This function adds STAT/JOBS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ JOBS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, JOBS
  EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
  JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
  EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
  JOBS_type = replace(JOBS_type, "_", ""	)
  JOBS_type = trim(JOBS_type)
  JOBS_type = split(JOBS_type)
  For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
    If JOBS_part <> "" then
      first_letter = ucase(left(JOBS_part, 1))
      other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
      new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
    End if
  Next
  EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
  jobs_hourly_wage = replace(jobs_hourly_wage, "_", "")   'trimming any underscores
' Navigates to the FS PIC
    EMWriteScreen "X", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
		EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
		jobs_SNAP_prospective_amt = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
    EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
    transmit
'Navigats to GRH PIC
	EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
	IF GRH_PIC_check = "GRH" THEN
		EMWriteScreen "X", 19, 71
		transmit
		EMReadScreen GRH_JOBS_amt, 8, 16, 69
		GRH_JOBS_amt = trim(GRH_JOBS_amt)
		EMReadScreen GRH_pay_frequency, 1, 3, 63
		EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
		GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
		PF3
	END IF
'  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    retro_JOBS_amt = trim(retro_JOBS_amt)
'  Reads the information on the prospective side of JOBS
	EMReadScreen prospective_JOBS_amt, 8, 17, 67
	prospective_JOBS_amt = trim(prospective_JOBS_amt)
'  Reads the information about health care off of HC Income Estimator
    EMReadScreen pay_frequency, 1, 18, 35
	EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
	IF HC_income_est_check = "Est" Then 'this is the old position
		EMWriteScreen "X", 19, 54
	ELSE								'this is the new position
		EMWriteScreen "X", 19, 48
	END IF
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit

  EMReadScreen JOBS_ver, 1, 6, 38
  EMReadScreen JOBS_income_end_date, 8, 9, 49
	'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
  If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
  If IsDate(JOBS_income_end_date) = True then
    variable_name_for_JOBS = variable_name_for_JOBS & new_JOBS_type & "(ended " & JOBS_income_end_date & "); "
  Else
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
	If GRH_pay_frequency = "1" then GRH_pay_frequency = "monthly"
    If GRH_pay_frequency = "2" then GRH_pay_frequency = "semimonthly"
    If GRH_pay_frequency = "3" then GRH_pay_frequency = "biweekly"
    If GRH_pay_frequency = "4" then GRH_pay_frequency = "weekly"
    variable_name_for_JOBS = variable_name_for_JOBS & "EI from " & trim(new_JOBS_type) & ", " & JOBS_month  & " amts:; "
    If SNAP_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- SNAP PIC: $" & SNAP_JOBS_amt & "/" & snap_pay_frequency & ", SNAP PIC Prospective: $" & jobs_SNAP_prospective_amt & ", calculated " & date_of_pic_calc & "; "
    If GRH_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- GRH PIC: $" & GRH_JOBS_amt & "/" & GRH_pay_frequency & ", calculated " & GRH_date_of_pic_calc & "; "
	If retro_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- Retrospective: $" & retro_JOBS_amt & " total; "
    IF prospective_JOBS_amt <> "" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Prospective: $" & prospective_JOBS_amt & " total; "
    IF isnumeric(jobs_hourly_wage) THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Hourly Wage: $" & jobs_hourly_wage & "; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF MAXIS_footer_month = current_month_for_hc_est THEN
	IF HC_JOBS_amt <> "________" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- HC Inc Est: $" & HC_JOBS_amt & "/" & pay_frequency & "; "
    END IF
	If JOBS_ver = "N" or JOBS_ver = "?" then variable_name_for_JOBS = variable_name_for_JOBS & "- No proof provided for this panel; "
  End if
end function

function add_OTHR_to_variable(OTHR_variable)
'--- This function adds STAT/OTHR data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ OTHR_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, OTHR
  EMReadScreen OTHR_type, 16, 6, 43
  OTHR_type = trim(OTHR_type)
  EMReadScreen OTHR_amt, 10, 8, 40
  OTHR_amt = trim(OTHR_amt)
  OTHR_amt = "$" & OTHR_amt
  OTHR_variable = OTHR_variable & trim(OTHR_type) & ", (" & OTHR_amt & ").; "
  new_OTHR_type = ""
end function

function add_RBIC_to_variable(variable_name_for_RBIC)
'--- This function adds STAT/RBIC data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ RBIC_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, RBIC
	EMReadScreen RBIC_month, 5, 20, 55
	RBIC_month = replace(RBIC_month, " ", "/")
	EMReadScreen RBIC_type, 14, 5, 48
	RBIC_type = trim(RBIC_type)
	EMReadScreen RBIC01_pro_amt, 8, 10, 62
	RBIC01_pro_amt = trim(RBIC01_pro_amt)
	EMReadScreen RBIC02_pro_amt, 8, 11, 62
	RBIC02_pro_amt = trim(RBIC02_pro_amt)
	EMReadScreen RBIC03_pro_amt, 8, 12, 62
	RBIC03_pro_amt = trim(RBIC03_pro_amt)
	EMReadScreen RBIC01_retro_amt, 8, 10, 47
	IF RBIC01_retro_amt <> "________" THEN RBIC01_retro_amt = trim(RBIC01_retro_amt)
	EMReadScreen RBIC02_retro_amt, 8, 11, 47
	IF RBIC02_retro_amt <> "________" THEN RBIC02_retro_amt = trim(RBIC02_retro_amt)
	EMReadScreen RBIC03_retro_amt, 8, 12, 47
	IF RBIC03_retro_amt <> "________" THEN RBIC03_retro_amt = trim(RBIC03_retro_amt)
	EMReadScreen RBIC_group_01, 17, 10, 25
		RBIC_group_01 = replace(RBIC_group_01, " __", "")
		RBIC_group_01 = replace(RBIC_group_01, " ", ", ")
	EMReadScreen RBIC_group_02, 17, 11, 25
		RBIC_group_02 = replace(RBIC_group_02, " __", "")
		RBIC_group_02 = replace(RBIC_group_02, " ", ", ")
	EMReadScreen RBIC_group_03, 17, 12, 25
		RBIC_group_03 = replace(RBIC_group_03, " __", "")
		RBIC_group_03 = replace(RBIC_group_03, " ", ", ")

	EMReadScreen RBIC_01_verif, 1, 10, 76
	IF RBIC_01_verif = "N" THEN
		RBIC01_pro_amt = RBIC01_pro_amt & ", not verified"
		RBIC01_retro_amt = RBIC01_retro_amt & ", not verified"
	END IF

	EMReadScreen RBIC_02_verif, 1, 11, 76
	IF RBIC_02_verif = "N" THEN
		RBIC02_pro_amt = RBIC02_pro_amt & ", not verified"
		RBIC02_retro_amt = RBIC02_retro_amt & ", not verified"
	END IF

	EMReadScreen RBIC_03_verif, 1, 12, 76
	IF RBIC_03_verif = "N" THEN
		RBIC03_pro_amt = RBIC03_pro_amt & ", not verified"
		RBIC03_retro_amt = RBIC03_retro_amt & ", not verified"
	END IF

	RBIC_expense_row = 15
	DO
		EMReadScreen RBIC_expense_type, 13, RBIC_expense_row, 28
		RBIC_expense_type = trim(RBIC_expense_type)
		EMReadScreen RBIC_expense_amt, 8, RBIC_expense_row, 62
		RBIC_expense_amt = trim(RBIC_expense_amt)
		EMReadScreen RBIC_expense_verif, 1, RBIC_expense_row, 76
		IF RBIC_expense_type <> "" THEN
			total_RBIC_expenses = total_RBIC_expenses & "- " & RBIC_expense_type & ", $" & RBIC_expense_amt
			IF RBIC_expense_verif <> "N" THEN
				total_RBIC_expenses = total_RBIC_expenses & "; "
			ELSE
				total_RBIC_expenses = total_RBIC_expenses & ", not verified; "
			END IF
			RBIC_expense_row = RBIC_expense_row + 1
			IF RBIC_expense_row = 19 THEN
				PF20
				EMReadScreen RBIC_last_page, 21, 24, 2
				RBIC_expense_row = 15
			END IF
		END IF
	LOOP UNTIL RBIC_expense_type = "" OR RBIC_last_page = "THIS IS THE LAST PAGE"
	EMReadScreen RBIC_ver, 1, 10, 76
	If RBIC_ver = "N" then RBIC_ver = ", no proof provided"
	EMReadScreen RBIC_end_date, 8, 6, 68
	RBIC_end_date = replace(RBIC_end_date, " ", "/")
	If isdate(RBIC_end_date) = True then
		variable_name_for_RBIC = variable_name_for_RBIC & trim(RBIC_type) & " RBIC, ended " & RBIC_end_date & RBIC_ver & "; "
	Else
		IF left(RBIC01_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Prospective, ($" & RBIC01_pro_amt & "); "
		IF left(RBIC01_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Retrospective, ($" & RBIC01_retro_amt & "); "
		IF left(RBIC02_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Prospective, ($" & RBIC02_pro_amt & "); "
		IF left(RBIC02_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Retrospective, ($" & RBIC02_retro_amt & "); "
		IF left(RBIC03_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Prospective, ($" & RBIC03_pro_amt & "); "
		IF left(RBIC03_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Retrospective, ($" & RBIC03_retro_amt & "); "
		IF total_RBIC_expenses <> "" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC Expenses:; " & total_RBIC_expenses
	End if
end function

function add_REST_to_variable(REST_variable)
'--- This function adds STAT/REST data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ REST_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, REST
  EMReadScreen REST_type, 16, 6, 41
  REST_type = trim(REST_type)
  EMReadScreen REST_amt, 10, 8, 41
  REST_amt = trim(REST_amt)
  REST_amt = "$" & REST_amt
  REST_variable = REST_variable & trim(REST_type) & ", (" & REST_amt & ").; "
  new_REST_type = ""
end function

function add_SECU_to_variable(SECU_variable)
'--- This function adds STAT/SECU data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ SECU_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, SECU
  EMReadScreen SECU_amt, 8, 10, 52
  SECU_amt = trim(SECU_amt)
  SECU_amt = "$" & SECU_amt
  EMReadScreen SECU_type, 2, 6, 50
  EMReadScreen SECU_location, 20, 8, 50
  SECU_location = replace(SECU_location, "_", "")
  SECU_location = split(SECU_location)
  For each SECU_part in SECU_location
    If SECU_part <> "" then
      first_letter = ucase(left(SECU_part, 1))
      other_letters = LCase(right(SECU_part, len(SECU_part) -1))
      If len(a) > 3 then
        new_SECU_location = new_SECU_location & b & c & " "
      Else
        new_SECU_location = new_SECU_location & a & " "
      End if
    End if
  Next
  EMReadScreen SECU_ver, 1, 11, 50
  If SECU_ver = "1" then SECU_ver = "agency form provided"
  If SECU_ver = "2" then SECU_ver = "source doc provided"
  If SECU_ver = "3" then SECU_ver = "verified via phone"
  If SECU_ver = "5" then SECU_ver = "other doc verified"
  If SECU_ver = "N" then SECU_ver = "no proof provided"
  SECU_variable = SECU_variable & SECU_type & " at " & new_SECU_location & " (" & SECU_amt & "), " & SECU_ver & ".; "
  new_SECU_location = ""
end function

function add_UNEA_to_variable(variable_name_for_UNEA)
'--- This function adds STAT/UNEA data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ UNEA_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, UNEA
  EMReadScreen UNEA_month, 5, 20, 55
  UNEA_month = replace(UNEA_month, " ", "/")
  EMReadScreen UNEA_type, 16, 5, 40
  If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
  If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
  If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
  UNEA_type = trim(UNEA_type)
  EMReadScreen UNEA_ver, 1, 5, 65
  EMReadScreen UNEA_income_end_date, 8, 7, 68
  If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
  If IsDate(UNEA_income_end_date) = True then
    variable_name_for_UNEA = variable_name_for_UNEA & UNEA_type & " (ended " & UNEA_income_end_date & "); "
  Else
    EMReadScreen UNEA_amt, 8, 18, 68
    UNEA_amt = trim(UNEA_amt)
      EMWriteScreen "X", 10, 26
      transmit
      EMReadScreen SNAP_UNEA_amt, 8, 17, 56
      SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
      EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
      transmit
      EMReadScreen retro_UNEA_amt, 8, 18, 39
      retro_UNEA_amt = trim(retro_UNEA_amt)
	EMReadScreen prosp_UNEA_amt, 8, 18, 68
	prosp_UNEA_amt = trim(prosp_UNEA_amt)
      EMWriteScreen "X", 6, 56
      transmit
      EMReadScreen HC_UNEA_amt, 8, 9, 65
      HC_UNEA_amt = trim(HC_UNEA_amt)
      EMReadScreen pay_frequency, 1, 10, 63
      transmit
      If HC_UNEA_amt = "________" then
        EMReadScreen HC_UNEA_amt, 8, 18, 68
        HC_UNEA_amt = trim(HC_UNEA_amt)
        pay_frequency = "mo budgeted prospectively"
    End If
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
    variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amts:; "
    If SNAP_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
    If retro_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
    If prosp_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF MAXIS_footer_month = current_month_for_hc_est THEN
    	If HC_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- HC Inc Est: $" & HC_UNEA_amt & "/" & pay_frequency & "; "
    END IF
    If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "
  End if
end function

function assess_button_pressed()
'--- This fuction will review the button pressed on a dialog with tabs to go to the correct dialog if the do loop structure is used.
'===== Keywords: DIALOGS, NAVIGATES
    If ButtonPressed = dlg_one_button Then
        pass_one = false
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = true
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_two_button Then
        pass_one = true
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_three_button Then
        pass_one = true
        pass_two = true
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_four_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_five_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_six_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_seven_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_eight_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_nine_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_ten_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = true
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = false
        show_ten = true

        tab_button = True
    End If
end function


function assign_county_address_variables(address_line_01, address_line_02)
'--- This function will assign an address to a variable selected from the interview_location variable in the Appt Letter script.
'~~~~~ address_line_01: 1st line of address (street address) from new_office_array
'~~~~~ address_line_02: 2nd line of address (city/state/zip) from new_office_array
'===== Keywords: MAXIS, APPT LETTER, ADDRESS
	For each office in county_office_array				'Splits the county_office_array, which is set by the config program and declared earlier in this file
		If instr(office, interview_location) <> 0 then		'If the name of the office is found in the "interview_location" variable, which is contained in the MEMO - appt letter script.
			new_office_array = split(office, "|")		'Split the office into its own array
			address_line_01 = new_office_array(0)		'Line 1 of the address is the first part of this array
			address_line_02 = new_office_array(1)		'Line 2 of the address is the second part of this array
		End if
	Next
end function

function attn()
 '--- This function sends or hits the ESC (escape) key.
  '===== Keywords: MAXIS, MMIS, PRISM, ESC
  EMSendKey "<attn>"
  EMWaitReady -1, 0
end function

function autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
 '--- This function autofills information for all HH members idenified from the HH_member_array from a selected MAXIS panel into an edit box in a dialog.
 '~~~~~ HH_member_array: array of HH members from function HH_member_custom_dialog(HH_member_array). User selects which HH members are added to array.
 '~~~~~ read_panel_from: first four characters because we use separate handling for HCRE-retro. This is something that should be fixed someday!!!!!!!!!
 '~~~~~ variable_written_to: the variable used by the editbox you wish to autofill.
 '===== Keywords: MAXIS, autofill, HH_member_array
  call navigate_to_MAXIS_screen("STAT", left(panel_read_from, 4))

  'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
  EMReadScreen panel_total_check, 6, 2, 73
  IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

  If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then
      Do
        'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
        EMReadScreen support_coop_code, 1, 4, 73
        If support_coop_code = "N" then
          EMReadScreen caregiver_ref_nbr, 2, 4, 47
          If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated.
        End if
        'Then it gets info on the ABPS themself.
        EMReadScreen ABPS_current, 45, 10, 30
        If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
        ABPS_current = replace(ABPS_current, "  First:", ",")
        ABPS_current = replace(ABPS_current, "_", "")
		ABPS_current = trim(ABPS_current)
        ABPS_current = split(ABPS_current)
        For each ABPS_part in ABPS_current
          first_letter = ucase(left(ABPS_part, 1))
          other_letters = LCase(right(ABPS_part, len(ABPS_part) -1))
          If len(ABPS_part) > 1 then
            new_ABPS_current = new_ABPS_current & first_letter & other_letters & " "
          Else
            new_ABPS_current = new_ABPS_current & ABPS_part & " "
          End if
        Next
        ABPS_row = 15 'Setting variable for do...loop
        Do
          Do 'Using a do...loop to determine which MEMB numbers are with this parent
            EMReadScreen child_ref_nbr, 2, ABPS_row, 35
            If child_ref_nbr <> "__" then
              amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
              children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
            End if
            ABPS_row = ABPS_row + 1
          Loop until ABPS_row > 17		'End of the row
          EMReadScreen more_check, 7, 19, 66
          If more_check = "More: +" then
            EMSendKey "<PF20>"
            EMWaitReady 0, 0
            ABPS_row = 15
          End if
        Loop until more_check <> "More: +"
        'Cleaning up the "children_for_ABPS" variable to be more readable
		If children_for_ABPS = "" Then
			stop_message = "The script you are running " & replace(name_of_script, ".vbs", "") & " is attempting to read information from ABPS. This ABPS panel does not have any children listed. Review the STAT panels, particularly about parental relationships (ABPS/PARE). This panel needs update, or may need to be deleted."
			script_end_procedure(stop_message)
		End If
        children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
        children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around
        if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
        if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
        variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
        'Resetting variables for the do...loop in case this function runs again
        new_ABPS_current = ""
        amt_of_children_for_ABPS = 0
        children_for_ABPS = ""
        'Checking to see if it needs to run again, if it does it transmits or else the loop stops
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
      'Combining the two variables (support coop and the variable written to)
      variable_written_to = support_coop & variable_written_to
    End if
  Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCI_total, 1, 2, 78
      If ACCI_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCI_to_variable(variable_written_to)
          EMReadScreen ACCI_panel_current, 1, 2, 73
          If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
        Loop until cint(ACCI_panel_current) = cint(ACCI_total)
      End if
    Next
  Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCT_total, 2, 2, 78
	  ACCT_total = trim(ACCT_total)   'deleting space if one digit.
      If ACCT_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCT_to_variable(variable_written_to)
          EMReadScreen ACCT_panel_current, 2, 2, 72
          ACCT_panel_current = trim(ACCT_panel_current)
          If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
        Loop until cint(ACCT_panel_current) = cint(ACCT_total)
      End if
    Next
  ElseIf panel_read_from = "ACUT" Then '----------------------------------------------------------------------------------------------------ACUT
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        transmit
        EMReadScreen ACUT_total, 1, 2, 78
        If ACUT_total <> 0 then
            EMReadScreen share_yn, 1, 6, 42
            EMReadScreen retro_heat_verif, 1, 10, 35
            EMReadScreen retro_heat_amount, 8, 10, 41
            EMReadScreen retro_air_verif, 1, 11, 35
            EMReadScreen retro_air_amount, 8, 11, 41
            EMReadScreen retro_elec_verif, 1, 12, 35
            EMReadScreen retro_elec_amount, 8, 12, 41
            EMReadScreen retro_fuel_verif, 1, 13, 35
            EMReadScreen retro_fuel_amount, 8, 13, 41
            EMReadScreen retro_garbage_verif, 1, 14, 35
            EMReadScreen retro_garbage_amount, 8, 14, 41
            EMReadScreen retro_water_verif, 1, 15, 35
            EMReadScreen retro_water_amount, 8, 15, 41
            EMReadScreen retro_sewer_verif, 1, 16, 35
            EMReadScreen retro_sewer_amount, 8, 16, 41
            EMReadScreen retro_other_verif, 1, 17, 35
            EMReadScreen retro_other_amount, 8, 17, 41

            EMReadScreen prosp_heat_verif, 1, 10, 55
            EMReadScreen prosp_heat_amount, 8, 10, 61
            EMReadScreen prosp_air_verif, 1, 11, 55
            EMReadScreen prosp_air_amount, 8, 11, 61
            EMReadScreen prosp_elec_verif, 1, 12, 55
            EMReadScreen prosp_elec_amount, 8, 12, 61
            EMReadScreen prosp_fuel_verif, 1, 13, 55
            EMReadScreen prosp_fuel_amount, 8, 13, 61
            EMReadScreen prosp_garbage_verif, 1, 14, 55
            EMReadScreen prosp_garbage_amount, 8, 14, 61
            EMReadScreen prosp_water_verif, 1, 15, 55
            EMReadScreen prosp_water_amount, 8, 15, 61
            EMReadScreen prosp_sewer_verif, 1, 16, 55
            EMReadScreen prosp_sewer_amount, 8, 16, 61
            EMReadScreen prosp_other_verif, 1, 17, 55
            EMReadScreen prosp_other_amount, 8, 17, 61

            EMReadScreen dwp_phone_yn, 1, 18, 55

            variable_written_to = "Actutal Utilitiy Expense for M" & HH_member
            If share_yn = "Y" Then variable_written_to = variable_written_to & " - this expense is shared."
            If retro_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (retro) $" & trim(retro_heat_amount) & " - Verif: " & retro_heat_verif & "."
            If prosp_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (prosp) $" & trim(prosp_heat_amount) & " - Verif: " & prosp_heat_verif & "."
            If retro_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (retro) $" & trim(retro_air_amount) & " - Verif: " & retro_air_verif & "."
            If prosp_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (prosp) $" & trim(prosp_air_amount) & " - Verif: " & prosp_air_verif & "."
            If retro_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (retro) $" & trim(retro_elec_amount) & " - Verif: " & retro_elec_verif & "."
            If prosp_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (prosp) $" & trim(prosp_elec_amount) & " - Verif: " & prosp_elec_verif & "."
            If retro_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (retro) $" & trim(retro_fuel_amount) & " - Verif: " & retro_fuel_verif & "."
            If prosp_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (prosp) $" & trim(prosp_fuel_amount) & " - Verif: " & prosp_fuel_verif & "."
            If retro_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (retro) $" & trim(retro_garbage_amount) & " - Verif: " & retro_garbage_verif & "."
            If prosp_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (prosp) $" & trim(prosp_garbage_amount) & " - Verif: " & prosp_garbage_verif & "."
            If retro_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (retro) $" & trim(retro_water_amount) & " - Verif: " & retro_water_verif & "."
            If prosp_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (prosp) $" & trim(prosp_water_amount) & " - Verif: " & prosp_water_verif & "."
            If retro_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (retro) $" & trim(retro_sewer_amount) & " - Verif: " & retro_sewer_verif & "."
            If prosp_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (prosp) $" & trim(prosp_sewer_amount) & " - Verif: " & prosp_sewer_verif & "."
            If retro_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (retro) $" & trim(retro_other_amount) & " - Verif: " & retro_other_verif & "."
            If prosp_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (prosp) $" & trim(prosp_other_amount) & " - Verif: " & prosp_other_verif & "."
            If dwp_phone_yn = "Y" Then variable_written_to = variable_written_to & " Standard DWP Phone allowance of $35."

        End If
    Next
  Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
    EMReadScreen addr_line_01, 22, 6, 43
    EMReadScreen addr_line_02, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 12, 9, 43
    variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
    variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
  Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
    EMReadScreen AREP_name, 37, 4, 32
    AREP_name = replace(AREP_name, "_", "")
    AREP_name = split(AREP_name)
    For each word in AREP_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
    EMReadScreen BILS_amt, 1, 2, 78
    If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
  Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen BUSI_total, 1, 2, 78
      If BUSI_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_BUSI_to_variable(variable_written_to)
          EMReadScreen BUSI_panel_current, 1, 2, 73
          If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
        Loop until cint(BUSI_panel_current) = cint(BUSI_total)
      End if
    Next
  Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen CARS_total, 2, 2, 78
	  CARS_total = trim(CARS_total)
      If CARS_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_CARS_to_variable(variable_written_to)
          EMReadScreen CARS_panel_current, 2, 2, 72
		  CARS_panel_current = trim(CARS_panel_current)
          If cint(CARS_panel_current) < cint(CARS_total) then transmit
        Loop until cint(CARS_panel_current) = cint(CARS_total)
      End if
    Next
  Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen cash_amt, 8, 8, 39
      cash_amt = trim(cash_amt)
      If cash_amt <> "________" then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
      End if
    Next
  Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen support_amt, 8, 10, 63
      support_amt = trim(support_amt)
      If support_amt <> "________" then
        EMReadScreen support_ver, 1, 10, 36
        If support_ver = "?" or support_ver = "N" then
          support_ver = ", no proof provided"
        Else
          support_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
      End if
      EMReadScreen alimony_amt, 8, 11, 63
      alimony_amt = trim(alimony_amt)
      If alimony_amt <> "________" then
        EMReadScreen alimony_ver, 1, 11, 36
        If alimony_ver = "?" or alimony_ver = "N" then
          alimony_ver = ", no proof provided"
        Else
          alimony_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
      End if
      EMReadScreen tax_dep_amt, 8, 12, 63
      tax_dep_amt = trim(tax_dep_amt)
      If tax_dep_amt <> "________" then
        EMReadScreen tax_dep_ver, 1, 12, 36
        If tax_dep_ver = "?" or tax_dep_ver = "N" then
          tax_dep_ver = ", no proof provided"
        Else
          tax_dep_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
      End if
      EMReadScreen other_COEX_amt, 8, 13, 63
      other_COEX_amt = trim(other_COEX_amt)
      If other_COEX_amt <> "________" then
        EMReadScreen other_COEX_ver, 1, 13, 36
        If other_COEX_ver = "?" or other_COEX_ver = "N" then
          other_COEX_ver = ", no proof provided"
        Else
          other_COEX_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
      End if
    Next
  Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadScreen DCEX_total, 1, 2, 78
      If DCEX_total <> 0 then
		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		Do
			DCEX_row = 11
			Do
				EMReadScreen expense_amt, 8, DCEX_row, 63
				expense_amt = trim(expense_amt)
				If expense_amt <> "________" then
					EMReadScreen child_ref_nbr, 2, DCEX_row, 29
					EMReadScreen expense_ver, 1, DCEX_row, 41
					If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
						expense_ver = ", no proof provided"
					Else
						expense_ver = ""
					End if
					variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
				End if
				DCEX_row = DCEX_row + 1
			Loop until DCEX_row = 17
			EMReadScreen DCEX_panel_current, 1, 2, 73
			If cint(DCEX_panel_current) < cint(DCEX_total) then transmit
		Loop until cint(DCEX_panel_current) = cint(DCEX_total)
	  End if
    Next
  Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DIET_row = 8 'Setting this variable for the next do...loop
      EMReadScreen DIET_total, 1, 2, 78
      If DIET_total <> 0 then
        DIET = DIET & "Member " & HH_member & "- "
        Do
          EMReadScreen diet_type, 2, DIET_row, 40
          EMReadScreen diet_proof, 1, DIET_row, 51
          If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then
            diet_proof = ", no proof provided"
          Else
            diet_proof = ""
          End if
          If diet_type = "01" then diet_type = "High Protein"
          If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
          If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
          If diet_type = "04" then diet_type = "Lo Cholesterol"
          If diet_type = "05" then diet_type = "High Residue"
          If diet_type = "06" then diet_type = "Preg/Lactation"
          If diet_type = "07" then diet_type = "Gluten Free"
          If diet_type = "08" then diet_type = "Lactose Free"
          If diet_type = "09" then diet_type = "Anti-Dumping"
          If diet_type = "10" then diet_type = "Hypoglycemic"
          If diet_type = "11" then diet_type = "Ketogenic"
          If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
          DIET_row = DIET_row + 1
        Loop until DIET_row = 19
      End if
    Next
  Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
    For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadscreen DISA_total, 1, 2, 78
	  IF DISA_total <> 0 THEN
		'Reads and formats CASH/GRH disa status
		EMReadScreen CASH_DISA_status, 2, 11, 59
		EMReadScreen CASH_DISA_verif, 1, 11, 69
		IF CASH_DISA_status = "01" or CASH_DISA_status = "02" or CASH_DISA_status = "03" OR CASH_DISA_status = "04" THEN CASH_DISA_status = "RSDI/SSI certified"
		IF CASH_DISA_status = "06" THEN CASH_DISA_status = "SMRT/SSA pends"
		IF CASH_DISA_status = "08" THEN CASH_DISA_status = "Certified Blind"
		IF CASH_DISA_status = "09" THEN CASH_DISA_status = "Ill/Incap"
		IF CASH_DISA_status = "10" THEN CASH_DISA_status = "Certified disabled"
		IF CASH_DISA_verif = "?" OR CASH_DISA_verif = "N" THEN
			CASH_DISA_verif = ", no proof provided"
		ELSE
			CASH_DISA_verif = ""
		END IF

		'Reads and formats SNAP disa status
		EmreadScreen SNAP_DISA_status, 2, 12, 59
		EMReadScreen SNAP_DISA_verif, 1, 12, 69
		IF SNAP_DISA_status = "01" or SNAP_DISA_status = "02" or SNAP_DISA_status = "03" OR SNAP_DISA_status = "04" THEN SNAP_DISA_status = "RSDI/SSI certified"
		IF SNAP_DISA_status = "08" THEN SNAP_DISA_status = "Certified Blind"
		IF SNAP_DISA_status = "09" THEN SNAP_DISA_status = "Ill/Incap"
		IF SNAP_DISA_status = "10" THEN SNAP_DISA_status = "Certified disabled"
		IF SNAP_DISA_status = "11" THEN SNAP_DISA_status = "VA determined PD disa"
		IF SNAP_DISA_status = "12" THEN SNAP_DISA_status = "VA (other accept disa)"
		IF SNAP_DISA_status = "13" THEN SNAP_DISA_status = "Cert RR Ret Disa & on MEDI"
		IF SNAP_DISA_status = "14" THEN SNAP_DISA_status = "Other Govt Perm Disa Ret Bnft"
		IF SNAP_DISA_status = "15" THEN SNAP_DISA_status = "Disability from MINE list"
		IF SNAP_DISA_status = "16" THEN SNAP_DISA_status = "Unable to p&p own meal"
		IF SNAP_DISA_verif = "?" OR SNAP_DISA_verif = "N" THEN
			SNAP_DISA_verif = ", no proof provided"
		ELSE
			SNAP_DISA_verif = ""
		END IF

		'Reads and formats HC disa status/verif
		EMReadScreen HC_DISA_status, 2, 13, 59
		EMReadScreen HC_DISA_verif, 1, 13, 69
		If HC_DISA_status = "01" or HC_DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
		If HC_DISA_status = "06" then HC_DISA_status = "SMRT/SSA pends"
		If HC_DISA_status = "08" then HC_DISA_status = "Certified blind"
		If HC_DISA_status = "10" then HC_DISA_status = "Certified disabled"
		If HC_DISA_status = "11" then HC_DISA_status = "Spec cat- disa child"
		If HC_DISA_status = "20" then HC_DISA_status = "TEFRA- disabled"
		If HC_DISA_status = "21" then HC_DISA_status = "TEFRA- blind"
		If HC_DISA_status = "22" then HC_DISA_status = "MA-EPD"
		If HC_DISA_status = "23" then HC_DISA_status = "MA/waiver"
		If HC_DISA_status = "24" then HC_DISA_status = "SSA/SMRT appeal pends"
		If HC_DISA_status = "26" then HC_DISA_status = "SSA/SMRT disa deny"
		IF HC_DISA_verif = "?" OR HC_DISA_verif = "N" THEN
			HC_DISA_verif = ", no proof provided"
		ELSE
			HC_DISA_verif = ""
		END IF
		'cleaning to make variable to write
		IF CASH_DISA_status = "__" THEN
			CASH_DISA_status = ""
		ELSE
			IF CASH_DISA_status = SNAP_DISA_status THEN
				SNAP_DISA_status = "__"
				CASH_DISA_status = "CASH/SNAP: " & CASH_DISA_status & " "
			ELSE
				CASH_DISA_status = "CASH: " & CASH_DISA_status & " "
			END IF
		END IF
		IF SNAP_DISA_status = "__" THEN
			SNAP_DISA_status = ""
		ELSE
			SNAP_DISA_status = "SNAP: " & SNAP_DISA_status & " "
		END IF
		IF HC_DISA_status = "__" THEN
			HC_DISA_status = ""
		ELSE
			HC_DISA_status = "HC: " & HC_DISA_status & " "
		END IF
		'Adding verif code info if N or ?
		IF CASH_DISA_verif <> "" THEN CASH_DISA_status = CASH_DISA_status & CASH_DISA_verif & " "
		IF SNAP_DISA_verif <> "" THEN SNAP_DISA_status = SNAP_DISA_status & SNAP_DISA_verif & " "
		IF HC_DISA_verif <> "" THEN HC_DISA_status = HC_DISA_status & HC_DISA_verif & " "
		'Creating final variable
		IF CASH_DISA_status <> "" THEN FINAL_DISA_status = CASH_DISA_status
		IF SNAP_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & SNAP_DISA_status
		IF HC_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & HC_DISA_status

		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		variable_written_to = variable_written_to & FINAL_DISA_status & "; "
	  END IF
    Next
  Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
    row = 14
    Do
      EMReadScreen reference_numbers_current_row, 40, row, 39
      reference_numbers = reference_numbers + reference_numbers_current_row
      row = row + 1
    Loop until row = 18
    reference_numbers = replace(reference_numbers, "  ", " ")
    reference_numbers = split(reference_numbers)
    For each member in reference_numbers
      If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
    Next
    EATS_info = trim(EATS_info)
    if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
    If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
 Elseif panel_read_from = "EMPS" then '----------------------------------------------------------------------------------------------------EMPS
    	For each HH_member in HH_member_array
  		'blanking out variables for the next HH member
  		EMPS_info = ""
  		ES_exemptions = ""
  		ES_info = ""
  		EMWriteScreen HH_member, 20, 76
  		EMWriteScreen "01", 20, 79
  		transmit
  		EMReadScreen EMPS_total, 1, 2, 78
  		If EMPS_total <> 0 then
  			'orientation info (EMPS_info variable)-------------------------------------------------------------------------
  			EMReadScreen EMPS_orientation_date, 8, 5, 39
  			IF EMPS_orientation_date = "__ __ __" then
  				EMPS_orientation_date = "none"
  			ElseIf EMPS_orientation_date <> "__ __ __" then
  				EMPS_orientation_date = replace(EMPS_orientation_date, " ", "/")
  				EMPS_info = EMPS_info & " Fin orient: " & EMPS_orientation_date & ","
  			END IF
  	  		EMReadScreen EMPS_orientation_attended, 1, 5, 65
  			IF EMPS_orientation_attended <> "_" then EMPS_info = EMPS_info & " Attended orient: " & EMPS_orientation_attended & ","
  			'Good cause (EMPS_info variable)
  			EMReadScreen EMPS_good_cause, 2, 5, 79
  			IF EMPS_good_cause <> "__" then
  				If EMPS_good_cause = "01" then EMPS_good_cause = "01-No Good Cause"
  				If EMPS_good_cause = "02" then EMPS_good_cause = "02-No Child Care"
  				If EMPS_good_cause = "03" then EMPS_good_cause = "03-Ill or Injured"
  				If EMPS_good_cause = "04" then EMPS_good_cause = "04-Care Ill/Incap. Family Member"
  				If EMPS_good_cause = "05" then EMPS_good_cause = "05-Lack of Transportation"
  				If EMPS_good_cause = "06" then EMPS_good_cause = "06-Emergency"
  				If EMPS_good_cause = "07" then EMPS_good_cause = "07-Judicial Proceedings"
  				If EMPS_good_cause = "08" then EMPS_good_cause = "08-Conflicts with Work/School"
  				If EMPS_good_cause = "09" then EMPS_good_cause = "09-Other Impediments"
  				If EMPS_good_cause = "10" then EMPS_good_cause = "10-Special Medical Criteria "
  				If EMPS_good_cause = "20" then EMPS_good_cause = "20-Exempt--Only/1st Caregiver Employed 35+ Hours"
  				If EMPS_good_cause = "21" then EMPS_good_cause = "21-Exempt--2nd Caregiver Employed 20+ Hours"
  				If EMPS_good_cause = "22" then EMPS_good_cause = "22-Exempt--Preg/Parenting Caregiver < Age 20"
  				If EMPS_good_cause = "23" then EMPS_good_cause = "23-Exempt--Special Medical Criteria"
  				IF EMPS_good_cause <> "__" then EMPS_info = EMPS_info & " Good cause: " & EMPS_good_cause & ","
  			END IF

  			'sanction dates (EMPS_info variable)
  			EMReadScreen EMPS_sanc_begin, 8, 6, 39
  			If EMPS_sanc_begin <> "__ 01 __" then
  				EMPS_sanc_begin = replace(EMPS_sanc_begin, "_", "/")
  				sanction_date = sanction_date & EMPS_sanc_begin
  			END IF
  			EMReadScreen EMPS_sanc_end, 8, 6, 65
  			If EMPS_sanc_end <> "__ 01 __" then
  				EMPS_sanc_end = replace(EMPS_sanc_end, "_", "/")
  				sanction_date = sanction_date & "-" & EMPS_sanc_end
  			END IF
  			IF sanction_date <> "" then EMPS_info = EMPS_info & " Sanction dates: " & sanction_date & ","
  			'cleaning up ES_info variable
  			If right(EMPS_info, 1) = "," then EMPS_info = left(EMPS_info, len(EMPS_info) - 1)
  			IF trim(EMPS_info) <> "" then EMPS_info = EMPS_info & "."

  			'other sanction dates (ES_exemptions variable)--------------------------------------------------------------------------------
  			'special medical criteria
			EMReadScreen EMPS_memb_at_home, 1, 8, 76
  			IF EMPS_memb_at_home <> "N" then
				If EMPS_memb_at_home = "1" then EMPS_memb_at_home = "Home-Health/Waiver service"
				IF EMPS_memb_at_home = "2" then EMPS_memb_at_home = "Child w/ severe emotional dist"
				IF EMPS_memb_at_home = "3" then EMPS_memb_at_home = "Adult/Serious Persistent MI"
				ES_exemptions = ES_exemptions & " Special med criteria: " & EMPS_memb_at_home & ","
  			END IF

			EMReadScreen EMPS_care_family, 1, 9, 76
  			IF EMPS_care_family = "Y" then ES_exemptions = ES_exemptions & " Care of ill/incap memb: " & EMPS_care_family & ","
  			EMReadScreen EMPS_crisis, 1, 10, 76
  			IF EMPS_crisis = "Y" then ES_exemptions = ES_exemptions & " Family crisis: " & EMPS_crisis & ","

			'hard to employ
			EMReadScreen EMPS_hard_employ, 2, 11, 76
  			IF EMPS_hard_employ <> "NO" then
				IF EMPS_hard_employ = "IQ" then EMPS_hard_employ = "IQ tested at < 80"
				IF EMPS_hard_employ = "LD" then EMPS_hard_employ = "Learning Disabled"
				IF EMPS_hard_employ = "MI" then EMPS_hard_employ = "Mentally ill"
				IF EMPS_hard_employ = "DD" then EMPS_hard_employ = "Dev Disabled"
				IF EMPS_hard_employ = "UN" then EMPS_hard_employ = "Unemployable"
				ES_exemptions = ES_exemptions & " Hard to employ: " & EMPS_hard_employ & ","
  			END IF

  			'EMPS under 1 coding and dates used(ES_exemptions variable)
  			EMReadScreen EMPS_under1, 1, 12, 76
  			IF EMPS_under1 = "Y" then
  				ES_exemptions = ES_exemptions & " FT child under 1: " & EMPS_under1 & ","
  				EMWriteScreen "X", 12, 39
  				transmit
  				MAXIS_row = 7
  				MAXIS_col = 22
  				DO
  					EMReadScreen exemption_date, 9, MAXIS_row, MAXIS_col
  					If trim(exemption_date) = "" then exit do
  					If exemption_date <> "__ / ____" then
  						child_under1_dates = child_under1_dates & exemption_date & ", "
  						MAXIS_col = MAXIS_col + 11
  						If MAXIS_col = 66 then
  							MAXIS_row = MAXIS_row + 1
  							MAXIS_col = 22
  						END IF
  					END IF
  				LOOP until exemption_date = "__ / ____" or (MAXIS_row = 9 and MAXIS_col = 66)
  				PF3
  				'cleaning up excess comma at the end of child_under1_dates variable
  				If right(child_under1_dates,  2) = ", " then child_under1_dates = left(child_under1_dates, len(child_under1_dates) - 2)
  				If trim(child_under1_dates) = "" then child_under1_dates = " N/A"
  				ES_exemptions = ES_exemptions & " Child under 1 exeption dates: " & child_under1_dates & ","
  			END IF

  			'cleaning up ES_exemptions variable
  			If right(ES_exemptions, 1) = "," then ES_exemptions = left(ES_exemptions, len(ES_exemptions) - 1)
  			IF trim(ES_exemptions) <> "" then ES_exemptions = ES_exemptions & "."

  			'Reading ES Information (for ES_info variable)
  			EMReadScreen ES_status, 40, 15, 40
  			ES_status = trim(ES_status)
  			IF ES_status <> "" then ES_info = ES_info & " ES status: " & ES_status & ","
  			EMReadScreen ES_referral_date, 8, 16, 40
  			If ES_referral_date <> "__ __ __" then
  				ES_referral_date = replace(ES_referral_date, " ", "/")
  				ES_info = ES_info & " ES referral date: " & ES_referral_date & ","
  			END IF

  			EMReadScreen DWP_plan_date, 8, 17, 40
  			IF DWP_plan_date <> "__ __ __" then
  				DWP_plan_date = replace(DWP_plan_date, "_", "/")
  				ES_info = ES_info & " DWP plan date: " & DWP_plan_date & ","
  			END IF

			EMReadScreen minor_ES_option, 2, 16, 76
			If minor_ES_option <> "__" then
				IF minor_ES_option = "SC" then minor_ES_option = "Secondary Education"
				IF minor_ES_option = "EM" then minor_ES_option = "Employment"
				ES_info = ES_info & " 18/19 yr old ES option: " & minor_ES_option & ","
			END if

			'cleaning up ES_info variable
  			If right(ES_info, 1) = "," then ES_info = left(ES_info, len(ES_info) - 1)

  			variable_written_to = variable_written_to & "Member " & HH_member & "- "
  			variable_written_to = variable_written_to & EMPS_info & ES_exemptions & ES_info & "; "
  		END IF
  	next
  Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen FACI_total, 1, 2, 78
      If FACI_total <> 0 then
        row = 14
        Do
          EMReadScreen date_in_check, 4, row, 53
		  EMReadScreen date_in_month_day, 5, row, 47
          EMReadScreen date_out_check, 4, row, 77
		  date_in_month_day = replace(date_in_month_day, " ", "/") & "/"
          If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
          If row > 18 then
            EMReadScreen FACI_page, 1, 2, 73
            If FACI_page = FACI_total then
              FACI_status = "Not in facility"
            Else
              transmit
              row = 14
            End if
          End if
        Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
        EMReadScreen client_FACI, 30, 6, 43
        client_FACI = replace(client_FACI, "_", "")
        FACI_array = split(client_FACI)
        For each a in FACI_array
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            new_FACI = new_FACI & b & c & " "
          End if
        Next
        client_FACI = new_FACI
        If FACI_status = "Not in facility" then
          client_FACI = ""
        Else
          variable_written_to = variable_written_to & "Member " & HH_member & "- "
          variable_written_to = variable_written_to & client_FACI & " Date in: " & date_in_month_day & date_in_check & "; "
        End if
      End if
    Next
  Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
	For each HH_member in HH_member_array
	  EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	  If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      fmed_row = 9 'Setting this variable for the next do...loop
      EMReadScreen fmed_total, 1, 2, 78
      If fmed_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
		  use_expense = False					'<--- Used to determine if an FMED expense that has an end date is going to be counted.
          EMReadScreen fmed_type, 2, fmed_row, 25
          EMReadScreen fmed_proof, 2, fmed_row, 32
          EMReadScreen fmed_amt, 8, fmed_row, 70
		  EMReadScreen fmed_end_date, 5, fmed_row, 60		'reading end date to see if this one even gets added.
		  IF fmed_end_date <> "__ __" THEN
			fmed_end_date = replace(fmed_end_date, " ", "/01/")
			fmed_end_date = dateadd("M", 1, fmed_end_date)
			fmed_end_date = dateadd("D", -1, fmed_end_date)
			IF datediff("D", date, fmed_end_date) > 0 THEN use_expense = True		'<--- If the end date of the FMED expense is the current month or a future month, the expense is going to be counted.
		  END IF
		  If fmed_end_date = "__ __" OR use_expense = TRUE then					'Skips entries with an end date or end dates in the past.
            If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then
              fmed_proof = ", no proof provided"
            Else
              fmed_proof = ""
            End if
            If fmed_amt = "________" then
              fmed_amt = ""
            Else
              fmed_amt = " ($" & trim(fmed_amt) & ")"
            End if
            If fmed_type = "01" then fmed_type = "Nursing Home"
            If fmed_type = "02" then fmed_type = "Hosp/Clinic"
            If fmed_type = "03" then fmed_type = "Physicians"
            If fmed_type = "04" then fmed_type = "Prescriptions"
            If fmed_type = "05" then fmed_type = "Ins Premiums"
            If fmed_type = "06" then fmed_type = "Dental"
            If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
            If fmed_type = "08" then fmed_type = "Vision Care"
            If fmed_type = "09" then fmed_type = "Medicare Prem"
            If fmed_type = "10" then fmed_type = "Mo. Spdwn Amt/Waiver Obl"
            If fmed_type = "11" then fmed_type = "Home Care"
            If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
            If fmed_type = "15" then fmed_type = "Medi Part D premium"
            If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
			IF fmed_end_date <> "__ __" THEN					'<--- If there is a counted FMED expense with a future end date, the script will modify the way that end date is displayed.
				fmed_end_date = datepart("M", fmed_end_date) & "/" & right(datepart("YYYY", fmed_end_date), 2)		'<--- Begins pulling apart fmed_end_date to format it to human speak.
				IF left(fmed_end_date, 1) <> "0" THEN fmed_end_date = "0" & fmed_end_date
				variable_written_to = left(variable_written_to, len(variable_written_to) - 2) & ", counted through " & fmed_end_date & "; "			'<--- Putting variable_written_to back together with FMED expense end date information.
			END IF
          End if
          fmed_row = fmed_row + 1
          If fmed_row = 15 then
            PF20
            fmed_row = 9
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
          End if
        Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
      End if
    Next
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
    EMReadScreen variable_written_to, 8, 10, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
    If isdate(variable_written_to) = False then variable_written_to = ""
  Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
    EMReadScreen variable_written_to, 5, 10, 64
    If isdate(variable_written_to) = True then
      variable_written_to = replace(variable_written_to, " ", "/01/")
      If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
        variable_written_to = variable_written_to
      Else
        variable_written_to = ""
      End if
    End if
  Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
    EMReadScreen HEST_total, 1, 2, 78
    If HEST_total <> 0 then
      EMReadScreen heat_air_check, 6, 13, 75
      If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
      EMReadScreen electric_check, 6, 14, 75
      If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
      EMReadScreen phone_check, 6, 15, 75
      If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
    End if
  Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen IMIG_total, 1, 2, 78
      If IMIG_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen IMIG_type, 30, 6, 48
        variable_written_to = variable_written_to & trim(IMIG_type) & "; "
      End if
    Next
  Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
    EMReadScreen INSA_amt, 1, 2, 78
    If INSA_amt <> 0 then
      'Runs once per INSA screen
		For i = 1 to INSA_amt step 1
			insurance_name = ""
			'Goes to the correct screen
			EMWriteScreen "0" & i, 20, 79
			transmit
			'Gather Insurance Name
			EMReadScreen INSA_name, 38, 10, 38
			INSA_name = replace(INSA_name, "_", "")
			INSA_name = split(INSA_name)
			For each word in INSA_name
				If trim(word) <> "" then
						first_letter_of_word = ucase(left(word, 1))
						rest_of_word = LCase(right(word, len(word) -1))
						If len(word) > 4 then
							insurance_name = insurance_name & first_letter_of_word & rest_of_word & " "
						Else
							insurance_name = insurance_name & word & " "
						End if
				End if
			Next
			'Create a list of members covered by this insurance
			INSA_row = 15 : INSA_col = 30
			insured_count = 0
			member_list = ""
			Do
				EMReadScreen insured_member, 2, INSA_row, INSA_col
				If insured_member <> "__" then
					if member_list = "" then member_list = insured_member
					if member_list <> "" then member_list = member_list & ", " & insured_member
					INSA_col = INSA_col + 4
					If INSA_col = 70 then
						INSA_col = 30 : INSA_row = 16
					End If
				End If
			loop until insured_member = "__"
			'Retain "variable_written_to" as is while also adding members covered by the insurance policy
			'Example - "Members: 01, 03, 07 are covered by Blue Cross Blue Shield; "
			variable_written_to = variable_written_to & "Members: " & member_list & " are covered by " & trim(insurance_name) & "; "
		Next
		'This will loop and add the above statement for all insurance policies listed
	End if
  Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_JOBS_to_variable(variable_written_to)
          EMReadScreen JOBS_panel_current, 1, 2, 73
          If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
        Loop until cint(JOBS_panel_current) = cint(JOBS_total)
      End if
    Next
  Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen MEDI_amt, 1, 2, 78
      If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
    Next
  Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen rel_to_applicant, 2, 10, 42
      EMReadScreen client_age, 3, 8, 76
      If client_age = "   " then client_age = 0
      If cint(client_age) >= 21 or rel_to_applicant = "02" then
        number_of_adults = number_of_adults + 1
      Else
        number_of_children = number_of_children + 1
      End if
    Next
    If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
    If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
    If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
  Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen citizen, 1, 11, 49
      If citizen = "Y" then citizen = "US citizen"
      If citizen = "N" then citizen = "non-citizen"
      EMReadScreen citizenship_ver, 2, 11, 78
      EMReadScreen SSA_MA_citizenship_ver, 1, 12, 49
      If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
      If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
      If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
      variable_written_to = variable_written_to & "Member " & HH_member & "- "
      variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
    Next
  ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
    EMReadScreen variable_written_to, 8, 6, 39
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen OTHR_total, 2, 2, 78
	  OTHR_total = trim(OTHR_total)
      If OTHR_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_OTHR_to_variable(variable_written_to)
          EMReadScreen OTHR_panel_current, 2, 2, 72
		  OTHR_panel_current = trim(OTHR_panel_current)
          If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
        Loop until cint(OTHR_panel_current) = cint(OTHR_total)
      End if
    Next
  Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen panel_amt, 1, 2, 78
      If panel_amt <> "0" then
        PBEN = PBEN & "Member " & HH_member & "- "
        row = 8
        Do
          EMReadScreen PBEN_type, 12, row, 28
          EMReadScreen PBEN_disp, 1, row, 77
          If PBEN_disp = "A" then PBEN_disp = " appealing"
          If PBEN_disp = "D" then PBEN_disp = " denied"
          If PBEN_disp = "E" then PBEN_disp = " eligible"
          If PBEN_disp = "P" then PBEN_disp = " pends"
          If PBEN_disp = "N" then PBEN_disp = " not applied yet"
          If PBEN_disp = "R" then PBEN_disp = " refused"
          If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
          row = row + 1
        Loop until row = 14
      End if
    Next
    If PBEN <> "" then variable_written_to = variable_written_to & PBEN
  Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen PREG_total, 1, 2, 78
      If PREG_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen PREG_due_date, 8, 10, 53
        If PREG_due_date = "__ __ __" then
          PREG_due_date = "unknown"
        Else
          PREG_due_date = replace(PREG_due_date, " ", "/")
        End if
        variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
      End if
    Next
  Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
    row = 6
    Do
      EMReadScreen appl_prog_date, 8, row, 33
      If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
      row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    variable_written_to = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
      if CDate(appl_prog_date_array(i)) > variable_written_to then
        variable_written_to = CDate(appl_prog_date_array(i))
      End if
    next
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen RBIC_total, 1, 2, 78
      If RBIC_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_RBIC_to_variable(variable_written_to)
          EMReadScreen RBIC_panel_current, 1, 2, 73
          If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
        Loop until cint(RBIC_panel_current) = cint(RBIC_total)
      End if
    Next
  Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen REST_total, 2, 2, 78
	  REST_total = trim(REST_total)
      If REST_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_REST_to_variable(variable_written_to)
          EMReadScreen REST_panel_current, 2, 2, 72
		  REST_panel_current = trim(REST_panel_current)
          If cint(REST_panel_current) < cint(REST_total) then transmit
        Loop until cint(REST_panel_current) = cint(REST_total)
      End if
    Next
  Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
    EMReadScreen variable_written_to, 8, 13, 37
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
	For each HH_member in HH_member_array
			EMWriteScreen HH_member, 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen school_type, 2, 7, 40							'Reading the school type code and converting it into words
			If school_type = "01" then school_type = "elementary school"
			If school_type = "11" then school_type = "middle school"
			If school_type = "02" then school_type = "high school"
			If school_type = "03" then school_type = "GED"
			If school_type = "07" then school_type = "IEP"
			If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
			If school_type = "12" then school_type = "adult basic education"
			If school_type = "13" then school_type = "English as a 2nd language"
			If school_type = "06" or school_type = "__" or school_type = "?_" then  'if the school type is blank, child not in school, or postponed default type to blank.
				school_type = ""
			Else
				EMReadScreen SCHL_ver, 2, 6, 63
				If SCHL_ver = "?_" or SCHL_ver = "NO" then								'If the verification field is postponed or NO it defaults to no proof provided
					school_proof_type = ", no proof provided"
				Else
					school_proof_type = ""
				End if
				EMReadScreen FS_eligibility_status_SCHL, 2, 16, 63				'Reading the FS eligibility status and converting it to words
				IF FS_eligibility_status_SCHL = "01" THEN FS_eligibility_status_SCHL = ", FS Elig Status: < 18 or 50+"
				IF FS_eligibility_status_SCHL = "02" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Disabled"
				IF FS_eligibility_status_SCHL = "03" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Attenting Higher Ed or Attending < 1/2"
				IF FS_eligibility_status_SCHL = "04" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Employed 20 Hours/Wk"
				IF FS_eligibility_status_SCHL = "05" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Fed/State Work Study Program"
				IF FS_eligibility_status_SCHL = "06" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant Under 6"
				IF FS_eligibility_status_SCHL = "07" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant 6-11, daycare not available"
				IF FS_eligibility_status_SCHL = "09" THEN FS_eligibility_status_SCHL = ", FS Elig Status: WIA, TAA, TRA, or FSET placement"
				IF FS_eligibility_status_SCHL = "10" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Full Time Single Parent with Child under 12"
				IF FS_eligibility_status_SCHL = "99" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Eligible"
				IF FS_eligibility_status_SCHL = "__" or FS_eligibility_status_SCHL = "?_" THEN FS_eligibility_status_SCHL = ""
				'formatting the output variable for the function
				variable_written_to = variable_written_to & "Member " & HH_member & "- "
				variable_written_to = variable_written_to & school_type & school_proof_type & FS_eligibility_status_SCHL & "; "
			End if
		Next
  Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SECU_total, 2, 2, 78
	  SECU_total = trim(SECU_total)
      If SECU_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_SECU_to_variable(variable_written_to)
          EMReadScreen SECU_panel_current, 2, 2, 72
		  SECU_panel_current = trim(SECU_panel_current)
          If cint(SECU_panel_current) < cint(SECU_total) then transmit
        Loop until cint(SECU_panel_current) = cint(SECU_total)
      End if
    Next
  Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SHEL_total, 1, 2, 78
      If SHEL_total <> 0 then
        member_number_designation = "Member " & HH_member & "- "
        row = 11
        Do
          EMReadScreen SHEL_amount, 8, row, 56
          If SHEL_amount <> "________" then
            EMReadScreen SHEL_type, 9, row, 24
            EMReadScreen SHEL_proof_check, 2, row, 67
            If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then
              SHEL_proof = ", no proof provided"
            Else
              SHEL_proof = ""
            End if
            SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ". ;"
          End if
          row = row + 1
        Loop until row = 19
        variable_written_to = variable_written_to & member_number_designation & SHEL_expense
      End if
      SHEL_expense = ""
    Next
 Elseif panel_read_from = "SWKR" then '---------------------------------------------------------------------------------------------------SWKR
    EMReadScreen SWKR_name, 35, 6, 32
    SWKR_name = replace(SWKR_name, "_", "")
    SWKR_name = split(SWKR_name)
    For each word in SWKR_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen STWK_total, 1, 2, 78
      If STWK_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen STWK_verification, 1, 7, 63
        If STWK_verification = "N" then
          STWK_verification = ", no proof provided"
        Else
          STWK_verification = ""
        End if
        EMReadScreen STWK_employer, 30, 6, 46
        STWK_employer = replace(STWK_employer, "_", "")
        STWK_employer = split(STWK_employer)
        For each STWK_part in STWK_employer
          If STWK_part <> "" then
            first_letter = ucase(left(STWK_part, 1))
            other_letters = LCase(right(STWK_part, len(STWK_part) -1))
            If len(STWK_part) > 3 then
              new_STWK_employer = new_STWK_employer & first_letter & other_letters & " "
            Else
              new_STWK_employer = new_STWK_employer & STWK_part & " "
            End if
          End if
        Next
        EMReadScreen STWK_income_stop_date, 8, 8, 46
        If STWK_income_stop_date = "__ __ __" then
          STWK_income_stop_date = "at unknown date"
        Else
          STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
        End if
      EMReadScreen voluntary_quit, 1, 10, 46
	vol_quit_info = ", Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
      End if
      new_STWK_employer = "" 'clearing variable to prevent duplicates
    Next
  Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen UNEA_total, 1, 2, 78
      If UNEA_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_UNEA_to_variable(variable_written_to)
          EMReadScreen UNEA_panel_current, 1, 2, 73
          If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
        Loop until cint(UNEA_panel_current) = cint(UNEA_total)
      End if
    Next
  Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
    EMReadScreen wreg_total, 1, 2, 78
    IF wreg_total <> "0" THEN
	EmWriteScreen "X", 13, 57
	transmit
	 bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
	  bene_yr_row = 10
       abawd_counted_months = 0
       second_abawd_period = 0
 	 month_count = 0
 	   DO
	   		'establishing variables for specific ABAWD counted month dates
	 		If bene_mo_col = "19" then counted_date_month = "01"
	 		If bene_mo_col = "23" then counted_date_month = "02"
	 		If bene_mo_col = "27" then counted_date_month = "03"
	 		If bene_mo_col = "31" then counted_date_month = "04"
	 		If bene_mo_col = "35" then counted_date_month = "05"
	 		If bene_mo_col = "39" then counted_date_month = "06"
	 		If bene_mo_col = "43" then counted_date_month = "07"
	 		If bene_mo_col = "47" then counted_date_month = "08"
	 		If bene_mo_col = "51" then counted_date_month = "09"
	 		If bene_mo_col = "55" then counted_date_month = "10"
	 		If bene_mo_col = "59" then counted_date_month = "11"
	 		If bene_mo_col = "63" then counted_date_month = "12"
	 		'reading to see if a month is counted month or not
  		  	EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
			'counting and checking for counted ABAWD months
			IF is_counted_month = "X" or is_counted_month = "M" THEN
				EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
				abawd_counted_months_string = counted_date_month & "/" & counted_date_year
				abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
				abawd_counted_months = abawd_counted_months + 1				'adding counted months
			END IF

			'declaring & splitting the abawd months array
			If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
			abawd_months_array = Split(abawd_info_list, ",")

			'counting and checking for second set of ABAWD months
			IF is_counted_month = "Y" or is_counted_month = "N" THEN
				EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
				second_abawd_period = second_abawd_period + 1				'adding counted months
				second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
				second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
			END IF

			'declaring & splitting the second set of abawd months array
			If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
			second_months_array = Split(second_set_info_list,",")

			bene_mo_col = bene_mo_col - 4
    		IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
   	     		bene_mo_col = 63
   	   	    END IF
    		month_count = month_count + 1
  	   LOOP until month_count = 36
  		PF3

	EmreadScreen read_WREG_status, 2, 8, 50
	If read_WREG_status = "03" THEN  WREG_status = "WREG = incap"
	If read_WREG_status = "04" THEN  WREG_status = "WREG = resp for incap HH memb"
	If read_WREG_status = "05" THEN  WREG_status = "WREG = age 60+"
	If read_WREG_status = "06" THEN  WREG_status = "WREG = < age 16"
	If read_WREG_status = "07" THEN  WREG_status = "WREG = age 16-17, live w/prnt/crgvr"
	If read_WREG_status = "08" THEN  WREG_status = "WREG = resp for child < 6 yrs old"
	If read_WREG_status = "09" THEN  WREG_status = "WREG = empl 30 hrs/wk or equiv"
	If read_WREG_status = "10" THEN  WREG_status = "WREG = match grant part"
	If read_WREG_status = "11" THEN  WREG_status = "WREG = rec/app for unemp ins"
	If read_WREG_status = "12" THEN  WREG_status = "WREG = in schl, train prog or higher ed"
	If read_WREG_status = "13" THEN  WREG_status = "WREG = in CD prog"
	If read_WREG_status = "14" THEN  WREG_status = "WREG = rec MFIP"
	If read_WREG_status = "20" THEN  WREG_status = "WREG = pend/rec DWP or WB"
	If read_WREG_status = "22" THEN  WREG_status = "WREG = app for SSI"
	If read_WREG_status = "15" THEN  WREG_status = "WREG = age 16-17 not live w/ prnt/crgvr"
	If read_WREG_status = "16" THEN  WREG_status = "WREG = 50-59 yrs old"
	If read_WREG_status = "21" THEN  WREG_status = "WREG = resp for child < 18"
	If read_WREG_status = "17" THEN  WREG_status = "WREG = rec RCA or GA"
	If read_WREG_status = "18" THEN  WREG_status = "WREG = provide home schl"
	If read_WREG_status = "30" THEN  WREG_status = "WREG = mand FSET part"
	If read_WREG_status = "02" THEN  WREG_status = "WREG = non-coop w/ FSET"
	If read_WREG_status = "33" THEN  WREG_status = "WREG = non-coop w/ referral"
	If read_WREG_status = "__" THEN  WREG_status = "WREG = blank"

	EmreadScreen read_abawd_status, 2, 13, 50
	If read_abawd_status = "01" THEN  abawd_status = "ABAWD = work reg exempt."
    	If read_abawd_status = "02" THEN  abawd_status = "ABAWD = < age 18."
	If read_abawd_status = "03" THEN  abawd_status = "ABAWD = age 50+."
	If read_abawd_status = "04" THEN  abawd_status = "ABAWD = crgvr of minor child."
	If read_abawd_status = "05" THEN  abawd_status = "ABAWD = pregnant."
	If read_abawd_status = "06" THEN  abawd_status = "ABAWD = emp ave 20 hrs/wk."
	If read_abawd_status = "07" THEN  abawd_status = "ABAWD = work exp participant."
	If read_abawd_status = "08" THEN  abawd_status = "ABAWD = othr E & T service."
	If read_abawd_status = "09" THEN  abawd_status = "ABAWD = reside in waiver area."
	IF read_abawd_status = "10" AND abawd_counted_months = "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo."
	Elseif read_abawd_status = "10" AND second_abawd_period = "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & "."
	Elseif read_abawd_status = "10" AND second_abawd_period <> "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & ". Counted second set months: " & second_set_info_list & "."
	END IF
	If read_abawd_status = "11" THEN  abawd_status = "ABAWD = Using second set of ABAWD months. Counted second set months: " & second_set_info_list & "."
	If read_abawd_status = "12" THEN  abawd_status = "ABAWD = RCA or GA recip."
	'If read_abawd_status = "13" THEN  abawd_status = "ABAWD = ABAWD banked months."
	If read_abawd_status = "__" THEN  abawd_status = "ABAWD = blank"

	variable_written_to = variable_written_to & "Member " & HH_member & "- " & WREG_status & ", " & abawd_status & "; "
     END IF
    Next
  End if
  variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
  if right(variable_written_to, 1) = ";" then variable_written_to = left(variable_written_to, len(variable_written_to) - 1)
end function

function back_to_SELF()
'--- This function will return back to the 'SELF' menu or the MAXIS home menu
'===== Keywords: MAXIS, SELF, navigate
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
end function

function bypass_database_busy_msg()
'--- This function will transmit past the occasional message that a database record is busy when sending a case through background.
'===== Keywords: MAXIS, SELF, navigate
    EMReadScreen database_busy, 31, 4, 44
    If database_busy = "A MAXIS database record is busy" Then transmit
end function

function cancel_confirmation()
'--- This function asks if the user if they want to cancel. If you say yes, the script will end. If no, the dialog will appear for the user again.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
	End if
end function

function cancel_without_confirmation()
'--- This function ends a script after a user presses cancel. There is no confirmation message box but the end message for statistical information that cancel was pressed.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
        script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
        'Left the If...End If in the tier in case we want more stats or error handling, or if we need specialty processing for workflows
    End if
end function

function change_client_name_to_FML(client_name)
'--- This function changes the format of a participant name. client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".
'~~~~~ client_name: variable used within the script for name to be converted
'===== Keywords: PRISM, MAXIS, name, change

	client_name = trim(client_name)
	length = len(client_name)

	'Adds handling for names that have no spaces or 1 space
	If Instr(client_name, ", ") then
		position = InStr(client_name, ", ")
		first_name = Right(client_name, length-position - 1)
	elseif Instr(client_name, ",") then
		position = InStr(client_name, ",")                           '
		first_name = Right(client_name, length-position)
	END if
	last_name = Left(client_name, position - 1)

	'final formating of the client name
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
end function

function change_date_to_soonest_working_day(date_to_change, forward_or_back)
'--- This function will change a date that is on a weekend or Hennepin County holiday to the next working date before the date provided, the date will remain the same if it is not a holiday or weekend.
'~~~~~ date_to_change: variable in the form of a date - this will change once the function is called
'~~~~~ forward_or_back: varibale that must be either 'forward' or 'back' as a string to select which way your want to shift the date.
'===== Keywords: MAXIS, date, change
	forward_or_back = UCase(forward_or_back)
	If forward_or_back = "BACK" Then
	    Do
	        If WeekdayName(WeekDay(date_to_change)) = "Saturday" Then date_to_change = DateAdd("d", -1, date_to_change)
	        If WeekdayName(WeekDay(date_to_change)) = "Sunday" Then date_to_change = DateAdd("d", -2, date_to_change)
	        is_holiday = FALSE
	        For each holiday in HOLIDAYS_ARRAY
	            If holiday = date_to_change Then
	                is_holiday = TRUE
	                date_to_change = DateAdd("d", -1, date_to_change)
	            End If
	        Next
	    Loop until is_holiday = FALSE
	End If
	If forward_or_back = "FORWARD" Then
	    Do
	        If WeekdayName(WeekDay(date_to_change)) = "Saturday" Then date_to_change = DateAdd("d", 2, date_to_change)
	        If WeekdayName(WeekDay(date_to_change)) = "Sunday" Then date_to_change = DateAdd("d", 1, date_to_change)
	        is_holiday = FALSE
	        For each holiday in HOLIDAYS_ARRAY
	            If holiday = date_to_change Then
	                is_holiday = TRUE
	                date_to_change = DateAdd("d", 1, date_to_change)
	            End If
	        Next
	    Loop until is_holiday = FALSE
	End If

end function

function changelog_display()
'--- This function determines if the user has been informed of a change to a script, and if not will display a mesage box with the script's change log information
'===== Keywords: MAXIS, PRISM, change, info, information
	If name_of_script = "ACTIONS - DEU-MATCH CLEARED CC.vbs" or name_of_script = "ACTIONS - DEU-MATCH CLEARED CC" Then script_end_procedure_with_error_report("This script is no longer supported by the BlueZone Script team and cannot be run. PLease reach out to the BlueZone Script Team with any questions.")
	If changelog_enabled = "" Then changelog_enabled = true
	If changelog_enabled <> false Then
		'Needs to determine MyDocs directory before proceeding.
		Set wshshell = CreateObject("WScript.Shell")
		user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

		'Now determines name of file
		local_changelog_path = user_myDocs_folder & "scripts-local-changelog-entries.txt"

		'Before doing comparisons, it needs to see what the most recent item added to the list was.
		last_item_added_to_changelog = split(changelog(0), " | ")

		With objFSO

			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream

			'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

			If .FileExists(local_changelog_path) = False then
				'Setting the object to open the text file for appending the new data
				Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)

				'Write the contents of the text file
				objTextStream.WriteLine date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Close the object so it can be opened again shortly
				objTextStream.Close

				'Since the file was new, we can simply exit the function
				exit function
			End if

			'Setting the object to open the text file for reading the data already in the file
			Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

			'Reading the entire text file into a string
			every_line_in_text_file = objTextStream.ReadAll

			'Splitting the text file contents into an array which will be sorted
			local_changelog_array = split(every_line_in_text_file, vbNewLine)

			'Looks to see if the script has been used before!
			'for each local_changelog_item in local_changelog_array
			for i = 0 to ubound(local_changelog_array)
				If local_changelog_array(i) <> "" then 'some are likely blank
					'splits the local_changelog_array(i) into an array: 0 -> date, 1 -> name_of_script, 2 -> text_of_change
					local_changelog_item_array = split(local_changelog_array(i), " | ")

					'Looking to see if the script is in fact in the local changelog list. If it is, we will then check the text against the listed changes to see what needs to be displayed.
					if local_changelog_item_array(1) = name_of_script then
						script_in_local_changelog = true
						if local_changelog_item_array(2) <> last_item_added_to_changelog(1) then
							display_changelog = true
							local_changelog_text_of_change = trim(local_changelog_item_array(2))
							line_in_local_changelog_array_to_delete = i
						Else
							display_changelog = false
						End if
					End if
				End if
			next

			'Close the file
			objTextStream.Close

			'If the script is not in the local changelog, it needs to be added. If this is the case, it shouldn't display the changelog at all, because it'll be the first time the script was run.
			If script_in_local_changelog <> true then

				'Setting the object to open the text file for appending the new data
				Set objTextStream = .OpenTextFile(local_changelog_path, ForAppending, true)

				'Write the contents of the text file
				objTextStream.WriteLine date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Close the file and clean up
				objTextStream.Close

				'Setting this to false. We don't want to display the changelog if the script has never been added to the local list of changelog events
				display_changelog = false

			End if

			'So, if the script IS in the local changelog, and needs to be displayed, it takes special handling to ensure that's done.
			If display_changelog = true then

				'Splitting the changelog into different variables for making things prettier
				For each changelog_entry in changelog
					date_of_change = left(changelog_entry, instr(changelog_entry, " | ") - 1)
					scriptwriter_of_change = trim(right(changelog_entry, len(changelog_entry) - instrrev(changelog_entry, "|") ))
					text_of_change = replace(replace(replace(changelog_entry, scriptwriter_of_change, ""), date_of_change, ""), " | ", "")

					'If the text_of_change is the same as that stored in the local changelog, that means the user is up-to-date to this point, and the script should exit without displaying any more updates. Otherwise, add it to the contents.
					if trim(text_of_change) = trim(local_changelog_text_of_change) then
					 	exit for
					else
                        text_of_change = replace(text_of_change, "##~##", vbCR)
                        If name_of_script = "Functions Library" Then
                            changelog_msgbox = changelog_msgbox & "-----" & cdate(date_of_change) & "-----" & vbNewLine & text_of_change & vbNewLine & vbNewLine & "Thank you!" & vbNewLine & "The BlueZone Script Team" & vbNewLine & vbNewLine
                        Else
                            changelog_msgbox = changelog_msgbox & "-----" & cdate(date_of_change) & "-----" & vbNewLine & text_of_change & vbNewLine & "Completed by " & scriptwriter_of_change & vbNewLine & vbNewLine
                        End If
					end if

				Next

				If changelog_msgbox <> "" then
                    If name_of_script = "Functions Library" Then
                        message_of_change = MsgBox("Script Announcement: " & vbNewLine & vbNewLine & changelog_msgbox, vbSystemModal, "BZST Communication")
                        'MsgBox "Script Announcement: " & vbNewLine & vbNewLine & changelog_msgbox
                    Else
                        message_of_change = MsgBox("Recent changes in this script: " & vbNewLine & vbNewLine & changelog_msgbox, vbSystemModal, "BZST Changes to Script")
                        'MsgBox "Recent changes in this script: " & vbNewLine & vbNewLine & changelog_msgbox
                    End If
				End if

				'Now we need to determine what the most recent change is, in order to add this to our text file
				string_to_enter_into_local_changelog = date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Lastly, if it displayed a changelog, it should go through and update the record to remove the old entry and replace it with this one.
				Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)						'Opening the file one last time
				for i = 0 to ubound(local_changelog_array)
					If i = line_in_local_changelog_array_to_delete then local_changelog_array(i) = string_to_enter_into_local_changelog
					if local_changelog_array(i) <> "" then objTextStream.WriteLine local_changelog_array(i)
				next

			end if

			'Close the file
			objTextStream.Close
		End with
	End If

end function

function changelog_update(date_of_change, text_of_change, scriptwriter_of_change)
'--- This function adds the change to the scripts to the user change log to be displayed in function changelog_display()
'~~~~~ date_of_change: date the change was made/committed to the script file. Surround date in ""
'~~~~~ text_of_change: information about the change to the script that users statewide will see. Please be clear about your updates. You can write several sentences. Surround text in "".
'~~~~~ scriptwriter_of_change: scriptwriter name and county seperated by a comma. Surround name and county name with "".
'===== Keywords: MAXIS, PRISM, change, info, information
	If changelog_enabled = "" Then changelog_enabled = true
	If changelog_enabled <> false Then
		ReDim Preserve changelog(UBound(changelog) + 1)
		changelog(ubound(changelog)) = date_of_change & " | " & text_of_change & " | " & scriptwriter_of_change
	End If
end function

function check_for_MAXIS(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
                BeginDialog Password_dialog, 0, 0, 156, 55, "Password Dialog"
                ButtonGroup ButtonPressed
                OkButton 45, 35, 50, 15
                CancelButton 100, 35, 50, 15
                Text 5, 5, 150, 25, "You have passworded out. Please enter your password, then press OK to continue. Press CANCEL to stop the script. "
                EndDialog
                Do
                    Do
                        dialog Password_dialog
                        cancel_confirmation
                    Loop until ButtonPressed = -1
                    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
                Loop until are_we_passworded_out = false					'loops until user passwords back in
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

function check_for_MMIS(end_script)
'--- This function checks to ensure the user is in a MMIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MMIS screen.
'===== Keywords: MMIS, production, script_end_procedure
	Do
		transmit
		row = 1
		col = 1
		EMSearch "MMIS", row, col
		IF row <> 1 then
			If end_script = True then
				script_end_procedure("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again.")
			Else
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 216, 55, "MMIS Dialog"
                ButtonGroup ButtonPressed
                OkButton 125, 35, 40, 15
                CancelButton 170, 35, 40, 15
                Text 5, 5, 210, 25, "You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press CANCEL to exit the script."
                EndDialog
                Do
                    Dialog Dialog1
                    cancel_without_confirmation
                Loop until ButtonPressed = -1
			End if
		End if
	Loop until row = 1
end function

function check_for_password(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Transmit 'transmitting to see if the password screen appears
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function check_for_password_without_transmit(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function clear_line_of_text(row, start_column)
'--- This function clears out a single line of text
'~~~~~ row: coordinate of row to clear
'~~~~~ start_column: coordinate of column to start clearing
'===== Keywords: MAXIS, PRISM, production, clear
  EMSetCursor row, start_column
  EMSendKey "<EraseEof>"
  EMWaitReady 0, 0
end function

function confirm_docs_accepted_in_ecf(closing_msg)
'--- This function asks the worker if they have accepted in ECF the documents processed while using a script.
'~~~~~ closing_msg: the end message for display at the script end
'===== Keywords: MAXIS, ECF, statistics, reminder, script_end_procedure
    'This function will be called in script_end_procedure and script_end_procedure_with_error_report so that in can apply to any script with the script_that_handles_documents set to TRUE
    If script_that_handles_documents = TRUE Then                'This variable should be defined in the the script to identify if the script should use this functionality
        confirm_ecf_updated = MsgBox("Since this script notes the processing of documents, this is a reminder to correctly accept the documents in ECF." & vbNewLine & vbNewLine & "As a part of holistic processing, we want to be sure all case and administrative actions are completed. PRO TIP - accepting documents in ECF is a part of the activity report and show work completed." & vbNewLine & vbNewLine & "Did you remember to accept your documents?" & vbNewLine & "If you didn't, do it NOW and still press 'yes'.", vbQuestion + vbSystemModal + vbYesNo, "Please Accept Documents in ECF")         'This is the the msessage shown to the script user asking if they have accepted the documents in ECF

        'The function then updates the end display message used in script_end_procedure to add the information to the message to the worker and to add the information to the statistics.
        If confirm_ecf_updated = vbNo Then closing_msg = closing_msg & vbNewLine & "Documents were not accepted in ECF - please go accept documents you have worked on."
        If confirm_ecf_updated = vbYes Then closing_msg = closing_msg & vbNewLine & "Thank you for accepting your documents in ECF."
    End If
end function

function confirm_tester_information()
'--- Ask a tester to confirm the details we have for them. THIS FUNCTION IS CALLED IN THE FUNCTIONS LIBRARY
'===== Keywords: Testing, Infrastucture

    'Now the script will look to see if this user is a tester that needs their information confirmed.
    Leave_Confirmation = FALSE                              'This will allow the user to cancel the update if they desire
    For each tester in tester_array                         'looping through all of the testers
        If user_ID_for_validation = tester.tester_id_number Then            'If the person who is running the script is a tester
            If tester.tester_confirmed = FALSE Then                         'If the information is not saved as confirmed.
                'Script user is asked if they can confirm their information now.
                confirm_testing_now = MsgBox("Hello " & tester.tester_first_name & "! Thank you for agreeing to test BlueZone Scripts. You are an invaluable part of our process and the development of new tools and scripts." & vbNewLine & vbNewLine & "To be sure the testing functionality works correctly, we need to be sure we have the correct information about you. We need you to confrim a few details, this will take less than 5 minutes." & vbNewLine & vbNewLine & "Do you have time to confirm your information now?", vbQuestion + vbYesNo, "Confirm Tester Detail")

                If confirm_testing_now = vbYes Then     'If they select 'Yes' the script will run the dialogs to confirm information
                    show_initial_dialog = TRUE          'this is set to show the initial dialog because there are 2 dialogs that loop together and once we pass the first, we don't want to see it again
					for each user_prog in tester.tester_programs
						If user_prog = "SNAP" Then snap_checkbox = checked
						If user_prog = "GA" Then ga_checkbox = checked
						If user_prog = "MSA" Then msa_checkbox = checked
						If user_prog = "MFIP" Then mfip_checkbox = checked
						If user_prog = "DWP" Then dwp_checkbox = checked
						If user_prog = "GRH" Then grh_checkbox = checked
						If user_prog = "IMD" Then imd_checkbox = checked
						If user_prog = "MA" Then ma_checkbox = checked
						If user_prog = "MA-EPD" Then ma_epd_checkbox = checked
						If user_prog = "LTC" Then ltc_checkbox = checked
						If user_prog = "EA" Then ea_checkbox = checked
						If user_prog = "EGA" Then ega_checkbox = checked
						If user_prog = "LTH" Then lth_checkbox = checked
					next
					for each user_grp in tester.tester_groups
						If user_grp = "PSS" Then pss_checkbox = checked
						If user_grp = "QI" Then qi_checkbox = checked
						If user_grp = "YET" Then yet_checkbox = checked
						If user_grp = "AVS" Then avs_checkbox = checked
						If user_grp = "Sanctions" Then sanc_checkbox = checked
					next
                    Do
                        Do
                            err_msg = ""
                            Do
                                If show_initial_dialog = TRUE Then              'If we haven't seen the initial dialog yet
                                    err_msg = ""                                'blanking the error message'
                                    update_information = FALSE                  'defaulting this to false as this indicates if a change is needed.

                                    'If any of these properties are blank, we need to default them to needing an update
                                    If tester.tester_first_name = "" Then first_name_action = "Incorrect - Change"
                                    If tester.tester_last_name = "" Then last_name_action = "Incorrect - Change"
                                    If tester.tester_email = "" Then email_action = "Incorrect - Change"
                                    If tester.tester_id_number = "" Then id_number_action = "Incorrect - Change"
                                    If tester.tester_x_number = "" Then x_number_action = "Incorrect - Change"
                                    If tester.tester_supervisor_name = "" Then supervisor_action = "Incorrect - Change"
                                    If tester.tester_population = "" Then population_action = "Incorrect - Change"
                                    If tester.tester_region = "" Then region_action = "Incorrect - Change"

                                    the_dialog = ""        'reset for safety

									BeginDialog the_dialog, 0, 0, 371, 260, "Detailed Tester Information"      'first dialog just lists the properties
									  ButtonGroup ButtonPressed
									    OkButton 315, 240, 50, 15
									  Text 60, 15, 40, 10, "First Name:"
									  Text 60, 35, 40, 10, "Last Name:"
									  Text 50, 55, 50, 10, "Email Address:"
									  Text 10, 75, 90, 10, "Hennepin County ID (WF#):"
									  Text 40, 95, 60, 10, "MAXIS X-Number:"
									  Text 40, 115, 60, 10, "Supervisor Name:"
									  Text 40, 135, 60, 10, "Population/Team:"
									  Text 110, 15, 105, 10, tester.tester_first_name
									  Text 110, 35, 105, 10, tester.tester_last_name
									  Text 110, 55, 140, 10, tester.tester_email
									  Text 110, 75, 60, 10, tester.tester_id_number
									  Text 110, 95, 60, 10, tester.tester_x_number
									  Text 110, 115, 150, 10, tester.tester_supervisor_name
									  Text 110, 135, 60, 10, tester.tester_population
									  DropListBox 285, 10, 80, 45, "Correct"+chr(9)+"Incorrect - Change", first_name_action
									  DropListBox 285, 30, 80, 45, "Correct"+chr(9)+"Incorrect - Change", last_name_action
									  DropListBox 285, 50, 80, 45, "Correct"+chr(9)+"Incorrect - Change", email_action
									  DropListBox 285, 70, 80, 45, "Correct"+chr(9)+"Incorrect - Change", id_number_action
									  DropListBox 285, 90, 80, 45, "Correct"+chr(9)+"Incorrect - Change", x_number_action
									  DropListBox 285, 110, 80, 45, "Correct"+chr(9)+"Incorrect - Change", supervisor_action
									  DropListBox 285, 130, 80, 45, "Correct"+chr(9)+"Incorrect - Change", population_action
									  Text 10, 240, 130, 15, "Please reach out to the BlueZone Script team with any questions."
									  CheckBox 110, 160, 30, 10, "SNAP", snap_checkbox
									  CheckBox 145, 160, 25, 10, "GA", ga_checkbox
									  CheckBox 145, 170, 25, 10, "MSA", msa_checkbox
									  CheckBox 175, 160, 30, 10, "MFIP", mfip_checkbox
									  CheckBox 175, 170, 30, 10, "DWP", dwp_checkbox
									  CheckBox 210, 160, 25, 10, "GRH", grh_checkbox
									  CheckBox 210, 170, 25, 10, "IMD", imd_checkbox
									  CheckBox 240, 160, 25, 10, "MA", ma_checkbox
									  CheckBox 265, 160, 40, 10, "MA-EPD", ma_epd_checkbox
									  CheckBox 265, 170, 25, 10, "LTC", ltc_checkbox
									  CheckBox 305, 160, 25, 10, "EA", ea_checkbox
									  CheckBox 305, 170, 25, 10, "EGA", ega_checkbox
									  CheckBox 335, 160, 25, 10, "LTH", lth_checkbox
									  CheckBox 110, 200, 25, 10, "PSS", pss_checkbox
									  CheckBox 145, 200, 25, 10, "QI", qi_checkbox
									  CheckBox 175, 200, 25, 10, "YET", yet_checkbox
									  CheckBox 210, 200, 50, 10, "AVS Access", avs_checkbox
									  CheckBox 265, 200, 45, 10, "Sanctions", sanc_checkbox
									  Text 110, 220, 115, 10, "List any other processing groups:"
									  EditBox 225, 215, 135, 15, other_groups_reported
									  GroupBox 40, 150, 325, 35, "Programs"
									  GroupBox 40, 190, 325, 45, "Groups"
									EndDialog

                                    Dialog the_dialog          'showing the dialog
                                    If ButtonPressed = 0 Then       'cancelling the confirmation functionality without cancelling the script run
                                        MsgBox "We will cancel your confirmation of information. The script you selected will continue at this time. Please confirm your information at a future time."
                                        Leave_Confirmation = TRUE
                                        Exit Do
                                    End If


                                    'These properties MUST be filled in and if they are blank, we need to know what they are - mandatory fields
                                    If tester.tester_first_name = "" AND first_name_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since FIRST NAME is blank, this information must be updated. Select 'Incorrect - Change' for First Name."
                                    If tester.tester_last_name = "" AND last_name_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since LAST NAME is blank, this information must be updated. Select 'Incorrect - Change' for Last Name."
                                    If tester.tester_email = "" AND email_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since EMAIL is blank, this information must be updated. Select 'Incorrect - Change' for Email."
                                    If tester.tester_id_number = "" AND id_number_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since HENNEPIN ID (WF#) is blank, this information must be updated. Select 'Incorrect - Change' for WF-Number."
                                    If tester.tester_x_number = "" AND x_number_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since MAXIS ID  (X-NUMBER) is blank, this information must be updated. Select 'Incorrect - Change' for X-Number."
                                    If tester.tester_supervisor_name = "" AND supervisor_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since SUPERVISOR is blank, this information must be updated. Select 'Incorrect - Change' for Supervisor."
                                    If tester.tester_population = "" AND population_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since POPULATION is blank, this information must be updated. Select 'Incorrect - Change' for Population."
                                    ' If tester.tester_region = "" AND region_action <> "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since  is blank, this information must be updated. Select 'Incorrect - Change' for ."

                                    If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine &  err_msg       'showing the error message

                                End If
                            Loop until err_msg = ""

                            If Leave_Confirmation = TRUE Then Exit Do   'If the user pressed cancel before, this leaves this loop so we don't see another dialog
                            If first_name_action = "Incorrect - Change" Then update_information = TRUE          'If anything was selected to change, we need to run the update dialog
                            If last_name_action = "Incorrect - Change" Then update_information = TRUE
                            If email_action = "Incorrect - Change" Then update_information = TRUE
                            If id_number_action = "Incorrect - Change" Then update_information = TRUE
                            If x_number_action = "Incorrect - Change" Then update_information = TRUE
                            If supervisor_action = "Incorrect - Change" Then update_information = TRUE
                            If population_action = "Incorrect - Change" Then update_information = TRUE
                            If region_action = "Incorrect - Change" Then update_information = TRUE
                            If groups_action = "Incorrect - Change" Then update_information = TRUE
                            err_msg = ""

                            show_initial_dialog = FALSE             'setting this to false so if we loop we don't see the first dialog again

                            the_dialog = ""        'resetting for safetly
                            If update_information = TRUE Then                   'If a change was indicated in dialog 1, we show this  new dialog with the update fields
                                BeginDialog the_dialog, 0, 0, 371, 260, "Detailed Tester Information"
                                  ButtonGroup ButtonPressed
                                    OkButton 315, 240, 50, 15
                                  Text 60, 15, 40, 10, "First Name:"
                                  Text 60, 35, 40, 10, "Last Name:"
                                  Text 50, 55, 50, 10, "Email Address:"
                                  Text 10, 75, 90, 10, "Hennepin County ID (WF#):"
                                  Text 40, 95, 60, 10, "MAXIS X-Number:"
                                  Text 40, 115, 60, 10, "Supervisor Name:"
                                  Text 40, 135, 60, 10, "Population/Team:"

                                  Text 10, 240, 130, 15, "Please reach out to the BlueZone Script team with any questions."
                                  If first_name_action = "Incorrect - Change" Then
                                    new_first_name = tester.tester_first_name
                                    EditBox 110, 10, 105, 15, new_first_name
                                  Else
                                    Text 110, 15, 105, 10, tester.tester_first_name
                                  End If
                                  If last_name_action = "Incorrect - Change" Then
                                    new_last_name = tester.tester_last_name
                                    EditBox 110, 30, 105, 15, new_last_name
                                  Else
                                    Text 110, 35, 105, 10, tester.tester_last_name
                                  End If
                                  If email_action = "Incorrect - Change" Then
                                    new_email = tester.tester_email
                                    EditBox 110, 50, 140, 15, new_email
                                  Else
                                    Text 110, 55, 140, 10, tester.tester_email
                                  End If
                                  If id_number_action = "Incorrect - Change" Then
                                    new_id_number = tester.tester_id_number
                                    EditBox 110, 70, 60, 15, new_id_number
                                  Else
                                    Text 110, 75, 60, 10, tester.tester_id_number
                                  End If
                                  If x_number_action = "Incorrect - Change" Then
                                    new_x_number = tester.tester_x_number
                                    EditBox 110, 90, 60, 15, new_x_number
                                  Else
                                    Text 110, 95, 60, 10, tester.tester_x_number
                                  End If
                                  If supervisor_action = "Incorrect - Change" Then
                                    new_supervisor_name = tester.tester_supervisor_name
                                    EditBox 110, 110, 150, 15, new_supervisor_name
                                  Else
                                    Text 110, 115, 150, 10, tester.tester_supervisor_name
                                  End If
                                  If population_action = "Incorrect - Change" Then
                                    new_population = tester.tester_population
                                    EditBox 110, 130, 60, 15, new_population
                                  Else
                                    Text 110, 135, 60, 10, tester.tester_population
                                  End If
								  CheckBox 110, 160, 30, 10, "SNAP", snap_checkbox
								  CheckBox 145, 160, 25, 10, "GA", ga_checkbox
								  CheckBox 145, 170, 25, 10, "MSA", msa_checkbox
								  CheckBox 175, 160, 30, 10, "MFIP", mfip_checkbox
								  CheckBox 175, 170, 30, 10, "DWP", dwp_checkbox
								  CheckBox 210, 160, 25, 10, "GRH", grh_checkbox
								  CheckBox 210, 170, 25, 10, "IMD", imd_checkbox
								  CheckBox 240, 160, 25, 10, "MA", ma_checkbox
								  CheckBox 265, 160, 40, 10, "MA-EPD", ma_epd_checkbox
								  CheckBox 265, 170, 25, 10, "LTC", ltc_checkbox
								  CheckBox 305, 160, 25, 10, "EA", ea_checkbox
								  CheckBox 305, 170, 25, 10, "EGA", ega_checkbox
								  CheckBox 335, 160, 25, 10, "LTH", lth_checkbox
								  CheckBox 110, 200, 25, 10, "PSS", pss_checkbox
								  CheckBox 145, 200, 25, 10, "QI", qi_checkbox
								  CheckBox 175, 200, 25, 10, "YET", yet_checkbox
								  CheckBox 210, 200, 50, 10, "AVS Access", avs_checkbox
								  CheckBox 265, 200, 45, 10, "Sanctions", sanc_checkbox
								  Text 110, 220, 115, 10, "List any other processing groups:"
								  EditBox 225, 215, 135, 15, other_groups_reported
								  GroupBox 40, 150, 325, 35, "Programs"
								  GroupBox 40, 190, 325, 45, "Groups"
                                EndDialog

                                Dialog the_dialog
                                If ButtonPressed = 0 Then                       'If user presses cancel this will cancel the functionality but not the script run.
                                    MsgBox "We will cancel your confirmation of information. The script you selected will continue at this time. Please confirm your information at a future time."
                                    Leave_Confirmation = TRUE
                                    Exit Do
                                End If

                                'Mandating these properties to be completed.
                                If new_first_name = "" AND first_name_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* FIRST NAME is blank, this information is required for testers. Update the detail for First Name."
                                If new_last_name = "" AND last_name_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* LAST NAME is blank, this information is required for testers. Update the detail for Last Name."
                                If new_email = "" AND email_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* EMAIL is blank, this information is required for testers. Update the detail for Email."
                                If new_id_number = "" AND id_number_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* HENNEPIN ID (WF#) is blank, this information is required for testers. Update the detail for WF-Number."
                                If new_x_number = "" AND x_number_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* MAXIS ID  (X-NUMBER) is blank, this information is required for testers. Update the detail for X-Number."
                                If new_supervisor_name = "" AND supervisor_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* SUPERVISOR is blank, this information is required for testers. Update the detail for Supervisor."
                                If new_population = "" AND population_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* POPULATION is blank, this information is required for testers. Update the detail for Population."
								If snap_checkbox = unchecked AND ga_checkbox = unchecked AND msa_checkbox = unchecked AND mfip_checkbox = unchecked AND dwp_checkbox = unchecked AND grh_checkbox = unchecked AND imd_checkbox = unchecked AND ma_checkbox = unchecked AND ma_epd_checkbox = unchecked AND ltc_checkbox = unchecked AND ea_checkbox = unchecked AND ega_checkbox = unchecked AND lth_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* You must select at least one program that you work in."
                                ' If new_region = "" AND region_action = "Incorrect - Change" Then err_msg = err_msg & vbNewLine & "* Since  is blank, this information is required for testers. Update the detail for ."

                                If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine &  err_msg

                            End If
                        Loop until err_msg = ""
                        If Leave_Confirmation = TRUE Then Exit Do
                        Call check_for_password(are_we_passworded_out)
                    Loop until are_we_passworded_out = FALSE
                    If Leave_Confirmation = TRUE Then Exit For

                    Message_Information = "Testing Confirmation Email"          'Setting the email information/text

                    If update_information = TRUE Then
                        Message_Information = Message_Information & vbNewLine & "*** Updates to Information made ***" & vbNewLine
                        If first_name_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New First Name: " & new_first_name & " - Change From: " & tester.tester_first_name
                        If last_name_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Last Name: " & new_last_name & " - Change From: " & tester.tester_last_name
                        If email_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Email: " & new_email & " - Change From: " & tester.tester_email
                        If id_number_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New ID Number: " & new_id_number & " - Change From: " & tester.tester_id_number
                        If x_number_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New X Number: " & new_x_number & " - Change From: " & tester.tester_x_number
                        If supervisor_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Supervisor: " & new_supervisor_name & " - Change From: " & tester.tester_supervisor_name
                        If population_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Population: " & new_population & " - Change From: " & tester.tester_population
                        If region_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Region: " & new_region & " - Change From: " & tester.tester_region
                        If groups_action = "Incorrect - Change" Then Message_Information = Message_Information & vbNewLine & "New Groups: " & new_groups & " - Change From: " & join(tester.tester_groups, ",")
                    Else
                        Message_Information = Message_Information & vbNewLine & "No updates indicated - tester information for " & tester.tester_full_name & " is correct and can be updated as confirmed."
                    End If
					Message_Information = Message_Information & vbNewLine & vbNewLine & "PROGRAMS:"

					Message_Information = Message_Information & vbNewLine & "SNAP - "
					If snap_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "GA - "
					If ga_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "MSA - "
					If msa_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "MFIP - "
					If mfip_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "DWP - "
					If dwp_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "GRH - "
					If grh_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "IMD - "
					If imd_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "MA - "
					If ma_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "MA-EPD - "
					If ma_epd_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "LTC - "
					If ltc_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "EA - "
					If ea_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "EGA - "
					If ega_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "LTH - "
					If lth_checkbox = checked Then Message_Information = Message_Information & "YES"

					Message_Information = Message_Information & vbNewLine & vbNewLine & "GROUPS:"

					Message_Information = Message_Information & vbNewLine & "PSS - "
					If pss_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "QI - "
					If qi_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "YET - "
					If yet_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "AVS - "
					If avs_checkbox = checked Then Message_Information = Message_Information & "YES"
					Message_Information = Message_Information & vbNewLine & "SANCTION - "
					If sanc_checkbox = checked Then Message_Information = Message_Information & "YES"
					If Trim(other_groups_reported) <> "" Then Message_Information = Message_Information & vbNewLine & "Other Groups: " & other_groups_reported

                    Message_Information = Message_Information & vbNewLine & vbNewLine & "Thank you for taking the time to confirm your information. We will update your information in the script records and you will no longer be asked to confirm your information. (We have to do this manually so there may be a delay but we will update as soon as we are able.)" & vbNewLine
                    tester.send_tester_email FALSE, Message_Information         ''sending the email
                End If
            End If
        End If
    Next
end function

function convert_array_to_droplist_items(array_to_convert, output_droplist_box)
'--- This function converts an array into a droplist to be used within dialog
'~~~~~ array_to_convert: name of the array
'~~~~~ output_droplist_box: name of droplist variant/variable
'===== Keywords: MAXIS, PRISM, production, array, droplist
	For each item in array_to_convert
		If output_droplist_box = "" then
			output_droplist_box = item
		Else
			output_droplist_box = output_droplist_box & chr(9) & item
		End if
	Next
end function

function convert_date_into_MAXIS_footer_month(date_to_convert, footer_month_input, footer_year_input)
'--- This function converts a date (MM/DD/YY or MM/DD/YYYY format) into a separate footer month and footer year variables.
'~~~~~ date_to_convert: variable name of date you want to convert
'~~~~~ footer_month_input: footer month to convert the date into
'~~~~~ footer_year_input: footer year to convert the date into
'===== Keywords: MAXIS, production, array, droplist, convert
	footer_month_input = DatePart("m", date_to_convert)										'Uses DatePart function to copy the month from date_to_convert into the MAXIS_footer_month variable.
	footer_month_input = Right("0" & footer_month_input, 2)		                                'Uses Len function to determine if the MAXIS_footer_month is a single digit month. If so, it adds a 0, which MAXIS needs.
	footer_year_input = DatePart("yyyy", date_to_convert)									'Uses DatePart function to copy the year from date_to_convert into the MAXIS_footer_year variable.
	footer_year_input = Right(footer_year_input, 2)											'Uses Right function to reduce the MAXIS_footer_year variable to it's right 2 characters (allowing for a 2 digit footer year).
end function

function convert_digit_to_excel_column(col_in_excel)
'--- This function converts a numeric digit to an Excel column, up to 104 digits (columns).
'~~~~~ col_in_excel: must be a numeric, cannot exceed 104. Do not put in "".
'===== Keywords: MAXIS, PRISM, convert, Excel
	'Create string with the alphabet

	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	'Assigning a letter, based on that column. Uses "mid" function to determine it. If number > 26, it handles by adding a letter (per Excel).
	convert_digit_to_excel_column = Mid(alphabet, col_in_excel, 1)
	If col_in_excel >= 27 and col_in_excel < 53 then convert_digit_to_excel_column = "A" & Mid(alphabet, col_in_excel - 26, 1)
	If col_in_excel >= 53 and col_in_excel < 79 then convert_digit_to_excel_column = "B" & Mid(alphabet, col_in_excel - 52, 1)
	If col_in_excel >= 79 and col_in_excel < 105 then convert_digit_to_excel_column = "C" & Mid(alphabet, col_in_excel - 78, 1)
    If col_in_excel >= 105 and col_in_excel < 131 then convert_digit_to_excel_column = "D" & Mid(alphabet, col_in_excel - 104, 1)
    If col_in_excel >= 131 and col_in_excel < 157 then convert_digit_to_excel_column = "E" & Mid(alphabet, col_in_excel - 130, 1)
    If col_in_excel >= 157 and col_in_excel < 183 then convert_digit_to_excel_column = "F" & Mid(alphabet, col_in_excel - 156, 1)
    If col_in_excel >= 183 and col_in_excel < 209 then convert_digit_to_excel_column = "G" & Mid(alphabet, col_in_excel - 182, 1)
    If col_in_excel >= 209 and col_in_excel < 235 then convert_digit_to_excel_column = "H" & Mid(alphabet, col_in_excel - 208, 1)
	'Closes script if the number gets too high (very rare circumstance, just errorproofing)
	If col_in_excel >= 235 then script_end_procedure("This script is only able to assign excel columns to 234 distinct digits. You've exceeded this number, and this script cannot continue.")
end function

Function create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
'--- This function is used to grab all active X numbers according to the supervisor X number(s) inputted
'~~~~~ array_name: name of array that will contain all the supervisor's staff x numbers
'~~~~~ supervisor_array: list of supervisor's x numbers seperated by comma
'===== Keywords: MAXIS, array, supervisor, worker number, create
	CALL navigate_to_MAXIS_screen("REPT", "USER")  	'Getting to REPT/USER
	'Sorting by supervisor
	PF5
	PF5
	'Reseting array_name
	array_name = ""
	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)
			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
End Function

function create_array_of_all_active_x_numbers_in_county(array_name, county_code)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	call navigate_to_MAXIS_screen("REPT", "USER")

	'Hitting PF5 to force sorting, which allows directly selecting a county
	PF5

	'Inserting county
	EMWriteScreen county_code, 21, 6
	transmit

	'Declaring the MAXIS row
	MAXIS_row = 7

	'Blanking out array_name in case this has been used already in the script
	array_name = ""

	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
			array_name = trim(array_name & " " & worker_ID)				'writing to variable
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then
			PF8			'getting to next screen
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank
	array_name = split(array_name)
end function

function create_mainframe_friendly_date(date_variable, screen_row, screen_col, year_type)
'--- This function creates a mainframe friendly date. This can be used for both year formats and input spacing.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'~~~~~ year_type: formatting to export date year as "YY" or "YYYY"
'===== Keywords: MAXIS, PRISM, MMIS, date, create
	var_month = datepart("m", date_variable)
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month & "/", screen_row, screen_col
	var_day = datepart("d", date_variable)
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day & "/", screen_row, screen_col + 3
	If year_type = "YY" then
		var_year = right(datepart("yyyy", date_variable), 2)
	ElseIf year_type = "YYYY" then
		var_year = datepart("yyyy", date_variable)
	Else
		MsgBox "Year type entered incorrectly. Fourth parameter of function create_mainframe_friendly_date should read ""YYYY"" or ""YY"". The script will now stop."
		StopScript
	END IF
	EMWriteScreen var_year, screen_row, screen_col + 6
end function

function create_MAXIS_friendly_date(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM DD YY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	If len(var_month) = 1 then var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	If len(var_day) = 1 then var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
end function

function create_MAXIS_friendly_date_three_spaces_between(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM  DD  YY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))		'determines the date based on the variable length: month
	If len(var_month) = 1 then var_month = "0" & var_month				'adds a '0' in front of a single digit month
	EMWriteScreen var_month, screen_row, screen_col					'writes in var_month at coordinates set in function line
	var_day = datepart("d", dateadd("d", variable_length, date_variable)) 		'determines the date based on the variable length: day
	If len(var_day) = 1 then var_day = "0" & var_day 				'adds a '0' in front of a single digit day
	EMWriteScreen var_day, screen_row, screen_col + 5 				'writes in var_day at coordinates set in function line, and starts 5 columns into date field in MAXIS
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable)) 	'determines the date based on the variable length: year
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 10 			'writes in var_year at coordinates set in function line , and starts 5 columns into date field in MAXIS
end function

function create_MAXIS_friendly_date_with_YYYY(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM DD YYYY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen var_year, screen_row, screen_col + 6
end function

function create_MAXIS_friendly_phone_number(phone_number_variable, screen_row, screen_col)
'--- This function creates a MAXIS friendly phone number
'~~~~~ phone_number_variable: the name of the variable to output
'~~~~~ screen_row: row to start writing phone number
'~~~~~ screen_col: column to start writing phone number
'===== Keywords: MAXIS, date, create
	WITH (new RegExp)                                                            	'Uses RegExp to bring in special string functions to remove the unneeded strings
                .Global = True                                                   	'I don't know what this means but David made it work so we're going with it
                .Pattern = "\D"                                                	 	'Again, no clue. Just do it.
                phone_number_variable = .Replace(phone_number_variable, "")    	 	'This replaces the non-digits of the phone number with nothing. That leaves us with a bunch of numbers
	END WITH
	EMWriteScreen left(phone_number_variable, 3), screen_row, screen_col 		'writes in left 3 digits of the phone number in variable
	EMWriteScreen mid(phone_number_variable, 4, 3), screen_row, screen_col + 6	'writes in middle 3 digits of the phone number in variable
	EMWriteScreen right(phone_number_variable, 4), screen_row, screen_col + 12	'writes in right 4 digits of the phone number in variable
end function

FUNCTION create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
'--- This function creates a an outlook appointment
'~~~~~ (appt_date): date of the appointment
'~~~~~ (appt_start_time): start time of the appointment - format example: "08:00 AM"
'~~~~~ (appt_end_time): end time of the appointment - format example: "08:00 AM"
'~~~~~ (appt_subject): subject of the email in quotations or a variable
'~~~~~ (appt_body): body of the email in quotations or a variable
'~~~~~ (appt_location): name of location in quotations or a variable
'~~~~~ (appt_reminder): reminder for appointment. Set to TRUE or FALSE
'~~~~~ (reminder_in_minutes): enter the number of minutes prior to the appointment to set the reminder. Set as 0 if at the time of the appoint. Set to "" if appt_reminder is set to FALSE
'~~~~~ (appt_category): can be left "" or assgin to the set the name of the category in quotations
'===== Keywords: MAXIS, PRISM, create, outlook, appointment

	'Assigning needed numbers as variables for readability
	olAppointmentItem = 1
	olRecursDaily = 0

	'Creating an Outlook object item
	Set objOutlook = CreateObject("Outlook.Application")
	Set objAppointment = objOutlook.CreateItem(olAppointmentItem)

	'Assigning individual appointment options
	objAppointment.Start = appt_date & " " & appt_start_time		'Start date and time are carried over from parameters
	objAppointment.End = appt_date & " " & appt_end_time			'End date and time are carried over from parameters
	objAppointment.AllDayEvent = False 								'Defaulting to false for this. Perhaps someday this can be true. Who knows.
	objAppointment.Subject = appt_subject							'Defining the subject from parameters
	objAppointment.Body = appt_body									'Defining the body from parameters
	objAppointment.Location = appt_location							'Defining the location from parameters
	If appt_reminder = FALSE then									'If the reminder parameter is false, it skips the reminder, otherwise it sets it to match the number here.
		objAppointment.ReminderSet = False
	Else
		objAppointment.ReminderSet = True
		objAppointment.ReminderMinutesBeforeStart = reminder_in_minutes
	End if
	objAppointment.Categories = appt_category						'Defines a category
	objAppointment.Save												'Saves the appointment
END FUNCTION

Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'--- This function creates a an outlook appointment
'~~~~~ (email_recip): email address for recipeint - seperated by semicolon
'~~~~~ (email_recip_CC): email address for recipeints to cc - seperated by semicolon
'~~~~~ (email_subject): subject of email in quotations or a variable
'~~~~~ (email_body): body of email in quotations or a variable
'~~~~~ (email_attachment): set as "" if no email or file location
'~~~~~ (send_email): set as TRUE or FALSE
'===== Keywords: MAXIS, PRISM, create, outlook, email

	'Setting up the Outlook application
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    If send_email = False then objMail.Display      'To display message only if the script is NOT sending the email for the user.

    'Adds the information to the email
    objMail.to = email_recip                        'email recipient
    objMail.cc = email_recip_CC                     'cc recipient
    objMail.Subject = email_subject                 'email subject
    objMail.Body = email_body                       'email body
    If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
    'Sends email
    If send_email = true then objMail.Send	                   'Sends the email
    Set objMail =   Nothing
    Set objOutlook = Nothing
End Function

function create_panel_if_nonexistent()
'--- This function creates a panel if a panel does not exist. This is currently only used within the FuncLib itself.
'~~~~~ (): keep this parameter empty
'===== Keywords: MAXIS, FuncLib only, create
	EMWriteScreen reference_number , 20, 76
	transmit
	EMReadScreen case_panel_check, 44, 24, 2
	If case_panel_check = "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen quantity_of_screens, 1, 2, 78
		If quantity_of_screens <> "0" then
			PF9
		ElseIf quantity_of_screens = "0" then
			EMWriteScreen "__", 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	ElseIf case_panel_check <> "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen error_scan, 80, 24, 1
		error_scan = trim(error_scan)
		EMReadScreen quantity_of_screens, 1, 2, 78
		If error_scan = "" and quantity_of_screens <> "0" then
			PF9
		ElseIf error_scan <> "" and quantity_of_screens <> "0" then
			'FIX ERROR HERE
			msgbox("Error: " & error_scan)
		ElseIf error_scan <> "" and quantity_of_screens = "0" then
			EMWriteScreen reference_number, 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	End If
end function

Function create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    'All 10-day cutoff dates are provided in POLI/TEMP TE19.132
    '--- This function creates and saves a TIKL message in DAIL/WRIT.
    '~~~~~ TIKL_text: Text that the TIKL message will say.
    '~~~~~ num_of_days: how many days the TIKL should be set for. Must be a numeric or a numeric value.
    '~~~~~ date_to_start: this determines which date to start counting the number of days to TIKL out from. Ex: date to use today's date, or Application_date to use within the CAF.
    '~~~~~ ten_day_adjust: True or False. True to adjust the TIKL date to the 1st day of the next month if after 10 day cutoff, False to NOT adjust to 10 day cutoff.
    '~~~~~ TIKL_note_text: This varible is determnined by the TIKL date and is to be used in the case note. Leave as 'TIKL_note_text' in the parameter.
    '===== Keywords: MAXIS, TIKL

    adjusted_date = False
    TIKL_date = DateAdd("D", num_of_days, date_to_start)    'Creates the TIKL date based on the number of days and date to start chosen by the user
    If cdate(TIKL_date) < date then
        msgbox "Unable to create TIKL, the TIKL date is a past date. Please manually track this case and action."   'fail-safe in case the TIKL date created is in the past. DAIL/WRIN does not allow past dates.
    Else
        If ten_day_adjust = True then
            TIKL_mo = right("0" & DatePart("m",    TIKL_date), 2) 'Creating new month and year variables to determine which ten day cut off date to use
            TIKL_yr = right(      DatePart("yyyy", TIKL_date), 2)

            IF TIKL_mo = "01" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #01/19/2023#
            ELSEIF TIKL_mo = "02" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #02/16/2023#
            ELSEIF TIKL_mo = "03" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #03/21/2023#
            ELSEIF TIKL_mo = "04" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #04/20/2023#
            ELSEIF TIKL_mo = "05" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #05/19/2023#
            ELSEIF TIKL_mo = "06" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #06/20/2023#
            ELSEIF TIKL_mo = "07" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #07/20/2023#
            ELSEIF TIKL_mo = "08" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #08/21/2023#
            ELSEIF TIKL_mo = "09" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #09/20/2023#
            ELSEIF TIKL_mo = "10" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #10/19/2023#
            ELSEIF TIKL_mo = "11" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #11/20/2023#
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "23" THEN
                ten_day_cutoff = #12/20/2023#
            ELSEIF TIKL_mo = "11" AND TIKL_yr = "22" THEN
                ten_day_cutoff = #11/18/2022#
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "22" THEN
                ten_day_cutoff = #12/21/2022#                                'last month of current year
            ELSE
            	missing_date = True 'in case TIKL time spans exceed 10 day cut off calendar.
            END IF

            If missing_date = True then
                TIKL_date = TIKL_date 'defaults to the date set by the user
            Else
                'Determining the TIKL date based on if past 10 day cut off or not.
                If cdate(TIKL_date) > cdate(ten_day_cutoff) then
                    'Date of the 1st of the next month where negative action can be taken is determined & becomes the TIKL_date
                    new_TIKL_mo = right(DatePart("m",    DateAdd("m", 1, TIKL_date)), 2)
                    new_TIKL_yr = right(DatePart("yyyy", DateAdd("m", 1, TIKL_date)), 2)
                    TIKL_date = new_TIKL_mo & "/01/" & new_TIKL_yr
                    adjusted_date = True
                End if
            End if
        End if
        'Creating the TIKL message
        Call navigate_to_MAXIS_screen("DAIL", "WRIT")
        call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18)    '0 is the date as all the adjustments are already determined.
        Call write_variable_in_TIKL(TIKL_text)
        PF3 'to save & exit

        If adjusted_date = True then
            TIKL_note_text = "* TIKL created for " & TIKL_date & ", the 1st day negative action could occur."
        Else
            TIKL_note_text = "* TIKL created for " & TIKL_date & ", " & num_of_days & " day return."
        End if
    End if
End Function

function date_array_generator(initial_month, initial_year, date_array)
'--- This function creates a series of dates (Example: for each footer month/year through current month plus 1)
'~~~~~ initial_month: first footer month
'~~~~~ initial_year: first footer year
'~~~~~ date_array: the name of the array that holds the dates/number of months to create dates for
'===== Keywords: MAXIS, create, date, array
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date

	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
end function

function determine_130_percent_of_FPG(footer_month, footer_year, hh_size, fpg_130_percent)
'--- This function outputs the dollar amount (as a number) of 130% FPG based on HH Size as needed by SNAP
'~~~~~ footer_month: relevant footer month - the calculation changes every Ocotber and we need to ensure we are pulling the correct amount
'~~~~~ footer_year: relevant footer year - the calculation changes every Ocotber and we need to ensure we are pulling the correct amount
'~~~~~ hh_size: NUMBER - the number of people in the SNAP unit
'~~~~~ fpg_130_percent: NUMBER - this will output a number with the amount of 130%  FPG based on footer month and HH Size
'===== Keywords: SNAP, calculation, Income Test
	month_to_review = footer_month & "/1/" & footer_year		'making this a date
	month_to_review = DateAdd("d", 0, month_to_review)

	If IsNumeric(hh_size) = True Then							'error handling to ensure that HH size is a number
		hh_size = hh_size*1
		If DateDiff("d", #10/1/2022#, month_to_review) >= 0 Then				'these are the associated amounts
			If hh_size = 1 Then fpg_130_percent = 1473
			If hh_size = 2 Then fpg_130_percent = 1984
			If hh_size = 3 Then fpg_130_percent = 2495
			If hh_size = 4 Then fpg_130_percent = 3007
			If hh_size = 5 Then fpg_130_percent = 3518
			If hh_size = 6 Then fpg_130_percent = 4029
			If hh_size = 7 Then fpg_130_percent = 4541
			If hh_size = 8 Then fpg_130_percent = 5052

			If hh_size > 8 Then fpg_130_percent = 5052 + (512 * (hh_size-8))
		ElseIf DateDiff("d", #10/1/2021#, month_to_review) >= 0 Then
			If hh_size = 1 Then fpg_130_percent = 1396
			If hh_size = 2 Then fpg_130_percent = 1888
			If hh_size = 3 Then fpg_130_percent = 2379
			If hh_size = 4 Then fpg_130_percent = 2871
			If hh_size = 5 Then fpg_130_percent = 3363
			If hh_size = 6 Then fpg_130_percent = 3855
			If hh_size = 7 Then fpg_130_percent = 4347
			If hh_size = 8 Then fpg_130_percent = 4839

			If hh_size > 8 Then fpg_130_percent = 4839 + (492 * (hh_size-8))
		End If
	End If
end function

function determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
'--- Function used to return booleans on case and program status based on CASE CURR information. There is no input informat but MAXIS_case_number needs to be defined.
'~~~~~ case_active: Outputs BOOLEAN of if the case is active in any MAXIS program
'~~~~~ case_pending: Outputs BOOLEAN of if the case is pending for any MAXIS Program
'~~~~~ case_rein: Outputs BOOLEAN of if the case is in REIN for any MAXIS Program
'~~~~~ family_cash_case: Outputs BOOLEAN of if the case is active or pending for any family cash program (MFIP or DWP)
'~~~~~ mfip_case: Outputs BOOLEAN of if the case is active or pending MFIP
'~~~~~ dwp_case: Outputs BOOLEAN of if the case is active or pending DWP
'~~~~~ adult_cash_case: Outputs BOOLEAN of if the case is active or pending any adult cash program (GA or MSA)
'~~~~~ ga_case: Outputs BOOLEAN of if the case is active or pending GA
'~~~~~ msa_case: Outputs BOOLEAN of if the case is active or pending MSA
'~~~~~ grh_case: Outputs BOOLEAN of if the case is active or pending GRH
'~~~~~ snap_case: Outputs BOOLEAN of if the case is active or pending SNAP
'~~~~~ ma_case: Outputs BOOLEAN of if the case is active or pending MA
'~~~~~ msp_case: Outputs BOOLEAN of if the case is active or pending any MSP
'~~~~~ emer_case: Outputs BOOLEAN of if the case is active or pending any Emergency Assistance
'~~~~~ unknown_cash_pending: BOOLEAN of if the case has a general 'CASH' program pending but it has not been defined
'~~~~~ unknown_hc_pending: BOOLEAN of if the case has a general 'HC' program pending but it has not been defined
'~~~~~ ga_status: Outputs the program status for GA - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ msa_status: Outputs the program status for MSA - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ mfip_status: Outputs the program status for MFIP - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ dwp_status: Outputs the program status for DWP - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ grh_status: Outputs the program status for GRH - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ snap_status: Outputs the program status for SNAP - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ ma_status: Outputs the program status for MA - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ msp_status: Outputs the program status for MSP - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ emer_status: Outputs the program status for EA/EGA - will be one of these options (ACTIVE, APP OPEN, APP CLOSE, INACTIVE, PENDING, REIN)
'~~~~~ msp_type: string of which MSP is active (QMB, SLMB, QI1)
'~~~~~ emer_type: string of which Emergency Assistance program is pending/issued (EA/EGA)
'~~~~~ list_active_programs: string of all the programs that appear active, app open, or app close
'~~~~~ list_pending_programs: string of all the programs that appear pending
'===== Keywords: MAXIS, case status, output, status
	EMReadScreen on_DAIL_check, 4, 2, 48 										'Read the top of the page to see if we are on DAIL DAIL - this will have different navigation to get to CASE CURR if we are on the DAIL
	EMReadScreen no_pop_up_on_DAIL, 9, 4, 24 									'Ensuring that there isn't a DAIL open on the page - which would inhibit direct navigating
	EMGetCursor dail_row, dail_col												'Navigating using the DAIL requires that the cursor has been set on the correct DAIL message, this will check to be sure the cursor is set to what is likely to be a DAIL navigation line
	'The function will try to use commands from DAIL to navigate to CASE/CURR
	'It makes sure on_DAIL is 'DAIL' and the Selection option is visible
	'The command line function is always in column 3 and betweet row 6 and 18
	If on_DAIL_check = "DAIL" and no_pop_up_on_DAIL = "Selection" and dail_col = 3 and dail_row > 5 and dail_row < 19 Then
		'IF USING THIS FUNCTION IN A DAIL SCRUBBER SCRIPT - BE SURE TO ENTER A SetCursor BEFORE CALLING THE FUNCTION - OR OTHERWISE ENSURE THE CURSOR IS ON THE DAIL
		EMSendKey "H"		'H is the command for CASE/CURR
		TRANSMIT
	End If
	'Now navigate to CASE CURR - if we are already there (by using the DAIL) this won't move as the function now checks to see if we are already at the screen
	Call navigate_to_MAXIS_screen("CASE", "CURR")           				'First the function will navigate to CASE/CURR so the inofrmation discovered is based on current status

    family_cash_case = FALSE                                					'defaulting all of the booleans
    adult_cash_case = FALSE
    ga_case = FALSE
    msa_case = FALSE
    mfip_case = FALSE
    dwp_case = FALSE
    grh_case = FALSE
    snap_case = FALSE
    ma_case = FALSE
    msp_case = FALSE
	emer_case = False
    case_active = FALSE
    case_pending = FALSE
	case_rein = FALSE
    unknown_cash_pending = FALSE
	unknown_hc_pending = FALSE
	ga_status = "INACTIVE"
	msa_status = "INACTIVE"
	mfip_status = "INACTIVE"
	dwp_status = "INACTIVE"
	grh_status = "INACTIVE"
	snap_status = "INACTIVE"
	ma_status = "INACTIVE"
	msp_status = "INACTIVE"
	emer_status = "INACTIVE"
	case_status = "INACTIVE"
	list_active_programs = ""
	list_pending_programs = ""

    'The function will use the same functionality for each program and search CASE:CURR to find the program deader for detail about the status.
    'If 'ACTIVE', 'APP CLOSE', 'APP OPEN', or 'PENDING' is listed after the header the function will mark the boolean for that program as 'TRUE'
    'If 'ACTIVE', 'APP CLOSE', or 'APP OPEN' is listed, the function will mark case_active as TRUE
    'If 'PENDING' is listed, the function wil mark case_pending as TRUE
	row = 1                                                 'First we will look at the main case stats
    col = 1
    EMSearch "Case:", row, col
    If row <> 0 Then
        EMReadScreen case_status, 15, row, col + 6
        case_status = trim(case_status)
    End If
    row = 1                                                 'looking for SNAP information
    col = 1
    EMSearch "FS:", row, col
    If row <> 0 Then
        EMReadScreen fs_status, 9, row, col + 4
        fs_status = trim(fs_status)
		snap_status = fs_status
        If fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN" Then
            snap_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "SNAP, "
        End If
        If fs_status = "PENDING" Then
            snap_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "SNAP, "
        End If
		If left(fs_status, 4) = "REIN" Then
			snap_case = TRUE
			case_rein = TRUE
		End If
    End If
    row = 1                                             'Looking for GRH information
    col = 1
    EMSearch "GRH:", row, col
    If row <> 0 Then
        EMReadScreen grh_status, 9, row, col + 5
        grh_status = trim(grh_status)
        If grh_status = "ACTIVE" or grh_status = "APP CLOSE" or grh_status = "APP OPEN" Then
            grh_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "GRH, "
        End If
        If grh_status = "PENDING" Then
            grh_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "GRH, "
        ENd If
		If left(grh_status, 4) = "REIN" Then
			grh_case = TRUE
			case_rein = TRUE
		End If
	End If
    row = 1                                             'Looking for MSA information
    col = 1
    EMSearch "MSA:", row, col
    If row <> 0 Then
        EMReadScreen ms_status, 9, row, col + 5
        ms_status = trim(ms_status)
		msa_status = ms_status
        If ms_status = "ACTIVE" or ms_status = "APP CLOSE" or ms_status = "APP OPEN" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "MSA, "
        End If
        If ms_status = "PENDING" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "MSA, "
        ENd If
		If left(ms_status, 4) = "REIN" Then
			msa_case = TRUE
			adult_cash_case = TRUE
			case_rein = TRUE
		End If
	End If
    row = 1                                             'Looking for GA information
    col = 1
    EMSearch " GA:", row, col
    If row <> 0 Then
        EMReadScreen ga_status, 9, row, col + 5
        ga_status = trim(ga_status)
        If ga_status = "ACTIVE" or ga_status = "APP CLOSE" or ga_status = "APP OPEN" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "GA, "
        End If
        If ga_status = "PENDING" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "GA, "
        ENd If
		If left(ga_status, 4) = "REIN" Then
			ga_case = TRUE
			adult_cash_case = TRUE
			case_rein = TRUE
		End If
	End If
    row = 1                                             'Looking for DWP information
    col = 1
    EMSearch "DWP:", row, col
    If row <> 0 Then
        EMReadScreen dw_status, 9, row, col + 4
        dw_status = trim(dw_status)
		dwp_status = dw_status
        If dw_status = "ACTIVE" or dw_status = "APP CLOSE" or dw_status = "APP OPEN" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "DWP, "
        End If
        If dw_status = "PENDING" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "DWP, "
        ENd If
		If left(dw_status, 4) = "REIN" Then
			dwp_case = TRUE
			family_cash_case = TRUE
			case_rein = TRUE
		End If
	End If
    row = 1                                             'Looking for MFIP information
    col = 1
    EMSearch "MFIP:", row, col
    If row <> 0 Then
        EMReadScreen mf_status, 9, row, col + 6
        mf_status = trim(mf_status)
		mfip_status = mf_status
        If mf_status = "ACTIVE" or mf_status = "APP CLOSE" or mf_status = "APP OPEN" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "MFIP, "
        End If
        If mf_status = "PENDING" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "MFIP, "
        ENd If
		If left(mf_status, 4) = "REIN" Then
			mfip_case = TRUE
			family_cash_case = TRUE
			case_rein = TRUE
		End If
	End If
    row = 1                                                 'Looking for a general 'Cash' header which means any kind of cash could be pending
    col = 1
    EMSearch "Cash:", row, col
    If row <> 0 Then
        EMReadScreen cash_status, 9, row, col + 6
        cash_status = trim(cash_status)
        If cash_status = "PENDING" Then
            unknown_cash_pending = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "CASH, "
        ENd If
    End If
	row = 1                                             'Looking for IV-E information
	col = 1
	EMSearch "IV-E:", row, col
	If row <> 0 Then
	    EMReadScreen ive_status, 9, row, col + 6
	    ive_status = trim(ive_status)
	    If ive_status = "ACTIVE" or ive_status = "APP CLOSE" or ive_status = "APP OPEN" Then list_active_programs = list_active_programs & "IV-E, "
	    If ive_status = "PENDING" Then list_pending_programs = list_pending_programs & "IV-E, "
	End If
	row = 1                                             'Looking for CCAP information
	col = 1
	EMSearch "CCAP", row, col
	If row <> 0 Then
	    EMReadScreen cca_status, 9, row, col + 6
	    cca_status = trim(cca_status)
	    If cca_status = "ACTIVE" or cca_status = "APP CLOSE" or cca_status = "APP OPEN" Then list_active_programs = list_active_programs & "CCAP, "
	    If cca_status = "PENDING" Then list_pending_programs = list_pending_programs & "CCAP, "
	End If
	row = 1                                                 'Looking for a general 'Cash' header which means any kind of cash could be pending
    col = 1
    EMSearch "HC:", row, col
    If row <> 0 Then
        EMReadScreen hc_status, 9, row, col + 4
        hc_status = trim(hc_status)
        If hc_status = "PENDING" Then
            unknown_hc_pending = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "HC, "
        ENd If
    End If
    row = 1                                             'Looking for MA information
    col = 1
    EMSearch "MA:", row, col
    If row <> 0 Then
        EMReadScreen ma_status, 9, row, col + 4
        ma_status = trim(ma_status)
        If ma_status = "ACTIVE" or ma_status = "APP CLOSE" or ma_status = "APP OPEN" Then
            ma_case = TRUE
            case_active = TRUE
			If InStr(list_active_programs, "HC") = 0 Then list_active_programs = list_active_programs & "HC, "
        End If
        If ma_status = "PENDING" Then
            ma_case = TRUE
            case_pending = TRUE
			If InStr(list_pending_programs, "HC") = 0 Then list_pending_programs = list_pending_programs & "HC, "
        End If
		If left(ma_status, 4) = "REIN" Then
			ma_case = TRUE
			case_rein = TRUE
		End If
	End If
    'MSA programs have different headers so we need to search for them all seperately'
    row = 1                                             'Looking for QMB information for MSA programs
    col = 1
    EMSearch "QMB:", row, col
    If row <> 0 Then
        EMReadScreen qm_status, 9, row, col + 5
        qm_status = trim(qm_status)
        If qm_status = "ACTIVE" or qm_status = "APP CLOSE" or qm_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
			If InStr(list_active_programs, "HC") = 0 Then list_active_programs = list_active_programs & "HC, "
        End If
        If qm_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
			If InStr(list_pending_programs, "HC") = 0 Then list_pending_programs = list_pending_programs & "HC, "
        End If
		If left(qm_status, 4) = "REIN" Then
			msp_case = TRUE
			case_rein = TRUE
		End If
    End If
    row = 1                                             'Looking for SLMB information for MSA programs
    col = 1
    EMSearch "SLMB:", row, col
    If row <> 0 Then
        EMReadScreen sl_status, 9, row, col + 6
        sl_status = trim(sl_status)
        If sl_status = "ACTIVE" or sl_status = "APP CLOSE" or sl_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
			If InStr(list_active_programs, "HC") = 0 Then list_active_programs = list_active_programs & "HC, "
        End If
        If sl_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
			If InStr(list_pending_programs, "HC") = 0 Then list_pending_programs = list_pending_programs & "HC, "
        End If
		If left(sl_status, 4) = "REIN" Then
			msp_case = TRUE
			case_rein = TRUE
		End If
    End If
    row = 1                                             'Looking for QI information for MSA programs
    col = 1
    EMSearch "QI:", row, col
    If row <> 0 Then
        EMReadScreen qi_status, 9, row, col + 5
        qi_status = trim(qi_status)
        If qi_status = "ACTIVE" or qi_status = "APP CLOSE" or qi_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
			If InStr(list_active_programs, "HC") = 0 Then list_active_programs = list_active_programs & "HC, "
        End If
        If qi_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
			If InStr(list_pending_programs, "HC") = 0 Then list_pending_programs = list_pending_programs & "HC, "
        End If
		If left(qi_status, 4) = "REIN" Then
			msp_case = TRUE
			case_rein = TRUE
		End If
    End If
	If qm_status = "ACTIVE" or sl_status = "ACTIVE" or qi_status = "ACTIVE" Then
		msp_status = "ACTIVE"
		If qm_status = "ACTIVE" Then msp_type = "QMB"
		If sl_status = "ACTIVE" Then msp_type = "SLMB"
		If qi_status = "ACTIVE" Then msp_type = "QI1"
	ElseIf qm_status = "APP OPEN" or sl_status = "APP OPEN" or qi_status = "APP OPEN" Then
		msp_status = "APP OPEN"
		If qm_status = "APP OPEN" Then msp_type = "QMB"
		If sl_status = "APP OPEN" Then msp_type = "SLMB"
		If qi_status = "APP OPEN" Then msp_type = "QI1"
	ElseIf qm_status = "APP CLOSE" or sl_status = "APP CLOSE" or qi_status = "APP CLOSE" Then
		msp_status = "APP CLOSE"
		If qm_status = "APP CLOSE" Then msp_type = "QMB"
		If sl_status = "APP CLOSE" Then msp_type = "SLMB"
		If qi_status = "APP CLOSE" Then msp_type = "QI1"
	ElseIf qm_status = "PENDING" or sl_status = "PENDING" or qi_status = "PENDING" Then
		msp_status = "PENDING"
		If qm_status = "PENDING" Then msp_type = "QMB"
		If sl_status = "PENDING" Then msp_type = "SLMB"
		If qi_status = "PENDING" Then msp_type = "QI1"
	ElseIf qm_status = "REIN" or sl_status = "REIN" or qi_status = "REIN" Then
		msp_status = "REIN"
		If qm_status = "REIN" Then msp_type = "QMB"
		If sl_status = "REIN" Then msp_type = "SLMB"
		If qi_status = "REIN" Then msp_type = "QI1"
	End If
	row = 1                                             'Looking for EMER information
    col = 1
    EMSearch "EGA:", row, col
    If row <> 0 Then
        EMReadScreen ega_status, 9, row, col + 4
        ega_status = trim(ega_status)
        If ega_status = "ACTIVE" or ega_status = "APP CLOSE" or ega_status = "APP OPEN" Then
            emer_case = TRUE
            case_active = TRUE
			list_active_programs = list_active_programs & "EGA, "
        End If
        If ega_status = "PENDING" Then
            dwp_case = TRUE
            emer_case = TRUE
            case_pending = TRUE
			list_pending_programs = list_pending_programs & "EGA, "
        ENd If
		If left(ega_status, 4) = "REIN" Then
			emer_case = TRUE
			case_rein = TRUE
		End If
	End If
	row = 1                                             'Looking for EMER information
	col = 1
	EMSearch "EA:", row, col
	If row <> 0 Then
		EMReadScreen ea_status, 9, row, col + 5
		ea_status = trim(ea_status)
		If ea_status = "ACTIVE" or ea_status = "APP CLOSE" or ea_status = "APP OPEN" Then
			emer_case = TRUE
			case_active = TRUE
			list_active_programs = list_active_programs & "EA, "
		End If
		If ea_status = "PENDING" Then
			emer_case = TRUE
			case_pending = TRUE
			list_pending_programs = list_pending_programs & "EA, "
		ENd If
		If left(ea_status, 4) = "REIN" Then
			emer_case = TRUE
			case_rein = TRUE
		End If
	End If
	If ega_status = "ACTIVE" or ea_status = "ACTIVE" Then
		emer_status = "ACTIVE"
		If ega_status = "ACTIVE" Then emer_type = "EGA"
		If ea_status = "ACTIVE" Then emer_type = "EA"
	ElseIf ega_status = "APP OPEN" or ea_status = "APP OPEN" Then
		emer_status = "APP OPEN"
		If ega_status = "APP OPEN" Then emer_type = "EGA"
		If ea_status = "APP OPEN" Then emer_type = "EA"
	ElseIf ega_status = "APP CLOSE" or ea_status = "APP CLOSE" Then
		emer_status = "APP CLOSE"
		If ega_status = "APP CLOSE" Then emer_type = "EGA"
		If ea_status = "APP CLOSE" Then emer_type = "EA"
	ElseIf ega_status = "PENDING" or ea_status = "PENDING" Then
		emer_status = "PENDING"
		If ega_status = "PENDING" Then emer_type = "EGA"
		If ea_status = "PENDING" Then emer_type = "EA"
	ElseIf ega_status = "REIN" or ea_status = "REIN" Then
		emer_status = "REIN"
		If ega_status = "REIN" Then emer_type = "EGA"
		If ea_status = "REIN" Then emer_type = "EA"
	End If

	'formatting the string of the list of active programs and pending programs
	list_active_programs = trim(list_active_programs)  'trims excess spaces of list_active_programs
	If right(list_active_programs, 1) = "," THEN list_active_programs = left(list_active_programs, len(list_active_programs) - 1)

	list_pending_programs = trim(list_pending_programs)  'trims excess spaces of list_pending_programs
	If right(list_pending_programs, 1) = "," THEN list_pending_programs = left(list_pending_programs, len(list_pending_programs) - 1)
End Function

function display_ADDR_information(update_addr, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)
'--- This function has a portion of dialog that can be inserted into a defined dialog. This does NOT have a 'BeginDialog' OR a dialog call. This can allow us to have the same display and update functionality of ADDR information in different scripts/dialogs
'~~~~~ update_addr: boolean - This parameter will determine if the dialog will be displayed with the panel information is 'edit mode' or not.
'~~~~~ notes_on_address: string - variable to enter information and details about the address information in an editbox
'~~~~~ resi_street_full: string - resident address information
'~~~~~ resi_city: string - resident address information
'~~~~~ resi_state: string - resident address information
'~~~~~ resi_zip: string - resident address information
'~~~~~ resi_county: string - resident address information
'~~~~~ addr_verif: string - the verification on the panel - can be updated here
'~~~~~ addr_homeless: string - information about address homeless status
'~~~~~ addr_reservation: string - information about address reservation status
'~~~~~ reservation_name: string - information about address - the specifc reservation name
'~~~~~ addr_living_sit: string - information about living situation
'~~~~~ mail_street_full: string - mailing address information
'~~~~~ mail_city: string - mailing address information
'~~~~~ mail_state: string - mailing address information
'~~~~~ mail_zip: string -  mailing address information
'~~~~~ addr_eff_date: string - the date from the panel indicated as the effective date of the address informaiton
'~~~~~ phone_one: string - phone number information
'~~~~~ phone_two: string - phone number information
'~~~~~ phone_three: string - phone number information
'~~~~~ type_one: string - information about the type of the phone number from the ADDR panel
'~~~~~ type_two: string - information about the type of the phone number from the ADDR panel
'~~~~~ type_three: string - information about the type of the phone number from the ADDR panel
'~~~~~ address_change_date: string - a date that can be entered in the dialog with the date of the change of address information
'~~~~~ update_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ save_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_mail_addr_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_one_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_two_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_three_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'===== Keywords: MAXIS, ADDR, dialog, update
	GroupBox 10, 35, 375, 95, "Residence Address"
	If update_addr = False Then
		Text 330, 35, 50, 10, addr_eff_date
		Text 70, 55, 305, 15, resi_street_full
		Text 70, 75, 105, 15, resi_city
		Text 205, 75, 110, 45, resi_state
		Text 340, 75, 35, 15, resi_zip
		Text 125, 95, 45, 45, addr_reservation
		Text 245, 95, 130, 15, reservation_name
		Text 125, 115, 45, 45, addr_homeless
		If addr_living_sit = "10 - Unknown" OR addr_living_sit = "Blank" Then
			DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
		Else
			Text 245, 115, 130, 45, addr_living_sit
		End If
		Text 70, 165, 305, 15, mail_street_full
		Text 70, 185, 105, 15, mail_city
		Text 205, 185, 110, 45, mail_state
		Text 340, 185, 35, 15, mail_zip
		Text 20, 240, 90, 15, phone_one
		Text 125, 240, 65, 45, type_one
		Text 20, 260, 90, 15, phone_two
		Text 125, 260, 65, 45, type_two
		Text 20, 280, 90, 15, phone_three
		Text 125, 280, 65, 45, type_three
		Text 325, 215, 50, 10, address_change_date
		Text 255, 245, 120, 10, resi_county
		Text 255, 280, 120, 10, addr_verif
		EditBox 10, 320, 375, 15, notes_on_address
		PushButton 290, 300, 95, 15, "Update Information", update_information_btn
	End If
	If update_addr = True Then
		EditBox 330, 30, 40, 15, addr_eff_date
		EditBox 70, 50, 305, 15, resi_street_full
		EditBox 70, 70, 105, 15, resi_city
		DropListBox 205, 70, 110, 45, ""+chr(9)+state_list, resi_state
		EditBox 340, 70, 35, 15, resi_zip
		DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", addr_reservation
		DropListBox 245, 90, 130, 15, "Select One..."+chr(9)+""+chr(9)+"BD - Bois Forte - Deer Creek"+chr(9)+"BN - Bois Forte - Nett Lake"+chr(9)+"BV - Bois Forte - Vermillion Lk"+chr(9)+"FL - Fond du Lac"+chr(9)+"GP - Grand Portage"+chr(9)+"LL - Leach Lake"+chr(9)+"LS - Lower Sioux"+chr(9)+"ML - Mille Lacs"+chr(9)+"PL - Prairie Island Community"+chr(9)+"RL - Red Lake"+chr(9)+"SM - Shakopee Mdewakanton"+chr(9)+"US - Upper Sioux"+chr(9)+"WE - White Earth", reservation_name
		DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", addr_homeless
		DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
		EditBox 70, 160, 305, 15, mail_street_full
		EditBox 70, 180, 105, 15, mail_city
		DropListBox 205, 180, 110, 45, ""+chr(9)+state_list, mail_state
		EditBox 340, 180, 35, 15, mail_zip
		EditBox 20, 240, 90, 15, phone_one
		DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+""+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_one
		EditBox 20, 260, 90, 15, phone_two
		DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+""+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_two
		EditBox 20, 280, 90, 15, phone_three
		DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+""+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_three
		EditBox 325, 210, 50, 15, address_change_date
		ComboBox 255, 245, 120, 45, county_list_smalll+chr(9)+resi_county, resi_county
		DropListBox 255, 280, 120, 45, "SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed"+chr(9)+"Blank", addr_verif
		EditBox 10, 320, 375, 15, notes_on_address
		PushButton 290, 300, 95, 15, "Save Information", save_information_btn
	End If

	PushButton 325, 145, 50, 10, "CLEAR", clear_mail_addr_btn
	PushButton 205, 240, 35, 10, "CLEAR", clear_phone_one_btn
	PushButton 205, 260, 35, 10, "CLEAR", clear_phone_two_btn
	PushButton 205, 280, 35, 10, "CLEAR", clear_phone_three_btn
	Text 250, 35, 80, 10, "ADDR effective date:"
	Text 20, 55, 45, 10, "House/Street"
	Text 45, 75, 20, 10, "City"
	Text 185, 75, 20, 10, "State"
	Text 325, 75, 15, 10, "Zip"
	Text 20, 95, 100, 10, "Do you live on a Reservation?"
	Text 180, 95, 60, 10, "If yes, which one?"
	Text 30, 115, 90, 10, "Client Indicates Homeless:"
	Text 185, 115, 60, 10, "Living Situation?"
	GroupBox 10, 135, 375, 70, "Mailing Address"
	Text 20, 165, 45, 10, "House/Street"
	Text 45, 185, 20, 10, "City"
	Text 185, 185, 20, 10, "State"
	Text 325, 185, 15, 10, "Zip"
	GroupBox 10, 210, 235, 90, "Phone Number"
	Text 20, 225, 50, 10, "Number"
	Text 125, 225, 25, 10, "Type"
	Text 255, 215, 60, 10, "Date of Change:"
	Text 255, 235, 75, 10, "County of Residence:"
	Text 255, 270, 75, 10, "ADDR Verification:"
	Text 10, 310, 75, 10, "Additional Notes:"
end function

function display_HEST_information(update_hest, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, notes_on_hest, update_information_btn, save_information_btn)
'--- This function has a portion of dialog that can be inserted into a defined dialog. This does NOT have a 'BeginDialog' OR a dialog call. This can allow us to have the same display and update functionality of HEST information in different scripts/dialogs
'~~~~~ update_hest: boolean - This parameter will determine if the dialog will be displayed with the panel information is 'edit mode' or not.
'~~~~~ all_persons_paying: string - detail of all the people that are responsible for paying utilities - from the HEST panel
'~~~~~ choice_date: string - formatted as a date - the date from the HEST panel
'~~~~~ actual_initial_exp: string - information from the panel of the actual expense amount in the initial month
'~~~~~ retro_heat_ac_yn: string - as 'Y' or 'N' or blank - indicator of is retro heat/ac paid
'~~~~~ retro_heat_ac_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ retro_heat_ac_amt: number - amount of the SUA for heat/ac - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ retro_electric_yn: string - as 'Y' or 'N' or blank - indicator of is retro electric is paid
'~~~~~ retro_electric_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ retro_electric_amt: number - amount of the SUA for electric - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ retro_phone_yn: string - as 'Y' or 'N' or blank - indicator of is retro phone is paid
'~~~~~ retro_phone_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ retro_phone_amt: number - amount of the SUA for phone - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ prosp_heat_ac_yn: string - as 'Y' or 'N' or blank - indicator of is prosp heat/ac is paid
'~~~~~ prosp_heat_ac_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ prosp_heat_ac_amt: number - amount of the SUA for heat/ac - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ prosp_electric_yn: string - as 'Y' or 'N' or blank - indicator of is prosp electric is paid
'~~~~~ prosp_electric_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ prosp_electric_amt: number - amount of the SUA for electric - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ prosp_phone_yn: string - as 'Y' or 'N' or blank - indicator of is prosp phone is paid
'~~~~~ prosp_phone_units: string - 2 digit as a number - this indicates the number of units that split this expense
'~~~~~ prosp_phone_amt: number - amount of the SUA for phone - read from the panel or calculated by navigate_HEST_buttons as detailed by HEST_standards functions
'~~~~~ total_utility_expense: number - calculated by access_HEST_panel or navigate_HEST_buttons of the total the SUA allowed by what is paid
'~~~~~ update_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ save_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'===== Keywords: MAXIS, HEST, dialog, update
	If update_hest = False Then
		Text 75, 30, 145, 10, all_persons_paying
	    Text 75, 50, 50, 10, choice_date
	    Text 125, 70, 50, 10, actual_initial_exp
	    Text 70, 125, 40, 10, retro_heat_ac_yn
	    Text 115, 125, 20, 10, retro_heat_ac_units
	    Text 150, 125, 45, 10, retro_heat_ac_amt
	    Text 240, 125, 40, 10, prosp_heat_ac_yn
	    Text 285, 125, 20, 10, prosp_heat_ac_units
	    Text 320, 125, 45, 10, prosp_heat_ac_amt
	    Text 70, 145, 40, 10, retro_electric_yn
	    Text 115, 145, 20, 10, retro_electric_units
	    Text 150, 145, 45, 10, retro_electric_amt
	    Text 240, 145, 40, 10, prosp_electric_yn
	    Text 285, 145, 20, 10, prosp_electric_units
	    Text 320, 145, 45, 10, prosp_electric_amt
	    Text 70, 165, 40, 10, retro_phone_yn
	    Text 115, 165, 20, 10, retro_phone_units
	    Text 150, 165, 45, 10, retro_phone_amt
	    Text 240, 165, 40, 10, prosp_phone_yn
	    Text 285, 165, 20, 10, prosp_phone_units
	    Text 320, 165, 45, 10, prosp_phone_amt
		Text 55, 185, 150, 10, "Total Counted Utility Expense: $" & total_utility_expense
		EditBox 10, 220, 370, 15, notes_on_hest

		PushButton 280, 185, 95, 15, "Update Information", update_information_btn
	End If
	If update_hest = True Then
		EditBox 75, 25, 145, 15, all_persons_paying
	    EditBox 75, 45, 50, 15, choice_date
	    EditBox 125, 65, 50, 15, actual_initial_exp
	    DropListBox 65, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_heat_ac_yn
	    DropListBox 235, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_heat_ac_yn
	    DropListBox 65, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_electric_yn
	    DropListBox 235, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_electric_yn
	    DropListBox 65, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_phone_yn
	    DropListBox 235, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_phone_yn
		EditBox 10, 220, 370, 15, notes_on_hest
		PushButton 280, 185, 95, 15, "Save Information", save_information_btn
	End If
	GroupBox 10, 15, 370, 190, "Utility Information (SUA - Standard Utility Allowance)"
	Text 15, 30, 60, 10, "Persons Paying:"
	Text 15, 50, 55, 10, "FS Choice Date:"
	Text 15, 70, 110, 10, "Actual Expense In Initial Month: $ "
	Text 20, 125, 30, 10, "Heat/Air:"
	Text 20, 145, 30, 10, "Electric:"
	Text 25, 165, 25, 10, "Phone:"
	GroupBox 55, 85, 150, 95, "Retrospective"
	Text 65, 105, 20, 10, "(Y/N)"
	Text 110, 100, 20, 20, "#/FS Units"
	Text 150, 105, 30, 10, "Amount"
	GroupBox 225, 85, 150, 95, "Prospective"
	Text 235, 105, 20, 10, "(Y/N)"
	Text 280, 100, 20, 20, "#/FS Units"
	Text 320, 105, 25, 10, "Amount"
	Text 10, 210, 75, 10, "Additional Notes:"
end function

function display_HOUSING_CHANGE_information(housing_questions_step, shel_update_step, notes_on_address, original_resi_street_full, original_resi_city, original_resi_state, original_resi_zip, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, THE_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, hh_comp_change, hest_heat_ac_yn_code, hest_electric_yn_code, hest_phone_yn_code, new_total_shelter_expense_amount, people_paying_SHEL, verif_detail, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)
'--- This function has a portion of dialog that can be inserted into a defined dialog. This does NOT have a 'BeginDialog' OR a dialog call. This can allow us to have the same display and update functionality of HEST information in different scripts/dialogs
'~~~~~ housing_questions_step - number - to identify which step of the dialog process we are at
'~~~~~ shel_update_step - number - to identify which step of the housing exxpense update process we are at
'~~~~~ notes_on_address - variable - string for entering additional notes
'~~~~~ original_resi_street_full - string - residence address read from MAXIS originally
'~~~~~ original_resi_city - string - residence address read from MAXIS originally
'~~~~~ original_resi_state - string - residence address read from MAXIS originally
'~~~~~ original_resi_zip - string - residence address read from MAXIS originally
'~~~~~ resi_street_full - string - residence address - this may change
'~~~~~ resi_city - string - residence address - this may change
'~~~~~ resi_state - string - residence address - this may change
'~~~~~ resi_zip - string - residence address - this may change
'~~~~~ resi_county - string - residence address - this may change
'~~~~~ addr_verif - string - residence address - this may change
'~~~~~ addr_homeless - string - residence address - this may change
'~~~~~ addr_reservation - string - residence address - this may change
'~~~~~ reservation_name - string - residence address - this may change
'~~~~~ addr_living_sit - string - residence address - this may change
'~~~~~ mail_street_full - string - mailing address - this may change
'~~~~~ mail_city - string - mailing address - this may change
'~~~~~ mail_state - string - mailing address - this may change
'~~~~~ mail_zip - string - mailing address - this may change
'~~~~~ addr_eff_date - string - address from ADDR panel
'~~~~~ phone_one - string - phone information - this may change
'~~~~~ phone_two - string - phone information - this may change
'~~~~~ phone_three - string - phone information - this may change
'~~~~~ type_one - string - phone information - this may change
'~~~~~ type_two - string - phone information - this may change
'~~~~~ type_three - string - phone information - this may change
'~~~~~ address_change_date - date - enterd in the dialog
'~~~~~ update_information_btn - number - button definition
'~~~~~ save_information_btn - number - button definition
'~~~~~ clear_mail_addr_btn - number - button definition
'~~~~~ clear_phone_one_btn - number - button definition
'~~~~~ clear_phone_two_btn - number - button definition
'~~~~~ clear_phone_three_btn - number - button definition
'~~~~~ household_move_yn - string - dialog input answer
'~~~~~ household_move_everyone_yn - string - dialog input answer
'~~~~~ move_date - date - dialog input answer
'~~~~~ shel_change_yn - string - dialog input answer
'~~~~~ shel_verif_received_yn - string - dialog input answer
'~~~~~ shel_start_date - date - dialog input answer
'~~~~~ shel_shared_yn - string - dialog input answer
'~~~~~ shel_subsidized_yn - string - dialog input answer
'~~~~~ total_current_rent - string - information from SHEL
'~~~~~ all_rent_verif - string - information from SHEL
'~~~~~ total_current_lot_rent - string - information from SHEL
'~~~~~ all_lot_rent_verif - string - information from SHEL
'~~~~~ total_current_garage - string - information from SHEL
'~~~~~ all_mortgage_verif - string - information from SHEL
'~~~~~ total_current_insurance - string - information from SHEL
'~~~~~ all_insurance_verif - string - information from SHEL
'~~~~~ total_current_taxes - string - information from SHEL
'~~~~~ all_taxes_verif - string - information from SHEL
'~~~~~ total_current_room - string - information from SHEL
'~~~~~ all_room_verif - string - information from SHEL
'~~~~~ total_current_mortgage - string - information from SHEL
'~~~~~ all_garage_verif - string - information from SHEL
'~~~~~ total_current_subsidy - string - information from SHEL
'~~~~~ all_subsidy_verif - string - information from SHEL
'~~~~~ shel_change_type - string - dialog input answer
'~~~~~ hest_heat_ac_yn - string - dialog input answer
'~~~~~ hest_electric_yn - string - dialog input answer
'~~~~~ hest_ac_on_electric_yn - string - dialog input answer
'~~~~~ hest_heat_on_electric_yn - string - dialog input answer
'~~~~~ hest_phone_yn - string - dialog input answer
'~~~~~ update_addr_button - number - button definition
'~~~~~ addr_or_shel_change_notes - string - dialog input answer
'~~~~~ view_addr_update_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ view_shel_update_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ view_shel_details_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ what_is_the_living_arrangement - string - dialog input answer
'~~~~~ unit_owned - string - dialog input answer
'~~~~~ new_total_rent_amount - string - dialog input answer
'~~~~~ new_total_mortgage_amount - string - dialog input answer
'~~~~~ new_total_lot_rent_amount - string - dialog input answer
'~~~~~ new_total_room_amount - string - dialog input answer
'~~~~~ new_room_payment_frequency - string - dialog input answer
'~~~~~ new_mortgage_have_escrow_yn - string - dialog input answer
'~~~~~ new_morgage_insurance_amount - string - dialog input answer
'~~~~~ new_excess_insurance_yn - string - dialog input answer
'~~~~~ new_total_tax_amount - string - dialog input answer
'~~~~~ new_rent_subsidy_yn - string - dialog input answer
'~~~~~ new_renter_insurance_amount - string - dialog input answer
'~~~~~ new_renters_insurance_required_yn - string - dialog input answer
'~~~~~ new_total_garage_amount - string - dialog input answer
'~~~~~ new_garage_rent_required_yn - string - dialog input answer
'~~~~~ new_vehicle_insurance_amount - string - dialog input answer
'~~~~~ new_total_insurance_amount - string - dialog input answer
'~~~~~ new_total_subsidy_amount - string - dialog input answer
'~~~~~ new_SHEL_paid_to_name - string - dialog input answer
'~~~~~ other_person_checkbox - 1 or 0 - dialog input checkbox
'~~~~~ other_person_name - string - dialog input answer
'~~~~~ payment_split_evenly_yn - string - dialog input answer
'~~~~~ THE_ARRAY - an ARRAY - of all SHEL panels on the case - filled with access_SHEL_panel
'~~~~~ person_age_const - number - constant used in the ARRAY where the person's age is saved
'~~~~~ person_shel_checkbox - number - constant used in the ARRAY where a checkboxx detail is saved
'~~~~~ shel_ref_number_const - number - constant used in the ARRAY where the reference number of the person is saved
'~~~~~ new_shel_pers_total_amt_const - number - constant used in the ARRAY where the amount of shelter expense paid is saved
'~~~~~ new_shel_pers_total_amt_type_const - number - constant used in the ARRAY where the type (dollars or percent) is indicated
'~~~~~ other_new_shel_total_amt - string - dialog input answer
'~~~~~ other_new_shel_total_amt_type - string - dialog input answer
'~~~~~ new_rent_verif - string - dialog input answer
'~~~~~ new_lot_rent_verif - string - dialog input answer
'~~~~~ new_mortgage_verif - string - dialog input answer
'~~~~~ new_insurance_verif - string - dialog input answer
'~~~~~ new_taxes_verif - string - dialog input answer
'~~~~~ new_room_verif - string - dialog input answer
'~~~~~ new_garage_verif - string - dialog input answer
'~~~~~ new_subsidy_verif - string - dialog input answer
'~~~~~ hh_comp_change - string - dialog input answer
'~~~~~ hest_heat_ac_yn_code - string of y or n - determined by thee function for the code that indicates if this expense is paid
'~~~~~ hest_electric_yn_code - string of y or n - determined by thee function for the code that indicates if this expense is paid
'~~~~~ hest_phone_yn_code - string of y or n - determined by thee function for the code that indicates if this expense is paid
'~~~~~ new_total_shelter_expense_amount - number - as a string - the total of the new shelter expense
'~~~~~ people_paying_SHEL - string - a list of the persons entered into the dialog as paying the expense
'~~~~~ verif_detail - string - collection of the verifs entered
'~~~~~ housing_change_continue_btn - number - button definition
'~~~~~ housing_change_overview_btn - number - button definition
'~~~~~ housing_change_addr_update_btn - number - button definition
'~~~~~ housing_change_shel_update_btn - number - button definition
'~~~~~ housing_change_shel_details_btn - number - button definition
'~~~~~ housing_change_review_btn - number - button definition
'~~~~~ enter_shel_one_btn - number - button definition
'~~~~~ enter_shel_two_btn - number - button definition
'~~~~~ enter_shel_three_btn - number - button definition
'===== Keywords: MAXIS, ADDR, SHEL, HEST, dialog, update
	yes_no_list = "?"+chr(9)+"Yes"+chr(9)+"No"
	x_pos = 345
	If view_shel_details_dlg = True Then
		shel_det_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	If view_shel_update_dlg = True Then
		shel_up_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	If view_addr_update_dlg = True Then
		addr_up_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	overview_x_pos = x_pos

	GroupBox 10, 10, 460, 355, "Change in HOUSING Information"

	If housing_questions_step = 1 Then
		Text overview_x_pos + 10, 10, 60, 13, "OVERVIEW"

		GroupBox 15, 25, 450, 75, "Address"
		Text 75, 45, 95, 10, "Did the household move?"
		DropListBox 170, 40, 45, 45, yes_no_list, household_move_yn
		Text 25, 65, 145, 10, "Did everyone in the household move with?"
		DropListBox 170, 60, 45, 45, yes_no_list, household_move_everyone_yn
		Text 75, 85, 90, 10, "What date did they move?"
		EditBox 170, 80, 45, 15, move_date
		Text 255, 40, 95, 10, "Current Residence Address:"
		Text 265, 55, 190, 10, resi_street_full
		Text 265, 70, 190, 10, resi_city & ", " & left(resi_state, 2) & " " & resi_zip

		GroupBox 15, 105, 450, 70, "Housing Cost"
		Text 40, 125, 130, 10, "Is there a change to the housing cost?"
		DropListBox 170, 120, 45, 45, yes_no_list, shel_change_yn

		Text 40, 145, 125, 10, "What date did the new expense start?"
		EditBox 170, 140, 45, 15, shel_start_date

		Text 265, 115, 95, 10, "Current Housing Costs"
		Text 280, 130, 35, 10, " Rent: "
		Text 305, 130, 30, 10, "$ " & total_current_rent
		Text 375, 130, 40, 10, " Taxes: "
		Text 405, 130, 30, 10, "$ " & total_current_taxes
		Text 270, 140, 45, 10, "Lot Rent: "
		Text 305, 140, 30, 10, "$ " & total_current_lot_rent
		Text 375, 140, 40, 10, " Room: "
		Text 405, 140, 30, 10, "$ " & total_current_room
		Text 265, 150, 50, 10, " Mortgage: "
		Text 305, 150, 30, 10, "$ " & total_current_mortgage
		Text 370, 150, 45, 10, " Garage: "
		Text 405, 150, 30, 10, "$ " & total_current_garage
		Text 265, 160, 50, 10, "Insurance: "
		Text 305, 160, 30, 10, "$ " & total_current_insurance
		Text 370, 160, 45, 10, "Subsidy: "
		Text 405, 160, 30, 10, "$ " & total_current_subsidy

		GroupBox 15, 180, 450, 115, "Utilities Expense"
	    Text 25, 195, 275, 10, "Is the household responsible to paythe Heat Expense or Air Conditioner Expense?"
	    DropListBox 295, 190, 45, 45, yes_no_list, hest_heat_ac_yn
	    Text 25, 215, 180, 10, "Is the household responsible to pay electric expense?"
	    DropListBox 210, 210, 45, 45, yes_no_list, hest_electric_yn
	    Text 40, 230, 235, 10, "If yes, is there any AC plugged into that is used at any point in the year?"
	    DropListBox 280, 225, 45, 45, yes_no_list, hest_ac_on_electric_yn
	    Text 40, 250, 235, 10, "If yes, does this include any heat source during any point in the year?"
	    DropListBox 280, 245, 45, 45, yes_no_list, hest_heat_on_electric_yn
	    Text 25, 270, 145, 10, "Is anyone responsible to PAY for a phone?"
	    DropListBox 170, 265, 45, 45, yes_no_list, hest_phone_yn
	    Text 30, 280, 230, 10, "(Free phone plans without a payment requirement cannot be counted.)"
	End If

	If housing_questions_step = 2 Then
		Text addr_up_x_pos + 5, 10, 60, 10, "ADDR UPDATE"

		Text 15, 25, 450, 10, "STEP 2 - ADDR UPDATES  -  Enter new address information here:"

		Call display_ADDR_information(True, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)
	End If
	If housing_questions_step = 3 Then
		Text shel_up_x_pos + 5, 10, 60, 10, "SHEL UPDATE"

		Text 15, 25, 450, 10, "STEP 3 - SHEL UPDATES"

		If shel_update_step > 0 Then
			Text 20, 45, 95, 10, "What is the living situation?"
		    DropListBox 115, 40, 125, 45, "Select One..."+chr(9)+"Apartment or Townhouse"+chr(9)+"House"+chr(9)+"Trailer Home/Mobile Home"+chr(9)+"Room Only"+chr(9)+"Shelter"+chr(9)+"Hotel"+chr(9)+"Vehicle"+chr(9)+"Other", what_is_the_living_arrangement
		    Text 250, 45, 120, 10, "Does the household own the home?"
		    DropListBox 370, 40, 90, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", unit_owned
		    PushButton 410, 55, 50, 10, "Enter", enter_shel_one_btn
		End If

		If shel_update_step > 1 Then
			GroupBox 15, 65, 450, 155, "Payment Details"
			If (what_is_the_living_arrangement = "Apartment or Townhouse" OR what_is_the_living_arrangement = "House") Then
				If unit_owned = "No" Then
				    Text 20, 80, 105, 10, "What is the total rent amount?"
				    EditBox 130, 75, 50, 15, new_total_rent_amount
					Text 225, 80, 100, 10, "Who is the expense paid to?"
					EditBox 325, 75, 135, 15, new_SHEL_paid_to_name
				    Text 20, 100, 195, 10, "Does the household receive a subsidy for the rent amount?"
				    DropListBox 220, 95, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_rent_subsidy_yn
					Text 290, 100, 75, 10, "Subsidy Amount:"
					EditBox 365, 95, 50, 15, new_total_subsidy_amount
				    Text 20, 120, 150, 10, "What is the amount of any renters insurance?"
				    EditBox 175, 115, 50, 15, new_renter_insurance_amount
				    Text 260, 120, 135, 10, "Is this insurance required per the lease?"
				    DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_renters_insurance_required_yn
				    Text 20, 140, 130, 10, "What is the amount of the garage rent?"
				    EditBox 150, 135, 50, 15, new_total_garage_amount
				    Text 250, 140, 145, 10, "Is this garage rental required per the lease?"
				    DropListBox 400, 135, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_garage_rent_required_yn

				End If

				If unit_owned = "Yes" Then
					Text 20, 80, 115, 10, "What is the total mortgage amount?"
					EditBox 140, 75, 50, 15, new_total_mortgage_amount
					Text 230, 80, 170, 10, "Does this payment include all taxes and insturance?"
					DropListBox 400, 75, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_mortgage_have_escrow_yn
					Text 20, 100, 150, 10, "How much insurance do you pay seperately?"
					EditBox 170, 95, 40, 15, new_morgage_insurance_amount
					Text 220, 100, 195, 10, "Do have more insurance than required by the mortgage?"
					DropListBox 400, 95, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_excess_insurance_yn
					Text 20, 120, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 115, 50, 15, new_total_tax_amount
					Text 20, 140, 100, 10, "Who is the mortgage paid to?"
					EditBox 120, 135, 135, 15, new_SHEL_paid_to_name
				End If
			ElseIf what_is_the_living_arrangement = "Trailer Home/Mobile Home" Then
				If unit_owned = "No" Then
					Text 20, 80, 110, 10, "What is the total unit rent amount?"
					EditBox 135, 75, 50, 15, new_total_rent_amount
					Text 20, 100, 105, 10, "What is the lot rent Amount?"
					EditBox 130, 95, 50, 15, new_total_lot_rent_amount
					Text 20, 120, 150, 10, "What is the amount of any renters insurance?"
					EditBox 175, 115, 50, 15, new_renter_insurance_amount
					Text 260, 120, 135, 10, "Is this insurance required per the lease?"
					DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_renters_insurance_required_yn

					Text 20, 140, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 135, 50, 15, new_total_tax_amount
					Text 225, 140, 100, 10, "Who is the expense paid to?"
					EditBox 325, 135, 135, 15, new_SHEL_paid_to_name
				End If

				If unit_owned = "Yes" Then
					Text 20, 80, 115, 10, "What is the total mortgage amount?"
					EditBox 140, 75, 50, 15, new_total_mortgage_amount
					Text 230, 80, 170, 10, "Does this payment include all taxes and insturance?"
					DropListBox 400, 75, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_mortgage_have_escrow_yn
					Text 20, 100, 105, 10, "What is the lot rent Amount?"
					EditBox 135, 95, 50, 15, new_total_lot_rent_amount
					Text 20, 120, 150, 10, "How much insurance do you pay seperately?"
					EditBox 170, 115, 40, 15, new_morgage_insurance_amount
					Text 220, 120, 195, 10, "Do have more insurance than required by the mortgage?"
					DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_excess_insurance_yn
					Text 20, 140, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 135, 50, 15, new_total_tax_amount
					Text 225, 140, 100, 10, "Who is the expense paid to?"
					EditBox 325, 135, 135, 15, new_SHEL_paid_to_name
				End If
			ElseIf what_is_the_living_arrangement = "Room Only" Then
				Text 20, 80, 115, 10, "What is the total room rent amount?"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 100, 10, "Who is the room rent paid to?"
				EditBox 120, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Shelter" Then
				Text 20, 80, 115, 10, "What is the cost for the shelter"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 120, 10, "Who is the shelter expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Hotel" Then
				Text 20, 80, 115, 10, "What is the cost for the hotel room?"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 120, 10, "Who is the hotel expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Vehicle" Then
				Text 20, 80, 115, 10, "How much insurance do you pay?"
				EditBox 135, 75, 40, 15, new_vehicle_insurance_amount
				Text 20, 100, 120, 10, "Who is the vehicle expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Other" Then
				Text 25, 80, 35, 10, "Rent: $"
				EditBox 55, 75, 30, 15, new_total_rent_amount
				Text 120, 80, 60, 10, "Mortgage: $"
				EditBox 165, 75, 30, 15, new_total_mortgage_amount
				Text 230, 80, 50, 10, "Lot Rent: $"
				EditBox 275, 75, 30, 15, new_total_lot_rent_amount
				Text 330, 80, 35, 10, "Room: $"
				EditBox 365, 75, 30, 15, new_total_room_amount

				Text 20, 100, 40, 10, "Taxes: $"
				EditBox 55, 95, 30, 15, new_total_tax_amount
				Text 125, 100, 45, 10, "Garage: $"
				EditBox 165, 95, 30, 15, new_total_garage_amount
				Text 225, 100, 55, 10, "Insurance: $"
				EditBox 275, 95, 30, 15, new_total_insurance_amount
				Text 330, 100, 35, 10, "Subsidy: $"
				EditBox 365, 95, 30, 15, new_total_subsidy_amount
				Text 20, 120, 120, 10, "Who is the housing expense paid to?"
				EditBox 140, 115, 135, 15, new_SHEL_paid_to_name
			End If
			GroupBox 20, 155, 440, 65, "Check the box for each person responsible for the housing payment:"
			x_pos = 30
			y_pos = 170
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_age_const, the_membs) >= 18 Then
					CheckBox 30, 170, 80, 10, "MEMB " & THE_ARRAY(shel_ref_number_const, the_membs), THE_ARRAY(person_shel_checkbox, the_membs)
					x_pos = x_pos + 125
					If x_pos = 200 Then
						y_pos = y_pos + 15
					End If
				End If
			next
			CheckBox 290, 170, 145, 10, "Someone outside the household. Name:", other_person_checkbox
			EditBox 305, 180, 150, 15, other_person_name
			Text 25, 205, 200, 10, "Is the payment split evenly among all the responsible parties?"
			DropListBox 230, 200, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", payment_split_evenly_yn
			PushButton 405, 205, 50, 10, "Enter", enter_shel_two_btn
		End If

		If shel_update_step > 2 Then
		    GroupBox 15, 225, 450, 50, "How is the payment split?"
			x_pos = 25
			y_pos = 240
			If new_rent_subsidy_yn = "Yes" Then
				Text x_pos, y_pos, 60, 10, "Subsidy pays: "
				EditBox x_pos + 65, y_pos - 5, 50, 15, new_total_subsidy_amount
				Text x_pos + 120, y_pos, 55, 10, "dollars"
				x_pos = x_pos + 195
				If x_pos = 415 Then
					x_pos = 25
					y_pos = y_pos + 20
				End If
			End If
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
					Text x_pos, y_pos, 60, 10, "MEMB " & THE_ARRAY(shel_ref_number_const, the_membs) & " pays: "
					EditBox x_pos + 65, y_pos - 5, 50, 15, THE_ARRAY(new_shel_pers_total_amt_const, the_membs)
					DropListBox x_pos + 120, y_pos - 5, 55, 45, "dollars"+chr(9)+"percent", THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs)
					x_pos = x_pos + 195
					If x_pos = 415 Then
						x_pos = 25
						y_pos = y_pos + 20
					End If
				End If
			next
			If other_person_checkbox = checked Then
				Text x_pos, y_pos, 100, 10, "Other: " & other_person_name & " pays: "
				EditBox x_pos + 105, y_pos - 5, 50, 15, other_new_shel_total_amt
				DropListBox x_pos + 160, y_pos - 5, 55, 45, "dollars"+chr(9)+"percent", other_new_shel_total_amt_type
			End If
		    PushButton 410, 260, 50, 10, "Enter", enter_shel_three_btn
		End If

		If shel_update_step > 3 Then
		    GroupBox 15, 280, 450, 55, "Is the housing expense verified?"
			x_pos = 25
			y_pos = 300

			If new_total_rent_amount <> "" AND new_total_rent_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total RENT of $" & new_total_rent_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_rent_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If

			If new_total_lot_rent_amount <> "" AND new_total_lot_rent_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total LOT RENT of $" & new_total_lot_rent_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_lot_rent_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_mortgage_amount <> "" AND new_total_mortgage_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total MORTGAGE of $" & new_total_mortgage_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_mortgage_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_insurance_amount <> "" AND new_total_insurance_amount <> "0" Then
				Text x_pos, y_pos, 120, 10, "Total INSURANCE of $" & new_total_insurance_amount & " verification:"
				DropListBox x_pos + 125, y_pos - 5, 80, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_insurance_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_tax_amount <> "" AND new_total_tax_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total TAXES of $" & new_total_tax_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_taxes_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_room_amount <> "" AND new_total_room_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total TOOM of $" & new_total_room_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_room_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_garage_amount <> "" AND new_total_garage_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total GARAGE of $" & new_total_garage_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_garage_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_subsidy_amount <> "" AND new_total_subsidy_amount <> "0" Then
				Text x_pos, y_pos, 120, 10, "Total SUBSIDY of $" & new_total_subsidy_amount & " verification:"
				DropListBox x_pos + 125, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_subsidy_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
		End If
	End If
	If housing_questions_step = 4 Then
		Text 420, 10, 60, 10, "REVIEW"

		Text 15, 25, 450, 10, "STEP 5 - REVIEW AND CONFIRM"

		GroupBox 15, 40, 450, 75, "Household Address Change"
	    If household_move_yn = "Yes" Then Text 25, 55, 150, 10, "The household address changed on " & move_date
		If household_move_yn = "No" Then Text 25, 55, 150, 10, "The household address has not changed."
	    Text 35, 70, 65, 10, "Previous Address"
	    Text 40, 85, 155, 10, original_resi_street_full
	    Text 40, 95, 155, 10, original_resi_city & ", " & left(original_resi_state, 2) & " " & original_resi_zip
	    Text 255, 70, 65, 10, "New Address"
	    Text 260, 85, 155, 10, resi_street_full
	    Text 260, 95, 155, 10, resi_city & ", " & left(resi_state, 2) & " " & resi_zip

	    GroupBox 15, 120, 450, 55, "Household Composition Change"
	    Text 25, 140, 165, 10, "List any changes in the household composition:"
	    EditBox 25, 150, 435, 15, hh_comp_change

	    GroupBox 15, 180, 450, 30, "Standard Utility Allowance - SUA"
		hest_heat_ac_yn_code = "N"
		hest_electric_yn_code = "N"
		hest_phone_yn_code = "N"
		If hest_heat_ac_yn = "Yes" Then hest_heat_ac_yn_code = "Y"
		If hest_ac_on_electric_yn = "Yes" Then hest_heat_ac_yn_code = "Y"
		If hest_heat_on_electric_yn = "Yes" Then hest_heat_ac_yn_code = "Y"
		If hest_electric_yn = "Yes" Then hest_electric_yn_code = "Y"
		If hest_phone_yn = "Yes" Then hest_phone_yn_code = "Y"
	    Text 25, 195, 50, 10, "Heat/AC - " & hest_heat_ac_yn_code
	    Text 110, 195, 50, 10, "Electric - "  & hest_electric_yn_code
	    Text 195, 195, 50, 10, "Phone - " & hest_phone_yn_code

	    GroupBox 15, 215, 450, 70, "Housing Expense Change"
	    Text 25, 230, 100, 10, "New Total Housing Expense:"
	    Text 130, 230, 50, 10, "$" & new_total_shelter_expense_amount
		Text 200, 230, 75, 10, "Effective Date:"
	    Text 275, 230, 50, 10, shel_start_date

	    Text 95, 245, 30, 10, "Paid by:"
	    Text 130, 245, 330, 10, people_paying_SHEL
	    Text 85, 260, 40, 10, "Verification:"
	    Text 130, 260, 330, 10, verif_detail

	End If

	Text 20, 350, 55, 10, "Additional Notes:"
	EditBox 80, 345, 385, 15, addr_or_shel_change_notes

	If housing_questions_step <> 1 Then PushButton overview_x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	If view_addr_update_dlg = True AND housing_questions_step <> 2 Then PushButton addr_up_x_pos, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	If view_shel_update_dlg = True AND housing_questions_step <> 3 Then PushButton shel_up_x_pos, 8, 60, 13, "SHEL UPDATE", housing_change_shel_update_btn
	If err_msg = "" AND housing_questions_step <> 4 Then PushButton 405, 8, 60, 13, "REVIEW", housing_change_review_btn

	If housing_questions_step <> 4 Then PushButton 390, 325, 70, 10, "CONTINUE", housing_change_continue_btn
end function

function display_SHEL_information(update_shel, show_totals, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, total_paid_to, percent_paid_by_household, percent_paid_by_others, total_current_rent, total_current_lot_rent, total_current_mortgage, total_current_insurance, total_current_taxes, total_current_room, total_current_garage, total_current_subsidy, update_information_btn, save_information_btn, const_memb_buttons, clear_all_btn, view_total_shel_btn, update_household_percent_button)
'--- This function has a portion of dialog that can be inserted into a defined dialog. This does NOT have a 'BeginDialog' OR a dialog call. This can allow us to have the same display and update functionality of SHEL information in different scripts/dialogs
'~~~~~ update_shel: boolean - indicating if the dialog information should be in edit mode or not
'~~~~~ show_totals: boolean - indicates if we are looking at the case total information or the Member specific information
'~~~~~ SHEL_ARRAY: The name of the array used for the all the MEMBER panel information, this is in line with the function access_SHEL_panel
'~~~~~ selection: number - This defnies which of the member information from the array should be displayed - defined in navigate_SHEL_buttons
'~~~~~ const_shel_member: number - constant - the defined constant for the array - the member number information
'~~~~~ const_shel_exists: number - constant - the defined constant for the array - boolean - if a SHEL panel exists
'~~~~~ const_hud_sub_yn: number - constant - the defined constant for the array - code from SHEL - if HUD Subsidy exists
'~~~~~ const_shared_yn: number - constant - the defined constant for the array - code from SHEL - if the expense is shared
'~~~~~ const_paid_to: number - constant - the defined constant for the array - from SHEL - who the expense is paid to
'~~~~~ const_rent_retro_amt: number - constant - the defined constant for the array - number - rent amount
'~~~~~ const_rent_retro_verif: number - constant - the defined constant for the array - string - rent verif
'~~~~~ const_rent_prosp_amt: number - constant - the defined constant for the array - number - rent amount
'~~~~~ const_rent_prosp_verif: number - constant - the defined constant for the array - string - rent verif
'~~~~~ const_lot_rent_retro_amt: number - constant - the defined constant for the array - number - lot rent amount
'~~~~~ const_lot_rent_retro_verif: number - constant - the defined constant for the array - string - lot rent verif
'~~~~~ const_lot_rent_prosp_amt: number - constant - the defined constant for the array - number - lot rent amount
'~~~~~ const_lot_rent_prosp_verif: number - constant - the defined constant for the array - string - lot rent verif
'~~~~~ const_mortgage_retro_amt: number - constant - the defined constant for the array - number - mortgage amount
'~~~~~ const_mortgage_retro_verif: number - constant - the defined constant for the array - string - mortgage verif
'~~~~~ const_mortgage_prosp_amt: number - constant - the defined constant for the array - number - mortgage amount
'~~~~~ const_mortgage_prosp_verif: number - constant - the defined constant for the array - string - mortgage verif
'~~~~~ const_insurance_retro_amt: number - constant - the defined constant for the array - number - insurance amount
'~~~~~ const_insurance_retro_verif: number - constant - the defined constant for the array - string - insurance verif
'~~~~~ const_insurance_prosp_amt: number - constant - the defined constant for the array - number - insurance amount
'~~~~~ const_insurance_prosp_verif: number - constant - the defined constant for the array - string - insurance verif
'~~~~~ const_tax_retro_amt: number - constant - the defined constant for the array - number - tax amount
'~~~~~ const_tax_retro_verif: number - constant - the defined constant for the array - string - tax verif
'~~~~~ const_tax_prosp_amt: number - constant - the defined constant for the array - number - tax amount
'~~~~~ const_tax_prosp_verif: number - constant - the defined constant for the array - string - tax verif
'~~~~~ const_room_retro_amt: number - constant - the defined constant for the array - number - room amount
'~~~~~ const_room_retro_verif: number - constant - the defined constant for the array - string - room verif
'~~~~~ const_room_prosp_amt: number - constant - the defined constant for the array - number - room amount
'~~~~~ const_room_prosp_verif: number - constant - the defined constant for the array - string - room verif
'~~~~~ const_garage_retro_amt: number - constant - the defined constant for the array - number - garage amount
'~~~~~ const_garage_retro_verif: number - constant - the defined constant for the array - string - garage verif
'~~~~~ const_garage_prosp_amt: number - constant - the defined constant for the array - number - garage amount
'~~~~~ const_garage_prosp_verif: number - constant - the defined constant for the array - string - garage verif
'~~~~~ const_subsidy_retro_amt: number - constant - the defined constant for the array - number - subsidy amount
'~~~~~ const_subsidy_retro_verif: number - constant - the defined constant for the array - string - subsidy verif
'~~~~~ const_subsidy_prosp_amt: number - constant - the defined constant for the array - number - subsidy amount
'~~~~~ const_subsidy_prosp_verif: number - constant - the defined constant for the array - string - subsidy verif
'~~~~~ total_paid_to: who is the expense paid to - from reading all the panels on the case
'~~~~~ percent_paid_by_household: number - how much of the expense is paid by the household
'~~~~~ percent_paid_by_others: number - how much of the expese is paid by others
'~~~~~ total_current_rent: numberr - the total rent amount in all panels on the case
'~~~~~ total_current_lot_rent: numberr - the total lot rent amount in all panels on the case
'~~~~~ total_current_mortgage: numberr - the total mortgage amount in all panels on the case
'~~~~~ total_current_insurance: numberr - the total insurance amount in all panels on the case
'~~~~~ total_current_taxes: numberr - the total tax amount in all panels on the case
'~~~~~ total_current_room: numberr - the total room amount in all panels on the case
'~~~~~ total_current_garage: numberr - the total garage amount in all panels on the case
'~~~~~ total_current_subsidy: numberr - the total subsidy amount in all panels on the case
'~~~~~ update_information_btn: number - defined button
'~~~~~ save_information_btn: number - defined button
'~~~~~ const_memb_buttons: number - constant - the defined constant for the array - defined button
'~~~~~ clear_all_btn: number - defined button
'~~~~~ view_total_shel_btn: number - defined button
'~~~~~ update_household_percent_button: number - defined button
'===== Keywords: MAXIS, Dialog, SHEL
	Text 10, 10, 360, 10, "Review the Shelter informaiton known with the client. If it needs updating, press this button to make changes:"

	If show_totals = True Then
		Text 415, 253, 65, 10, "TOTAL SHEL"

		If update_shel = True Then
			EditBox 105, 25, 165, 15, total_paid_to
			EditBox 125, 45, 20, 15, total_paid_by_household
			EditBox 125, 65, 20, 15, total_paid_by_others
			EditBox 105, 105, 45, 15, total_current_rent
			EditBox 105, 125, 45, 15, total_current_lot_rent
			EditBox 105, 145, 45, 15, total_current_mortgage
			EditBox 105, 165, 45, 15, total_current_insurance
			EditBox 105, 185, 45, 15, total_current_taxes
			EditBox 105, 205, 45, 15, total_current_room
			EditBox 105, 225, 45, 15, total_current_garage
			EditBox 105, 245, 45, 15, total_current_subsidy
			PushButton 400, 265, 75, 15, "Save Information", save_information_btn
		End If
		If update_shel = False Then
			Text 105, 30, 165, 10, total_paid_to
			Text 125, 50, 20, 10, total_paid_by_household
			Text 125, 70, 20, 10, total_paid_by_others
			Text 105, 110, 45, 10, total_current_rent
			Text 105, 130, 45, 10, total_current_lot_rent
			Text 105, 150, 45, 10, total_current_mortgage
			Text 105, 170, 45, 10, total_current_insurance
			Text 105, 190, 45, 10, total_current_taxes
			Text 105, 210, 45, 10, total_current_room
			Text 105, 230, 45, 10, total_current_garage
			Text 105, 250, 45, 10, total_current_subsidy
			PushButton 400, 265, 75, 15, "Update Information", update_information_btn
		End If
		Text 15, 30, 90, 10, "Housing Expense Paid to"
		Text 15, 50, 100, 10, "Expense Paid by Household"
		Text 145, 50, 50, 10, "% (percent)"
		' PushButton 210, 41, 125, 13, "MANAGE HOUSEHOLD PERCENT", update_household_percent_button
		Text 15, 70, 100, 10, "Expense Paid by Someone Else"
		Text 145, 70, 50, 10, "% (percent)"
		Text 105, 90, 120, 10, "Current Total Amount"
		Text 80, 110, 20, 10, "Rent:"
	    Text 70, 130, 30, 10, "Lot Rent:"
	    Text 65, 150, 35, 10, "Mortgage:"
	    Text 65, 170, 40, 10, "Insurance:"
	    Text 75, 190, 25, 10, "Taxes:"
	    Text 75, 210, 25, 10, "Room:"
	    Text 75, 230, 30, 10, "Garage:"
	    Text 70, 250, 30, 10, "Subsidy:"

	Else
		PushButton 400, 250, 75, 15, "TOTAL SHEL", view_total_shel_btn

		If update_shel = True Then
			EditBox 105, 25, 165, 15, SHEL_ARRAY(const_paid_to, selection)
			DropListBox 165, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_hud_sub_yn, selection)
			DropListBox 310, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_shared_yn, selection)
			EditBox 105, 95, 45, 15, SHEL_ARRAY(const_rent_retro_amt, selection)
			DropListBox 155, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_retro_verif, selection)
			EditBox 255, 95, 45, 15, SHEL_ARRAY(const_rent_prosp_amt, selection)
			DropListBox 305, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_prosp_verif, selection)
			EditBox 105, 115, 45, 15, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
			DropListBox 155, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_retro_verif, selection)
			EditBox 255, 115, 45, 15, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
			DropListBox 305, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
			EditBox 105, 135, 45, 15, SHEL_ARRAY(const_mortgage_retro_amt, selection)
			DropListBox 155, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_retro_verif, selection)
			EditBox 255, 135, 45, 15, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
			DropListBox 305, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_prosp_verif, selection)
			EditBox 105, 155, 45, 15, SHEL_ARRAY(const_insurance_retro_amt, selection)
			DropListBox 155, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_retro_verif, selection)
			EditBox 255, 155, 45, 15, SHEL_ARRAY(const_insurance_prosp_amt, selection)
			DropListBox 305, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_prosp_verif, selection)
			EditBox 105, 175, 45, 15, SHEL_ARRAY(const_tax_retro_amt, selection)
			DropListBox 155, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_retro_verif, selection)
			EditBox 255, 175, 45, 15, SHEL_ARRAY(const_tax_prosp_amt, selection)
			DropListBox 305, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_prosp_verif, selection)
			EditBox 105, 195, 45, 15, SHEL_ARRAY(const_room_retro_amt, selection)
			DropListBox 155, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_retro_verif, selection)
			EditBox 255, 195, 45, 15, SHEL_ARRAY(const_room_prosp_amt, selection)
			DropListBox 305, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_prosp_verif, selection)
			EditBox 105, 215, 45, 15, SHEL_ARRAY(const_garage_retro_amt, selection)
			DropListBox 155, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_retro_verif, selection)
			EditBox 255, 215, 45, 15, SHEL_ARRAY(const_garage_prosp_amt, selection)
			DropListBox 305, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_prosp_verif, selection)
			EditBox 105, 235, 45, 15, SHEL_ARRAY(const_subsidy_retro_amt, selection)
			DropListBox 155, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_retro_verif, selection)
			EditBox 255, 235, 45, 15, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
			DropListBox 305, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_prosp_verif, selection)
			PushButton 400, 265, 75, 15, "Save Information", save_information_btn
		End If
		If update_shel = False Then
			Text 105, 30, 165, 10, SHEL_ARRAY(const_paid_to, selection)
			Text 165, 50, 40, 10, SHEL_ARRAY(const_hud_sub_yn, selection)
			Text 310, 50, 40, 10, SHEL_ARRAY(const_shared_yn, selection)
			Text 105, 100, 45, 10, SHEL_ARRAY(const_rent_retro_amt, selection)
			Text 160, 100, 70, 10, SHEL_ARRAY(const_rent_retro_verif, selection)
			Text 255, 100, 45, 10, SHEL_ARRAY(const_rent_prosp_amt, selection)
			Text 310, 100, 70, 10, SHEL_ARRAY(const_rent_prosp_verif, selection)
			Text 105, 120, 45, 10, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
			Text 160, 120, 70, 10, SHEL_ARRAY(const_lot_rent_retro_verif, selection)
			Text 255, 120, 45, 10, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
			Text 310, 120, 70, 10, SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
			Text 105, 140, 45, 10, SHEL_ARRAY(const_mortgage_retro_amt, selection)
			Text 160, 140, 70, 10, SHEL_ARRAY(const_mortgage_retro_verif, selection)
			Text 255, 140, 45, 10, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
			Text 310, 140, 70, 10, SHEL_ARRAY(const_mortgage_prosp_verif, selection)
			Text 105, 160, 45, 10, SHEL_ARRAY(const_insurance_retro_amt, selection)
			Text 160, 160, 70, 10, SHEL_ARRAY(const_insurance_retro_verif, selection)
			Text 255, 160, 45, 10, SHEL_ARRAY(const_insurance_prosp_amt, selection)
			Text 310, 160, 70, 10, SHEL_ARRAY(const_insurance_prosp_verif, selection)
			Text 105, 180, 45, 10, SHEL_ARRAY(const_tax_retro_amt, selection)
			Text 160, 180, 70, 10, SHEL_ARRAY(const_tax_retro_verif, selection)
			Text 255, 180, 45, 10, SHEL_ARRAY(const_tax_prosp_amt, selection)
			Text 310, 180, 70, 10, SHEL_ARRAY(const_tax_prosp_verif, selection)
			Text 105, 200, 45, 10, SHEL_ARRAY(const_room_retro_amt, selection)
			Text 160, 200, 70, 10, SHEL_ARRAY(const_room_retro_verif, selection)
			Text 255, 200, 45, 10, SHEL_ARRAY(const_room_prosp_amt, selection)
			Text 310, 200, 70, 10, SHEL_ARRAY(const_room_prosp_verif, selection)
			Text 105, 220, 45, 10, SHEL_ARRAY(const_garage_retro_amt, selection)
			Text 160, 220, 70, 10, SHEL_ARRAY(const_garage_retro_verif, selection)
			Text 255, 220, 45, 10, SHEL_ARRAY(const_garage_prosp_amt, selection)
			Text 310, 220, 70, 10, SHEL_ARRAY(const_garage_prosp_verif, selection)
			Text 105, 240, 45, 10, SHEL_ARRAY(const_subsidy_retro_amt, selection)
			Text 160, 240, 70, 10, SHEL_ARRAY(const_subsidy_retro_verif, selection)
			Text 255, 240, 45, 10, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
			Text 310, 240, 70, 10, SHEL_ARRAY(const_subsidy_prosp_verif, selection)
			PushButton 400, 265, 75, 15, "Update Information", update_information_btn
		End If

		PushButton 325, 30, 70, 13, "CLEAR ALL", clear_all_btn
	    Text 15, 30, 90, 10, "Housing Expense Paid to"
		Text 105, 50, 60, 10, "HUD Subsidized"
	    Text 225, 50, 85, 10, "Housing Expense Shared"
	    GroupBox 15, 65, 380, 190, "Housing Expense Amounts"
	    Text 105, 75, 65, 10, "Retrospective"
	    Text 255, 75, 65, 10, "Prospective"
	    Text 105, 85, 30, 10, "Amount"
	    Text 255, 85, 25, 10, "Amount"
	    Text 160, 85, 20, 10, "Verif"
	    Text 310, 85, 20, 10, "Verif"
		Text 80, 100, 20, 10, "Rent:"
	    Text 70, 120, 30, 10, "Lot Rent:"
	    Text 65, 140, 35, 10, "Mortgage:"
	    Text 65, 160, 40, 10, "Insurance:"
	    Text 75, 180, 25, 10, "Taxes:"
	    Text 75, 200, 25, 10, "Room:"
	    Text 75, 220, 30, 10, "Garage:"
	    Text 70, 240, 30, 10, "Subsidy:"

	End If
	y_pos = 25
	For the_member = 0 to UBound(SHEL_ARRAY, 2)
		If the_member = selection Then
			Text 416, y_pos + 2, 60, 10, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member)
			y_pos = y_pos + 15
		Else
			PushButton 400, y_pos, 75, 13, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member), SHEL_ARRAY(const_memb_buttons, the_member)
			y_pos = y_pos + 15
		End If
	Next
end function

function dynamic_calendar_dialog(selected_dates_array, month_to_use, text_prompt, one_date_only, disable_weekends, disable_month_change, start_date, end_date)
'--- This function creates a dynamic calendar that users can select dates from to be used in scheduleing. This is used in BULK - REVS SCRUBBER.
'~~~~~ selected_dates_array:the output array it will contain dates in MM/DD/YY format
'~~~~~ month_to_use: this can be MM/YY or MM/DD/YY format as long as it is considered a date it will work.
'~~~~~ one_date_only: this is a True/false parameter which will restrict the function to only allow one date to be selected if set to TRUE
'~~~~~ disable_weekends: this is a True/false parameter which will restrict the selection of weekends if set to TRUE
'~~~~~ disable_month_change: this is a True/false parameter which will restrict the selection of different months if set to TRUE
'~~~~~ start_date & end_date: this will provide a range of dates which cannot be selected. These are to be entered as numbers. For example start_date = 3 and end_date = 14 the days 3 through 14 will be unavailable to select
'===== Keywords: MAXIS, PRISM, create, date, calendar, dialog
	'dimming array to display the dates
	DIM display_dates_array
	DO
		full_date_to_display = ""					'resetting variables to make sure loops work properly.
		selected_dates_array = ""
		'Determining the number of days in the calendar month.
		display_month = DatePart("M", month_to_use) & "/01/" & DatePart("YYYY", month_to_use)			'Converts whatever the month_to_use variable is to a MM/01/YYYY format
		num_of_days = DatePart("D", (DateAdd("D", -1, (DateAdd("M", 1, display_month)))))								'Determines the number of days in a month by using DatePart to get the day of the last day of the month, and just using the day variable gives us a total

		'Redeclares the available dates array to be sized appropriately (with the right amount of dates) and another dimension for whether-or-not it was selected
		Redim display_dates_array(num_of_days, 0)


		'Actually displays the dialog
        dialog1 = ""
		BeginDialog dialog1, 0, 0, 280, 190, "Select Date(s)"
			Text 5, 10, 265, 50, text_prompt
			'This next part`creates a line showing the month displayed"
			Text 120, 70, 55, 10, (MonthName(DatePart("M", display_month)) & " " & DatePart("YYYY", display_month))

			'Defining the vertical position starting point for the for...next which displays dates in the dialog
			vertical_position = 85
			'This for...next displays dates in the dialog, and has checkboxes for available dates (defined in-code as dates before the 8th)
			for day_to_display = 1 to num_of_days																						'From first day of month to last day of month...
				full_date_to_display = (DatePart("M", display_month) & "/" & day_to_display & "/" & DatePart("YYYY", display_month))		'Determines the full date to display in the dialog. It needs the full date to determine the day-of-week (we obviously don't want weekends)
				horizontal_position = 15 + (40 * (WeekDay(full_date_to_display) - 1))													'horizontal position of this is the weekday numeric value (1-7) * 40, minus 1, and plus 15 pixels
				IF WeekDay(full_date_to_display) = vbSunday AND day_to_display <> 1 THEN vertical_position = vertical_position + 15		'If the day of the week isn't Sunday and the day isn't first of the month, kick the vertical position up another 15 pixels

				'This blocks out anything that's an unavailable date, currently defined as any date before the 8th. Other dates display as a checkbox.
				IF day_to_display <= end_date AND day_to_display >= start_date THEN
					Text horizontal_position, vertical_position, 30, 10, " X " & day_to_display
					display_dates_array(day_to_display, 0) = unchecked 'unchecking so selections cannot be made the range between start_date and end_date
				ELSE
					IF (disable_weekends = TRUE AND WeekDay(full_date_to_display) = vbSunday) OR (disable_weekends = TRUE AND WeekDay(full_date_to_display) = vbSaturday) THEN		'If the weekends are disabled this will change them to text rather than checkboxes
						Text horizontal_position, vertical_position, 30, 10, " X " & day_to_display
					ELSE
						CheckBox horizontal_position, vertical_position, 35, 10, day_to_display, display_dates_array(day_to_display, 0)
					END IF
				END IF
			NEXT
			ButtonGroup ButtonPressed
			OkButton 175, 170, 50, 15
			CancelButton 225, 170, 50, 15
			IF disable_month_change = FALSE THEN
				PushButton 85, 65, 20, 15, "<", prev_month_button
				PushButton 180, 65, 20, 15, ">", next_month_button
			END IF
		EndDialog

		IF one_date_only = TRUE THEN										'if only one date is allowed to be selected the script will act one way. Else it will allow for an large array of dates from a month to be build.
			DO
				selected_dates_array = ""									' declaring array at start of do loop.
				Dialog
				cancel_confirmation
				IF ButtonPressed = prev_month_button THEN month_to_use = dateadd("M", -1, month_to_use)				'changing the month_to_use based on previous or next month
				IF ButtonPressed = next_month_button THEN month_to_use = dateadd("M", 1, month_to_use)				'this will allow us to get to a new month when the dialog is rebuild.
				FOR i = 0 to num_of_days																			'checking each checkbox in the array to see what dates were selected.
					IF display_dates_array(i, 0) = 1 THEN 															'if the date has been checked
						IF len(DatePart("M", month_to_use)) = 1 THEN												'adding a leading 0 to the month if needed
							output_month = "0" & DatePart("M", month_to_use)
						ELSE
							output_month =  DatePart("M", month_to_use)
						END IF
						IF len(i) = 1 THEN 																			'building the output array with dates in MM/DD/YY format
							selected_dates_array = selected_dates_array & output_month & "/0" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
						ELSE
							selected_dates_array = selected_dates_array & output_month & "/" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
						END IF
					END IF
				NEXT
				selected_dates_array = selected_dates_array & "end"						'this will allow us to delete the extra entry in the array
				selected_dates_array = replace(selected_dates_array, ";end", "")
				selected_dates_array = Split(selected_dates_array, ";")					'splitting array
			IF Ubound(selected_dates_array) <> 0 AND (buttonpressed <> prev_month_button or buttonpressed <> next_month_button) THEN msgbox "Please select just one date."
			LOOP until Ubound(selected_dates_array) = 0
		ELSE
			Dialog
			cancel_confirmation
			IF ButtonPressed = prev_month_button THEN month_to_use = dateadd("M", -1, month_to_use)					'changing the month_to_use based on previous or next month
			IF ButtonPressed = next_month_button THEN month_to_use = dateadd("M", 1, month_to_use)					'this will allow us to get to a new month when the dialog is rebuild.
			FOR i = 0 to num_of_days																				'checking each checkbox in the array to see what dates were selected.
				IF display_dates_array(i, 0) = 1 THEN 																'if the date has been checked
					IF len(DatePart("M", month_to_use)) = 1 THEN 													'adding a leading 0 to the month if needed
						output_month = "0" & DatePart("M", month_to_use)
					ELSE
						output_month =  DatePart("M", month_to_use)
					END IF
					IF len(i) = 1 THEN 																				'building the output array with dates in MM/DD/YY format addding leading 0 to DD if needed.
						selected_dates_array = selected_dates_array & output_month & "/0" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
					ELSE
						selected_dates_array = selected_dates_array & output_month & "/" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
					END IF
				END IF
			NEXT
			selected_dates_array = selected_dates_array & "end"							'this will allow us to delete the extra entry in the array
			selected_dates_array = replace(selected_dates_array, ";end", "")
			selected_dates_array = Split(selected_dates_array, ";")						'splitting array
		END IF
	LOOP until buttonpressed = -1								'looping until someone hits the ok button, this makes the previous and next buttons work.
end function

function excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = alerts_status
end function

Function File_Exists(file_name, does_file_exist)
'--- This function will check if a file exists or not, and will output a boolean.
'~~~~~ file_name: variable for the name of the file you are searching.
'~~~~~ does_file_exist: boolean that is putput based on if file is found or not. Do not rename.
'===== Keywords: objFSO, file, boolean
    ' Set objFSO is done on lines 9-10 of Funclib
    If (objFSO.FileExists(file_name)) Then
        does_file_exist = True
    Else
      does_file_exist = False
    End If
End Function

function file_selection_system_dialog(file_selected, file_extension_restriction)
'--- This function allows a user to select a file to be opened in a script
'~~~~~ file_selected: variable for the name of the file
'~~~~~ file_extension_restriction: restricts all other file type besides allowed file type. Example: ".csv" only allows a CSV file to be accessed.
'===== Keywords: MAXIS, MMIS, PRISM, file
	'Creates a Windows Script Host object
	Set wShell=CreateObject("WScript.Shell")

	'This loops until the right file extension is selected. If it isn't specified (= ""), it'll always exit here.
	Do
		'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
		Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE ><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

		'Creates the file_selected variable from the exit
		file_selected = oExec.StdOut.ReadLine

		'If no file is selected the script will stop
		If file_selected = "" then stopscript

		'If the rightmost characters of the file selected don't match what was in the file_extension_restriction argument, it'll tell the user. Otherwise the loop (and function) ends.
		If right(file_selected, len(file_extension_restriction)) <> file_extension_restriction then MsgBox "You've entered an incorrect file type. The allowable file type is: " & file_extension_restriction & "."
	Loop until right(file_selected, len(file_extension_restriction)) = file_extension_restriction
end function

function find_variable(opening_string, variable_name, length_of_variable)
'--- This function finds a string on a page in BlueZone
'~~~~~ opening_string: string to search for
'~~~~~ variable_name: variable name of the string
'~~~~~ length_of_variable: length of the string
'===== Keywords: MAXIS, MMIS, PRISM, find
  row = 1
  col = 1
  EMSearch opening_string, row, col
  If row <> 0 then EMReadScreen variable_name, length_of_variable, row, col + len(opening_string)
end function

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result, approval_found)
'--- Function to find and navigate to the last approved version of ELIG. YOU SHOULD NAVIGATE TO THE CORRECT ELIG RESULTS FIRST
'~~~~~ cmd_row: NUMBER enter the row the COMMAND line is on (this is different for different programs)
'~~~~~ cmd_col: NUMBER enter the column thhe COMMAND Line is on (this is different for different programs)
'~~~~~ version_number: outputs a the version number that it found as the last approved
'~~~~~ version_date: outputs the process date for the version it found
'~~~~~ version_result: outputs the ELIG/INELIG information for the approved version
'~~~~~ approval_found: BOOLEAN - If an appoved version was found
'===== Keywords: MAXIS, find, ELIG
	Call write_value_and_transmit("99", cmd_row, cmd_col)			'opening the pop-up with all versions listed.
	approval_found = True											'default the approval to being found

	row = 7															'this is  the first row of the pop-up'
	Do
		EMReadScreen elig_version, 2, row, 22						'reading the information about the version
		EMReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do				'If it was 'APPROVED' this is the most recent version that is appoved and we have all the information

		row = row + 1												'go to the next row'
	Loop until approval_status = ""									'once we hit a blank, there are no more vversions

	Call clear_line_of_text(18, 54)									''erasing the version entry as it defaults when the pop-up opens
	If approval_status = "" Then									'if no APPROVAL was found, then we leave without navigating and changing the found to false
		approval_found = false
		PF3
	Else
		Call write_value_and_transmit(elig_version, 18, 54)			'if an approval was found, we navigate to it and save the information to the output variables.
		version_number = "0" & elig_version
		version_date = elig_date
		version_result = elig_result
	End If
end function

function find_MAXIS_worker_number(x_number)
'--- This function finds a MAXIS worker's X number
'~~~~~ x_number: worker number variable
'===== Keywords: MAXIS, worker number
	EMReadScreen SELF_check, 4, 2, 50		'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN				'if on the self screen then x # is read from coordinates
		EMReadScreen x_number, 7, 22, 8
	ELSE
		Call find_variable("PW: ", x_number, 7)	'if not, then the PW: variable is searched to find the worker #
		If isnumeric(MAXIS_worker_number) = true then 	 'making sure that the worker # is a number
			MAXIS_worker_number = x_number				'delcares the MAXIS_worker_number to be the x_number
		End if
	END if
end function

function find_user_name(the_person_running_the_script)
'--- This function finds the outlook name of the person using the script
'~~~~~ the_person_running_the_script:the variable for the person's name to output
'===== Keywords: MAXIS, worker name, email signature
	Set objOutlook = CreateObject("Outlook.Application")
	Set the_person_running_the_script = objOutlook.GetNamespace("MAPI").CurrentUser
	the_person_running_the_script = the_person_running_the_script & ""
	Set objOutlook = Nothing
end function

'This function fixes the case for a phrase. For example, "ROBERT P. ROBERTSON" becomes "Robert P. Robertson".
'	It capitalizes the first letter of each word.
function fix_case(phrase_to_split, smallest_length_to_skip)										'Ex: fix_case(client_name, 3), where 3 means skip words that are 3 characters or shorter
	phrase_to_split = split(phrase_to_split)													'splits phrase into an array
	For each word in phrase_to_split															'processes each word independently
		If word <> "" then																		'Skip blanks
			first_character = ucase(left(word, 1))												'grabbing the first character of the string, making uppercase and adding to variable
			remaining_characters = LCase(right(word, len(word) -1))								'grabbing the remaining characters of the string, making lowercase and adding to variable
			If len(word) > smallest_length_to_skip then											'skip any strings shorter than the smallest_length_to_skip variable
				output_phrase = output_phrase & first_character & remaining_characters & " "	'output_phrase is the output of the function, this combines the first_character and remaining_characters
			Else
				output_phrase = output_phrase & word & " "										'just pops the whole word in if it's shorter than the smallest_length_to_skip variable
			End if
		End if
	Next
	phrase_to_split = output_phrase																'making the phrase_to_split equal to the output, so that it can be used by the rest of the script.
end function

function fix_case_for_name(name_variable)
'--- This function takes in a client's name and outputs the name (accounting for hyphenated surnames) with Ucase first character & and lcase the rest. This is like fix_case but this function is a bit more specific for names
'~~~~~ name_variable: should be client_name for function to work
'===== Keywords: MAXIS, MMIS, PRISM, name, case
	name_variable = split(name_variable, " ")
	FOR EACH client_name IN name_variable
		IF client_name <> "" THEN
			IF InStr(client_name, "-") = 0 THEN
				client_name = UCASE(left(client_name, 1)) & LCASE(right(client_name, len(client_name) - 1))
				output_variable = output_variable & " " & client_name
			ELSE				'When the client has a hyphenated surname
				hyphen_location = InStr(client_name, "-")
				first_part = left(client_name, hyphen_location - 1)
				first_part = UCASE(left(first_part, 1)) & LCASE(right(first_part, len(first_part) - 1))
				second_part = right(client_name, len(client_name) - hyphen_location)
				second_part = UCASE(left(second_part, 1)) & LCASE(right(second_part, len(second_part) - 1))
				output_variable = output_variable & " " & first_part & "-" & second_part
			END IF
		END IF
	NEXT
	name_variable = output_variable
end function

Function generate_client_list(list_for_dropdown, initial_text)
'--- This function creates a variable formatted for a DropListBox or ComboBox in a dialog to have all the clients on a case as an option.
'~~~~~ list_for_dropdown: the variable to put in the dialog for the list
'~~~~~ initial_text: the words to have in the top position of the list
'===== Keywords: MAXIS, DIALOG, CLIENTS
	memb_row = 5
    list_for_dropdown = initial_text

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next
End Function

function gather_case_benefits_details(months_to_go_back, run_from_client_contact)
'--- This function reviews a case to read and display issuance information for current and past months.
'~~~~~ months_to_go_back: A number that counts how many past months to review
'~~~~~ run_from_client_contact: determining if this is being run from client contact
'===== Keywords: MAXIS, DIALOG, CLIENTS

	'setting the constants for the array used wihin the function. These are not passed through.
    const fn_footer_month_const    = 0
    const fn_footer_year_const     = 1
    const fn_snap_issued_const     = 2
    const fn_snap_recoup_const     = 3
    const fn_ga_issued_const       = 4
    const fn_ga_recoup_const       = 5
    const fn_msa_issued_const      = 6
    const fn_msa_recoup_const      = 7
    const fn_mf_mf_issued_const    = 8
    const fn_mf_mf_recoup_const    = 9
    const fn_mf_fs_issued_const    = 10
    const fn_mf_hg_issued_const    = 11
    const fn_dwp_issued_const      = 12
    const fn_dwp_recoup_const      = 13
    const fn_emer_issued_const     = 14
    const fn_emer_prog_const       = 15
    const fn_grh_issued_const      = 16
    const fn_grh_recoup_const      = 17
    const fn_no_issuance_const     = 18
    const fn_last_const            = 25

    Dim ISSUED_BENEFITS_ARRAY()			'defning the array used in the function to save the past benefits amounts

    complete_script_run_btn = 50				'defning the button numbers to ensure they don't get mixed up
    run_pa_verif_reqquest_btn = 100
    run_client_contact_btn = 110
    change_lookback_month_count_btn = 200
    elig_fs_btn = 300
    elig_ga_btn = 310
    elig_msa_btn = 320
    elig_mfip_btn = 340
    elig_dwp_btn = 350
    elig_grh_btn = 360
    view_by_month_btn = 400
    view_by_prog_btn = 410

	'this function will gather details from INQQB and save them into the defined array.
    Call read_inqb_for_all_issuances(months_to_go_back, beginning_footer_month, ISSUED_BENEFITS_ARRAY, fn_footer_month_const, fn_footer_year_const, fn_snap_issued_const, fn_snap_recoup_const, fn_ga_issued_const, fn_ga_recoup_const, fn_msa_issued_const, fn_msa_recoup_const, fn_mf_mf_issued_const, fn_mf_mf_recoup_const, fn_mf_fs_issued_const, fn_mf_hg_issued_const, fn_dwp_issued_const, fn_dwp_recoup_const, fn_emer_issued_const, fn_emer_prog_const, fn_grh_issued_const, fn_grh_recoup_const, fn_no_issuance_const, fn_last_const, snap_found, ga_found, msa_found, mfip_found, dwp_found, grh_found)

    MAXIS_footer_month = CM_plus_1_mo                              'setting the footermonth to the current month
    MAXIS_footer_year = CM_plus_1_yr

	'determining the program information
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
    Call Back_to_SELF

	'This section will read ELIG for CM+1 for any program that is ACTIVE or will be ACTIVE next month. This will display the ongoing benefit in the dialog
    If snap_status = "ACTIVE" or snap_status = "APP OPEN" Then					'SNAP'
        call navigate_to_MAXIS_screen("ELIG", "FS  ")
        Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("FSSM", 19, 70)

        EMReadScreen snap_benefit_monthly_fs_allotment, 10, 8, 71
        EMReadScreen snap_benefit_prorated_amt, 		10, 9, 71
        EMReadScreen snap_benefit_prorated_date,		8, 9, 58
        EMReadScreen snap_benefit_amt, 					10, 13, 71

        snap_benefit_monthly_fs_allotment = trim(snap_benefit_monthly_fs_allotment)
        snap_benefit_prorated_amt = trim(snap_benefit_prorated_amt)
        snap_benefit_prorated_date = trim(snap_benefit_prorated_date)
        ongoing_snap_amount = trim(snap_benefit_amt)

        Call Back_to_SELF
    End If
    If ga_status = "ACTIVE" or ga_status = "APP OPEN" Then						'GA
        call navigate_to_MAXIS_screen("ELIG", "GA  ")
        Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("GASM", 20, 70)

        EMReadScreen ga_elig_summ_monthly_grant, 10, 9, 71
        EMReadScreen ga_elig_summ_amount_to_be_paid, 10, 14, 71

        ga_elig_summ_monthly_grant = trim(ga_elig_summ_monthly_grant)
        ongoing_ga_amount = trim(ga_elig_summ_amount_to_be_paid)

        Call Back_to_SELF
    End If
    If msa_status = "ACTIVE" or msa_status = "APP OPEN" Then					'MMSA
        call navigate_to_MAXIS_screen("ELIG", "MSA ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("MSSM", 20, 71)

        EMReadScreen msa_elig_summ_grant, 9, 11, 72
        EMReadScreen msa_elig_summ_current_payment, 9, 17, 72

        msa_elig_summ_grant = trim(msa_elig_summ_grant)
        ongoing_msa_amount = trim(msa_elig_summ_current_payment)

        Call Back_to_SELF
    End If
    If mfip_status = "ACTIVE" or mfip_status = "APP OPEN" Then					'MFIP
        call navigate_to_MAXIS_screen("ELIG", "MFIP")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("MFSM", 20, 71)

        EMReadScreen mfip_case_summary_grant_amount, 10, 11, 71
        EMReadScreen mfip_case_summary_net_grant_amount, 10, 13, 71
        EMReadScreen mfip_case_summary_cash_portion, 10, 14, 71
        EMReadScreen mfip_case_summary_food_portion, 10, 15, 71
        EMReadScreen mfip_case_summary_housing_grant, 10, 16, 71

        mfip_case_summary_grant_amount = trim(mfip_case_summary_grant_amount)
        mfip_case_summary_net_grant_amount = trim(mfip_case_summary_net_grant_amount)
        ongoing_mfip_cash_amount = trim(mfip_case_summary_cash_portion)
        ongoing_mfip_food_amount = trim(mfip_case_summary_food_portion)
        ongoing_mfip_hg_amount = trim(mfip_case_summary_housing_grant)

        Call Back_to_SELF
    End If
    If dwp_status = "ACTIVE" or dwp_status = "APP OPEN" Then					'DWP
        call navigate_to_MAXIS_screen("ELIG", "DWP ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("DWSM", 20, 71)


        EMReadScreen dwp_case_summary_grant_amount, 10, 10, 71
        EMReadScreen dwp_case_summary_net_grant_amount, 10, 12, 71
        EMReadScreen dwp_case_summary_shelter_benefit_portion, 10, 13, 71
        EMReadScreen dwp_case_summary_personal_needs_portion, 10, 14, 71

        dwp_case_summary_grant_amount = trim(dwp_case_summary_grant_amount)
        ongoing_dwp_amount = trim(dwp_case_summary_net_grant_amount)
        dwp_case_summary_shelter_benefit_portion = trim(dwp_case_summary_shelter_benefit_portion)
        dwp_case_summary_personal_needs_portion = trim(dwp_case_summary_personal_needs_portion)

        Call Back_to_SELF
    End If
    If grh_status = "ACTIVE" or ga_stagrh_statustus = "APP OPEN" Then			'GRH
        call navigate_to_MAXIS_screen("ELIG", "GRH ")
        Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        Call write_value_and_transmit("GRSM", 20, 71)

        EMReadScreen ongoing_grh_amount_one, 		9, 12, 31
        EMReadScreen ongoing_grh_amount_two, 		9, 12, 50

        ongoing_grh_amount_one = trim(ongoing_grh_amount_one)
        ongoing_grh_amount_two = trim(ongoing_grh_amount_two)

        Call Back_to_SELF
    End If

	'the dialog can show information either sorted by program or by month. This defaults it to sorting by program and sets the functionality to switch.
    view_by_program = 1
    view_by_month = 2
    dialog_history_view = view_by_program
	'Looping to show the dialog
    Do
        Do
            programs_with_no_cm_plus_one_issuance = ""							'setting up a list of programs with no issuance next month'
            If ongoing_snap_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", SNAP"
            If ongoing_ga_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", GA"
            If ongoing_msa_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", MSA"
            If ongoing_mfip_cash_amount = "" and ongoing_mfip_food_amount = "" and ongoing_mfip_hg_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", MFIP"
            If ongoing_dwp_amount = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", DWP"
            If ongoing_grh_amount_one = "" Then programs_with_no_cm_plus_one_issuance = programs_with_no_cm_plus_one_issuance & ", GRH"
            if left(programs_with_no_cm_plus_one_issuance, 1) = "," Then programs_with_no_cm_plus_one_issuance = right(programs_with_no_cm_plus_one_issuance, len(programs_with_no_cm_plus_one_issuance)-1)
            programs_with_no_cm_plus_one_issuance = trim(programs_with_no_cm_plus_one_issuance)

            programs_with_no_past_issuance = ""									'setting up a list of programs with no issuance in past months
            If snap_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", SNAP"
            If ga_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", GA"
            If msa_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", MSA"
            If mfip_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", MFIP"
            If dwp_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", DWP"
            If grh_found = False Then programs_with_no_past_issuance = programs_with_no_past_issuance & ", GRH"
            if left(programs_with_no_past_issuance, 1) = "," Then programs_with_no_past_issuance = right(programs_with_no_past_issuance, len(programs_with_no_past_issuance)-1)
            programs_with_no_past_issuance = trim(programs_with_no_past_issuance)

			'This part is to determine the size of the groupbox and the dialog length
            prog_count = 1
            If mfip_found = True Then prog_count = prog_count + 1
            If snap_found = True Then prog_count = prog_count + 1
            If ga_found = True Then prog_count = prog_count + 1
            If msa_found = True Then prog_count = prog_count + 1
            If dwp_found = True Then prog_count = prog_count + 1
            If grh_found = True Then prog_count = prog_count + 1
            prog_len_multiplier = prog_count/2
            prog_len_multiplier = INT(prog_len_multiplier)

            If dialog_history_view = view_by_program Then
                grp_bx_len = 45
                grp_bx_len = grp_bx_len + prog_len_multiplier * 15
                no_issuance_months = ""
                For each_inqb_item = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                    If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_inqb_item) = False Then grp_bx_len = grp_bx_len + 10 * prog_len_multiplier
                    If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_inqb_item) = True Then no_issuance_months = no_issuance_months & ", " & ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_inqb_item) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_inqb_item)
                Next
                If left(no_issuance_months, 1) = "," Then no_issuance_months = right(no_issuance_months, len(no_issuance_months)-1)
                no_issuance_months = trim(no_issuance_months)
                If no_issuance_months <> "" Then grp_bx_len = grp_bx_len + 15
            End If


            If dialog_history_view = view_by_month Then
                grp_bx_len = 55
                no_issuance_months = ""
                For each_inqb_item = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                    If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_inqb_item) = False Then grp_bx_len = grp_bx_len + 10
                    If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_inqb_item) = True Then no_issuance_months = no_issuance_months & ", " & ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_inqb_item) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_inqb_item)
                Next
                If left(no_issuance_months, 1) = "," Then no_issuance_months = right(no_issuance_months, len(no_issuance_months)-1)
                no_issuance_months = trim(no_issuance_months)
                If no_issuance_months <> "" Then grp_bx_len = grp_bx_len + 15
            End If
            dlg_len = 160 + grp_bx_len

			'defining the dialog
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 441, dlg_len, "Case " & MAXIS_case_number & " Issuance Details"
             ButtonGroup ButtonPressed
                EditBox 500, 600, 50, 15, fake_edit_box
                GroupBox 10, 10, 420, 105, "Current Approval Amounts"
                Text 20, 25, 180, 10, "Based on ELIG for current month plus 1  (" & CM_plus_1_mo & "/" & CM_plus_1_yr &")"

                x_pos = 30
                If ongoing_snap_amount <> "" Then
                    Text x_pos, 40, 25, 10, "SNAP"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_snap_amount
                    PushButton x_pos, 80, 35, 10, "ELIG/FS", elig_fs_btn
                    x_pos = x_pos + 60
                End If
                If ongoing_ga_amount <> "" Then
                    Text x_pos, 40, 25, 10, "GA"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_ga_amount
                    PushButton x_pos, 80, 35, 10, "ELIG/GA", elig_ga_btn
                    x_pos = x_pos + 60
                End If
                If ongoing_msa_amount <> "" Then
                    Text x_pos, 40, 25, 10, "MSA"
                    Text x_pos+5, 50, 30, 10, "$ " & ongoing_msa_amount
                    PushButton x_pos, 80, 40, 10, "ELIG/MSA", elig_msa_btn
                    x_pos = x_pos + 65
                End If
                If ongoing_mfip_cash_amount <> "" or ongoing_mfip_food_amount <> "" or ongoing_mfip_hg_amount <> "" Then
                    Text x_pos, 40, 25, 10, "MFIP"
                    Text x_pos+5, 50, 60, 10, "MF - $ " & ongoing_mfip_cash_amount
                    Text x_pos+5, 60, 60, 10, "FS - $ " & ongoing_mfip_food_amount
                    Text x_pos+5, 70, 60, 10, "HG - $ " & ongoing_mfip_hg_amount
                    PushButton x_pos, 80, 45, 10, "ELIG/MFIP", elig_mfip_btn
                    x_pos = x_pos + 70
                End If
                If ongoing_dwp_amount <> "" Then
                    Text x_pos, 40, 25, 10, "DWP"
                    Text 300, 50, 30, 10, "$ " & ongoing_dwp_amount
                    PushButton x_pos, 80, 40, 10, "ELIG/DWP", elig_dwp_btn
                    x_pos = x_pos + 65
                End If
                If ongoing_grh_amount_one <> "" Then
                    Text x_pos, 40, 25, 10, "GRH"
                    Text x_pos+5, 50, 45, 10, "One - $ " & ongoing_grh_amount_one
                    if ongoing_grh_amount_two <> "" Then Text x_pos+5, 60, 45, 10, "Two - $ " & ongoing_grh_amount_two
                    PushButton x_pos, 80, 40, 10, "ELIG/GRH", elig_grh_btn
                End If
                Text 140, 100, 280, 10, "No Eligibility for: " & programs_with_no_cm_plus_one_issuance
                '
                GroupBox 10, 125, 420, grp_bx_len, "Past Issuance Amounts"
                Text 25, 140, 200, 10, "Information going back " & months_to_go_back & " months from " & beginning_footer_month & " to " & CM_mo & "/" & CM_yr
                PushButton 265, 135, 160, 15, "Change the Number of Months to Go Back", change_lookback_month_count_btn

                x_pos = 30
                y_pos = 155
                no_issue_month_found = false
                If no_issuance_months <> "" Then
                    no_issue_month_found = True
                    Text 30, y_pos, 200, 10, "No issuances for " & no_issuance_months
                    y_pos = y_pos + 15
                End If

                If dialog_history_view = view_by_program Then
                    y_pos_reset = y_pos

                    If mfip_found = True Then
                        Text x_pos, y_pos, 35, 10, "MFIP"
                        y_pos = y_pos + 10

                        For each_mf_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_mf_issue) = False Then
                                month_info = ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_mf_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_mf_issue)
                                If ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, each_mf_issue) = "" and ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, each_mf_issue) = "" and ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, each_mf_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, month_info & "  .  . None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_mf_mf_recoup_const, each_mf_issue) <> "" Then Text x_pos+10, y_pos, 200, 10, month_info & "  .  . Cash $ " & ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, each_mf_issue) & "  -  Food $  " & ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, each_mf_issue) & "  -  HG $  " & ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, each_mf_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_mf_mf_recoup_const, each_mf_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_mf_mf_recoup_const, each_mf_issue) = "" Then Text x_pos+10, y_pos, 200, 10, month_info & "  .  . Cash $ " & ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, each_mf_issue) & "  -  Food $  " & ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, each_mf_issue) & "  -  HG $  " & ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, each_mf_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If

                    If snap_found = True Then
                        Text x_pos, y_pos, 35, 10, "SNAP"
                        y_pos = y_pos + 10

                        For each_fs_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_fs_issue) = False Then
                                If ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, each_fs_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_fs_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_fs_issue) & "  .  .  .  None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_snap_recoup_const, each_fs_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_fs_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_fs_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, each_fs_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_snap_recoup_const, each_fs_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_snap_recoup_const, each_fs_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_fs_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_fs_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, each_fs_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If

                    If ga_found = True Then
                        Text x_pos, y_pos, 35, 10, "GA"
                        y_pos = y_pos + 10

                        For each_ga_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_ga_issue) = False Then
                                If ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, each_ga_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_ga_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_ga_issue) & "  .  .  .  None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_ga_recoup_const, each_ga_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_ga_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_ga_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, each_ga_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_ga_recoup_const, each_ga_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_ga_recoup_const, each_ga_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_ga_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_ga_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, each_ga_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If

                    If msa_found = True Then
                        Text x_pos, y_pos, 35, 10, "MSA"
                        y_pos = y_pos + 10

                        For each_msa_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_msa_issue) = False Then
                                If ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, each_msa_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_msa_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_msa_issue) & "  .  .  .  None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_msa_recoup_const, each_msa_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_msa_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_msa_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, each_msa_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_msa_recoup_const, each_msa_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_msa_recoup_const, each_msa_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_msa_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_msa_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, each_msa_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If

                    If dwp_found = True Then
                        Text x_pos, y_pos, 35, 10, "DWP"
                        y_pos = y_pos + 10

                        For each_dwp_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_dwp_issue) = False Then
                                If ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, each_dwp_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_dwp_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_dwp_issue) & "  .  .  .  None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_dwp_recoup_const, each_dwp_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_dwp_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_dwp_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, each_dwp_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_dwp_recoup_const, each_dwp_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_dwp_recoup_const, each_dwp_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_dwp_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_dwp_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, each_dwp_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If

                    If grh_found = True Then
                        Text x_pos, y_pos, 35, 10, "GRH"
                        y_pos = y_pos + 10

                        For each_grh_issue = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                            If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_grh_issue) = False Then
                                If ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, each_grh_issue) = "" Then
                                    Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_grh_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_grh_issue) & "  .  .  .  None"
                                Else
                                    If ISSUED_BENEFITS_ARRAY(fn_grh_recoup_const, each_grh_issue) <> "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_grh_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_grh_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, each_grh_issue) & "        Recoup: $ " & ISSUED_BENEFITS_ARRAY(fn_grh_recoup_const, each_grh_issue)
                                    If ISSUED_BENEFITS_ARRAY(fn_grh_recoup_const, each_grh_issue) = "" Then Text x_pos+10, y_pos, 150, 10, ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_grh_issue) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_grh_issue) & "  .  .  .  $ " & ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, each_grh_issue)
                                End If
                                y_pos = y_pos + 10
                            End If
                        Next
                        y_pos_end = y_pos
                        If x_pos = 30 Then
                            x_pos = 260
                            y_pos = y_pos_reset
                        ElseIf x_pos = 260 Then
                            x_pos = 30
                            y_pos_reset = y_pos + 5
                            y_pos = y_pos_reset
                        End If
                    End If
                End If

                If dialog_history_view = view_by_month Then
                    For each_issue_mo = 0 to UBound(ISSUED_BENEFITS_ARRAY, 2)
                        If ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, each_issue_mo) = False Then
                            month_info = ISSUED_BENEFITS_ARRAY(fn_footer_month_const, each_issue_mo) & "/" & ISSUED_BENEFITS_ARRAY(fn_footer_year_const, each_issue_mo)
                            beneits_info = ""
                            If ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, each_issue_mo) <> "" OR ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, each_issue_mo) <> "" OR ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  MFIP - (MF $ " & ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, each_issue_mo) & ", FS $  " & ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, each_issue_mo) & ", HG $  " & ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, each_issue_mo) & ")    |  "
                            If ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  SNAP - $ " & ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, each_issue_mo) & "    |  "
                            If ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  GA - $ " & ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, each_issue_mo) & "    |  "
                            If ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  MSA - $ " & ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, each_issue_mo) & "    |  "
                            If ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  DWP - $ " & ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, each_issue_mo) & "    |  "
                            If ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, each_issue_mo) <> "" Then beneits_info = beneits_info & "  GRH - $ " & ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, each_issue_mo) & "    |  "
                            If right(beneits_info, 7) = "    |  " Then beneits_info = left(beneits_info, len(beneits_info)-7)
                            Text 20, y_pos, 400, 10, month_info & "  .  .  .  " & beneits_info

                            y_pos = y_pos + 10

                        End If
                    Next
                    y_pos = y_pos + 5
                End if
                Text 135, 110+grp_bx_len, 295, 10, "No Issuances for: " & programs_with_no_past_issuance

                If dialog_history_view = view_by_program Then PushButton 20, 110+grp_bx_len, 100, 12, "View History by Month", view_by_month_btn
                If dialog_history_view = view_by_month Then PushButton 20, 110+grp_bx_len, 100, 12, "View History by Program", view_by_prog_btn

                If run_from_client_contact = False Then
                    PushButton 15, dlg_len-25, 160, 15, "Run NOTICES - PA Verifications Request", run_pa_verif_reqquest_btn
                    PushButton 185, dlg_len-25, 135, 15, "Run NOTES - Client Contact", run_client_contact_btn
					PushButton 330, dlg_len-25, 100, 15, "End Script Run", complete_script_run_btn
                End If
				If run_from_client_contact = True Then
					PushButton 330, dlg_len-25, 100, 15, "Return to Client Contact", complete_script_run_btn
				End If
            EndDialog

            dialog Dialog1
            If run_from_client_contact = False Then cancel_without_confirmation
			If run_from_client_contact = True Then cancel_confirmation

            If ButtonPressed = view_by_month_btn Then dialog_history_view = view_by_month
            If ButtonPressed = view_by_prog_btn Then dialog_history_view = view_by_program

            If ButtonPressed = run_pa_verif_reqquest_btn Then Call run_from_GitHub(script_repository & "notices/pa-verif-request.vbs" )
            If ButtonPressed = run_client_contact_btn Then Call run_from_GitHub(script_repository & "notes/client-contact.vbs" )
            If ButtonPressed = complete_script_run_btn Then ButtonPressed = -1

            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = False			'anything after here within the function moves in MAXIS and we must be passworded in.

        If ButtonPressed = change_lookback_month_count_btn Then					'if the button is pressed to change the number of months, a small dialog appars to enter a new number in.
            months_to_go_back = months_to_go_back & ""							'dialogs need strings
            Do
                Do
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 141, 80, "Lookback Months Update"
                      EditBox 100, 35, 25, 15, months_to_go_back
                      ButtonGroup ButtonPressed
                        OkButton 75, 55, 50, 15
                      Text 10, 10, 130, 15, "How many months should the script search for issuance amounts?"
                      Text 30, 40, 70, 10, "Months to look back:"
                    EndDialog

                    dialog Dialog1
                    cancel_confirmation

                    If IsNumeric(months_to_go_back) = False Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please review the number of months you have entered." & vbCr & vbCr &"This needs to be a number."

                Loop until IsNumeric(months_to_go_back) = True
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = False
            months_to_go_back = months_to_go_back * 1							'math needs number

			'rereading INQB and creating a new array of the past issuance information
            Call read_inqb_for_all_issuances(months_to_go_back, beginning_footer_month, ISSUED_BENEFITS_ARRAY, fn_footer_month_const, fn_footer_year_const, fn_snap_issued_const, fn_snap_recoup_const, fn_ga_issued_const, fn_ga_recoup_const, fn_msa_issued_const, fn_msa_recoup_const, fn_mf_mf_issued_const, fn_mf_mf_recoup_const, fn_mf_fs_issued_const, fn_mf_hg_issued_const, fn_dwp_issued_const, fn_dwp_recoup_const, fn_emer_issued_const, fn_emer_prog_const, fn_grh_issued_const, fn_grh_recoup_const, fn_no_issuance_const, fn_last_const, snap_found, ga_found, msa_found, mfip_found, dwp_found, grh_found)

            ButtonPressed = change_lookback_month_count_btn						'reset the button to make sure the dialog doesn't end.
        End If

        MAXIS_footer_month = CM_plus_1_mo                              'setting the footermonth to the current month
        MAXIS_footer_year = CM_plus_1_yr
		'Goind to ELIG for any program if the ELIG button was pressed.
        If ButtonPressed = elig_fs_btn Then
            call navigate_to_MAXIS_screen("ELIG", "FS  ")
            Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If

        If ButtonPressed = elig_ga_btn Then
            call navigate_to_MAXIS_screen("ELIG", "GA  ")
            Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If

        If ButtonPressed = elig_msa_btn Then
            call navigate_to_MAXIS_screen("ELIG", "MSA ")
            Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If

        If ButtonPressed = elig_mfip_btn Then
            call navigate_to_MAXIS_screen("ELIG", "MFIP")
            Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If

        If ButtonPressed = elig_dwp_btn Then
            call navigate_to_MAXIS_screen("ELIG", "DWP ")
            Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If

        If ButtonPressed = elig_grh_btn Then
            call navigate_to_MAXIS_screen("ELIG", "GRH ")
            Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
        End If
    Loop until ButtonPressed = -1												'This will keep going until 'Enter' is pressed or one of the 'End' buttons
	'No output - the dialog just ends - this is why there is no pass through
end function

function get_county_code()
'--- This function determines county_name from worker_county_code, and asks for it if it's blank
'===== Keywords: MAXIS, MMIS, PRISM, county
	If left(code_from_installer, 2) = "PT" then 'special handling for Pine Tech
		worker_county_code = "PWVTS"
	Else
		If worker_county_code = "MULTICOUNTY" or worker_county_code = "" then 		'If the user works for many counties (i.e. SWHHS) or isn't assigned (i.e. a scriptwriter) it asks.
			Do
				two_digit_county_code_variable = inputbox("Select the county to proxy as. Ex: ''01''")
				If two_digit_county_code_variable = "" then stopscript
				If len(two_digit_county_code_variable) <> 2 or isnumeric(two_digit_county_code_variable) = False then MsgBox "Your county proxy code should be two digits and numeric."
			Loop until len(two_digit_county_code_variable) = 2 and isnumeric(two_digit_county_code_variable) = True
			worker_county_code = "X1" & two_digit_county_code_variable
			If two_digit_county_code_variable = "91" then worker_county_code = "PW"	'For DHS folks without proxy
		End If
	End if

    'Determining county name
    if worker_county_code = "x101" then
        county_name = "Aitkin County"
    elseif worker_county_code = "x102" then
        county_name = "Anoka County"
    elseif worker_county_code = "x103" then
        county_name = "Becker County"
    elseif worker_county_code = "x104" then
        county_name = "Beltrami County"
    elseif worker_county_code = "x105" then
        county_name = "Benton County"
    elseif worker_county_code = "x106" then
        county_name = "Big Stone County"
    elseif worker_county_code = "x107" then
        county_name = "Blue Earth County"
    elseif worker_county_code = "x108" then
        county_name = "Brown County"
    elseif worker_county_code = "x109" then
        county_name = "Carlton County"
    elseif worker_county_code = "x110" then
        county_name = "Carver County"
    elseif worker_county_code = "x111" then
        county_name = "Cass County"
    elseif worker_county_code = "x112" then
        county_name = "Chippewa County"
    elseif worker_county_code = "x113" then
        county_name = "Chisago County"
    elseif worker_county_code = "x114" then
        county_name = "Clay County"
    elseif worker_county_code = "x115" then
        county_name = "Clearwater County"
    elseif worker_county_code = "x116" then
        county_name = "Cook County"
    elseif worker_county_code = "x117" then
        county_name = "Cottonwood County"
    elseif worker_county_code = "x118" then
        county_name = "Crow Wing County"
    elseif worker_county_code = "x119" then
        county_name = "Dakota County"
    elseif worker_county_code = "x120" then
        county_name = "Dodge County"
    elseif worker_county_code = "x121" then
        county_name = "Douglas County"
    elseif worker_county_code = "x122" then
        county_name = "Faribault County"
    elseif worker_county_code = "x123" then
        county_name = "Fillmore County"
    elseif worker_county_code = "x124" then
        county_name = "Freeborn County"
    elseif worker_county_code = "x125" then
        county_name = "Goodhue County"
    elseif worker_county_code = "x126" then
        county_name = "Grant County"
    elseif worker_county_code = "x127" then
        county_name = "Hennepin County"
    elseif worker_county_code = "x128" then
        county_name = "Houston County"
    elseif worker_county_code = "x129" then
        county_name = "Hubbard County"
    elseif worker_county_code = "x130" then
        county_name = "Isanti County"
    elseif worker_county_code = "x131" then
        county_name = "Itasca County"
    elseif worker_county_code = "x132" then
        county_name = "Jackson County"
    elseif worker_county_code = "x133" then
        county_name = "Kanabec County"
    elseif worker_county_code = "x134" then
        county_name = "Kandiyohi County"
    elseif worker_county_code = "x135" then
        county_name = "Kittson County"
    elseif worker_county_code = "x136" then
        county_name = "Koochiching County"
    elseif worker_county_code = "x137" then
        county_name = "Lac Qui Parle County"
    elseif worker_county_code = "x138" then
        county_name = "Lake County"
    elseif worker_county_code = "x139" then
        county_name = "Lake of the Woods County"
    elseif worker_county_code = "x140" then
        county_name = "LeSueur County"
    elseif worker_county_code = "x141" then
        county_name = "Lincoln County"
    elseif worker_county_code = "x142" then
        county_name = "Lyon County"
    elseif worker_county_code = "x143" then
        county_name = "Mcleod County"
    elseif worker_county_code = "x144" then
        county_name = "Mahnomen County"
    elseif worker_county_code = "x145" then
        county_name = "Marshall County"
    elseif worker_county_code = "x146" then
        county_name = "Martin County"
    elseif worker_county_code = "x147" then
        county_name = "Meeker County"
    elseif worker_county_code = "x148" then
        county_name = "Mille Lacs County"
    elseif worker_county_code = "x149" then
        county_name = "Morrison County"
    elseif worker_county_code = "x150" then
        county_name = "Mower County"
    elseif worker_county_code = "x151" then
        county_name = "Murray County"
    elseif worker_county_code = "x152" then
        county_name = "Nicollet County"
    elseif worker_county_code = "x153" then
        county_name = "Nobles County"
    elseif worker_county_code = "x154" then
        county_name = "Norman County"
    elseif worker_county_code = "x155" then
        county_name = "Olmsted County"
    elseif worker_county_code = "x156" then
        county_name = "Otter Tail County"
    elseif worker_county_code = "x157" then
        county_name = "Pennington County"
    elseif worker_county_code = "x158" then
        county_name = "Pine County"
    elseif worker_county_code = "x159" then
        county_name = "Pipestone County"
    elseif worker_county_code = "x160" then
        county_name = "Polk County"
    elseif worker_county_code = "x161" then
        county_name = "Pope County"
    elseif worker_county_code = "x162" then
        county_name = "Ramsey County"
    elseif worker_county_code = "x163" then
        county_name = "Red Lake County"
    elseif worker_county_code = "x164" then
        county_name = "Redwood County"
    elseif worker_county_code = "x165" then
        county_name = "Renville County"
    elseif worker_county_code = "x166" then
        county_name = "Rice County"
    elseif worker_county_code = "x167" then
        county_name = "Rock County"
    elseif worker_county_code = "x168" then
        county_name = "Roseau County"
    elseif worker_county_code = "x169" then
        county_name = "St. Louis County"
    elseif worker_county_code = "x170" then
        county_name = "Scott County"
    elseif worker_county_code = "x171" then
        county_name = "Sherburne County"
    elseif worker_county_code = "x172" then
        county_name = "Sibley County"
    elseif worker_county_code = "x173" then
        county_name = "Stearns County"
    elseif worker_county_code = "x174" then
        county_name = "Steele County"
    elseif worker_county_code = "x175" then
        county_name = "Stevens County"
    elseif worker_county_code = "x176" then
        county_name = "Swift County"
    elseif worker_county_code = "x177" then
        county_name = "Todd County"
    elseif worker_county_code = "x178" then
        county_name = "Traverse County"
    elseif worker_county_code = "x179" then
        county_name = "Wabasha County"
    elseif worker_county_code = "x180" then
        county_name = "Wadena County"
    elseif worker_county_code = "x181" then
        county_name = "Waseca County"
    elseif worker_county_code = "x182" then
        county_name = "Washington County"
    elseif worker_county_code = "x183" then
        county_name = "Watonwan County"
    elseif worker_county_code = "x184" then
        county_name = "Wilkin County"
    elseif worker_county_code = "x185" then
        county_name = "Winona County"
    elseif worker_county_code = "x186" then
        county_name = "Wright County"
    elseif worker_county_code = "x187" then
        county_name = "Yellow Medicine County"
    elseif worker_county_code = "x188" then
        county_name = "Mille Lacs Band"
    elseif worker_county_code = "x192" then
        county_name = "White Earth Nation"
    elseif worker_county_code = "PWVTS" then
    	county_name = "Pine Tech"
    end if
end function

function get_state_name_from_state_code(state_code, state_name, include_state_code)
'--- function to use the 2 digit code to get the name of the state
'~~~~~ state_code: string - 2 digit state code
'~~~~~ state_name: string - this will output the name of the state based on the entered code
'~~~~~ include_state_code: boolean - enter TRUE here to have the state_name output with the state_code, eg. MN Minnesota
'===== Keywords: format, variable
    If state_code = "NB" Then state_name = "MN Newborn"							'This is the list of all the states connected to the code.
    If state_code = "FC" Then state_name = "Foreign Country"
    If state_code = "UN" Then state_name = "Unknown"
    If state_code = "AL" Then state_name = "Alabama"
    If state_code = "AK" Then state_name = "Alaska"
    If state_code = "AZ" Then state_name = "Arizona"
    If state_code = "AR" Then state_name = "Arkansas"
    If state_code = "CA" Then state_name = "California"
    If state_code = "CO" Then state_name = "Colorado"
    If state_code = "CT" Then state_name = "Connecticut"
    If state_code = "DE" Then state_name = "Delaware"
    If state_code = "DC" Then state_name = "District Of Columbia"
    If state_code = "FL" Then state_name = "Florida"
    If state_code = "GA" Then state_name = "Georgia"
    If state_code = "HI" Then state_name = "Hawaii"
    If state_code = "ID" Then state_name = "Idaho"
    If state_code = "IL" Then state_name = "Illnois"
    If state_code = "IN" Then state_name = "Indiana"
    If state_code = "IA" Then state_name = "Iowa"
    If state_code = "KS" Then state_name = "Kansas"
    If state_code = "KY" Then state_name = "Kentucky"
    If state_code = "LA" Then state_name = "Louisiana"
    If state_code = "ME" Then state_name = "Maine"
    If state_code = "MD" Then state_name = "Maryland"
    If state_code = "MA" Then state_name = "Massachusetts"
    If state_code = "MI" Then state_name = "Michigan"
	If state_code = "MN" Then state_name = "Minnesota"
    If state_code = "MS" Then state_name = "Mississippi"
    If state_code = "MO" Then state_name = "Missouri"
    If state_code = "MT" Then state_name = "Montana"
    If state_code = "NE" Then state_name = "Nebraska"
    If state_code = "NV" Then state_name = "Nevada"
    If state_code = "NH" Then state_name = "New Hampshire"
    If state_code = "NJ" Then state_name = "New Jersey"
    If state_code = "NM" Then state_name = "New Mexico"
    If state_code = "NY" Then state_name = "New York"
    If state_code = "NC" Then state_name = "North Carolina"
    If state_code = "ND" Then state_name = "North Dakota"
    If state_code = "OH" Then state_name = "Ohio"
    If state_code = "OK" Then state_name = "Oklahoma"
    If state_code = "OR" Then state_name = "Oregon"
    If state_code = "PA" Then state_name = "Pennsylvania"
    If state_code = "RI" Then state_name = "Rhode Island"
    If state_code = "SC" Then state_name = "South Carolina"
    If state_code = "SD" Then state_name = "South Dakota"
    If state_code = "TN" Then state_name = "Tennessee"
    If state_code = "TX" Then state_name = "Texas"
    If state_code = "UT" Then state_name = "Utah"
    If state_code = "VT" Then state_name = "Vermont"
    If state_code = "VA" Then state_name = "Virginia"
    If state_code = "WA" Then state_name = "Washington"
    If state_code = "WV" Then state_name = "West Virginia"
    If state_code = "WI" Then state_name = "Wisconsin"
    If state_code = "WY" Then state_name = "Wyoming"
    If state_code = "PR" Then state_name = "Puerto Rico"
    If state_code = "VI" Then state_name = "Virgin Islands"

    If include_state_code = TRUE Then state_name = state_code & " " & state_name	'This adds the code to the state name if seelected
end function

function get_this_script_started(script_index, end_script, month_to_use)
'--- WORK IN PROGRESS - This function has the primary functionality needed at the begining of an individual script run.
'~~~~~ script_index: this should just be 'script_index' and indicates the number of the script in the COMPLETE LIST OF SCRIPTS.
'~~~~~ end_script: If NOT in MAXIS (passworded out) should the script end.
'~~~~~ month_to_use: default value of the footer month and year - Options: 'MAXIS MONTH', 'CM', 'CM PLUS 1', 'CM PLUS 2', 'CM MINUS 1', 'CM MINUS 2'
'~~~~~
'===== Keywords: MAXIS, dialog,
	EMConnect ""
	Call check_for_MAXIS(end_script)
	Call MAXIS_case_number_finder(MAXIS_case_number)

	month_to_use = UCase(month_to_use)
	If month_to_use = "MAXIS MONTH" Then Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
	If month_to_use = "CM" Then
		MAXIS_footer_month = CM_mo
		MAXIS_footer_year = CM_yr
	End If
	If month_to_use = "CM PLUS 1" Then
		MAXIS_footer_month = CM_plus_1_mo
		MAXIS_footer_year = CM_plus_1_yr
	End If
	If month_to_use = "CM PLUS 2" Then
		MAXIS_footer_month = CM_plus_2_mo
		MAXIS_footer_year = CM_plus_2_yr
	End If
	If month_to_use = "CM MINUS 1" Then
		MAXIS_footer_month = right("0" &             DatePart("m",           DateAdd("m", -1, date)            ), 2)
		MAXIS_footer_year =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
	End If
	If month_to_use = "CM MINUS 2" Then
		MAXIS_footer_month = right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
		MAXIS_footer_year =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
	End If

	' MsgBox "The script running is:" & vbCR & "Category - " & script_array(script_index).category & vbCr & "Name - " & script_array(script_index).script_name

	'Showing the case number dialog
	Do
		DO
			err_msg = ""

			'Initial dialog to gather case number and footer month.
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 236, 195, "Case number dialog"
			  EditBox 70, 105, 65, 15, MAXIS_case_number
			  EditBox 70, 125, 20, 15, MAXIS_footer_month
			  EditBox 95, 125, 20, 15, MAXIS_footer_year
			  EditBox 70, 145, 160, 15, Worker_signature
			  ButtonGroup ButtonPressed
				OkButton 125, 175, 50, 15
				CancelButton 180, 175, 50, 15
				PushButton 165, 85, 60, 10, "INSTRUCTIONS", instructions_btn
				PushButton 115, 160, 115, 10, "SAVE MY WORKER SIGNATURE", save_worker_sig
			  GroupBox 10, 5, 220, 95, "Currently Running "
			  Text 20, 20, 210, 10, "Script: " & script_array(script_index).script_name
			  Text 30, 30, 195, 10, "from the " & script_array(script_index).category & " category"
			  Text 20, 45, 50, 10, "Description:"
			  Text 25, 55, 200, 25, script_array(script_index).description
			  Text 20, 110, 45, 10, "Case number:"
			  Text 20, 130, 45, 10, "Footer Month:"
			  Text 10, 150, 60, 10, "Worker Signature"
			  Text 125, 130, 25, 10, "mm/yy"
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = instructions_btn Then
				err_msg = "LOOP"
				call open_URL_in_browser(script_array(script_index).SharePoint_instructions_URL)
			ElseIf ButtonPressed = save_worker_sig Then
				err_msg = "LOOP"
			Else
				' MsgBox MAXIS_case_number
		        Call validate_MAXIS_case_number(err_msg, "*")
		        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			End If
		LOOP UNTIL err_msg = ""
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false

end function

Function get_to_RKEY()
'--- This function will get the user back to the main MMIS selection screen RKEY. You will need to already be in MMIS. Navigate_to_MMIS_region
'~~~~~ HH_member_array: leave blank
'===== Keywords: MMIS, navigate
    EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
    IF MMIS_panel_check <> "RKEY" THEN
        attempt = 1
        DO
            If MMIS_case_number = "" Then Call MMIS_case_number_finder(MMIS_case_number)
            PF6
            EMReadScreen MMIS_panel_check, 4, 1, 52
            attempt = attempt + 1
            If attempt = 15 Then Exit Do
        Loop Until MMIS_panel_check = "RKEY"
    End If
    EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
    IF MMIS_panel_check <> "RKEY" THEN
    	DO
    		PF6
    		EMReadScreen session_terminated_check, 18, 1, 7
    	LOOP until session_terminated_check = "SESSION TERMINATED"

        'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
        EMWriteScreen "MW00", 1, 2
        transmit
        transmit

        EMReadScreen MMIS_menu, 24, 3, 30
	    If MMIS_menu = "GROUP SECURITY SELECTION" Then
            row = 1
            col = 1
            EMSearch " C3", row, col
            If row <> 0 Then
                EMWriteScreen "X", row, 4
                transmit
            Else
                row = 1
                col = 1
                EMSearch " C4", row, col
                If row <> 0 Then
                    EMWriteScreen "X", row, 4
                    transmit
                Else
                    script_end_procedure_with_error_report("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
                End If
            End If

            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "X", row, col - 3
            transmit
        Else
            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "X", row, col - 3
            transmit
        End If
    END IF
End Function

Function hest_standards(heat_AC_amt, electric_amt, phone_amt, date_variable)
'--- This function determines the SUA - Standard Utility Allowance based on the date selected. This changes each October CM.18.15.09 at: https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00181509
'~~~~~ heat_AC_amt: Heat/AC expense variable. Recommended to keep as heat_AC_amt.
'~~~~~ electric_amt: Electric expense variable. Recommended to keep as electric_amt.
'~~~~~ phone_amt: Phone expense variable. Recommended to keep as phone_amt.
'~~~~~ date_variable: This is the date you need to compare to when measuring against the October date. Generally this is the application_date.
'===== Keywords: MAXIS, member, array, dialog
    If DateDiff("d",date_variable,#10/01/2022#) <= 0 then
        'October 2022 -- Amounts for applications on or AFTER 10/01/2022
        heat_AC_amt = 586
        electric_amt = 185
        phone_amt = 55
    Elseif DateDiff("d",date_variable,#10/01/2022#) > 0 then
        'October 2021 -- Amounts for applications BEFORE 10/01/2022
        heat_AC_amt = 488
        electric_amt = 149
        phone_amt = 56
    End if
End Function

function HH_member_custom_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
    transmit

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
        EMReadScreen access_denied_check, 13, 24, 2
        'MsgBox access_denied_check
        If access_denied_check = "ACCESS DENIED" Then
            PF10
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
        Else
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
    		last_name = trim(replace(last_name, "_", "")) & " "
    		first_name = trim(replace(first_name, "_", "")) & " "
    		mid_initial = replace(mid_initial, "_", "")
        End If
		client_string = ref_nbr & last_name & first_name & mid_initial
		client_array = client_array & client_string & "|"
		transmit
	    Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_array = TRIM(client_array)
	test_array = split(client_array, "|")
	total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

	DIM all_client_array()
	ReDim all_clients_array(total_clients, 1)

	FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
		Interim_array = split(client_array, "|")
		all_clients_array(x, 0) = Interim_array(x)
		all_clients_array(x, 1) = 1
	NEXT

	BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		Text 10, 5, 105, 10, "Household members to look at:"
		FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
			IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		NEXT
		ButtonGroup ButtonPressed
		OkButton 185, 10, 50, 15
		CancelButton 185, 30, 50, 15
	ENDDIALOG
													'runs the dialog that has been dynamically created. Streamlined with new functions.
	Dialog HH_memb_dialog
	If buttonpressed = 0 then stopscript
	check_for_maxis(True)

	HH_member_array = ""

	FOR i = 0 to total_clients
		IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
			IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
				'msgbox all_clients_
				HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
			END IF
		END IF
	NEXT

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
	HH_member_array = SPLIT(HH_member_array, " ")
end function

function is_date_holiday_or_weekend(date_to_review, boolean_variable)
'--- This function reviews a date to determine if it falls on a weekend or Hennepin County holiday
'~~~~~ date_to_review: this should be in the form of a date
'~~~~~ boolean_variable: this returns TRUE if the date is a weekend or holiday, FALSE if it is not
'==== Keywords: MAXIS, dates, boolean
    non_working_day = FALSE
    day_of_week = WeekdayName(WeekDay(date_to_review))
    If day_of_week = "Saturday" OR day_of_week = "Sunday" Then non_working_day = TRUE
    For each holiday in HOLIDAYS_ARRAY
        If holiday = date_to_review Then non_working_day = TRUE
    Next
    boolean_variable = non_working_day
end function

function MAXIS_background_check()
'--- This function checks to see if a user is in background
'===== Keywords: MAXIS, background
	Do
		call navigate_to_MAXIS_screen("STAT", "SUMM")
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			PF3
			Pause 2
		End if
	Loop until SELF_check <> "SELF"
end function

function MAXIS_case_number_finder(variable_for_MAXIS_case_number)
'--- This function finds the MAXIS case number if listed on a MAXIS screen
'~~~~~ variable_for_MAXIS_case_number: this should be <code>MAXIS_case_number</code>
'===== Keywords: MAXIS, case number
	EMReadScreen variable_for_SELF_check, 4, 2, 50
	IF variable_for_SELF_check = "SELF" then
		EMReadScreen variable_for_MAXIS_case_number, 8, 18, 43
		variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
		variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
	ELSE
		row = 1
		col = 1
		EMSearch "Case Nbr:", row, col
		If row <> 0 then
			EMReadScreen variable_for_MAXIS_case_number, 8, row, col + 10
			variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
			variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
		END IF
	END IF

end function

function MAXIS_dialog_navigation()
'--- This function navigates to various panels in MAXIS. You need to name your buttons using the button names in the function.
'===== Keywords: MAXIS, dialog, navigation
	'This part works with the prev/next buttons on several of our dialogs. You need to name your buttons prev_panel_button, next_panel_button, prev_memb_button, and next_memb_button in order to use them.
	EMReadScreen STAT_check, 4, 20, 21
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = prev_memb_button then
			HH_memb_row = HH_memb_row - 1
			EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
			If isnumeric(prev_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen prev_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		ELSEIF ButtonPressed = next_memb_button then
			HH_memb_row = HH_memb_row + 1
			EMReadScreen next_HH_memb, 2, HH_memb_row, 3
			If isnumeric(next_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen next_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		End if
    End if

'This part takes care of remaining navigation buttons, designed to go to a single panel using the maxis case number.
    'CASE
	If ButtonPressed = ADHI_button then call navigate_to_MAXIS_screen("CASE", "ADHI") 'Address History (ADHI)
    If ButtonPressed = CURR_button then call navigate_to_MAXIS_screen("CASE", "CURR") 'Current Case Display (CURR)
	If ButtonPressed = HISC_button then call navigate_to_MAXIS_screen("CASE", "HISC")'Cash Payment History (HISC)
	If ButtonPressed = HISF_button then call navigate_to_MAXIS_screen("CASE", "HISF") 'FS Payment History (HISF)
    If ButtonPressed = NOTC_button then call navigate_to_MAXIS_screen("CASE", "NOTC") 'Notice Display (NOTC)
    If ButtonPressed = NOTE_button then call navigate_to_MAXIS_screen("CASE", "NOTE") 'Case Notes (NOTE)
	If ButtonPressed = PERS_button then call navigate_to_MAXIS_screen("CASE", "PERS") 'Person Status (PERS)
    'CCOL
    If ButtonPressed = CADR_button then call navigate_to_MAXIS_screen("CCOL", "CADR") 'Claim Person Address (CADR)
    If ButtonPressed = CBIL_button then call navigate_to_MAXIS_screen("CCOL", "CBIL") 'Bill Inquiry- By Person (CBIL)
    If ButtonPressed = CDEM_button then call navigate_to_MAXIS_screen("CCOL", "CDEM") 'Demand Letter Inquiry (CDEM)
    If ButtonPressed = CLAM_button then call navigate_to_MAXIS_screen("CCOL", "CLAM") 'Manual Entry of a Claim (CLAM)
    If ButtonPressed = CLDL_button then call navigate_to_MAXIS_screen("CCOL", "CLDL") 'Claim Repayment Agreement (CLDL)
    If ButtonPressed = CLIC_button then call navigate_to_MAXIS_screen("CCOL", "CLIC") 'Claim Inquiry- By Case (CLIC)
    If ButtonPressed = CLIP_button then call navigate_to_MAXIS_screen("CCOL", "CLIP") 'Claim Inquiry - By Person (CLIP)
	If ButtonPressed = CLIV_button then call navigate_to_MAXIS_screen("CCOL", "CLIV") 'Claim Inquiry - By Vendor (CLIV)
 	If ButtonPressed = CLRA_button then call navigate_to_MAXIS_screen("CCOL", "CLRA") 'Claim Repayment Agreement (CLRA)
    If ButtonPressed = CLSM_button then call navigate_to_MAXIS_screen("CCOL", "CLSM") 'Claim Summary/Maintenance (CLSM)
    If ButtonPressed = CRAA_button then call navigate_to_MAXIS_screen("CCOL", "CRAA") 'Claim Receipts and Adjust (CRAA)
    If ButtonPressed = CTHC_button then call navigate_to_MAXIS_screen("CCOL", "CTHC") 'Claim Trans Hist By Claim (CTHC)
    If ButtonPressed = CTHE_button then call navigate_to_MAXIS_screen("CCOL", "CTHE") 'Claim Trans Paid By Case (CTHE)
    If ButtonPressed = CTHP_button then call navigate_to_MAXIS_screen("CCOL", "CTHP") 'Claim Trans Paid By Person (CTHP)
	If ButtonPressed = CTHV_button then call navigate_to_MAXIS_screen("CCOL", "CTHV") 'Claim Trans Hist By Vendor (CTHV)
    If ButtonPressed = CTHW_button then call navigate_to_MAXIS_screen("CCOL", "CTHW") 'Claim Trans Log By Worker (CTHW)
    If ButtonPressed = CTOP_button then call navigate_to_MAXIS_screen("CCOL", "CTOP") 'Claim TOP Status (CTOP)
    If ButtonPressed = JGMT_button then call navigate_to_MAXIS_screen("CCOL", "JGMT") 'Judgment Tracking/Mgmt (JGMT)
    If ButtonPressed = PRMX_button then call navigate_to_MAXIS_screen("CCOL", "PRMX") 'Entry of PreMAXIS Claim (PRMX)
	'DAIL
	If ButtonPressed = CLMS_button then call navigate_to_MAXIS_screen("DAIL", "CLMS") 'Claims Report (CLMS)
	If ButtonPressed = COLA_button then call navigate_to_MAXIS_screen("DAIL", "COLA") 'COLA Report (COLA)
    If ButtonPressed = CSES_button then call navigate_to_MAXIS_screen("DAIL", "CSES") 'Child Support Report (CSES)
    If ButtonPressed = DAIL_button then call navigate_to_MAXIS_screen("DAIL", "DAIL") 'Daily Report (DAIL)
    If ButtonPressed = ELIG_button then call navigate_to_MAXIS_screen("DAIL", "ELIG") 'Eligibility Results Report (ELIG)
    If ButtonPressed = IEVS_button then call navigate_to_MAXIS_screen("DAIL", "IEVS") 'Interface Report (IEVS)
	If ButtonPressed = INFO_button then call navigate_to_MAXIS_screen("DAIL", "INFO") 'Information Message Report (INFO)
    If ButtonPressed = IV-E_button then call navigate_to_MAXIS_screen("DAIL", "IV-E") 'Select SSIS (IV-E)
    If ButtonPressed = MA_button then call navigate_to_MAXIS_screen("DAIL", "MA")     'Medical Assistance Report (MA)
    If ButtonPressed = MEC2_button then call navigate_to_MAXIS_screen("DAIL", "MEC2") 'MEG' Report (MEC2)
    If ButtonPressed = PARI_button then call navigate_to_MAXIS_screen("DAIL", "PARI") 'Paris Interstate (PARI)
    If ButtonPressed = PEPR_button then call navigate_to_MAXIS_screen("DAIL", "PEPR") 'Periodic Processing Report (PEPR)
    If ButtonPressed = PICK_button then call navigate_to_MAXIS_screen("DAIL", "PICK") 'Select Combination (PICK)
    If ButtonPressed = TIKL_button then call navigate_to_MAXIS_screen("DAIL", "TIKL") 'Tickler Report (TIKL)
    If ButtonPressed = WF1_button then call navigate_to_MAXIS_screen("DAIL", "WF1")   'Work Force 1 Report (WF1)
    If ButtonPressed = WRIT_button then call navigate_to_MAXIS_screen("DAIL", "WRIT") 'Write Ticklers (WRIT)
	'ELIG-DWP
	If ButtonPressed = ELIG_DWP_button then call navigate_to_MAXIS_screen("ELIG", "DWP_")
	If ButtonPressed = DWPR_button then call navigate_to_MAXIS_screen("ELIG", "DWPR") 'Person Results (DWPR)
	If ButtonPressed = DWCR_button then call navigate_to_MAXIS_screen("ELIG", "DWCR") 'Case Results (DWCR)
	If ButtonPressed = MDWB1_button then call navigate_to_MAXIS_screen("ELIG", "DWB1") 'Budget Sum Part 1 (DWB1)
	If ButtonPressed = DWB2_button then call navigate_to_MAXIS_screen("ELIG", "DWB2") 'Budget Sum Part 2 (DWB2)
	If ButtonPressed = DWSM_button then call navigate_to_MAXIS_screen("ELIG", "DWSM") 'Case Summary (DWSM)
	'ELIG-FS
	If ButtonPressed = ELIG_FS_button then call navigate_to_MAXIS_screen("ELIG", "FS__")
	If ButtonPressed = FSPR_button then call navigate_to_MAXIS_screen("ELIG", "FSPR") 'Person Results (FSPR)
	If ButtonPressed = FSCR_button then call navigate_to_MAXIS_screen("ELIG", "FSCR") 'Case Results (FSCR)
	If ButtonPressed = FSB1_button then call navigate_to_MAXIS_screen("ELIG", "FSB1") 'Monthly Budget 1 (FSB1)
	If ButtonPressed = FSB2_button then call navigate_to_MAXIS_screen("ELIG", "FSB2") 'Monthly Budget 2 (FSB2)
	If ButtonPressed = FSSM_button then call navigate_to_MAXIS_screen("ELIG", "FSSM") 'Case Summary (FSSM)
	'ELIG-GA
	If ButtonPressed = ELIG_GA_button then call navigate_to_MAXIS_screen("ELIG", "GA__")
	If ButtonPressed = GAPR_button then call navigate_to_MAXIS_screen("ELIG", "GAPR") 'Person Results (GAPR)
	If ButtonPressed = GACR_button then call navigate_to_MAXIS_screen("ELIG", "GACR") 'Case Results (GACR)
	If ButtonPressed = GAB1_button then call navigate_to_MAXIS_screen("ELIG", "GAB1") 'Monthly Budget 1 (GAB1)
	If ButtonPressed = GAB2_button then call navigate_to_MAXIS_screen("ELIG", "GAB2") 'Monthly Budget 2 (GAB2)
	If ButtonPressed = GASM_button then call navigate_to_MAXIS_screen("ELIG", "GASM") 'Case Summary (GASM)
	'ELIG-HC
	If ButtonPressed = ELIG_HC_button then call navigate_to_MAXIS_screen("ELIG", "HC__")
	If ButtonPressed = HHMM_button then call navigate_to_MAXIS_screen("ELIG", "HHMM") 'HC Member List (HHMM)
	If ButtonPressed = BSUM_button then call navigate_to_MAXIS_screen("ELIG", "BSUM") 'Basic HC Budg Sum (BSUM)
	If ButtonPressed = BHSM_button then call navigate_to_MAXIS_screen("ELIG", "BHSM") 'Basic HC Sum/App (BHSM)
	'ELIG-MFIP
	If ButtonPressed = ELIG_MFIP_button then call navigate_to_MAXIS_screen("ELIG", "MFIP")
	If ButtonPressed = MFPR_button then call navigate_to_MAXIS_screen("ELIG", "MFPR") 'Person Results (MFPR)
	If ButtonPressed = MFCR_button then call navigate_to_MAXIS_screen("ELIG", "MFCR") 'Case Results (MFCR)
	If ButtonPressed = MFBF_button then call navigate_to_MAXIS_screen("ELIG", "MFBF") 'Budget Factors (MFBF)
	If ButtonPressed = MFB1_button then call navigate_to_MAXIS_screen("ELIG", "MFB1") 'Monthly Budget 1 (MFB1)
	If ButtonPressed = MFB2_button then call navigate_to_MAXIS_screen("ELIG", "MFB2") 'Monthly Budget 2 (MFB2)
	If ButtonPressed = MFSM_button then call navigate_to_MAXIS_screen("ELIG", "MFSM") 'Case Summary (MFSM)
	'ELIG-MSA
	If ButtonPressed = ELIG_MSA_button then call navigate_to_MAXIS_screen("ELIG", "MSA_")
	If ButtonPressed = MSPR_button then call navigate_to_MAXIS_screen("ELIG", "MSPR") 'Person Results (MSPR)
	If ButtonPressed = MSGR_button then call navigate_to_MAXIS_screen("ELIG", "MSGR") 'Case Results (MSGR)
	If ButtonPressed = MSCB_button then call navigate_to_MAXIS_screen("ELIG", "MSCB") 'SSI Type Case Budget (MSCB)
	If ButtonPressed = MSSM_button then call navigate_to_MAXIS_screen("ELIG", "MSSM") 'Summary (MSSM)
	'ELIG-OT
	IF ButtonPressed = ELIG_DENY_button then call navigate_to_MAXIS_screen("ELIG", "DENY") 'Cash Denial (DENY)
	IF ButtonPressed = ELIG_EMER_button then call navigate_to_MAXIS_screen("ELIG", "EMER") 'Emergency Assistance (EMER)
	IF ButtonPressed = ELIG_IVE_button then call navigate_to_MAXIS_screen("ELIG", "IVE") 'Title IV-E Foster Care (IVE)
	If ButtonPressed = ELIG_GRH_button then call navigate_to_MAXIS_screen("ELIG", "GRH_") 'Group Residential Housing (GRH)'
	IF ButtonPressed = ELIG_RCA_button then call navigate_to_MAXIS_screen("ELIG", "RCA") 'Refugee Cash Assistance (RCA)
	IF ButtonPressed = ELIG_SUMM_button then call navigate_to_MAXIS_screen("ELIG", "SUMM") 'Eligibility Results Summary (SUMM)
	'INFC
    If ButtonPressed = CSIA_button then call navigate_to_MAXIS_screen("INFC", "CSIA")  'Child Support Interface A (CSIA)
    If ButtonPressed = CSIB_button then call navigate_to_MAXIS_screen("INFC", "CSIB")  'Child Support Interface B (CSIB)
    If ButtonPressed = CSIC_button then call navigate_to_MAXIS_screen("INFC", "CSIC")  'Child Support Interface C (CSIC)
    If ButtonPressed = CSID_button then call navigate_to_MAXIS_screen("INFC", "CSID")  'Child Support Interface D (CSID)
    If ButtonPressed = eDRS_button then call navigate_to_MAXIS_screen("INFC", "eDRS")  'Electronic Disq. Recipient System (eDRS)
    If ButtonPressed = SSIS_button then call navigate_to_MAXIS_screen("INFC", "SSIS")  'Social Services Information System (SSIS)
    If ButtonPressed = SVES_button then call navigate_to_MAXIS_screen("INFC", "SVES")  'State Verification Exchange (SVES)
    If ButtonPressed = WF1M_button then call navigate_to_MAXIS_screen("INFC", "WF1M")  'WF1 Manual Referral (WF1M)
    If ButtonPressed = WORK_button then call navigate_to_MAXIS_screen("INFC", "WORK")  'Workforce One Referral (WORK)
	'MONY
    If ButtonPressed = CANC_button then call navigate_to_MAXIS_screen("MONY", "CANC")'Cancel/Stop Payment (CANC)
    If ButtonPressed = CHCK_button then call navigate_to_MAXIS_screen("MONY", "CHCK")'Check Request (CHCK)
    If ButtonPressed = DISB_button then call navigate_to_MAXIS_screen("MONY", "DISB")'Disbursement Method Entry (DISB)
    If ButtonPressed = INQB_button then call navigate_to_MAXIS_screen("MONY", "INQB") 'Benefit History (INQB)
    If ButtonPressed = INQD_button then call navigate_to_MAXIS_screen("MONY", "INQD") 'Disbursement History (INQD)
    If ButtonPressed = INQF_button then call navigate_to_MAXIS_screen("MONY", "INQF")'Allocated Funds (INQF)
    If ButtonPressed = INQT_button then call navigate_to_MAXIS_screen("MONY", "INQT")'Transaction Inquiry (INQT)
	If ButtonPressed = INQX_button then call navigate_to_MAXIS_screen("MONY", "INQX") 'Paymt Hist Select Criteria (INQX)
	If ButtonPressed = REPL_button then call navigate_to_MAXIS_screen("MONY", "REPL")'Replace Disbursements (REPL)
    If ButtonPressed = VNDA_button then call navigate_to_MAXIS_screen("MONY", "VNDA")'Vendor Authorizations (VNDA)
    If ButtonPressed = VNDS_button then call navigate_to_MAXIS_screen("MONY", "VNDS")'Vendor Search  (VNDS)
    If ButtonPressed = VNDW_button then call navigate_to_MAXIS_screen("MONY", "VNDW")'Warrants by Vendor (VNDW)
	'REPT
	If ButtonPressed = ACTV_button then call navigate_to_MAXIS_screen("REPT", "ACTV")  'Active Caseload (ACTV)
	If ButtonPressed = ARPT_button then call navigate_to_MAXIS_screen("REPT", "ARPT")  'Activity Reports (ARPT)
	If ButtonPressed = EOMC_button then call navigate_to_MAXIS_screen("REPT", "EOMC")  'End of Month Closures (EOMC)
	If ButtonPressed = FCRR_button then call navigate_to_MAXIS_screen("REPT", "FCRR")  'Foster Care Referral Rept. (FCRR)
	If ButtonPressed = FCRV_button then call navigate_to_MAXIS_screen("REPT", "FCRV")  'Foster Care Review Report (FCRV)
	If ButtonPressed = IVCN_button then call navigate_to_MAXIS_screen("REPT", "HCCC")  'Health Care Cases Converted (HCCC)
	If ButtonPressed = IEVC_button then call navigate_to_MAXIS_screen("REPT", "IEVC")  'County Verifications To-Do (IEVC)
	If ButtonPressed = IVCN_button then call navigate_to_MAXIS_screen("REPT", "IVCN")  'IV-E Cases To Convert (IVCN)
	If ButtonPressed = INAC_button then call navigate_to_MAXIS_screen("REPT", "INAC")  'Inactive Caseload (INAC)
	If ButtonPressed = INTR_button then call navigate_to_MAXIS_screen("REPT", "INTR")  'Interstate Match Report (INTR)
	If ButtonPressed = MAMS_button then call navigate_to_MAXIS_screen("REPT", "MAMS")  'MA Monthly Report Form (MAMS)
	If ButtonPressed = MFCM_button then call navigate_to_MAXIS_screen("REPT", "MFCM")  'MFIP Participant Case Management (MFCM)
	If ButtonPressed = MLAR_button then call navigate_to_MAXIS_screen("REPT", "MLAR")  'MIPPA LIS Appl (MLAR)
	If ButtonPressed = MONT_button then call navigate_to_MAXIS_screen("REPT", "MONT")'HRF Status Update (MONT)
	If ButtonPressed = MRSR_button then call navigate_to_MAXIS_screen("REPT", "MRSR")  'Monthly Reporters Status (MRSR)
	If ButtonPressed = PND1_button then call navigate_to_MAXIS_screen("REPT", "PND1")  'GAF I Pending Report (PND1)
	If ButtonPressed = PND2_button then call navigate_to_MAXIS_screen("REPT", "PND2")  'GAF II Pending Report (PND2)
	If ButtonPressed = REVS_button then call navigate_to_MAXIS_screen("REPT", "REVS")  'Review Dates (REVS)
	If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("REPT", "REVW")  'Review Report (REVW)
    'STAT
    If ButtonPressed = ABPS_button then call navigate_to_MAXIS_screen("STAT", "ABPS") 'Absent Parent Status (ABPS)
    If ButtonPressed = ACCI_button then call navigate_to_MAXIS_screen("STAT", "ACCI") 'Accident (ACCI)
    If ButtonPressed = ACCT_button then call navigate_to_MAXIS_screen("STAT", "ACCT") 'Accounts (ACCT)
    If ButtonPressed = ACUT_button then call navigate_to_MAXIS_screen("STAT", "ACUT") 'Actual Utility Expenses (ACUT)
    If ButtonPressed = ADDR_button then call navigate_to_MAXIS_screen("STAT", "ADDR") 'Address (ADDR)
    If ButtonPressed = ADME_button then call navigate_to_MAXIS_screen("STAT", "ADME") 'Add Member (ADME)
    If ButtonPressed = ALTP_button then call navigate_to_MAXIS_screen("STAT", "ALTP") 'Alternate Payee (ALTP)
    If ButtonPressed = ALIA_button then call navigate_to_MAXIS_screen("STAT", "ALIA") 'Alias (ALIA)
    If ButtonPressed = AREP_button then call navigate_to_MAXIS_screen("STAT", "AREP") 'Authorized Representative (AREP)
    If ButtonPressed = BILS_button then call navigate_to_MAXIS_screen("STAT", "BILS") 'Health Care Medical Exp (BILS)
    If ButtonPressed = BUDG_button then call navigate_to_MAXIS_screen("STAT", "BUDG") 'Budget Period Selection(BUDG)
    If ButtonPressed = BUSI_button then call navigate_to_MAXIS_screen("STAT", "BUSI") 'Self-Employment Income (BUSI)
    If ButtonPressed = CARS_button then call navigate_to_MAXIS_screen("STAT", "CARS") 'Vehicles (CARS)
    If ButtonPressed = CASH_button then call navigate_to_MAXIS_screen("STAT", "CASH") 'Cash (CASH)
    If ButtonPressed = COEX_button then call navigate_to_MAXIS_screen("STAT", "COEX") 'Court Ordered Expenses (COEX)
    If ButtonPressed = DCEX_button then call navigate_to_MAXIS_screen("STAT", "DCEX") 'Dependent Care Expenses (DCEX)
    If ButtonPressed = DFLN_button then call navigate_to_MAXIS_screen("STAT", "DFLN") 'Drug Felon (DFLN)
    If ButtonPressed = DIET_button then call navigate_to_MAXIS_screen("STAT", "DIET") 'Prescribed Diet (DIET)
    If ButtonPressed = DISA_button then call navigate_to_MAXIS_screen("STAT", "DISA") 'Disability (DISA)
    If ButtonPressed = DISQ_button then call navigate_to_MAXIS_screen("STAT", "DISQ") 'Disqualification (DISQ)
    If ButtonPressed = DSTT_button then call navigate_to_MAXIS_screen("STAT", "DSTT") 'Destitute Status (DSTT)
    If ButtonPressed = EATS_button then call navigate_to_MAXIS_screen("STAT", "EATS") 'Eating Groups (EATS)
    If ButtonPressed = EMMA_button then call navigate_to_MAXIS_screen("STAT", "EMMA") 'Medical Erner Programs (EMMA)
    If ButtonPressed = EMPS_button then call navigate_to_MAXIS_screen("STAT", "EMPS") 'Employment Services (EMPS)
    If ButtonPressed = FACI_button then call navigate_to_MAXIS_screen("STAT", "FACI") 'Facility (FACI)
    If ButtonPressed = FCFC_button then call navigate_to_MAXIS_screen("STAT", "FCFC")'Foster Care Facility (FCFC)
    If ButtonPressed = FCPL_button then call navigate_to_MAXIS_screen("STAT", "FCPL")'Foster Care Placement (FCPL)
    If ButtonPressed = FMED_button then call navigate_to_MAXIS_screen("STAT", "FMED") 'FS Elderly/Disa Medical Exp (FMED)
    If ButtonPressed = HCMI_button then call navigate_to_MAXIS_screen("STAT", "HCMI") 'Health Care Misc Info (HCMI)
    If ButtonPressed = HCRE_button then call navigate_to_MAXIS_screen("STAT", "HCRE") 'Heath Care Request (HCRE)
    If ButtonPressed = HEST_button then call navigate_to_MAXIS_screen("STAT", "HEST") 'Housing Exp Standard (HEST)
    If ButtonPressed = IMIG_button then call navigate_to_MAXIS_screen("STAT", "IMIG") 'Immigration Status (IMIG)
    If ButtonPressed = INSA_button then call navigate_to_MAXIS_screen("STAT", "INSA") 'Insurance (INSA)
    If ButtonPressed = JOBS_button then call navigate_to_MAXIS_screen("STAT", "JOBS") 'Job Income (JOBS)
    If ButtonPressed = LUMP_button then call navigate_to_MAXIS_screen("STAT", "LUMP") 'Lump Sum (LUMP)
    If ButtonPressed = MEDI_button then call navigate_to_MAXIS_screen("STAT", "MEDI") 'Medicare (MEDI)
    If ButtonPressed = MEMB_button then call navigate_to_MAXIS_screen("STAT", "MEMB") 'Household Member (MEMB)
    If ButtonPressed = MEMI_button then call navigate_to_MAXIS_screen("STAT", "MEMI") 'Additional Member Info (MEMI)
    If ButtonPressed = MISC_button then call navigate_to_MAXIS_screen("STAT", "MISC") 'Miscellaneous Case (MISC)
    If ButtonPressed = MMSA_button then call navigate_to_MAXIS_screen("STAT", "MMSA") 'Misc. MSA Information (MMSA)
    If ButtonPressed = MONT_button then call navigate_to_MAXIS_screen("STAT", "MONT") 'HRF Status Update (MONT)
    If ButtonPressed = MSUR_button then call navigate_to_MAXIS_screen("STAT", "MSUR") 'MNSure Eligibility Tracking (MSUR)
    If ButtonPressed = OTHR_button then call navigate_to_MAXIS_screen("STAT", "OTHR") 'Other Assets (OTHR)
    If ButtonPressed = PACT_button then call navigate_to_MAXIS_screen("STAT", "PACT") 'Program Action (PACT)
    If ButtonPressed = PARE_button then call navigate_to_MAXIS_screen("STAT", "PARE") 'Parent (PARE)
    If ButtonPressed = PBEN_button then call navigate_to_MAXIS_screen("STAT", "PBEN") 'Potential Benefits (PBEN)
    If ButtonPressed = PDED_button then call navigate_to_MAXIS_screen("STAT", "PDED") 'Program Deductions (PDED)
    If ButtonPressed = PREG_button then call navigate_to_MAXIS_screen("STAT", "PREG") 'Pregnancy (PREG)
    If ButtonPressed = PROG_button then call navigate_to_MAXIS_screen("STAT", "PROG") 'Program Designation (PROG)
    If ButtonPressed = RBIC_button then call navigate_to_MAXIS_screen("STAT", "RBIC") 'Room/Board Income (RBIC)
    If ButtonPressed = REMO_button then call navigate_to_MAXIS_screen("STAT", "REMO") 'Remove Member/Temp Abs (REMO)
    If ButtonPressed = RESI_button then call navigate_to_MAXIS_screen("STAT", "RESI") 'Case Residency Info (RESI)
    If ButtonPressed = REST_button then call navigate_to_MAXIS_screen("STAT", "REST") 'Real Estate (REST)
    If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("STAT", "REVW") 'Case Reviews (REVW)
    If ButtonPressed = SANC_button then call navigate_to_MAXIS_screen("STAT", "SANC") 'Sanction Tracking (SANC)
    If ButtonPressed = SCHL_button then call navigate_to_MAXIS_screen("STAT", "SCHL") 'School (SCHL)
    If ButtonPressed = SECU_button then call navigate_to_MAXIS_screen("STAT", "SECU") 'Securities (SECU)
    If ButtonPressed = SHEL_button then call navigate_to_MAXIS_screen("STAT", "SHEL") 'Shelter Expenses (SHEL)
    If ButtonPressed = SIBL_button then call navigate_to_MAXIS_screen("STAT", "SIBL") 'Sibling (SIBL)
    If ButtonPressed = SPON_button then call navigate_to_MAXIS_screen("STAT", "SPON") 'Sponsor Income & Assets (SPON)
    If ButtonPressed = SSRT_button then call navigate_to_MAXIS_screen("STAT", "SSRT") 'GRH Supplemental Service Rate (SSRT)
    If ButtonPressed = STEC_button then call navigate_to_MAXIS_screen("STAT", "STEC") 'Student Expenses (STEC)
    If ButtonPressed = STIN_button then call navigate_to_MAXIS_screen("STAT", "STIN") 'Student Income (STIN)
    If ButtonPressed = STWK_button then call navigate_to_MAXIS_screen("STAT", "STWK") 'Stop Work (STWK)
    If ButtonPressed = STRK_button then call navigate_to_MAXIS_screen("STAT", "STRK") 'Strike (STRK)
    If ButtonPressed = SWKR_button then call navigate_to_MAXIS_screen("STAT", "SWKR") 'Social Worker (SWKR)
    If ButtonPressed = TIME_button then call navigate_to_MAXIS_screen("STAT", "TIME") 'Time Tracking (TIME)
    If ButtonPressed = TRAC_button then call navigate_to_MAXIS_screen("STAT", "TRAC") 'Earned Income Disregard (TRAC)
    If ButtonPressed = TRAN_button then call navigate_to_MAXIS_screen("STAT", "TRAN") 'Transferred Assets (TRAN)
    If ButtonPressed = TRTX_button then call navigate_to_MAXIS_screen("STAT", "TRTX") 'Transition from Residential Trtmnt (TRTX)
    If ButtonPressed = TYPE_button then call navigate_to_MAXIS_screen("STAT", "TYPE") 'Assistance Type (TYPE)
    If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("STAT", "UNEA") 'Unearned Income (UNEA)
    If ButtonPressed = WKEX_button then call navigate_to_MAXIS_screen("STAT", "WKEX") 'Work Expenses (WKEX)
    If ButtonPressed = WREG_button then call navigate_to_MAXIS_screen("STAT", "WREG") 'Work Registration (WREG)
	'SPEC this error message comes up  CASE STATUS MUST BE INACTIVE TO ENTER 'ADDR', 'FACI' OR 'TRAC'
	If ButtonPressed = INAC_ADDR_button then call navigate_to_MAXIS_screen("SPEC", "ADDR") 'Address-Case Inactive (ADDR)
	If ButtonPressed = ADPT_button then call navigate_to_MAXIS_screen("SPEC", "ADPT") 'Adoption SSN Placement (ADPT)
	If ButtonPressed = INAC_FACI_button then call navigate_to_MAXIS_screen("SPEC", "FACI") 'Facility-Case Inactive (FACI)
	If ButtonPressed = LETR_button then call navigate_to_MAXIS_screen("SPEC", "LETR") 'Worker Selected Notices (LETR)
	If ButtonPressed = MEMO_button then call navigate_to_MAXIS_screen("SPEC", "MEMO") 'Client Memo (MEMO)'
	If ButtonPressed = INAC_TRAC_button then call navigate_to_MAXIS_screen("SPEC", "TRAC") '$30 & 1 /3 Tracking (TRAC)
	If ButtonPressed = XFER_button then call navigate_to_MAXIS_screen("SPEC", "XFER") 'Case Transfer (XFER)
	If ButtonPressed = WCOM_button then call navigate_to_MAXIS_screen("SPEC", "WCOM") 'Worker Comments (WCOM)
	'STAT Summaries
    If ButtonPressed = PNLP_button then call navigate_to_MAXIS_screen("STAT", "PNLP") 'Personal Summary (PNLP)
    If ButtonPressed = PNLI_button then call navigate_to_MAXIS_screen("STAT", "PNLI") 'Income Summary (PNLI)
    If ButtonPressed = PNLR_button then call navigate_to_MAXIS_screen("STAT", "PNLR") 'Resource Summary (PNLR)
    If ButtonPressed = PNLE_button then call navigate_to_MAXIS_screen("STAT", "PNLE") 'Expense Summary (PNLE)
    If ButtonPressed = SUMM_button then call navigate_to_MAXIS_screen("STAT", "SUMM") 'Edit Summary (SUMM)
    If ButtonPressed = ERRR_button then call navigate_to_MAXIS_screen("STAT", "ERRR") 'Error Prone Edit Summary (ERRR)
END FUNCTION


function MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'--- This function finds the MAXIS footer month/year for MAXIS cases in SELF, MAXIS panels or MEMO screens.
'~~~~~ MAXIS_footer_month: needs to be <code>MAXIS_footer_month</code>
'~~~~~ MAXIS_footer_year: needs to be <code>MAXIS_footer_year</code>
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50
    EMReadScreen MEMO_check, 4, 2, 47
    EMReadScreen casenote_check, 4, 2, 45
    Call find_variable("Function: ", MAXIS_function, 4)

	IF SELF_check = "SELF" THEN
		EMReadScreen MAXIS_footer_month, 2, 20, 43
		EMReadScreen MAXIS_footer_year, 2, 20, 46
	ELSEIF MEMO_check = "MEMO" or MEMO_check = "WCOM" Then
		EMReadScreen MAXIS_footer_month, 2, 19, 54
		EMReadScreen MAXIS_footer_year, 2, 19, 57
    ELSEIF casenote_check = "NOTE" then
    	EMReadScreen MAXIS_footer_month, 2, 20, 54
        EMReadScreen MAXIS_footer_year, 2, 20, 57
	ELSEIF MAXIS_function = "STAT" then
		EMReadScreen MAXIS_footer_month, 2, 20, 55
        EMReadScreen MAXIS_footer_year, 2, 20, 58
    ELSEIF MAXIS_function = "REPT" then
    		EMReadScreen MAXIS_footer_month, 2, 20, 54
            EMReadScreen MAXIS_footer_year, 2, 20, 57
    Else
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
    End if
end function

function MAXIS_footer_month_confirmation()
'--- This function is for checking and changing the footer month to the MAXIS_footer_month & MAXIS_footer_year selected by the user in the inital dialog if necessary
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50			'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN
		EMReadScreen panel_footer_month, 2, 20, 43
		EMReadScreen panel_footer_year, 2, 20, 46
	ELSE
		Call find_variable("Month: ", MAXIS_footer, 5)	'finding footer month and year if not on the SELF screen
		panel_footer_month = left(MAXIS_footer, 2)
		panel_footer_year = right(MAXIS_footer, 2)
		If row <> 0 then
  			panel_footer_month = panel_footer_month		'Establishing variables
			panel_footer_year =panel_footer_year
		END IF
	END IF
	panel_date = panel_footer_month & panel_footer_year		'creating new variable combining month and year for the date listed on the MAXIS panel
	dialog_date = MAXIS_footer_month & MAXIS_footer_year	'creating new variable combining the MAXIS_footer_month & MAXIS_footer_year to measure against the panel date
	IF panel_date <> dialog_date then 						'if dates are not equal
		back_to_SELF
		EMWriteScreen MAXIS_footer_month, 20, 43			'goes back to self and enters the date that the user selcted'
		EMWriteScreen MAXIS_footer_year, 20, 46
	END IF
end function

Function MMIS_case_number_finder(MMIS_case_number)
'--- This function finds the MAXIS case number if listed on a MMIS screen
'~~~~~ MMIS_case_number: this should be <code>MMIS_case_number</code>
'===== Keywords: MMIS, case number
    row = 1
    col = 1
    EMSearch "CASE NUMBER:", row, col
    If row <> 0 Then
        EMReadScreen MMIS_case_number, 8, row, col + 13
        MMIS_case_number = trim(MMIS_case_number)
    End If
    If MMIS_case_number = "" Then
        row = 1
        col = 1
        EMSearch "CASE NBR:", row, col
        If row <> 0 Then
            EMReadScreen MMIS_case_number, 8, row, col + 10
            MMIS_case_number = trim(MMIS_case_number)
        End If
    End If
    If MMIS_case_number = "" Then
        row = 1
        col = 1
        EMSearch "CASE:", row, col
        If row <> 0 Then
            EMReadScreen MMIS_case_number, 8, row, col + 6
            MMIS_case_number = trim(MMIS_case_number)
        End If
    End If
End Function

Function MMIS_panel_confirmation(panel_name, col)
'--- This function confirms that a user in on the correct MMIS panel.
'~~~~~ panel_name: name of the panel you are confirming as a string in ""
'~~~~~ col: The column which to start reading the panel name. For instance, this is usually 51 or 52 in MMIS.
'===== Keywords: MMIS, navigate, confirm
	Do
		EMReadScreen panel_check, 4, 1, col
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

function navigate_ADDR_buttons(update_addr, err_var, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, mail_street_full, mail_city, mail_state, mail_zip, phone_one, phone_two, phone_three, type_one, type_two, type_three)
'--- This function works with display_ADDR_information to manage the dialog movement
'~~~~~ update_addr: boolean - this will indicate if the dialog information is in 'edit mode' and is adjusted in this function by the button presses
'~~~~~ err_var: string - information output if any parameter if the detail is 'wrong' or missing for dialog entry
'~~~~~ update_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ save_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_mail_addr_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_one_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_two_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ clear_phone_three_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ mail_street_full: string - mailing address information
'~~~~~ mail_city: string - mailing address information
'~~~~~ mail_state: string - mailing address information
'~~~~~ mail_zip: string -  mailing address information
'~~~~~ phone_one: string - phone number information
'~~~~~ phone_two: string - phone number information
'~~~~~ phone_three: string - phone number information
'~~~~~ type_one: string - information about the type of the phone number from the ADDR panel
'~~~~~ type_two: string - information about the type of the phone number from the ADDR panel
'~~~~~ type_three: string - information about the type of the phone number from the ADDR panel
'===== Keywords: MMIS, ADDR, navigate, confirm. dialog
	If ButtonPressed = update_information_btn Then
		update_addr = TRUE
	ElseIf ButtonPressed = save_information_btn Then
		update_addr = FALSE
	Else
		update_addr = FALSE
	End If
	If update_addr = FALSE Then
		If type_one = "Select One..." Then type_one = ""
		If type_two = "Select One..." Then type_two = ""
		If type_three = "Select One..." Then type_three = ""
		If reservation_name = "Select One..." Then reservation_name = ""
	End If

	If ButtonPressed = clear_mail_addr_btn Then
		mail_street_full = ""
		mail_city = ""
		mail_state = ""
		mail_zip = ""
	End If
	If ButtonPressed = clear_phone_one_btn Then
		phone_one = ""
		type_one = ""
	End If
	If ButtonPressed = clear_phone_two_btn Then
		phone_two = ""
		type_two = ""
	End If
	If ButtonPressed = clear_phone_three_btn Then
		phone_three = ""
		type_three = ""
	End If
end function

function navigate_HEST_buttons(update_hest, err_var, update_information_btn, save_information_btn, choice_date, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, date_to_use_for_HEST_standards)
'--- This function works with display_HEST_information to manage the dialog movement
'~~~~~ update_hest: boolean - this will indicate if the dialog information is in 'edit mode' and is adjusted in this function by the button presses
'~~~~~ err_var: string - information output if any parameter if the detail is 'wrong' or missing for dialog entry
'~~~~~ update_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ save_information_btn: number - the button assignment should be set in the dialog to be able to identify which button was pressed
'~~~~~ choice_date: date - the variable to enter into the panel the FS choice date.
'~~~~~ retro_heat_ac_yn: string - the y/n selection for retro heat/ac - will be either 'Y', 'N', ''
'~~~~~ retro_heat_ac_units: string - two digit entry of the number of units responsible for this expense for retro heat/ac
'~~~~~ retro_heat_ac_amt: number - the expense amount of retro heat/ac listed on the HEST panel
'~~~~~ retro_electric_yn: string - the y/n selection for retro electric - will be either 'Y', 'N', ''
'~~~~~ retro_electric_units: string - two digit entry of the number of units responsible for this expense for retro electric
'~~~~~ retro_electric_amt: number - the expense amount of retro electric listed on the HEST panel
'~~~~~ retro_phone_yn: string - the y/n selection for retro phone - will be either 'Y', 'N', ''
'~~~~~ retro_phone_units: string - two digit entry of the number of units responsible for this expense for retro phone
'~~~~~ retro_phone_amt: number - the expense amount of retro phone listed on the HEST panel
'~~~~~ prosp_heat_ac_yn: string - the y/n selection for prosp heat/ac - will be either 'Y', 'N', ''
'~~~~~ prosp_heat_ac_units: string - two digit entry of the number of units responsible for this expense for prosp heat/ac
'~~~~~ prosp_heat_ac_amt: number - the expense amount of prosp heat/ac listed on the HEST panel
'~~~~~ prosp_electric_yn: string - the y/n selection for prosp electric - will be either 'Y', 'N', ''
'~~~~~ prosp_electric_units: string - two digit entry of the number of units responsible for this expense for prosp electric
'~~~~~ prosp_electric_amt: number - the expense amount of prosp electric listed on the HEST panel
'~~~~~ prosp_phone_yn: string - the y/n selection for prosp phone - will be either 'Y', 'N', ''
'~~~~~ prosp_phone_units: string - two digit entry of the number of units responsible for this expense for prosp phone
'~~~~~ prosp_phone_amt: number - the expense amount of prosp phone listed on the HEST panel
'~~~~~ total_utility_expense: number - the amount that will be budgeted for SNAP with SUA
'~~~~~ date_to_use_for_HEST_standards: date - this will indicate the date to use for HEST standards as they change every Oct 1
'===== Keywords: MMIS, HEST, navigate, confirm. dialog

	Call hest_standards(heat_AC_amt, electric_amt, phone_amt, date_to_use_for_HEST_standards)
	If ButtonPressed = update_information_btn Then
		update_hest = TRUE

		retro_heat_ac_amt = retro_heat_ac_amt & ""
		retro_electric_amt = retro_electric_amt & ""
		retro_phone_amt = retro_phone_amt & ""
		prosp_heat_ac_amt = prosp_heat_ac_amt & ""
		prosp_electric_amt = prosp_electric_amt & ""
		prosp_phone_amt = prosp_phone_amt & ""

	ElseIf ButtonPressed = save_information_btn Then
		update_hest = FALSE

		retro_heat_ac_amt = 0
		retro_heat_ac_units = ""
		retro_electric_amt = 0
		retro_electric_units = ""
		retro_phone_amt = 0
		retro_phone_units = ""
		prosp_heat_ac_amt = 0
		prosp_heat_ac_units = ""
		prosp_electric_amt = 0
		prosp_electric_units = ""
		prosp_phone_amt = 0
		prosp_phone_units = ""

		If retro_heat_ac_yn = "Y" Then
			retro_heat_ac_amt = heat_AC_amt
			retro_heat_ac_units = "01"
		End If
		If retro_electric_yn = "Y" Then
			retro_electric_amt = electric_amt
			retro_electric_units = "01"
		End If
		If retro_phone_yn = "Y" Then
			retro_phone_amt = phone_amt
			retro_phone_units = "01"
		End If
		If prosp_heat_ac_yn = "Y" Then
			prosp_heat_ac_amt = heat_AC_amt
			prosp_heat_ac_units = "01"
		End If
		If prosp_electric_yn = "Y" Then
			prosp_electric_amt = electric_amt
			prosp_electric_units = "01"
		End If
		If prosp_phone_yn = "Y" Then
			prosp_phone_amt = phone_amt
			prosp_phone_units = "01"
		End If

		total_utility_expense = 0
		If prosp_heat_ac_yn = "Y" Then
			total_utility_expense =  heat_AC_amt
		ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
			total_utility_expense =  electric_amt + phone_amt
		ElseIf prosp_electric_yn = "Y" Then
			total_utility_expense =  electric_amt
		Elseif prosp_phone_yn = "Y" Then
			total_utility_expense =  phone_amt
		End If

		If IsDate(choice_date) = False then
			update_hest = TRUE
			err_var = err_var & "* You must enter a date in the FS Choice Date field as that is required for the panel update."
		End If

	Else
		update_hest = FALSE
	End If
end function

function navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, THE_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_total_shelter_expense_amount, people_paying_SHEL, verif_detail, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)
'--- This function works with display_HOUSING CHANGE_information to manage the dialog movement
'~~~~~ err_msg - sting - generated by the function for error handling of dialog information
'~~~~~ housing_questions_step - number - to identify which step of the dialog process we are at
'~~~~~ shel_update_step - number - to identify which step of the housing exxpense update process we are at
'~~~~~ notes_on_address - variable - string for entering additional notes
'~~~~~ resi_street_full - string - residence address - this may change
'~~~~~ resi_city - string - residence address - this may change
'~~~~~ resi_state - string - residence address - this may change
'~~~~~ resi_zip - string - residence address - this may change
'~~~~~ resi_county - string - residence address - this may change
'~~~~~ addr_verif - string - residence address - this may change
'~~~~~ addr_homeless - string - residence address - this may change
'~~~~~ addr_reservation - string - residence address - this may change
'~~~~~ reservation_name - string - residence address - this may change
'~~~~~ addr_living_sit - string - residence address - this may change
'~~~~~ mail_street_full - string - mailing address - this may change
'~~~~~ mail_city - string - mailing address - this may change
'~~~~~ mail_state - string - mailing address - this may change
'~~~~~ mail_zip - string - mailing address - this may change
'~~~~~ addr_eff_date - string - address from ADDR panel
'~~~~~ phone_one - string - phone information - this may change
'~~~~~ phone_two - string - phone information - this may change
'~~~~~ phone_three - string - phone information - this may change
'~~~~~ type_one - string - phone information - this may change
'~~~~~ type_two - string - phone information - this may change
'~~~~~ type_three - string - phone information - this may change
'~~~~~ address_change_date - date - enterd in the dialog
'~~~~~ update_information_btn - number - button definition
'~~~~~ save_information_btn - number - button definition
'~~~~~ clear_mail_addr_btn - number - button definition
'~~~~~ clear_phone_one_btn - number - button definition
'~~~~~ clear_phone_two_btn - number - button definition
'~~~~~ clear_phone_three_btn - number - button definition
'~~~~~ household_move_yn - string - dialog input answer
'~~~~~ household_move_everyone_yn - string - dialog input answer
'~~~~~ move_date - date - dialog input answer
'~~~~~ shel_change_yn - string - dialog input answer
'~~~~~ shel_verif_received_yn - string - dialog input answer
'~~~~~ shel_start_date - date - dialog input answer
'~~~~~ shel_shared_yn - string - dialog input answer
'~~~~~ shel_subsidized_yn - string - dialog input answer
'~~~~~ total_current_rent - string - information from SHEL
'~~~~~ all_rent_verif - string - information from SHEL
'~~~~~ total_current_lot_rent - string - information from SHEL
'~~~~~ all_lot_rent_verif - string - information from SHEL
'~~~~~ total_current_garage - string - information from SHEL
'~~~~~ all_mortgage_verif - string - information from SHEL
'~~~~~ total_current_insurance - string - information from SHEL
'~~~~~ all_insurance_verif - string - information from SHEL
'~~~~~ total_current_taxes - string - information from SHEL
'~~~~~ all_taxes_verif - string - information from SHEL
'~~~~~ total_current_room - string - information from SHEL
'~~~~~ all_room_verif - string - information from SHEL
'~~~~~ total_current_mortgage - string - information from SHEL
'~~~~~ all_garage_verif - string - information from SHEL
'~~~~~ total_current_subsidy - string - information from SHEL
'~~~~~ all_subsidy_verif - string - information from SHEL
'~~~~~ shel_change_type - string - dialog input answer
'~~~~~ hest_heat_ac_yn - string - dialog input answer
'~~~~~ hest_electric_yn - string - dialog input answer
'~~~~~ hest_ac_on_electric_yn - string - dialog input answer
'~~~~~ hest_heat_on_electric_yn - string - dialog input answer
'~~~~~ hest_phone_yn - string - dialog input answer
'~~~~~ update_addr_button - number - button definition
'~~~~~ addr_or_shel_change_notes - string - dialog input answer
'~~~~~ view_addr_update_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ view_shel_update_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ view_shel_details_dlg - boolean - to determine what parts of the dialog to show
'~~~~~ what_is_the_living_arrangement - string - dialog input answer
'~~~~~ unit_owned - string - dialog input answer
'~~~~~ new_total_rent_amount - string - dialog input answer
'~~~~~ new_total_mortgage_amount - string - dialog input answer
'~~~~~ new_total_lot_rent_amount - string - dialog input answer
'~~~~~ new_total_room_amount - string - dialog input answer
'~~~~~ new_room_payment_frequency - string - dialog input answer
'~~~~~ new_mortgage_have_escrow_yn - string - dialog input answer
'~~~~~ new_morgage_insurance_amount - string - dialog input answer
'~~~~~ new_excess_insurance_yn - string - dialog input answer
'~~~~~ new_total_tax_amount - string - dialog input answer
'~~~~~ new_rent_subsidy_yn - string - dialog input answer
'~~~~~ new_renter_insurance_amount - string - dialog input answer
'~~~~~ new_renters_insurance_required_yn - string - dialog input answer
'~~~~~ new_total_garage_amount - string - dialog input answer
'~~~~~ new_garage_rent_required_yn - string - dialog input answer
'~~~~~ new_vehicle_insurance_amount - string - dialog input answer
'~~~~~ new_total_insurance_amount - string - dialog input answer
'~~~~~ new_total_subsidy_amount - string - dialog input answer
'~~~~~ new_SHEL_paid_to_name - string - dialog input answer
'~~~~~ other_person_checkbox - 1 or 0 - dialog input checkbox
'~~~~~ other_person_name - string - dialog input answer
'~~~~~ payment_split_evenly_yn - string - dialog input answer
'~~~~~ THE_ARRAY - an ARRAY - of all SHEL panels on the case - filled with access_SHEL_panel
'~~~~~ person_age_const - number - constant used in the ARRAY where the person's age is saved
'~~~~~ person_shel_checkbox - number - constant used in the ARRAY where a checkboxx detail is saved
'~~~~~ shel_ref_number_const - number - constant used in the ARRAY where the reference number of the person is saved
'~~~~~ new_shel_pers_total_amt_const - number - constant used in the ARRAY where the amount of shelter expense paid is saved
'~~~~~ new_shel_pers_total_amt_type_const - number - constant used in the ARRAY where the type (dollars or percent) is indicated
'~~~~~ other_new_shel_total_amt - string - dialog input answer
'~~~~~ other_new_shel_total_amt_type - string - dialog input answer
'~~~~~ new_total_shelter_expense_amount - number - as a string - the total of the new shelter expense
'~~~~~ people_paying_SHEL - string - a list of the persons entered into the dialog as paying the expense
'~~~~~ verif_detail - string - collection of the verifs entered
'~~~~~ new_rent_verif - string - dialog input answer
'~~~~~ new_lot_rent_verif - string - dialog input answer
'~~~~~ new_mortgage_verif - string - dialog input answer
'~~~~~ new_insurance_verif - string - dialog input answer
'~~~~~ new_taxes_verif - string - dialog input answer
'~~~~~ new_room_verif - string - dialog input answer
'~~~~~ new_garage_verif - string - dialog input answer
'~~~~~ new_subsidy_verif - string - dialog input answer
'~~~~~ housing_change_continue_btn - number - button definition
'~~~~~ housing_change_overview_btn - number - button definition
'~~~~~ housing_change_addr_update_btn - number - button definition
'~~~~~ housing_change_shel_update_btn - number - button definition
'~~~~~ housing_change_shel_details_btn - number - button definition
'~~~~~ housing_change_review_btn - number - button definition
'~~~~~ enter_shel_one_btn - number - button definition
'~~~~~ enter_shel_two_btn - number - button definition
'~~~~~ enter_shel_three_btn - number - button definition
'===== Keywords: MAXIS, HEST, SHEL, ADDR, navigate, confirm, dialog
	start_on_shel_questions = True
	If housing_questions_step <> 3 Then start_on_shel_questions = False

	If housing_questions_step = 3 Then

		If (what_is_the_living_arrangement = "Apartment or Townhouse" OR what_is_the_living_arrangement = "House") AND unit_owned = "Select One..." Then err_msg = err_msg & vbCr & "* For Apartment, House, or Mobile Home, you must select if the unit is owned or not to continue."
		If new_rent_subsidy_yn = "Select One..." Then err_msg = err_msg & vbCr & "* You must indiccate if the rent is subsidized."
		If trim(new_renter_insurance_amount) <> "" AND new_renter_insurance_amount <> "0" and new_renters_insurance_required_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have indicated a renters insurance amount, you must indicate if the if the insurance is required by the lease."
		If trim(new_total_garage_amount) <> "" AND new_total_garage_amount <> "0" and new_garage_rent_required_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have indicated a garage expense, you must indicate if the garage rent is required by the leease."
		If trim(new_total_mortgage_amount) <> "" AND new_total_mortgage_amount <> "0" and new_mortgage_have_escrow_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have entered a mortgage amount, you must indicate if there is an escrow with that mortgage payment (if taxes and insurance are included in the expense amount)."
		If trim(new_total_mortgage_amount) <> "" AND new_total_mortgage_amount <> "0" and new_excess_insurance_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since this has a mortgage expensee, you must indicate if the insurance paid has excess ccoverage."
		If new_mortgage_have_escrow_yn = "No" AND trim(new_morgage_insurance_amount) = "" Then err_msg = err_msg & vbCr & "* Since you indicated the mortgage does not have an escrow, you must indicate what the insurance amount is."
		If new_mortgage_have_escrow_yn = "No" AND trim(new_total_tax_amount) = "" Then err_msg = err_msg & vbCr & "* Since you indicated the mortgage does not have an escrow, you must indicate what the tax amount is."
		If trim(new_total_room_amount) <> "" AND new_total_room_amount <> "0" AND new_room_payment_frequency = "Select One..." Then err_msg = err_msg & vbCr & "* You must indicate the frequency of the room expense payment."

		total_current_rent = trim(total_current_rent)
		If total_current_rent = "" Then total_current_rent = 0
		total_current_rent = total_current_rent * 1
		total_current_lot_rent = trim(total_current_lot_rent)
		If total_current_lot_rent = "" Then total_current_lot_rent = 0
		total_current_lot_rent = total_current_lot_rent * 1
		total_current_garage = trim(total_current_garage)
		If total_current_garage = "" Then total_current_garage = 0
		total_current_garage = total_current_garage * 1
		total_current_insurance = trim(total_current_insurance)
		If total_current_insurance = "" Then total_current_insurance = 0
		total_current_insurance = total_current_insurance * 1
		total_current_taxes = trim(total_current_taxes)
		If total_current_taxes = "" Then total_current_taxes = 0
		total_current_taxes = total_current_taxes * 1
		total_current_room = trim(total_current_room)
		If total_current_room = "" Then total_current_room = 0
		total_current_room = total_current_room * 1
		total_current_mortgage = trim(total_current_mortgage)
		If total_current_mortgage = "" Then total_current_mortgage = 0
		total_current_mortgage = total_current_mortgage * 1
		total_current_subsidy = trim(total_current_subsidy)
		If total_current_subsidy = "" Then total_current_subsidy = 0
		total_current_subsidy = total_current_subsidy * 1

		If new_total_rent_amount = "" Then new_total_rent_amount = 0
		new_total_rent_amount = new_total_rent_amount * 1
		If new_renter_insurance_amount = "" Then new_renter_insurance_amount = 0
		new_renter_insurance_amount = new_renter_insurance_amount * 1
		If new_total_garage_amount = "" Then new_total_garage_amount = 0
		new_total_garage_amount = new_total_garage_amount * 1
		If new_total_mortgage_amount = "" Then new_total_mortgage_amount = 0
		new_total_mortgage_amount = new_total_mortgage_amount * 1
		If new_morgage_insurance_amount = "" Then new_morgage_insurance_amount = 0
		new_morgage_insurance_amount = new_morgage_insurance_amount * 1
		If new_total_tax_amount = "" Then new_total_tax_amount = 0
		new_total_tax_amount = new_total_tax_amount * 1
		If new_total_lot_rent_amount = "" Then new_total_lot_rent_amount = 0
		new_total_lot_rent_amount = new_total_lot_rent_amount * 1
		If new_total_room_amount = "" Then new_total_room_amount = 0
		new_total_room_amount = new_total_room_amount * 1
		If new_vehicle_insurance_amount = "" Then new_vehicle_insurance_amount = 0
		new_vehicle_insurance_amount = new_vehicle_insurance_amount * 1
		If new_total_insurance_amount = "" Then new_total_insurance_amount = 0
		new_total_insurance_amount = new_total_insurance_amount * 1
		If new_total_subsidy_amount = "" Then new_total_subsidy_amount = 0
		new_total_subsidy_amount = new_total_subsidy_amount * 1

		new_total_insurance_amount = new_morgage_insurance_amount + new_renter_insurance_amount + new_vehicle_insurance_amount

		new_total_shelter_expense_amount = 0
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_rent_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_garage_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_mortgage_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_tax_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_lot_rent_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_room_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_insurance_amount

		new_total_shelter_expense_amount = new_total_shelter_expense_amount - new_total_subsidy_amount

		people_paying_SHEL = ""
		number_of_people_paying = 0
		for the_membs = 0 to UBound(THE_ARRAY, 2)
			If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
				number_of_people_paying = number_of_people_paying + 1
				people_paying_SHEL = people_paying_SHEL & ", " & "MEMB " & THE_ARRAY(shel_ref_number_const, the_membs)
			End If
		Next
		If other_person_checkbox = checked Then
			number_of_people_paying = number_of_people_paying + 1
			people_paying_SHEL = people_paying_SHEL & ", " & other_person_name
		End If
		If new_rent_subsidy_yn = "Yes" Then people_paying_SHEL = people_paying_SHEL & ", SUBSIDY"
		If left(people_paying_SHEL, 1) = "," Then people_paying_SHEL = right(people_paying_SHEL, len(people_paying_SHEL) - 1)
		people_paying_SHEL = trim(people_paying_SHEL)

		If number_of_people_paying = 1 Then
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
					THE_ARRAY(new_shel_pers_total_amt_const, the_membs) = new_total_shelter_expense_amount & ""
					THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs) = "dollars"
				End If
			Next
			If other_person_checkbox = checked Then
				other_new_shel_total_amt = new_total_shelter_expense_amount & ""
				other_new_shel_total_amt_type = "dollars"
			End If
		ElseIf payment_split_evenly_yn = "Yes" Then
			If number_of_people_paying <> 0 Then
				amount_per_person = new_total_shelter_expense_amount/number_of_people_paying
				for the_membs = 0 to UBound(THE_ARRAY, 2)
					If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
						THE_ARRAY(new_shel_pers_total_amt_const, the_membs) = amount_per_person & ""
						THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs) = "dollars"
					End If
				Next
				If other_person_checkbox = checked Then
					other_new_shel_total_amt = amount_per_person & ""
					other_new_shel_total_amt_type = "dollars"
				End If
			End If
		End if
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_subsidy_amount

		verif_detail = ""
		If new_rent_verif <> "" Then verif_detail = verif_detail & ", Rent - " & trim(right(new_rent_verif, len(new_rent_verif)-4))
		If new_lot_rent_verif <> "" Then verif_detail = verif_detail & ", Lot Rent - " & trim(right(new_lot_rent_verif, len(new_lot_rent_verif)-4))
		If new_mortgage_verif <> "" Then verif_detail = verif_detail & ", Mortgage - " & trim(right(new_mortgage_verif, len(new_mortgage_verif)-4))
		If new_insurance_verif <> "" Then verif_detail = verif_detail & ", Insurance - " & trim(right(new_insurance_verif, len(new_insurance_verif)-4))
		If new_taxes_verif <> "" Then verif_detail = verif_detail & ", Taxes - " & trim(right(new_taxes_verif, len(new_taxes_verif)-4))
		If new_room_verif <> "" Then verif_detail = verif_detail & ", Room - " & trim(right(new_room_verif, len(new_room_verif)-4))
		If new_garage_verif <> "" Then verif_detail = verif_detail & ", Garage - " & trim(right(new_garage_verif, len(new_garage_verif)-4))
		If new_subsidy_verif <> "" Then verif_detail = verif_detail & ", Subisdy - " & trim(right(new_subsidy_verif, len(new_subsidy_verif)-4))
		If left(verif_detail, 1) = "," Then verif_detail = right(verif_detail, len(verif_detail) - 1)
		verif_detail = trim(verif_detail)

		new_total_rent_amount = new_total_rent_amount & ""
		new_renter_insurance_amount = new_renter_insurance_amount & ""
		new_total_garage_amount = new_total_garage_amount & ""
		new_total_mortgage_amount = new_total_mortgage_amount & ""
		new_morgage_insurance_amount = new_morgage_insurance_amount & ""
		new_total_tax_amount = new_total_tax_amount & ""
		new_total_lot_rent_amount = new_total_lot_rent_amount & ""
		new_total_room_amount = new_total_room_amount & ""
		new_vehicle_insurance_amount = new_vehicle_insurance_amount & ""
		new_total_insurance_amount = new_total_insurance_amount & ""
		new_total_subsidy_amount = new_total_subsidy_amount & ""

	End If

	view_shel_details_dlg = False
	If household_move_yn = "?" Then
		view_addr_update_dlg = "Unknown"
		err_msg = "STOP"
	End If

	If shel_change_yn = "?" Then
		view_shel_update_dlg = "Unknown"
		err_msg = "STOP"
	End If

	If household_move_yn = "Yes" Then
		view_addr_update_dlg = True
		shel_change_yn = "Yes"
		view_shel_update_dlg = True
		If shel_shared_yn = "Yes" Then view_shel_details_dlg = True
	End If
	If household_move_yn = "No" Then
		view_addr_update_dlg = False
		view_shel_details_dlg = False
	End If


	If err_msg = "" Then
		If ButtonPressed = housing_change_continue_btn Then
			housing_questions_step = housing_questions_step + 1

			If housing_questions_step = 2 and view_addr_update_dlg = False Then housing_questions_step = housing_questions_step + 1
			If housing_questions_step = 3 and view_shel_update_dlg = False Then housing_questions_step = housing_questions_step + 1

		End If
		If ButtonPressed = housing_change_overview_btn Then housing_questions_step = 1
		If ButtonPressed = housing_change_addr_update_btn Then housing_questions_step = 2
		If ButtonPressed = housing_change_shel_update_btn Then housing_questions_step = 3
		If ButtonPressed = housing_change_review_btn Then housing_questions_step = 4

		If housing_questions_step = 3 Then

			if start_on_shel_questions = False Then shel_update_step = 1

			If ButtonPressed = enter_shel_one_btn Then shel_update_step = 2
			If ButtonPressed = enter_shel_two_btn Then shel_update_step = 3
			If ButtonPressed = enter_shel_three_btn Then shel_update_step = 4

			total_current_rent = total_current_rent & ""
			total_current_lot_rent = total_current_lot_rent & ""
			total_current_garage = total_current_garage & ""
			total_current_insurance = total_current_insurance & ""
			total_current_taxes = total_current_taxes & ""
			total_current_room = total_current_room & ""
			total_current_mortgage = total_current_mortgage & ""
			total_current_subsidy = total_current_subsidy & ""
		End If
	End If
end function

function navigate_SHEL_buttons(update_shel, show_totals, err_var, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, update_information_btn, save_information_btn, const_memb_buttons, const_attempt_update, clear_all_btn, view_total_shel_btn, update_household_percent_button)
'--- This function works with display_SHEL_information to manage the dialog movement
'~~~~~ update_shel: boolean - indicating if the dialog information should be in edit mode or not
'~~~~~ show_totals: boolean - indicates if we are looking at the case total information or the Member specific information
'~~~~~ err_var: string - information output if any parameter if the detail is 'wrong' or missing for dialog entry
'~~~~~ SHEL_ARRAY: The name of the array used for the all the MEMBER panel information, this is in line with the function access_SHEL_panel
'~~~~~ selection: number - This defnies which of the member information from the array should be displayed - defined in navigate_SHEL_buttons
'~~~~~ const_shel_member: number - constant - the defined constant for the array - the member number information
'~~~~~ const_shel_exists: number - constant - the defined constant for the array - boolean - if a SHEL panel exists
'~~~~~ const_hud_sub_yn: number - constant - the defined constant for the array - code from SHEL - if HUD Subsidy exists
'~~~~~ const_shared_yn: number - constant - the defined constant for the array - code from SHEL - if the expense is shared
'~~~~~ const_paid_to: number - constant - the defined constant for the array - from SHEL - who the expense is paid to
'~~~~~ const_rent_retro_amt: number - constant - the defined constant for the array - number - rent amount
'~~~~~ const_rent_retro_verif: number - constant - the defined constant for the array - string - rent verif
'~~~~~ const_rent_prosp_amt: number - constant - the defined constant for the array - number - rent amount
'~~~~~ const_rent_prosp_verif: number - constant - the defined constant for the array - string - rent verif
'~~~~~ const_lot_rent_retro_amt: number - constant - the defined constant for the array - number - lot rent amount
'~~~~~ const_lot_rent_retro_verif: number - constant - the defined constant for the array - string - lot rent verif
'~~~~~ const_lot_rent_prosp_amt: number - constant - the defined constant for the array - number - lot rent amount
'~~~~~ const_lot_rent_prosp_verif: number - constant - the defined constant for the array - string - lot rent verif
'~~~~~ const_mortgage_retro_amt: number - constant - the defined constant for the array - number - mortgage amount
'~~~~~ const_mortgage_retro_verif: number - constant - the defined constant for the array - string - mortgage verif
'~~~~~ const_mortgage_prosp_amt: number - constant - the defined constant for the array - number - mortgage amount
'~~~~~ const_mortgage_prosp_verif: number - constant - the defined constant for the array - string - mortgage verif
'~~~~~ const_insurance_retro_amt: number - constant - the defined constant for the array - number - insurance amount
'~~~~~ const_insurance_retro_verif: number - constant - the defined constant for the array - string - insurance verif
'~~~~~ const_insurance_prosp_amt: number - constant - the defined constant for the array - number - insurance amount
'~~~~~ const_insurance_prosp_verif: number - constant - the defined constant for the array - string - insurance verif
'~~~~~ const_tax_retro_amt: number - constant - the defined constant for the array - number - tax amount
'~~~~~ const_tax_retro_verif: number - constant - the defined constant for the array - string - tax verif
'~~~~~ const_tax_prosp_amt: number - constant - the defined constant for the array - number - tax amount
'~~~~~ const_tax_prosp_verif: number - constant - the defined constant for the array - string - tax verif
'~~~~~ const_room_retro_amt: number - constant - the defined constant for the array - number - room amount
'~~~~~ const_room_retro_verif: number - constant - the defined constant for the array - string - room verif
'~~~~~ const_room_prosp_amt: number - constant - the defined constant for the array - number - room amount
'~~~~~ const_room_prosp_verif: number - constant - the defined constant for the array - string - room verif
'~~~~~ const_garage_retro_amt: number - constant - the defined constant for the array - number - garage amount
'~~~~~ const_garage_retro_verif: number - constant - the defined constant for the array - string - garage verif
'~~~~~ const_garage_prosp_amt: number - constant - the defined constant for the array - number - garage amount
'~~~~~ const_garage_prosp_verif: number - constant - the defined constant for the array - string - garage verif
'~~~~~ const_subsidy_retro_amt: number - constant - the defined constant for the array - number - subsidy amount
'~~~~~ const_subsidy_retro_verif: number - constant - the defined constant for the array - string - subsidy verif
'~~~~~ const_subsidy_prosp_amt: number - constant - the defined constant for the array - number - subsidy amount
'~~~~~ const_subsidy_prosp_verif: number - constant - the defined constant for the array - string - subsidy verif
'~~~~~ update_information_btn: number - defined button
'~~~~~ save_information_btn: number - defined button
'~~~~~ const_memb_buttons: number - constant - the defined constant for the array - defined button
'~~~~~ const_attempt_update: number - constant - the defined constant for the array - if there is a change to the information
'~~~~~ clear_all_btn: number - defined button
'~~~~~ view_total_shel_btn: number - defined button
'~~~~~ update_household_percent_button: number - defined button
'===== Keywords: MMIS, HEST, navigate, confirm. dialog
	If ButtonPressed = update_information_btn Then
		update_shel = TRUE
		update_attempted = True
		' MsgBox "In UPDATE button" & vbCr & vbCr & "Show totals - " & show_totals
	ElseIf ButtonPressed = save_information_btn Then
		update_shel = FALSE
	Else
		update_shel = FALSE
	End If

	If selection <> "" Then
		'REVIEWING THE INFORMATION IN THE ARRAY TO DETERMINE IF IT IS BLANK
		all_shel_details_blank = True

		If Trim(SHEL_ARRAY(const_paid_to, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_hud_sub_yn, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_shared_yn, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_prosp_verif, selection)) <> "" Then all_shel_details_blank = False

		If all_shel_details_blank = True Then SHEL_ARRAY(const_shel_exists, selection) = False

		If ButtonPressed = clear_all_btn Then
			SHEL_ARRAY(const_paid_to, selection) = ""
			SHEL_ARRAY(const_hud_sub_yn, selection) = ""
			SHEL_ARRAY(const_shared_yn, selection) = ""
			SHEL_ARRAY(const_rent_retro_amt, selection) = ""
			SHEL_ARRAY(const_rent_retro_verif, selection) = ""
			SHEL_ARRAY(const_rent_prosp_amt, selection) = ""
			SHEL_ARRAY(const_rent_prosp_verif, selection) = ""
			SHEL_ARRAY(const_lot_rent_retro_amt, selection) = ""
			SHEL_ARRAY(const_lot_rent_retro_verif, selection) = ""
			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = ""
			SHEL_ARRAY(const_lot_rent_prosp_verif, selection) = ""
			SHEL_ARRAY(const_mortgage_retro_amt, selection) = ""
			SHEL_ARRAY(const_mortgage_retro_verif, selection) = ""
			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = ""
			SHEL_ARRAY(const_mortgage_prosp_verif, selection) = ""
			SHEL_ARRAY(const_insurance_retro_amt, selection) = ""
			SHEL_ARRAY(const_insurance_retro_verif, selection) = ""
			SHEL_ARRAY(const_insurance_prosp_amt, selection) = ""
			SHEL_ARRAY(const_insurance_prosp_verif, selection) = ""
			SHEL_ARRAY(const_tax_retro_amt, selection) = ""
			SHEL_ARRAY(const_tax_retro_verif, selection) = ""
			SHEL_ARRAY(const_tax_prosp_amt, selection) = ""
			SHEL_ARRAY(const_tax_prosp_verif, selection) = ""
			SHEL_ARRAY(const_room_retro_amt, selection) = ""
			SHEL_ARRAY(const_room_retro_verif, selection) = ""
			SHEL_ARRAY(const_room_prosp_amt, selection) = ""
			SHEL_ARRAY(const_room_prosp_verif, selection) = ""
			SHEL_ARRAY(const_garage_retro_amt, selection) = ""
			SHEL_ARRAY(const_garage_retro_verif, selection) = ""
			SHEL_ARRAY(const_garage_prosp_amt, selection) = ""
			SHEL_ARRAY(const_garage_prosp_verif, selection) = ""
			SHEL_ARRAY(const_subsidy_retro_amt, selection) = ""
			SHEL_ARRAY(const_subsidy_retro_verif, selection) = ""
			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = ""
			SHEL_ARRAY(const_subsidy_prosp_verif, selection) = ""
			SHEL_ARRAY(const_shel_exists, selection) = False
		End If
	End If

	For memb_btn = 0 to UBound(SHEL_ARRAY, 2)
		If ButtonPressed = SHEL_ARRAY(const_memb_buttons, memb_btn) Then
			selection = memb_btn
			show_totals = False
		End If
	Next
	If selection <> "" Then
		If SHEL_ARRAY(const_shel_exists, selection) = False Then update_shel = True
		If update_shel = True Then
			SHEL_ARRAY(const_attempt_update, selection) = True
			update_attempted = True

			SHEL_ARRAY(const_rent_prosp_amt, selection) = SHEL_ARRAY(const_rent_prosp_amt, selection) & ""
			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = SHEL_ARRAY(const_lot_rent_prosp_amt, selection) & ""
			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = SHEL_ARRAY(const_mortgage_prosp_amt, selection) & ""
			SHEL_ARRAY(const_insurance_prosp_amt, selection) = SHEL_ARRAY(const_insurance_prosp_amt, selection) & ""
			SHEL_ARRAY(const_tax_prosp_amt, selection) = SHEL_ARRAY(const_tax_prosp_amt, selection) & ""
			SHEL_ARRAY(const_room_prosp_amt, selection) = SHEL_ARRAY(const_room_prosp_amt, selection) & ""
			SHEL_ARRAY(const_garage_prosp_amt, selection) = SHEL_ARRAY(const_garage_prosp_amt, selection) & ""
			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = SHEL_ARRAY(const_subsidy_prosp_amt, selection) & ""
		End If
	End If
	If ButtonPressed = view_total_shel_btn Then
		show_totals = True
		selection = ""
	End If
	If show_totals = True and update_shel = True Then
		total_paid_by_household = total_paid_by_household & ""
		total_paid_by_others = total_paid_by_others & ""
		total_current_rent = total_current_rent & ""
		total_current_lot_rent = total_current_lot_rent & ""
		total_current_mortgage = total_current_mortgage & ""
		total_current_insurance = total_current_insurance & ""
		total_current_taxes = total_current_taxes & ""
		total_current_room = total_current_room & ""
		total_current_garage = total_current_garage & ""
		total_current_subsidy = total_current_subsidy & ""
	End If
	' MsgBox "End NAVIGATE" & vbCr & vbCr & "Show totals - " & show_totals
end function

function navigate_to_MAXIS_screen_review_PRIV(function_to_go_to, command_to_go_to, is_this_priv)
'--- This function is to be used to navigate to a specific MAXIS screen and will check for privileged status
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'~~~~~ is_this_priv: This returns a true or false based on if the case appears to be privileged in MAXIS
'===== Keywords: MAXIS, navigate
  is_this_priv = FALSE
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    EMReadScreen locked_panel, 23, 2, 30
    IF locked_panel = "Program History Display" then PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = MAXIS_case_number and MAXIS_function = ucase(function_to_go_to) and STAT_note_check <> "NOTE" then
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen command_to_go_to, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen function_to_go_to, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen MAXIS_case_number, 18, 43
      EMWriteScreen MAXIS_footer_month, 20, 43
      EMWriteScreen MAXIS_footer_year, 20, 46
      EMWriteScreen command_to_go_to, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
	  EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	  If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits

	  EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it will return true as privileged response.
	  If priv_check = "PRIVIL" THEN is_this_priv = TRUE
    End if
  End if
end function

function navigate_to_MAXIS(maxis_mode)
'--- This function is to be used when navigating back to MAXIS from another function in BlueZone (MMIS, PRISM, INFOPAC, etc.)
'~~~~~ maxis_mode: This parameter needs to be "maxis_mode"
'===== Keywords: MAXIS, navigate
    attn
    Do
        EMReadScreen MAI_check, 3, 1, 33
        If MAI_check <> "MAI" then EMWaitReady 1, 1
    Loop until MAI_check = "MAI"

    EMReadScreen prod_check, 7, 6, 15
    IF prod_check = "RUNNING" THEN
        Call write_value_and_transmit("1", 2, 15)
    ELSE
        EMConnect"A"
        attn
        EMReadScreen prod_check, 7, 6, 15
        IF prod_check = "RUNNING" THEN
            Call write_value_and_transmit("1", 2, 15)
        ELSE
            EMConnect"B"
            attn
            EMReadScreen prod_check, 7, 6, 15
            IF prod_check = "RUNNING" THEN
                Call write_value_and_transmit("1", 2, 15)
            Else
                script_end_procedure("You do not appear to have Production mode running. This script will now stop. Please make sure you have production and MMIS open in the same session, and re-run the script.")
            END IF
        END IF
    END IF
end function

function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
'--- This function is to be used to navigate to a specific MAXIS screen
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'===== Keywords: MAXIS, navigate
	function_to_go_to = UCase(function_to_go_to)
	command_to_go_to = UCase(command_to_go_to)
	EMSendKey "<enter>"
	EMWaitReady 0, 0
	EMReadScreen MAXIS_check, 5, 1, 39
	If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
		already_at_the_correct_screen = False
		review_footer_month = False
		at_correct_footer_month = False

		EMReadScreen locked_panel, 23, 2, 30
		IF locked_panel = "Program History Display" then
			PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
		END IF
		row = 1
		col = 1
		EMSearch "Function: ", row, col
		If row <> 0 then
			EMReadScreen MAXIS_function, 4, row, col + 10
			EMReadScreen STAT_note_check, 4, 2, 45
			row = 1
			col = 1
			EMSearch "Case Nbr: ", row, col
			EMReadScreen current_case_number, 8, row, col + 10
			current_case_number = replace(current_case_number, "_", "")
			current_case_number = trim(current_case_number)
			EMReadScreen all_of_function_row, 80, row, 1
			footer_month_location = InStr(all_of_function_row, "Month")
			footer_month_location = footer_month_location + 7
			If function_to_go_to = "STAT" Then review_footer_month = True
			If function_to_go_to = "ELIG" Then review_footer_month = True
			If MAXIS_footer_month = "" OR MAXIS_footer_year = "" Then review_footer_month = False
			EMReadScreen current_footer_month, 2, row, footer_month_location
			EMReadScreen current_footer_year, 2, row, footer_month_location+3
			If review_footer_month = True Then
				If current_footer_month = MAXIS_footer_month AND current_footer_year = MAXIS_footer_year Then at_correct_footer_month = True
			Else
				at_correct_footer_month = True
			End If

			EMReadScreen all_of_row_two, 80, 2, 1
			If current_case_number = MAXIS_case_number and MAXIS_function = function_to_go_to AND InStr(all_of_row_two, command_to_go_to) <> 0 AND at_correct_footer_month = True Then already_at_the_correct_screen = True
		End if

		If already_at_the_correct_screen = False Then
			If current_case_number = MAXIS_case_number and MAXIS_function = function_to_go_to and STAT_note_check <> "NOTE" and at_correct_footer_month = True then
				row = 1
				col = 1
				EMSearch "Command: ", row, col
				EMWriteScreen command_to_go_to, row, col + 9
				EMSendKey "<enter>"
				EMWaitReady 0, 0
			Else
				Do
					EMSendKey "<PF3>"
					EMWaitReady 0, 0
					EMReadScreen SELF_check, 4, 2, 50
				Loop until SELF_check = "SELF"
				EMWriteScreen function_to_go_to, 16, 43
				EMWriteScreen "________", 18, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen MAXIS_footer_month, 20, 43
				EMWriteScreen MAXIS_footer_year, 20, 46
				EMWriteScreen command_to_go_to, 21, 70
				EMSendKey "<enter>"
				EMWaitReady 0, 0
				EMReadScreen abended_check, 7, 9, 27
				If abended_check = "abended" then
					EMSendKey "<enter>"
					EMWaitReady 0, 0
				End if
				EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
				If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
			End if
		End If
	End if
end function

function navigate_to_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
            BeginDialog Password_dialog, 0, 0, 156, 55, "Password Dialog"
            ButtonGroup ButtonPressed
            OkButton 45, 35, 50, 15
            CancelButton 100, 35, 50, 15
            Text 5, 5, 150, 25, "You have passworded out. Please enter your password, then press OK to continue. Press CANCEL to stop the script. "
            EndDialog
            Do
                Do
                    dialog Password_dialog
                    cancel_confirmation
                Loop until ButtonPressed = -1
			    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	 		Loop until are_we_passworded_out = false					'loops until user passwords back in
		End if
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		End Select

        'Now it finds the recipient file application feature and selects it.
        row = 1
        col = 1
        EMSearch menu_to_enter, row, col
        EMWriteScreen "X", row, col - 3
        transmit

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "X", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "X", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		End Select
        EMWaitReady 0, 0
	End If
end function

Function non_actionable_dails(actionable_dail)
'--- This function used to determine if a DAIL message is actionable or non-actionable as determined by the QI Team.
'~~~~~ actionable_dail: boolean variable to determine if message is actionable or not.
'===== Keywords: MAXIS, DAIL
    'Valuing variables used within the function
    this_month = CM_mo & " " & CM_yr
    next_month = CM_plus_1_mo & " " & CM_plus_1_yr
    CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

    actionable_dail = True    'Defaulting to True
    If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
        instr(dail_msg, "AP OF CHILD REF NBR:") OR _
        instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
        instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN LBUD IN THE MONTH") OR _
        instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN SBUD IN THE MONTH") OR _
        instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
        instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
        instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
        instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
        instr(dail_msg, "IS COOPERATING WITH CHILD SUPPORT") OR _
        instr(dail_msg, "IS NOT COOPERATING WITH CHILD SUPPORT") OR _
        instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
        instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
        instr(dail_msg, "REPORTED NAME CHG TO:") OR _
        instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
        instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
        instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
        instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
        instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
        instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
        instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
        instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
        instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
        instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
        instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
        instr(dail_msg, "WARNING MESSAGES EXIST") OR _
        instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
        instr(dail_msg, "APPLCT ID CHNGD") OR _
        instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
        instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
        instr(dail_msg, "CASE NOTE ENTERED BY") OR _
        instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
        instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
        instr(dail_msg, "CASE XFER") OR _
        instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
        instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
        instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
        instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
        instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
        instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
        instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
        instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
        instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
        instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
        instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
        instr(dail_msg, "PMI MERGED") OR _
        instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
        instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
        instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
        instr(dail_msg, "LAST GRADE COMPLETED") OR _
        instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
        instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
        instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
        instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
        instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
        instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
        instr(dail_msg, "GA: STATUS IS PENDING - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "GA: STATUS IS REIN OR SUSPEND - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
        instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
        instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
        instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
        instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
        instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
        instr(dail_msg, "SDNH NEW JOB DETAILS FOR MEMB 00") OR _
        instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
        instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
        instr(dail_msg, "SSN DIFFERS W/ CS RECORDS") OR _
        instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
        instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
        instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
        instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
        instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
        instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
        instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
        instr(dail_msg, "- TRANS #") OR _
        instr(dail_msg, "RSDI UPDATED - (REF") OR _
        instr(dail_msg, "SSI UPDATED - (REF") OR _
        instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then
            actionable_dail = False
        '----------------------------------------------------------------------------------------------------STAT EDITS older than Current Date
    Elseif dail_type = "STAT" or instr(dail_msg, "NEW FIAT RESULTS EXIST") then
        EmReadscreen stat_date, 8, dail_row, 39     'Stat date location
        If isdate(stat_date) = False then
            EmReadscreen alt_stat_date, 8, dail_row, 43 'fiat results date location
            If isdate(alt_stat_date) = True then
                stat_date = alt_stat_date
            End if
        End if
        If isdate(stat_date) = True then
            If DateDiff("d", stat_date, date) > 0 then
                actionable_dail = False     'Deleting any messages that were not created taday
            Else
                actionable_dail = True
            End if
        End if
    '----------------------------------------------------------------------------------------------------REMOVING PEPR messages not CM or CM + 1
    Elseif dail_type = "PEPR" then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
    '----------------------------------------------------------------------------------------------------clearing ELIG messages older than CM
    Elseif instr(dail_msg, "OVERPAYMENT POSSIBLE") or instr(dail_msg, "DISBURSE EXPEDITED SERVICE FS") or instr(dail_msg, "NEW FS VERSION MUST BE APPROVED") or instr(dail_msg, "APPROVE NEW ELIG RESULTS RECOUPMENT HAS INCREASED") or instr(dail_msg, "PERSON/S REQD FS NOT IN FS UNIT") then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
    '----------------------------------------------------------------------------------------------------SSN messages older than CM or CM +1
    Elseif instr(dail_msg, "SSN UNMATCHED DATA EXISTS") then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
    '----------------------------------------------------------------------------------------------------DISB CS messages older than CM or CM +1
    Elseif instr(dail_msg, "DISB CS") then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
    '----------------------------------------------------------------------------------------------------clearing Exempt IR TIKL's over 2 months old.
    Elseif instr(dail_msg, "%^% SENT THROUGH") then
        TIKL_date = cdate(TIKL_date)
        TIKL_date = right("0" & DatePart("m",dail_month), 2)
        if TIKL_date = CM_minus_2_mo then
            actionable_dail = False   ' delete the exempt IR message older than last month.
        Else
            actionable_dail = True
        End if
        '----------------------------------------------------------------------------------------------------MEC2
    Elseif dail_type = "MEC2" then
        if  instr(dail_msg, "RSDI END DATE") OR _
            instr(dail_msg, "SELF EMPLOYMENT REPORTED TO MEC²") OR _
            instr(dail_msg, "SSI REPORTED TO MEC²") OR _
            instr(dail_msg, "UNEMPLOYMENT INS") then
            actionable_dail = True            'Income based MEC2 messages will not be removed
        Else
            actionable_dail = False    'All other MEC2 messages can be deleted.
        End if
        '----------------------------------------------------------------------------------------------------SVES older than CM or CM + 1
    Elseif dail_type = "SVES" then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
        '----------------------------------------------------------------------------------------------------TIKL
    Elseif dail_type = "TIKL" then
        if instr(dail_msg, "VENDOR") OR instr(dail_msg, "VND") then
            actionable_dail = True        'Will not delete TIKL's with vendor information
        Else
            six_months = DateAdd("M", -6, date)
            If cdate(six_months) => cdate(dail_month) then
                actionable_dail = False     'Will delete any TIKL over 6 months old
            Else
                actionable_dail = True
            End if
        End if
    Else
        actionable_dail = True
    End if
End Function

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
    date_variable = dateadd("d", 0, date_variable)    'janky way to convert to a date, but hey it works.
    var_month     = right("0" & DatePart("m",    date_variable), 2)
    var_day       = right("0" & DatePart("d",    date_variable), 2)
    var_year      = right("0" & DatePart("yyyy", date_variable), 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function

function open_URL_in_browser(URL_to_open)
'--- This function is to be used to open a URL in user's default browser
'~~~~~ URL_to_open: web address to open
'===== Keywords: URL, open, web
	CreateObject("WScript.Shell").Run(URL_to_open)
end function

function PF1()
'--- This function sends or hits the PF1 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
end function

function PF2()
'--- This function sends or hits the PF2 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
end function

function PF3()
'--- This function sends or hits the PF3 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

function PF4()
'--- This function sends or hits the PF4 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
end function

function PF5()
'--- This function sends or hits the PF5 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
end function

function PF6()
'--- This function sends or hits the PF6 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
end function

function PF7()
'--- This function sends or hits the PF7 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
end function

function PF8()
'--- This function sends or hits the PF8 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9()
'--- This function sends or hits the PF9 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10()
'--- This function sends or hits the PF10 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

function PF11()
'--- This function sends or hits the PF11 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
end function

function PF12()
'--- This function sends or hits the PF12 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
end function

function PF13()
'--- This function sends or hits the PF13 key (SHIFT+F1).
 '===== Keywords: MAXIS, MMIS, PRISM, PF13
  EMSendKey "<PF13>"
  EMWaitReady 0, 0
end function

function PF14()
'--- This function sends or hits the PF14 key (SHIFT+F2).
 '===== Keywords: MAXIS, MMIS, PRISM, PF14
  EMSendKey "<PF14>"
  EMWaitReady 0, 0
end function

function PF15()
'--- This function sends or hits the PF15 key (SHIFT+F3).
 '===== Keywords: MAXIS, MMIS, PRISM, PF15
  EMSendKey "<PF15>"
  EMWaitReady 0, 0
end function

function PF16()
'--- This function sends or hits the PF16 key (SHIFT+F4).
 '===== Keywords: MAXIS, MMIS, PRISM, PF16
  EMSendKey "<PF16>"
  EMWaitReady 0, 0
end function

function PF17()
'--- This function sends or hits the PF17 key (SHIFT+F5).
 '===== Keywords: MAXIS, MMIS, PRISM, PF17
  EMSendKey "<PF17>"
  EMWaitReady 0, 0
end function

function PF18()
'--- This function sends or hits the PF18 key (SHIFT+F6).
 '===== Keywords: MAXIS, MMIS, PRISM, PF18
  EMSendKey "<PF18>"
  EMWaitReady 0, 0
end function

function PF19()
'--- This function sends or hits the PF19 key (SHIFT+F7).
 '===== Keywords: MAXIS, MMIS, PRISM, PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
end function

function PF20()
'--- This function sends or hits the PF20 key (SHIFT+F8).
 '===== Keywords: MAXIS, MMIS, PRISM, PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function PF21()
'--- This function sends or hits the PF21 key (SHIFT+F9).
 '===== Keywords: MAXIS, MMIS, PRISM, PF21
  EMSendKey "<PF21>"
  EMWaitReady 0, 0
end function

function PF22()
'--- This function sends or hits the PF22 key (SHIFT+F10).
 '===== Keywords: MAXIS, MMIS, PRISM, PF22
  EMSendKey "<PF22>"
  EMWaitReady 0, 0
end function

function PF23()
'--- This function sends or hits the PF23 key (SHIFT+F11).
 '===== Keywords: MAXIS, MMIS, PRISM, PF23
  EMSendKey "<PF23>"
  EMWaitReady 0, 0
end function

function PF24()
'--- This function sends or hits the PF24 key (SHIFT+F12).
 '===== Keywords: MAXIS, MMIS, PRISM, PF24
  EMSendKey "<PF24>"
  EMWaitReady 0, 0
end function

function proceed_confirmation(result_of_msgbox)
'--- This function asks the user if they want to proceed.
'~~~~~ result_of_msgbox: returns TRUE if Yes is pressed, and FALSE if No is pressed.
'===== Keywords: MAXIS, MMIS, PRISM, dialog, proceed, confirmation
	If ButtonPressed = -1 then
		proceed_confirm = MsgBox("Are you sure you want to proceed? Press Yes to continue, No to return to the previous screen, and Cancel to end the script.", vbYesNoCancel)
		If proceed_confirm = vbCancel then stopscript
		If proceed_confirm = vbYes then result_of_msgbox = TRUE
		If proceed_confirm = vbNo then result_of_msgbox = FALSE
	End if
end function

function provide_resources_information(case_number_known, create_case_note, note_detail_array, allow_cancel)
'--- Call the resources notifier dialog and display functionality. This can be added as part of any process.
'~~~~~ case_number_known: BOOLEAN - Indicate if the Case Number has already been confirmed earlier in the script run.
'~~~~~ create_case_note: BOOLEAN - Indicate if the function should create a case note.
'~~~~~ note_detail_array: Enter a variable to pass through an array with all of the case note detail lines. This way the detail can be entered within another note.
'~~~~~ allow_cancel: BOOLEAN - Indicate if the Cancel button and stopscript should be included in this function.
'===== Keywords: MAXIS, dialog, communication, Word, MEMO
	Dim MEMO_LINES_ARRAY()					'declaring an array so we can add MEMOs
	array_counter = 0						'setting the start of a counter
	no_resources_checkbox = unchecked		'making sure this is known even if it is not displayed, as it is optional
	'this variable is used to know if we have resources to send.

	Do
		DO
			Do
				err_msg = ""
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 211, 330, "Resources MEMO"			'dialog defined within the do loops because there is a nother in the larges loop
				  ButtonGroup ButtonPressed
				  	If case_number_known = FALSE Then
						EditBox 60, 5, 50, 15, MAXIS_case_number
						Text 10, 10, 50, 10, "Case number:"
					End If
					If allow_cancel = FALSE Then CheckBox 60, 25, 140, 10, "Check here if no resources are needed.", no_resources_checkbox
				    PushButton 150, 5, 50, 10, "Check All", check_all_button
					CheckBox 15, 50, 145, 10, "Document Submission Options", client_virtual_dropox_checkbox
					CheckBox 15, 65, 140, 10, "Community Action Partnership - CAP", cap_checkbox
					CheckBox 15, 80, 115, 10, "DHS MMIS Recipient HelpDesk", MMIS_helpdesk_checkbox
					CheckBox 15, 95, 180, 10, "DHS MNSure Helpdesk   * NOT FOR MA CLIENTS", MNSURE_helpdesk_checkbox
					CheckBox 15, 110, 145, 10, "Disability Hub (Disability Linkage Line)", disability_hub_checkbox
					CheckBox 15, 125, 175, 10, "Emergency Food Shelf Network (The Food Group)", emer_food_network_checkbox
					CheckBox 15, 140, 125, 10, "Emergency Mental Health Services", emer_mental_health_checkbox
					CheckBox 15, 155, 50, 10, "Front Door", front_door_checkbox
					CheckBox 15, 170, 75, 10, "Senior Linkage Line", sr_linkage_line_checkbox
					Text 15, 185, 185, 10, "Shelters for residents experiencing housing insecurity:"
					CheckBox 20, 195, 175, 10, "Family Shelters - Hennepin County Shelter Team", family_shelter_checkbox
					CheckBox 20, 210, 95, 10, "Single Adults Shelters", single_adults_shelter_checkbox
					CheckBox 20, 225, 110, 10, "Domestic Violence Shelter", domestic_violence_shelters_checkbox
					CheckBox 15, 240, 130, 10, "United Way First Call for Help (211)", united_way_checkbox
					CheckBox 15, 255, 35, 10, "WIC", wic_checkbox
					CheckBox 15, 270, 60, 10, "Xcel Energy", xcel_checkbox
					EditBox 80, 290, 125, 15, worker_signature
					If allow_cancel = TRUE Then
					    OkButton 100, 310, 50, 15
					    CancelButton 155, 310, 50, 15
					End If
					If allow_cancel = FALSE Then
						OkButton 155, 310, 50, 15
					End If
					GroupBox 5, 40, 200, 245, "Check any to send detail about the service to a client"
					Text 10, 295, 65, 10, "Worker signature:"
				EndDialog

				Dialog Dialog1			'showing the dialog

				resource_selected = False										'figuring out if any of the resources have been selected or not
				If client_virtual_dropox_checkbox = checked Then 	resource_selected = True
				If cap_checkbox = checked Then 						resource_selected = True
				If MMIS_helpdesk_checkbox = checked Then 			resource_selected = True
				If MNSURE_helpdesk_checkbox = checked Then 			resource_selected = True
				If disability_hub_checkbox = checked Then 			resource_selected = True
				If emer_food_network_checkbox = checked Then 		resource_selected = True
				If emer_mental_health_checkbox = checked Then 		resource_selected = True
				If front_door_checkbox = checked Then 				resource_selected = True
				If sr_linkage_line_checkbox = checked Then 			resource_selected = True
				If family_shelter_checkbox = checked Then 			resource_selected = True
				If single_adults_shelter_checkbox = checked Then 	resource_selected = True
				If domestic_violence_shelters_checkbox = checked Then resource_selected = True
				If united_way_checkbox = checked Then 				resource_selected = True
				If wic_checkbox = checked Then 						resource_selected = True
				If xcel_checkbox = checked Then 					resource_selected = True

				'dialog message handling
				If allow_cancel = TRUE Then
					cancel_without_confirmation
					If resource_selected = False Then err_msg = err_msg & vbNewLine & "You must select at least one resource."
				End If
				If allow_cancel = FALSE Then
					If resource_selected = False AND no_resources_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "Either select a resource or indicate none were needed by checking the box."
					If no_resources_checkbox = checked AND resource_selected = True Then err_msg = err_msg & vbNewLine & "You cannot indicate no resources AND indicate some resources. Review the checked boxes."
				End If
				If case_number_known = FALSE Then
					If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "You must fill in a valid case number." & vbNewLine
				End If
				If worker_signature = "" then err_msg = err_msg & "You must sign your case note." & vbNewLine
		        If ButtonPressed = check_all_button Then		'checking all the boxes if the button to check all the boxes is pushed
		            err_msg = "LOOP" & err_msg

					client_virtual_dropox_checkbox = checked
		            cap_checkbox = checked
		            MMIS_helpdesk_checkbox = checked
		            MNSURE_helpdesk_checkbox = checked
		            disability_hub_checkbox = checked
		            emer_food_network_checkbox = checked
		            emer_mental_health_checkbox = checked
		            front_door_checkbox = checked
		            sr_linkage_line_checkbox = checked
					family_shelter_checkbox = checked
					single_adults_shelter_checkbox = checked
					domestic_violence_shelters_checkbox = checked
		            united_way_checkbox = checked
					wic_checkbox = checked
		            xcel_checkbox = checked
		        End If
				IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN msgbox "Please resolve to continue:" & vbCr & vbCr & err_msg
			Loop until err_msg = ""
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false

		If no_resources_checkbox = unchecked Then		'If any resources are checked
			ReDim MEMO_LINES_ARRAY(0)					'resizing the array down to 0
			array_counter = 0							'resetting the array counter

			'Adding all of the resource information into the array so we know what the MEMO will look like.
			script_to_say = "Resource detail:" & vbNewLine
			MEMO_LINES_ARRAY(0) = "  ----Outside Resources - current as of " & date & "----"

			If client_virtual_dropox_checkbox = checked Then
				array_counter = array_counter + 6
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-5) = "*** Submitting Documents:"
				MEMO_LINES_ARRAY(array_counter-4) = "- Online at infokeep.hennepin.us or MNBenefits.mn.gov"
				MEMO_LINES_ARRAY(array_counter-3) = "  Use InfoKeep to upload documents directly to your case."
				MEMO_LINES_ARRAY(array_counter-2) = "- Mail, Fax, or Drop Boxes at Service Centers."
				MEMO_LINES_ARRAY(array_counter-1) = "  More Info: https://www.hennepin.us/economic-supports"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If cap_checkbox = checked Then
				array_counter = array_counter + 7
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-6) = "CAP - Community Action Partnership (Includes Energy Assist)"
		        MEMO_LINES_ARRAY(array_counter-5) = "Hours: Mon-Fri 8:00AM - 4:30PM Website: www.caphennepin.org"
		        MEMO_LINES_ARRAY(array_counter-4) = "Locations: Minneapolis Urban League   Phone: 952-930-3541"
		        MEMO_LINES_ARRAY(array_counter-3) = "           MN Council of Churches     Phone: 952-933-9639"
		        MEMO_LINES_ARRAY(array_counter-2) = "           Sabathani Community Center Phone: 952-930-3541"
		        MEMO_LINES_ARRAY(array_counter-1) = "           St. Louis Park             Phone: 952-933-9639"
			  	MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If MMIS_helpdesk_checkbox = checked Then
				array_counter = array_counter + 2
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-1) = "MN Health Care Recipient Help Desk - 651-431-2670"
			    MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If MNSURE_helpdesk_checkbox = checked Then
				array_counter = array_counter + 2
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-1) = "MNSure Helpdesk - 1-855-366-7873 (1-855-3MNSURE)"
			    MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If disability_hub_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "Disability Hub (formerly Disability Linkage Line)"
		        MEMO_LINES_ARRAY(array_counter-2) = "Phone: 1-866-333-2466 -Hrs: Mon - Fri 8:00AM - 5:00PM"
		        MEMO_LINES_ARRAY(array_counter-1) = "Website: disabilityhubmn.org"
		        MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If emer_food_network_checkbox = checked Then
				array_counter = array_counter + 3
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-2) = "The Food Group (formerly Emergency Food Network)"
		        MEMO_LINES_ARRAY(array_counter-1) = "Phone: 763-450-3860  - Website: thefoodgroupmn.org"
		        MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If emer_mental_health_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "Emergency Mental Health Services"
		        MEMO_LINES_ARRAY(array_counter-2) = "Adults 18 and older (COPE): 612-596-1223"
		        MEMO_LINES_ARRAY(array_counter-1) = "Children (Child Crisis Services): 612-348-2233"
		        MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If front_door_checkbox = checked Then
				array_counter = array_counter + 2
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-1) = "Hennepin County FRONT DOOR - 612-348-4111"
			    MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If sr_linkage_line_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "Senior Linkage Line"
		        MEMO_LINES_ARRAY(array_counter-2) = "Phone: 1-800-333-2433  - Hours: Mon - Fri 8:00 AM - 4:30 PM"
				MEMO_LINES_ARRAY(array_counter-1) = "Website: metroaging.org"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If family_shelter_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "Family Shelter Team  --  Phone: 612-348-9410"
				MEMO_LINES_ARRAY(array_counter-2) = "Hours: Mondays - Fridays: 8 AM - 11 PM"
				MEMO_LINES_ARRAY(array_counter-1) = "       Weekends/Holidays: 1 PM - 11PM"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If single_adults_shelter_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "Adult Shelter"
				MEMO_LINES_ARRAY(array_counter-2) = "Phone: 612-248-2350  --  Mon - Fri 8 AM - 4 PM"
				MEMO_LINES_ARRAY(array_counter-1) = "Phone: 211 (651-291-0211)  --  All Other Hours"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If domestic_violence_shelters_checkbox = checked Then
				array_counter = array_counter + 3
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-2) = "Domestic Violence Shelter"
				MEMO_LINES_ARRAY(array_counter-1) = "Day One Shelter  --  Phone: 866-223-1111"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If united_way_checkbox = checked Then
				array_counter = array_counter + 4
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-3) = "United Way First Call for Help (211) - DIAL 211"
		        MEMO_LINES_ARRAY(array_counter-2) = "Phone: 1-800-543-7709 OR 651-291-0211  - Available 24 Hrs"
		        MEMO_LINES_ARRAY(array_counter-1) = "Website: www.211unitedway.org"
		        MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If wic_checkbox = checked Then
				array_counter = array_counter + 3
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-2) = "WIC - Women, Infants, and Children"
				MEMO_LINES_ARRAY(array_counter-1) = "Phone: 612-348-6100"
				MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If
			If xcel_checkbox = checked Then
				array_counter = array_counter + 2
				ReDim Preserve MEMO_LINES_ARRAY(array_counter)
				MEMO_LINES_ARRAY(array_counter-1) = "Xcel Energy - 1-800-331-5262"
			    MEMO_LINES_ARRAY(array_counter) = "--   --   --   --   --   --   --   --   --   --   --"
			End If

			For each memo_line in MEMO_LINES_ARRAY
				script_to_say = script_to_say & vbNewLine & memo_line
			Next
			script_to_say = script_to_say & vbNewLine & vbNewLine & "Relay any of the above information to the client verbally now." & vbNewLine &_
			    "Then press OK and all of this detail will be added to a SPEC/MEMO so the client can have the information in writing."

			MsgBox script_to_say		'This shows the resources in a MSGBOX so the resources can be given verbally to the resident

			'selecting the way the resources should be sent to the resident
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 196, 70, "Dialog"
			  DropListBox 55, 30, 135, 45, "SPEC/MEMO"+chr(9)+"Word Document"+chr(9)+"Do Not Send", resource_method
			  ButtonGroup ButtonPressed
			    OkButton 140, 50, 50, 15
			  Text 5, 10, 195, 10, "How do you want to send this information to the resident?"
			EndDialog
			Do
			    dialog Dialog1

			    Call check_for_password(are_we_passworded_out)
			Loop until are_we_passworded_out = FALSE

			If array_counter > 29 AND resource_method = "SPEC/MEMO" Then
				MsgBox "MEMOs only allow for 30 lines and you have " & array_counter + 1 & " lines based on the resources you have selected." & vbCr & vbCr &_
					   "     --   --   --   --   --   --   --   --   --   --   --   --   --   --   --   --   --   --   --     " & vbCr & vbCr &_
					   "The checkbox dialog will appear and you can reselct the checkboxes to fit within the MEMO.", vbImportant, "MEMO too large"
			End If
		End If
		If resource_method = "Do Not Send" Then no_resources_checkbox = checked		'this sets the variable for the next part of the script to indicate if additional action is needed.
	Loop until array_counter < 30 OR resource_method <> "SPEC/MEMO"

	If no_resources_checkbox = unchecked Then
		If resource_method = "SPEC/MEMO" Then		'Creating a MEMO if that is the option selected
			Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, allow_cancel)	' start the memo writing process

			For each memo_line in MEMO_LINES_ARRAY		'using the array created above
				call write_variable_in_SPEC_MEMO(memo_line)
			Next
		    PF4
		End If

		If resource_method = "Word Document" Then		''Creating a WORD DOCUMENT if that is the option selected
		    '****writing the word document
		    Set objWord = CreateObject("Word.Application")
		    Const wdDialogFilePrint = 88
		    Const end_of_doc = 6
		    objWord.Caption = "Outside Resource Information"
		    objWord.Visible = True

		    Set objDoc = objWord.Documents.Add()
		    Set objSelection = objWord.Selection

		    objSelection.PageSetup.LeftMargin = 50
		    objSelection.PageSetup.RightMargin = 50
		    objSelection.PageSetup.TopMargin = 30
		    objSelection.PageSetup.BottomMargin = 25

		    todays_date = date & ""
		    objSelection.Font.Name = "Arial"
		    objSelection.Font.Size = "14"
		    objSelection.Font.Bold = TRUE
		    objSelection.TypeText "Outside Resource Information - Current as of "
		    objSelection.TypeText todays_date
		    objSelection.TypeParagraph()
		    objSelection.ParagraphFormat.SpaceAfter = 0

		    objSelection.Font.Size = "12"
		    objSelection.Font.Bold = FALSE
			If client_virtual_dropox_checkbox = checked Then
				objSelection.TypeText "You can submit documents Online at www.MNbenefits.mn.gov or" & vbCr
				objSelection.TypeText "Email with document attachment. EMAIL: hhsews@hennepin.us" & vbCr
				objSelection.TypeText " (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types)" & vbCr
				objSelection.TypeParagraph()
			End If
		    If cap_checkbox = checked Then
		        objSelection.TypeText "* CAP - Community Action Partnership (Inc. Energy Assist)" & vbCr
		        objSelection.TypeText "  Hours: Mon-Fri 8:00AM - 4:30PM Website: www.caphennepin.org" & vbCr
		        objSelection.TypeText "  Locations: Minneapolis Urban League   Phone: 952-930-3541" & vbCr
		        objSelection.TypeText "                    MN Council of Churches     Phone: 952-933-9639" & vbCr
		        objSelection.TypeText "                    Sabathani Community Center Phone: 952-930-3541" & vbCr
		        objSelection.TypeText "                    St. Louis Park             Phone: 952-933-9639" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If MMIS_helpdesk_checkbox = checked Then
		        objSelection.TypeText "* MN Health Care Recipient Help Desk - 651-431-2670" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If MNSURE_helpdesk_checkbox = checked Then
		        objSelection.TypeText "* MNSure Helpdesk - 1-855-366-7873 (1-855-3MNSURE)" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If disability_hub_checkbox = checked Then
		        objSelection.TypeText "* Disability Hub (formerly Disability Linkage Line)" & vbCr
		        objSelection.TypeText "    Phone: 1-866-333-2466 -Hrs: Mon - Fri 8:00AM - 5:00PM" & vbCr
		        objSelection.TypeText "    Website: disabilityhubmn.org" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If emer_food_network_checkbox = checked Then
		        objSelection.TypeText "* The Food Group (formerly Emergency Food Network)" & vbCr
		        objSelection.TypeText "     Phone: 763-450-3860  - Website: thefoodgroupmn.org" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If emer_mental_health_checkbox = checked Then
		        objSelection.TypeText "* Emergency Mental Health Services" & vbCr
		        objSelection.TypeText "       Adults 18 and older (COPE): 612-596-1223" & vbCr
		        objSelection.TypeText "       Children (Child Crisis Services): 612-348-2233" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If front_door_checkbox = checked Then
		        objSelection.TypeText "* Hennepin County FRONT DOOR - 612-348-4111" & vbCr
		        objSelection.TypeParagraph()
		    End If
		    If sr_linkage_line_checkbox = checked Then
		        objSelection.TypeText "* Senior Linkage Line" & vbCr
		        objSelection.TypeText "   Phone: 1-800-333-2433 - Hours: Mon - Fri 8:00AM - 4:30PM" & vbCr
		        objSelection.TypeText "   Currently has extended hours Mon - Thur 4:30PM - 6:30PM" & vbCr
		        objSelection.TypeText "   Website: metroaging.org" & vbCr
		        objSelection.TypeParagraph()
		    End If
			If family_shelter_checkbox = checked Then
				objSelection.TypeText "* Family Shelter Team" & vbCr
				objSelection.TypeText "   Phone: 612-348-9410" & vbCr
				objSelection.TypeText "   Hours: Mon - Fri 8 AM - 11 PM" & vbCr
				objSelection.TypeText "   Weekends/Holidays 1 PM - 11PM" & vbCr
				objSelection.TypeParagraph()
			End If
			If single_adults_shelter_checkbox = checked Then
				objSelection.TypeText "* Adult Shelter" & vbCr
				objSelection.TypeText "   Phone: 612-248-2350  --  Mon - Fri 8 AM - 4 PM" & vbCr
				objSelection.TypeText "   Phone: 211 (651-291-0211)  --  All Other Hours" & vbCr
				objSelection.TypeParagraph()
			End If
			If domestic_violence_shelters_checkbox = checked Then
				objSelection.TypeText "* Domestic Violence Shelter" & vbCr
				objSelection.TypeText "   Day One Shelter Phone: 866-223-1111" & vbCr
				objSelection.TypeParagraph()
			End If
		    If united_way_checkbox = checked Then
		        objSelection.TypeText "* United Way First Call for Help (211)" & vbCr
		        objSelection.TypeText "   Phone: 1-800-543-7709 OR 651- 291-0211 - 24 Hrs" & vbCr
		        objSelection.TypeText "   Website: www.211unitedway.org" & vbCr
		        objSelection.TypeParagraph()
		    End If
			If wic_checkbox = checked Then
				 objSelection.TypeText "* WIC - Women, Infants, and Children" & vbCr
				 objSelection.TypeText "   Phone: 612 348-6100" & vbCr
				 objSelection.TypeText "   Website: www.hennepin.us/residents/health-medical/wic-women-infants-children" & vbCr
				objSelection.TypeParagraph()
			End If
		    If xcel_checkbox = checked Then
		        objSelection.TypeText "* Xcel Energy - 1-800-331-5262" & vbCr
		        objSelection.TypeParagraph()
		    End If
		End If

		'create an array of the actions taken. This needs to be an array so we can send it out of the function if the function does not create a CASE:NOTEs
		note_detail_array = ""
		If resource_method = "SPEC/MEMO" Then note_detail_array = note_detail_array & "::* Information added to SPEC/MEMO to send in overnight batch."
		If resource_method = "Word Document" Then note_detail_array = note_detail_array & "::* Information added to Word Document for printing locally."

		If client_virtual_dropox_checkbox = checked Then note_detail_array = note_detail_array & "::* Client Virtual Dropbox."
		IF cap_checkbox = checked Then note_detail_array = note_detail_array & "::* Compunity Action Partnership - CAP (Energy Assistance)"
		IF MMIS_helpdesk_checkbox = checked Then note_detail_array = note_detail_array & "::* DHS MHCP Recipient HelpDesk"
		IF MNSURE_helpdesk_checkbox = checked Then note_detail_array = note_detail_array & "::* DHS MNSure HelpDesk"
		IF disability_hub_checkbox = checked Then note_detail_array = note_detail_array & "::* Disability Hub"
		IF emer_food_network_checkbox = checked Then note_detail_array = note_detail_array & "::* Emergency Food Network"
		IF emer_mental_health_checkbox = checked Then note_detail_array = note_detail_array & "::* Emergency Mental Health Services"
		IF front_door_checkbox = checked Then note_detail_array = note_detail_array & "::* Front Door"
		IF sr_linkage_line_checkbox = checked Then note_detail_array = note_detail_array & "::* Senior Linkage Line"
		IF family_shelter_checkbox = checked Then note_detail_array = note_detail_array & "::* Family Shelter Team"
		IF single_adults_shelter_checkbox = checked Then note_detail_array = note_detail_array & "::* Single Adulte Shelter Information"
		IF domestic_violence_shelters_checkbox = checked Then note_detail_array = note_detail_array & "::* Domestic Violence Shelter"
		IF united_way_checkbox = checked Then note_detail_array = note_detail_array & "::* United Way - 211"
		IF wic_checkbox = checked Then note_detail_array = note_detail_array & "::* WIC Information"
		IF xcel_checkbox = checked Then note_detail_array = note_detail_array & "::* Xcel Energy"

		If left(note_detail_array, 2) = "::" Then note_detail_array = right(note_detail_array, len(note_detail_array)-2)
		note_detail_array = split(note_detail_array, "::")

		If create_case_note = TRUE Then
			'Navigates to CASE/NOTE and starts a blank one
			start_a_blank_CASE_NOTE

			'Writes the case note--------------------------------------------
			call write_variable_in_CASE_NOTE("Outside resource information sent to client")

			For each note_line in note_detail_array
				Call write_variable_in_CASE_NOTE(note_line)
			Next

			call write_variable_in_CASE_NOTE("---")
			call write_variable_in_CASE_NOTE(worker_signature)
		End If
	Else
		note_detail_array = array()
	End If
end function

function read_boolean_from_excel(excel_place, script_variable)
'--- This function Will take the information in from the Excel cell and reformat it so that the script can use the information as a boolean
'~~~~~ excel_place: the cell value code - using 'objexcel.cells(r,c).value' format/information
'~~~~~ script_variable: whatever variable you want to use to store the information from this Excel location - this CAN be an array position.
'===== Keywords: MAXIS, Excel, output, boolean
	script_variable = trim(excel_place)
	script_variable = UCase(script_variable)

	If script_variable = "TRUE" Then script_variable = True
	If script_variable = "FALSE" Then script_variable = False
	'If this is not TRUE or FALSE, then it will just output what was in the cell all uppercase
end function

function read_inqb_for_all_issuances(months_to_go_back, beginning_footer_month, ISSUED_BENEFITS_ARRAY, fn_footer_month_const, fn_footer_year_const, fn_snap_issued_const, fn_snap_recoup_const, fn_ga_issued_const, fn_ga_recoup_const, fn_msa_issued_const, fn_msa_recoup_const, fn_mf_mf_issued_const, fn_mf_mf_recoup_const, fn_mf_fs_issued_const, fn_mf_hg_issued_const, fn_dwp_issued_const, fn_dwp_recoup_const, fn_emer_issued_const, fn_emer_prog_const, fn_grh_issued_const, fn_grh_recoup_const, fn_no_issuance_const, fn_last_const, snap_found, ga_found, msa_found, mfip_found, dwp_found, grh_found)
'--- This function works in association with gather_case_benefits_details to read hisoical issuance information from INQB
'~~~~~ months_to_go_back: NUMBER - information with how many months back to look
'~~~~~ beginning_footer_month - string of the first month to display in the dialog
'~~~~~ ISSUED_BENEFITS_ARRAY - This is the name of the array used to store past information
'~~~~~ fn_footer_month_const - CONSTANT for array details
'~~~~~ fn_footer_year_const - CONSTANT for array details
'~~~~~ fn_snap_issued_const - CONSTANT for array details
'~~~~~ fn_snap_recoup_const - CONSTANT for array details
'~~~~~ fn_ga_issued_const - CONSTANT for array details
'~~~~~ fn_ga_recoup_const - CONSTANT for array details
'~~~~~ fn_msa_issued_const - CONSTANT for array details
'~~~~~ fn_msa_recoup_const - CONSTANT for array details
'~~~~~ fn_mf_mf_issued_const - CONSTANT for array details
'~~~~~ fn_mf_mf_recoup_const - CONSTANT for array details
'~~~~~ fn_mf_fs_issued_const - CONSTANT for array details
'~~~~~ fn_mf_hg_issued_const - CONSTANT for array details
'~~~~~ fn_dwp_issued_const - CONSTANT for array details
'~~~~~ fn_dwp_recoup_const - CONSTANT for array details
'~~~~~ fn_emer_issued_const - CONSTANT for array details
'~~~~~ fn_emer_prog_const - CONSTANT for array details
'~~~~~ fn_grh_issued_const - CONSTANT for array details
'~~~~~ fn_grh_recoup_const - CONSTANT for array details
'~~~~~ fn_no_issuance_const - CONSTANT for array details
'~~~~~ fn_last_const - CONSTANT for array details
'~~~~~ snap_found - BOOLEAN - detailing if SNAP was found
'~~~~~ ga_found - BOOLEAN - detailing if SNAP was found
'~~~~~ msa_found - BOOLEAN - detailing if SNAP was found
'~~~~~ mfip_found - BOOLEAN - detailing if SNAP was found
'~~~~~ dwp_found - BOOLEAN - detailing if SNAP was found
'~~~~~ grh_found - BOOLEAN - detailing if SNAP was found
'===== Keywords: MAXIS, DIALOG, CLIENT

    ReDim ISSUED_BENEFITS_ARRAY(fn_last_const, 0)								'reset the array to blank it out from a previous run

    now_month = CM_mo & "/1/" & CM_yr											'setting the month from the footer months to have it start at the 1st
    now_month = DateAdd("d", 0, now_month)

    subtract_months = 0-months_to_go_back										'making this number negativ
    start_month = DateAdd("m", subtract_months, now_month)						'finding the first month to look at
    start_month_mo = right("00"&DatePart("m", start_month), 2)					'setting the footer month to start'
    start_month_yr = right(DatePart("yyyy", start_month), 2)
    beginning_footer_month = start_month_mo & "/" & start_month_yr				'string for display'
    month_to_review = start_month

    snap_found = False		'setting the default of these booleans'
    ga_found = False
    msa_found = False
    mfip_found = False
    dwp_found = False
    grh_found = False

    count_months = 0		'this is the array incrementer
    Do
        ReDim Preserve ISSUED_BENEFITS_ARRAY(fn_last_const, count_months)		'resize the array
        Call convert_date_into_MAXIS_footer_month(month_to_review, MAXIS_footer_month, MAXIS_footer_year)
        year_to_search = DatePart("yyyy", month_to_review)						'finding the month namme and year to search in INQB as these are written out
        year_to_search = year_to_search & ""
        month_to_search = MonthName(DatePart("m", month_to_review))
        ISSUED_BENEFITS_ARRAY(fn_footer_month_const, count_months) = MAXIS_footer_month		'saving the month information to the array
        ISSUED_BENEFITS_ARRAY(fn_footer_year_const, count_months) = MAXIS_footer_year
        ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = True

        Call back_to_SELF														'Navigating to INQB
        Call navigate_to_MAXIS_screen("MONY", "INQB")
		'We are going to read each row to find the corrent month and read the programs and amounts
        inqb_row = 6
        Do
            EMReadScreen inqb_month, 12, inqb_row, 3							'reading the month and year
            EMReadScreen inqb_year, 4, inqb_row, 16
            inqb_month = trim(inqb_month)
            If inqb_month = month_to_search and inqb_year = year_to_search Then	'if the month and year on the row match the ones we are looking for.
                EMReadScreen inqb_prog, 2, inqb_row, 23							'reading the details
                EMReadScreen inqb_amt, 10, inqb_row, 38
                EMReadScreen inqb_recoup, 10, inqb_row, 53
                EMReadScreen inqb_food, 10, inqb_row, 69
                EMReadScreen inqb_full, 77, inqb_row, 3							'I had to read the whole row to get the PROG information for some reason.

                If InStr(inqb_full, "FS") <> 0 Then								'If FS is listed, the script will add this detail to the SNAP place in the array
                    ISSUED_BENEFITS_ARRAY(fn_snap_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_snap_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    snap_found = True
                End If
                If InStr(inqb_full, "GA") <> 0 Then								'If GA is listed, the script will add this detail to the GA place in the array
                    ISSUED_BENEFITS_ARRAY(fn_ga_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_ga_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    ga_found = True
                End If
                If InStr(inqb_full, "MS") <> 0 Then								'If MS is listed, the script will add this detail to the MMSA place in the array
                    ISSUED_BENEFITS_ARRAY(fn_msa_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_msa_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    msa_found = True
                End If
                If InStr(inqb_full, "MF-MF") <> 0 Then							'If MF-MF is listed, the script will add this detail to the MFIP - MFF place in the array
                    ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_mf_mf_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    mfip_found = True
                End If
                If InStr(inqb_full, "MF-FS") <> 0 Then							'If MF-FS is listed, the script will add this detail to the MFIP - FS place in the array
                    ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    mfip_found = True
                End If
                If InStr(inqb_full, "MF-HG") <> 0 Then							'If MF-HG is listed, the script will add this detail to the MFIP - HG place in the array
                    ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    mfip_found = True
                End If
                If InStr(inqb_full, "DW") <> 0 Then								'If DW is listed, the script will add this detail to the DWP place in the array
                    ISSUED_BENEFITS_ARRAY(fn_dwp_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_dwp_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    dwp_found = True
                End If
                If InStr(inqb_full, "GR") <> 0 Then								'If GR is listed, the script will add this detail to the GRH place in the array
                    ISSUED_BENEFITS_ARRAY(fn_grh_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_grh_recoup_const, count_months) = trim(inqb_recoup)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                    grh_found = True
                End If
                If InStr(inqb_full, "EA") <> 0 Then								'If EA is listed, the script will add this detail to the AMER place in the array
                    ISSUED_BENEFITS_ARRAY(fn_emer_issued_const, count_months) = trim(inqb_amt)
                    ISSUED_BENEFITS_ARRAY(fn_no_issuance_const, count_months) = False
                End If
            End If

            inqb_row = inqb_row + 1												'going to the row
            EMReadScreen next_prog, 2, inqb_row, 23								'read for if we are at the end of the list
        Loop until next_prog = "  "
		'setting MFIP information to 0 if one is blank (IE MF-FS was issued but no MF-MF)
        If ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, count_months) = "" and (ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, count_months) <> "" OR ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, count_months) <> "") Then ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, count_months) = "0.00"
        If ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, count_months) = "" and (ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, count_months) <> "" OR ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, count_months) <> "") Then ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, count_months) = "0.00"
        If ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, count_months) = "" and (ISSUED_BENEFITS_ARRAY(fn_mf_fs_issued_const, count_months) <> "" OR ISSUED_BENEFITS_ARRAY(fn_mf_mf_issued_const, count_months) <> "") Then ISSUED_BENEFITS_ARRAY(fn_mf_hg_issued_const, count_months) = "0.00"
		PF3

        month_to_review = DateAdd("m", 1, month_to_review)						'going to the next month
        count_months = count_months + 1											'incrementing the Array
    Loop Until DateDiff("d", now_month, month_to_review) > 0					'going until we ahve read the currnt month
	Call Back_to_SELF
end function

function read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, rent_amt, rent_verif, lot_rent_amt, lot_rent_verif, mortgage_amt, mortgage_verif, insurance_amt, insurance_verif, taxes_amt, taxes_verif, room_amt, room_verif, garage_amt, garage_verif, subsidy_amt, subsidy_verif, total_shelter_expense, original_information)
'--- Function to read all of the SHEL panels and total everything listed on each panel to a final total.
'~~~~~ ref_numbers_with_panel: string of all member reference numbers that have a SHEL panel existing - seperated by "~"
'~~~~~ paid_to: string - of who the sheler expense is paid to. If there is more than one on different panels, this will say 'Multiple'
'~~~~~ rent_amt: number - the total of the prospective rent amount listed on all SHEL panels in the case
'~~~~~ rent_verif: string - the verification listed on the panel for the rent expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ lot_rent_amt: number - the total of the prospective lot rent amount listed on all SHEL panels in the case
'~~~~~ lot_rent_verif: string - the verification listed on the panel for the lot rent expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ mortgage_amt: number - the total of the prospective mortgage amount listed on all SHEL panels in the case
'~~~~~ mortgage_verif: string - the verification listed on the panel for the mortgage expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ insurance_amt: number - the total of the prospective insurance amount listed on all SHEL panels in the case
'~~~~~ insurance_verif: string - the verification listed on the panel for the insurance expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ taxes_amt: number - the total of the prospective taxes amount listed on all SHEL panels in the case
'~~~~~ taxes_verif: string - the verification listed on the panel for the taxes expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ room_amt: number - the total of the prospective room amount listed on all SHEL panels in the case
'~~~~~ room_verif: string - the verification listed on the panel for the room expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ garage_amt: number - the total of the prospective garage amount listed on all SHEL panels in the case
'~~~~~ garage_verif: string - the verification listed on the panel for the garage expense. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ subsidy_amt: number - the total of the prospective subsidy amount listed on all SHEL panels in the case
'~~~~~ subsidy_verif: string - the verification listed on the panel for the subsidy amount. If there are different verifications on different SHEL panels on the case, this will say 'Multiple'
'~~~~~ original_information: string - combination of all information read
'===== Keywords: MAXIS, SHEL
	'SEARCH THE LIST OF HOUSEHOLD MEMBERS TO SEARCH ALL SHEL PANELS
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadScreen access_denied_check, 13, 24, 2
		'MsgBox access_denied_check
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			last_name = "UNABLE TO FIND"
			first_name = " - Access Denied"
			mid_initial = ""
		Else
			client_list = client_list & ref_nbr & "|"
		End If
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_list = TRIM(client_list)				'making this list an array so we can go through it
	If right(client_list, 1) = "|" Then client_list = left(client_list, len(client_list) - 1)
	shel_ref_numbers_array = split(client_list, "|")

	rent_amt = 0			'setting the defaults of these parameters
	rent_verif = ""
	lot_rent_amt = 0
	lot_rent_verif = ""
	mortgage_amt = 0
	mortgage_verif = ""
	insurance_amt = 0
	insurance_verif = ""
	taxes_amt = 0
	taxes_verif = ""
	room_amt = 0
	room_verif = ""
	garage_amt = 0
	garage_verif = ""
	subsidy_amt = 0
	subsidy_verif = ""

	Call navigate_to_MAXIS_screen("STAT", "SHEL")								'going to SHEL

	For each memb_ref_number in shel_ref_numbers_array							'We are going to look at SHEL for each member.
		EMWriteScreen memb_ref_number, 20, 76
		transmit

		EMReadScreen shel_version, 1, 2, 78
		If shel_version = "1" Then												'if a SHEL panel exists, it will say '1' here
			ref_numbers_with_panel = ref_numbers_with_panel & "~" & memb_ref_number	'saving the member reference number to the list

		    EMReadScreen panel_paid_to,               25, 7, 50					'reading the panel paid to
		    panel_paid_to = replace(panel_paid_to, "_", "")						'formatting the string
			If paid_to = "" Then												'saving the information about who the SHEL expenses are paid to
				paid_to = panel_paid_to
			ElseIf paid_to <> panel_paid_to Then
				paid_to = "Multiple"
			End If

		    EMReadScreen rent_prosp_amt,        8, 11, 56						'reading rent amount and verification
		    EMReadScreen rent_prosp_verif,      2, 11, 67

		    rent_prosp_amt = replace(rent_prosp_amt, "_", "")					'formatiing the information about the amount ans adding it in to the running total
		    rent_prosp_amt = trim(rent_prosp_amt)
			If rent_prosp_amt = "" Then rent_prosp_amt = 0
			rent_prosp_amt = rent_prosp_amt * 1
			rent_amt = rent_amt + rent_prosp_amt

		    If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"		'formatiing the verif information for rent and adding it to the output parameter
		    If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
		    If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
		    If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Doc"
		    If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg, Neg Impact"
		    If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg, Pos Impact"
		    If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Verif"
			If rent_prosp_verif = "?_" Then rent_prosp_verif = "? - Delayed Verif"
		    If rent_prosp_verif = "__" Then rent_prosp_verif = ""
			If rent_verif = "" Then
				rent_verif = rent_prosp_verif
			ElseIf rent_verif <> rent_prosp_verif Then
				rent_verif = "Multiple"
			End If

		    EMReadScreen lot_rent_prosp_amt,    8, 12, 56						'reading lot rent amount and verification
		    EMReadScreen lot_rent_prosp_verif,  2, 12, 67

		    lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")			'formatiing the information about the amount and adding it in to the running total
		    lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
			If lot_rent_prosp_amt = "" Then lot_rent_prosp_amt = 0
			lot_rent_prosp_amt = lot_rent_prosp_amt * 1
			lot_rent_amt = lot_rent_amt + lot_rent_prosp_amt
		    If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"				'formatiing the verif information for rent and adding it to the output parameter
		    If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
		    If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
		    If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Doc"
		    If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg, Neg Impact"
		    If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg, Pos Impact"
		    If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Verif"
			If lot_rent_prosp_verif = "?_" Then lot_rent_prosp_verif = "? - Delayed Verif"
		    If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""
			If lot_rent_verif = "" Then
				lot_rent_verif = lot_rent_prosp_verif
			ElseIf lot_rent_verif <> lot_rent_prosp_verif Then
				lot_rent_verif = "Multiple"
			End If

		    EMReadScreen mortgage_prosp_amt,    8, 13, 56						'reading mortgage amount and verification
		    EMReadScreen mortgage_prosp_verif,  2, 13, 67

		    mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")			'formatiing the information about the amount and adding it in to the running total
		    mortgage_prosp_amt = trim(mortgage_prosp_amt)
			If mortgage_prosp_amt = "" Then mortgage_prosp_amt = 0
			mortgage_prosp_amt = mortgage_prosp_amt * 1
			mortgage_amt = mortgage_amt + mortgage_prosp_amt
		    If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"		'formatiing the verif information for rent and adding it to the output parameter
		    If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
		    If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Doc"
		    If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg, Neg Impact"
		    If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg, Pos Impact"
		    If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Verif"
			If mortgage_prosp_verif = "?_" Then mortgage_prosp_verif = "? - Delayed Verif"
		    If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""
			If mortgage_verif = "" Then
				mortgage_verif = mortgage_prosp_verif
			ElseIf mortgage_verif <> mortgage_prosp_verif Then
				mortgage_verif = "Multiple"
			End If

		    EMReadScreen insurance_prosp_amt,   8, 14, 56						'reading insurance amount and verification
		    EMReadScreen insurance_prosp_verif, 2, 14, 67

		    insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")			'formatiing the information about the amount and adding it in to the running total
		    insurance_prosp_amt = trim(insurance_prosp_amt)
			If insurance_prosp_amt = "" Then insurance_prosp_amt = 0
			insurance_prosp_amt = insurance_prosp_amt * 1
			insurance_amt = insurance_amt + insurance_prosp_amt
		    If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"		'formatiing the verif information for rent and adding it to the output parameter
		    If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Doc"
		    If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg, Neg Impact"
		    If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg, Pos Impact"
		    If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Verif"
			If insurance_prosp_verif = "?_" Then insurance_prosp_verif = "? - Delayed Verif"
		    If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""
			If insurance_verif = "" Then
				insurance_verif = insurance_prosp_verif
			ElseIf insurance_verif <> insurance_prosp_verif Then
				insurance_verif = "Multiple"
			End If

		    EMReadScreen tax_prosp_amt,         8, 15, 56						'reading tax amount and verification
		    EMReadScreen tax_prosp_verif,       2, 15, 67

		    tax_prosp_amt = replace(tax_prosp_amt, "_", "")						'formatiing the information about the amount and adding it in to the running total
		    tax_prosp_amt = trim(tax_prosp_amt)
			If tax_prosp_amt = "" Then tax_prosp_amt = 0
			tax_prosp_amt = tax_prosp_amt * 1
			taxes_amt = taxes_amt + tax_prosp_amt
		    If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"		'formatiing the verif information for rent and adding it to the output parameter
		    If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Doc"
		    If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg, Neg Impact"
		    If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg, Pos Impact"
		    If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Verif"
			If tax_prosp_verif = "?_" Then tax_prosp_verif = "? - Delayed Verif"
		    If tax_prosp_verif = "__" Then tax_prosp_verif = ""
			If taxes_verif = "" Then
				taxes_verif = tax_prosp_verif
			ElseIf taxes_verif <> tax_prosp_verif Then
				taxes_verif = "Multiple"
			End If

		    EMReadScreen room_prosp_amt,        8, 16, 56						'reading room amount and verification
		    EMReadScreen room_prosp_verif,      2, 16, 67

		    room_prosp_amt = replace(room_prosp_amt, "_", "")					'formatiing the information about the amount and adding it in to the running total
		    room_prosp_amt = trim(room_prosp_amt)
			If room_prosp_amt = "" Then room_prosp_amt = 0
			room_prosp_amt = room_prosp_amt * 1
			room_amt = room_amt + room_prosp_amt
		    If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"		'formatiing the verif information for rent and adding it to the output parameter
		    If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
		    If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
		    If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Doc"
		    If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg, Neg Impact"
		    If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg, Pos Impact"
		    If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Verif"
			If room_prosp_verif = "?_" Then room_prosp_verif = "? - Delayed Verif"
		    If room_prosp_verif = "__" Then room_prosp_verif = ""
			If room_verif = "" Then
				room_verif = room_prosp_verif
			ElseIf room_verif <> room_prosp_verif Then
				room_verif = "Multiple"
			End If

		    EMReadScreen garage_prosp_amt,      8, 17, 56						'reading garage amount and verification
		    EMReadScreen garage_prosp_verif,    2, 17, 67

		    garage_prosp_amt = replace(garage_prosp_amt, "_", "")				'formatiing the information about the amount and adding it in to the running total
		    garage_prosp_amt = trim(garage_prosp_amt)
			If garage_prosp_amt = "" Then garage_prosp_amt = 0
			garage_prosp_amt = garage_prosp_amt * 1
			garage_amt = garage_amt + garage_prosp_amt
		    If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"		'formatiing the verif information for rent and adding it to the output parameter
		    If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
		    If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
		    If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Doc"
		    If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg, Neg Impact"
		    If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg, Pos Impact"
		    If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Verif"
			If garage_prosp_verif = "?_" Then garage_prosp_verif = "? - Delayed Verif"
		    If garage_prosp_verif = "__" Then garage_prosp_verif = ""
			If garage_verif = "" Then
				garage_verif = garage_prosp_verif
			ElseIf garage_verif <> garage_prosp_verif Then
				garage_verif = "Multiple"
			End If

		    EMReadScreen subsidy_prosp_amt,     8, 18, 56						'reading subsidy amount and verification
		    EMReadScreen subsidy_prosp_verif,   2, 18, 67

		    subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")				'formatiing the information about the amount and adding it in to the running total
		    subsidy_prosp_amt = trim(subsidy_prosp_amt)
			If subsidy_prosp_amt = "" Then subsidy_prosp_amt = 0
			subsidy_prosp_amt = subsidy_prosp_amt * 1
			subsidy_amt = subsidy_amt + subsidy_prosp_amt
		    If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"		'formatiing the verif information for rent and adding it to the output parameter
		    If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
		    If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Doc"
		    If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Verif"
			If subsidy_prosp_verif = "?_" Then subsidy_prosp_verif = "? - Delayed Verif"
		    If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""
			If subsidy_verif = "" Then
				subsidy_verif = subsidy_prosp_verif
			ElseIf subsidy_verif <> subsidy_prosp_verif Then
				subsidy_verif = "Multiple"
			End If
		End If
	Next
	total_shelter_expense = rent_prosp_amt + lot_rent_prosp_amt + mortgage_prosp_amt + insurance_prosp_amt + tax_prosp_amt + room_prosp_amt + garage_prosp_amt

	If left(ref_numbers_with_panel, 1) = "~" Then ref_numbers_with_panel = right(ref_numbers_with_panel, len(ref_numbers_with_panel)-1)		'formatting the list of reference numbers with a SHEL panel
	'saving the information into the original information parameter for comparison
	original_information = rent_amt&"|"&rent_verif&"|"&lot_rent_amt&"|"&lot_rent_verif&"|"&mortgage_amt&"|"&mortgage_verif&"|"&insurance_amt&"|"&insurance_verif&"|"&taxes_amt&"|"&taxes_verif&"|"&room_amt&"|"&room_verif&"|"&garage_amt&"|"&garage_verif&"|"&subsidy_amt&"|"&subsidy_verif
end function

function reformat_phone_number(phone_number, format_needed)
'--- This function will take a phone number and output it with different formatting.
'~~~~~ phone_number: the number as it currently exists. This should be a 10 digit number
'~~~~~ format_needed: enter the format desired - this should use 111, 222, 3333 as the parts of the phone number - eg '( 111 ) 222 - 3333'
'===== Keywords: phone number, variable, format
	original_phone_number = phone_number				'saving the number to an unchanged variable
	phone_number = replace(phone_number, "-", "")		'removing all extra characters from the variable
	phone_number = replace(phone_number, "(", "")
	phone_number = replace(phone_number, ")", "")
	phone_number = replace(phone_number, " ", "")

	If len(phone_number) = 10 Then						'making sure this phone number is 10 digit.
		phone_part_one = left(phone_number, 3)			'getting each part of the phone number
		phone_part_two = mid(phone_number, 4, 3)
		phone_part_three = right(phone_number, 4)

		temp_phone = replace(format_needed, "111", phone_part_one)	'using the placeholders in the format variable to place the phone parts in
		temp_phone = replace(temp_phone, "222", phone_part_two)
		temp_phone = replace(temp_phone, "3333", phone_part_three)
		phone_number = temp_phone
	Else
		phone_number = original_phone_number			'if this phone variable is NOT 10 digit, it outputs the original variable.
	end If
end function

function run_another_script(script_path)
'--- This function runs another script from a specific file either stored locally or on the web.
'~~~~~ script_path: path of script to run
'===== Keywords: MAXIS, MMIS, PRISM, run, script, script path
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  ExecuteGlobal text_from_the_other_script
end function

function run_from_GitHub(url)
'--- This function runs a script file from GitHub
'~~~~~ url: web address of script file on GitHub
'===== Keywords: MAXIS, MMIS, PRISM, run, script, url, GitHub
	'Creates a list of items to remove from anything run from GitHub. This will allow for counties to use Option Explicit handling without fear.
	list_of_things_to_remove = array("OPTION EXPLICIT", _
									"option explicit", _
									"Option Explicit", _
									"dim case_number", _
									"DIM case_number", _
									"Dim case_number")
	If run_locally = "" or run_locally = False then					'Runs the script from GitHub if we're not set up to run locally.
		Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
		req.open "GET", url, False									'Attempts to open the URL
		req.send													'Sends request
		If req.Status = 200 Then									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			script_contents = req.responseText						'Empties the response into a variable called script_contents
			'Uses a for/next to remove the list_of_things_to_remove
			FOR EACH phrase IN list_of_things_to_remove
				script_contents = replace(script_contents, phrase, "")
			NEXT
			ExecuteGlobal script_contents							'Executes the remaining script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
											"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
											vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
			script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		call run_another_script(url)
	END IF
end function

function script_end_procedure(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then MsgBox closing_message '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Defaulting script success to successful
        SCRIPT_success = -1

		'Determines if the value of the MAXIS case number - BULK scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" then MAXIS_CASE_NUMBER = ""

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		IF using_SQL_database = TRUE then
    		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
		ELSE
			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
		END IF

        'Adds some data for users of the old database, but adds lots more data for users of the new.
        If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
    		'Opening usage_log and adding a record
    		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
    		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function

function script_end_procedure_with_error_report(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
    send_error_message = ""
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
        If testing_run = TRUE Then
            MsgBox(closing_message & vbNewLine & vbNewLine & "Since this script is in testing, please provide feedback")
            send_error_message = vbYes
        Else
            send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Script Run Completed")
        End If
    End If
    script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Defaulting script success to successful
        SCRIPT_success = -1

        'Determines if the value of the MAXIS case number - BULK scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" then MAXIS_CASE_NUMBER = ""

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		IF using_SQL_database = TRUE then
    		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
		ELSE
			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
		END IF

        'Adds some data for users of the old database, but adds lots more data for users of the new.
        If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
    		'Opening usage_log and adding a record
    		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
    		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if

    If send_error_message = vbYes Then
        'dialog here to gather more detail
        error_type = ""
        If testing_run = TRUE Then error_type = "TESTING RESPONSE"

        If trim(MAXIS_case_number) = "" Then
            If trim(MMIS_case_number) <> "" Then MAXIS_case_number = MMIS_case_number
        End If

        Do
            Do
                confirm_err = ""

                case_note_checkbox = unchecked
                stat_update_checkbox = unchecked
                date_checkbox = unchecked
                math_checkbox = unchecked
                tikl_checkbox = unchecked
                memo_wcom_checkbox = unchecked
                document_checkbox = unchecked
                missing_spot_checkbox = unchecked

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
                  Text 60, 35, 55, 10, MAXIS_case_number
                  ComboBox 220, 30, 175, 45, error_type+chr(9)+"BUG - something happened that was wrong"+chr(9)+"ENHANCEMENT - something could be done better"+chr(9)+"TYPO - grammatical/spelling type errors"+chr(9)+"DAIL - add support for this DAIL message.", error_type
                  EditBox 65, 50, 330, 15, error_detail
                  CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
                  CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
                  CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
                  CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
                  CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
                  CheckBox 95, 115, 70, 10, "MEMO or WCOM", memo_wcom_checkbox
                  CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
                  CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
                  EditBox 60, 140, 165, 15, worker_signature
                  ButtonGroup ButtonPressed
                    OkButton 290, 140, 50, 15
                    CancelButton 345, 140, 50, 15
                  Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
                  Text 5, 35, 50, 10, "Case Number:"
                  Text 125, 35, 95, 10, "What type of error occured?"
                  Text 5, 55, 60, 10, "Explain in detail:"
                  GroupBox 10, 75, 380, 60, "Common areas of issue"
                  Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
                  Text 10, 145, 50, 10, "Worker Name:"
                  Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
                EndDialog

                Dialog Dialog1

                If ButtonPressed = 0 Then
                    cancel_confirm_msg = MsgBox("An Error Report will NOT be sent as you pressed 'Cancel'." & vbNewLine & vbNewLine & "Is this what you would like to do?", vbQuestion + vbYesNo, "Confirm Cancel")
                    If cancel_confirm_msg = vbYes Then confirm_err = ""
                    If cancel_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If

                If ButtonPressed = -1 Then
                    full_text = "Error occurred on " & date & " at " & time
                    full_text = full_text & vbCr & "Error type - " & error_type
                    full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
                    full_text = full_text & vbCr & "Information: " & error_detail
                    If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

                    If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
                    If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
                    If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
                    If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
                    If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
                    If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
                    If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
                    If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

                    full_text = full_text & vbCr & "Closing message: " & closing_message
                    full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

                    send_confirm_msg = MsgBox("** This is what will be sent as an email to the BlueZone Script team:" & vbNewLine & vbNewLine & full_text & vbNewLine & vbNewLine & "*** Is this what you want to send? ***", vbQuestion + vbYesNo, "Confirm Error Report")

                    If send_confirm_msg = vbYes Then confirm_err = ""
                    If send_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If
            Loop until confirm_err = ""
            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
        LOOP UNTIL are_we_passworded_out = false
        'sent email here
        full_text = ""
        If ButtonPressed = -1 Then
            bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
            subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"

            full_text = "Error occurred on " & date & " at " & time
            full_text = full_text & vbCr & "Error type - " & error_type
            full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
            full_text = full_text & vbCr & "Information: " & error_detail
            If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

            If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
            If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
            If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
            If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
            If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
            If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
            If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
            If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

            full_text = full_text & vbCr & "Closing message: " & closing_message
            full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

            If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown

			attachment_here = ""
			If name_of_script = "NOTES - CAF.vbs" Then
				local_CAF_save_work_path = user_myDocs_folder & "caf-variables-" & MAXIS_case_number & "-info.txt"
				With objFSO
					If .FileExists(local_CAF_save_work_path) = True then
						attachment_here = local_CAF_save_work_path
					End if
				End With
			End If
			If name_of_script = "ACTIONS - INTERVIEW.vbs" Then
				local_interview_save_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
				With objFSO
					If .FileExists(local_interview_save_work_path) = True then
						attachment_here = local_interview_save_work_path
					End if
				End With
			End If

            Call create_outlook_email(bzt_email, "", subject_of_email, full_text, attachment_here, true)

            MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
        Else
            MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
        End If
    End If
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function

function select_testing_file(selection_type, the_selection, file_path, file_branch, force_error_reporting, allow_option)
'--- Divert the script to a testing file if run by a tester
'~~~~~ allow_option: Boolean to select if the tester can choose the testing version or not
'~~~~~ selection_type: this will indicate how the testers are being selected - valid options (ALL, GROUP, REGION, POPULATION, SCRIPT)
'~~~~~ the_selection: For the selection_type selected, indicate WHICH of these options you want selected
'~~~~~ file_path: where the testing file is located
'~~~~~ file_branch: which branch the file is in
'~~~~~ force_error_reporting: should the in-script error reporting automatically happen
'===== Keywords: MAXIS, PRISM, production, clear
	testing_script_name = ""
	file_path = replace(file_path, "\", "/")
	file_name_array = split(file_path, "/")
	testing_script_name = file_name_array(1)
	testing_script_name = replace(testing_script_name, ".vbs", "")
	testing_script_name = replace(testing_script_name, "-", " ")
	testing_script_name = UCase(file_name_array(0)) & " - " & UCase(testing_script_name)

    run_testing_file = FALSE
    If Instr(the_selection, ",") <> 0 Then
        selection_array = split(the_selection, ",")
    Else
        selection_array = array(the_selection)
    End If
    selection_type = UCase(selection_type)
    For each tester in tester_array
        If user_ID_for_validation = tester.tester_id_number Then
            Select Case selection_type

                Case "ALL"
                    run_testing_file = TRUE
                Case "GROUP" ' ADD OPTION FOR the_selection to be an array'
                    For each group in tester.tester_groups
                        For each selection in selection_array
                            selection = trim(selection)
                            If UCase(selection) = UCase(group) Then run_testing_file = TRUE
                            selected_group = group
                        Next
                    Next
                    selected_group = selection
                Case "REGION"
                    For each selection in selection_array
                        selection = trim(selection)
                        If UCase(selection) = UCase(tester.tester_region) Then run_testing_file = TRUE
                    Next
                Case "PROGRAM"
					For each prog in tester.tester_programs
						For each selection in selection_array
							selection = trim(selection)
							If UCase(selection) = UCase(prog) Then run_testing_file = TRUE
							selected_prog = prog
						Next
					Next
				Case "POPULATION"
                    For each selection in selection_array
                        selection = trim(selection)
                        If UCase(selection) = UCase(tester.tester_population) Then run_testing_file = TRUE
                    Next
                Case "SCRIPT"
                    For each each_script in tester.tester_scripts
                        If name_of_script = each_script Then run_testing_file = TRUE
                    Next
                Case Else
                    body_text = "The call of the function select_testing_file is using an invalid selection_type."
                    body_text = body_text & vbCr & "On script - " & testing_script_name & "."
                    body_text = body_text & vbCr & "The selection type of - " & selection_type & " was entered into the function call"
                    body_text = body_text & vbCr & "The only valid options are: ALL, SCRIPT, GROUP, PROGRAM, POPULATION, or REGION"
                    body_text = body_text & vbCr & "Review the script file particularly the call for the function select_testing_file."
                    Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", "FUNCTION ERROR - select_testing_file for " & testing_script_name, body_text, "", TRUE)
            End Select

            If tester.tester_population = "BZ" Then
                allow_option = TRUE
                run_testing_file = TRUE
                selection_type = "SCRIPTWRITER"
            End If

            If run_testing_file = TRUE and allow_option = TRUE Then
                continue_with_testing_file = MsgBox("You have been selected to test this script - " & testing_script_name & "." & vbNewLine & vbNewLine & "At this time you can select if you would like to run the testing file or the original file." & vbNewLine & vbNewLine & "** Would you like to test this script now?", vbQuestion + vbYesNo, "Use Testing File")
                If continue_with_testing_file = vbNo Then run_testing_file = FALSE
            End If

            If run_testing_file = TRUE Then
                tester.display_testing_message selection_type, the_selection, force_error_reporting
                ' Call tester.display_testing_message(selection_type, the_selection, force_error_reporting)
                If force_error_reporting = TRUE Then testing_run = TRUE
                If run_locally = true then
                    testing_script_url = "C:\MAXIS-scripts\" & file_path
					If file_branch <> "master" Then
						run_locally =  False
						testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/" & file_branch & "/" & file_path
					End If
                Else
                    testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/" & file_branch & "/" & file_path

                End If
                Call run_from_GitHub(testing_script_url)
            End If

        End If
    Next
end function

function sort_dates(dates_array)
'--- Takes an array of dates and reorders them to be chronological.
'~~~~~ dates_array: an array of dates only
'===== Keywords: MAXIS, date, order, list, array
    dim ordered_dates ()				'declaring a private array
    redim ordered_dates(0)				'setting the array with 1 parameter
    original_array_items_used = "~"		'creating a string
    days =  0							'starting a count
	'This is the BIG loop - with each loop through here, we add a new date to the ordered dates array
    do
        prev_date = ""					'blanking out the varibale 'prev_date' for each loop
        original_array_index = 0		'this will tell us WHERE in the original array the date is - which position - so we know to add only new ones.
        for each thing in dates_array	'looking at each item in the original array
            check_this_date = TRUE												'default to check the date
            new_array_index = 0													'startomg amptjer count and setting it to 0
            For each known_date in ordered_dates								'looking at the ones we already ordered
                if known_date = thing Then check_this_date = FALSE				'if it is already in the order, we don't need to review it again
                new_array_index = new_array_index + 1
            next
            if check_this_date = TRUE Then										'if this is a new date, review it - we are trying to find the EARLIEST date that has not been added
                if prev_date = "" Then											'if this is the first date we review
                    prev_date = thing											'set this as the date to compare
                    index_used = original_array_index							'setting it's place
                Else
                    if DateDiff("d", prev_date, thing) < 0 then					'If this date is before the last date, it is now the earliest date
                        prev_date = thing										'saving the date
                        index_used = original_array_index						'saving the position in the array
                    end if
                end if
            end if
            original_array_index = original_array_index + 1						'counting each position on each loop through the dates_array
        next
        if prev_date <> "" Then													'if the function found a NEW 'earliest date' we have to add it to the order
            redim preserve ordered_dates(days)									'resize the array of ordered dates
            ordered_dates(days) = prev_date										'putting this new earliest date into that newly added position
            original_array_items_used = original_array_items_used & index_used & "~"		'saving a string to know which positions in the original array have been used
            days = days + 1														'incrementing up for the next loop to make the ordered dates array bigger next time around
        end if
		'now we have to check for duplicate dates in the original array
        counter = 0																'setting a new incrementer to determine the position in the original array again
        For each thing in dates_array											'looking at each item in turn in the original array
            If InStr(original_array_items_used, "~" & counter & "~") = 0 Then	'If the array position has NOT been added to the string of all the array positions we need to check:
                For each new_date_thing in ordered_dates						'each item in the ordered dates array to see:
                    If thing = new_date_thing Then								'if the date of the original array matches the one we JUST added to the ordered dates (but remember is NOT in the position we have already added)'
                        original_array_items_used = original_array_items_used & counter & "~"	'saving that this date IS in the ordered dates array
                        days = days + 1											'resizing the number of items in the ordered dates array up
                    End If
                Next
            End If
            counter = counter + 1												'incrementing the position of the original array for comparing
        Next
    loop until days > UBOUND(dates_array)			'Once we have gone higher that the number of things in the original array - we are done

    dates_array = ordered_dates						'replacing the original array with the one with dates ordered
end function

function split_phone_number_into_parts(phone_variable, phone_left, phone_mid, phone_right)
'This function is to take the information provided as a phone number and split it up into the 3 parts
'--- This function is to take the information provided as a phone number and split it up into the 3 parts
'~~~~~ phone_variable: input the phone number here
'~~~~~ phone_left: string - first 3 digits of a 10 digit phone number - area code
'~~~~~ phone_mid: string - second 3 digits of a 10 digit phone number - the middle part
'~~~~~ phone_right: string - last 4 digits of a 10 digit phone number
'===== Keywords: MAXIS, case note, navigate, edit
    phone_variable = trim(phone_variable)
    If phone_variable <> "" Then
        phone_variable = replace(phone_variable, "(", "")						'formatting the phone variable to get rid of symbols and spaces
        phone_variable = replace(phone_variable, ")", "")
        phone_variable = replace(phone_variable, "-", "")
        phone_variable = replace(phone_variable, " ", "")
        phone_variable = trim(phone_variable)
        phone_left = left(phone_variable, 3)									'reading the certain sections of the variable for each part.
        phone_mid = mid(phone_variable, 4, 3)
        phone_right = right(phone_variable, 4)
        phone_variable = "(" & phone_left & ")" & phone_mid & "-" & phone_right
    End If
end function

function start_a_blank_CASE_NOTE()
'--- This function navigates user to a blank case note, presses PF9, and checks to make sure you're in edit mode (keeping you from writing all of the case note on an inquiry screen).
'===== Keywords: MAXIS, case note, navigate, edit
	call navigate_to_MAXIS_screen("CASE", "NOTE")
	DO
		PF9
		EMReadScreen case_note_check, 17, 2, 33
		EMReadScreen mode_check, 1, 20, 09
		If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then
            'msgbox "The script can't open a case note. Reasons may include:" & vbnewline & vbnewline & "* You may be in inquiry" & vbnewline & "* You may not have authorization to case note this case (e.g.: out-of-county case)" & vbnewline & vbnewline & "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
            BeginDialog Inquiry_Dialog, 0, 0, 241, 115, "CASE NOTE Cannot be Started"
              ButtonGroup ButtonPressed
                OkButton 185, 95, 50, 15
              Text 10, 10, 190, 10, "The script can't open a case note. Reasons may include:"
              Text 20, 25, 80, 10, "* You may be in inquiry"
              Text 20, 40, 185, 20, "* You may not have authorization to case note this case (e.g.: out-of-county case)"
              Text 5, 70, 225, 20, "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
            EndDialog
            Do
                Dialog Inquiry_Dialog
            Loop until ButtonPressed = -1
        End If
	Loop until (mode_check = "A" or mode_check = "E")
end function

function start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, end_script)
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'~~~~~ memo_opened: BOOLEAN to indicate if the MEMO started and is ready to write to
'~~~~~ search_for_arep_and_swkr: BOOLEAN to indicate if you want the function to search for if forms_to_arep and forms_to_swkr should be read from STAT
'~~~~~ forms_to_arep: Enter as 'Y' for this if the forms should go to the AREP if one is on the case
'~~~~~ forms_to_swkr: Enter as 'Y' for this if the forms should go to the SWKR if one is on the case
'~~~~~ send_to_other: Enter a 'Y' for this if we should send a MEMO to an 'Other' address - if this is 'Y' - the next parameters should be checked beforehand.
'~~~~~ other_name: Aderessee Name for OTHER
'~~~~~ other_street: Addressee Street Information
'~~~~~ other_city: Addressee City Information
'~~~~~ other_state: Addressee State Information - should be the State 2 letter code
'~~~~~ other_zip: Addressee Zip Information
'~~~~~ end_script: BOOLEAN to indicate if the function should end the script
'===== Keywords: MAXIS, notice, navigate, edit
	memo_opened = False
	If search_for_arep_and_swkr = True Then						'If the script has asked the function to check for AREP - some scripts already know
		call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
		EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
		call navigate_to_MAXIS_screen("STAT", "SWKR")         	'Navigates to STAT/SWKR to check and see if forms go to the SWKR
		EMReadscreen forms_to_swkr, 1, 15, 63                	'Reads for the "Forms to SWKR?" Y/N response on the panel.
	End If
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	start_memo_attempt = 1										'We are counting our attempts - this way it will try a few times - but won't get stuck in a loop
	Do
		PF5														'Creates a new MEMO.
		EMReadScreen case_stuck, 50, 24, 11						'Reading if the case is locked
		case_stuck = trim(case_stuck)
		case_stuck = replace(case_stuck, " ", "")
		start_memo_attempt = start_memo_attempt + 1
	Loop Until InStr(case_stuck, "LOCKED") = 0 OR start_memo_attempt = 50

	EMReadScreen memo_display_check, 12, 2, 33					'Once we think we have started a MEMO - check to see if we are in 'MEMO DISPLAY' - which is the basic SPEC/MEMO start place
	If memo_display_check = "Memo Display" then					'If we are still at MEMO DISPLAY' we did NOT start a MEMO
		memo_opened = False										'Setting this output to False so the Script knows we failed '
	Else
		row = 4                             					'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         					'Finding the place the AREP indicator is
		IF row > 4 THEN arep_row = row                      	'If it isn't 4, that means it was found. Logs the row it found the ALTREP string as arep_row

		row = 4                             					'Defining row and col for the search feature.
		col = 1
		EMSearch "SOCWKR", row, col         					'Finding the place the SWKR indicator is
		IF row > 4 THEN swkr_row = row                     		'If it isn't 4, that means it was found. Logs the row it found the SOCWKR string as swkr_row

		row = 4                             					'Defining row and col for the search feature.
		col = 1
		EMSearch "OTHER", row, col         						'Finding the place the OTHER Address indicator is
		IF row > 4 THEN other_row = row                     	'If it isn't 4, that means it was found. Logs the row it found the SOCWKR string as swkr_row

		EMWriteScreen "X", 5, 12                                        					'Initiates new memo to client
		IF forms_to_arep = "Y" AND arep_row <> "" THEN EMWriteScreen "X", arep_row, 12     	'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" AND swkr_row <> "" THEN EMWriteScreen "X", swkr_row, 12     	'If forms_to_swkr was "Y" (see above) it puts an X on the row SOCWKR was found.
		If send_to_other = "Y" AND other_row <> "" Then EMWriteScreen "X", other_row, 12	'If send_to_other was "Y" (see above) it puts an X on the row OTHER was found.

		transmit                                              	'Transmits to start the memo writing process
		If send_to_other = "Y" Then								'If we are sending to the 'OTHER' Address, the Address Information needs to be entered.
			other_street = trim(other_street)					'formatting

			EMWriteScreen other_name, 13, 24					'Entering the NAME
			If len(other_street) < 25 Then						'If the streeet information fits on one line - enters it here
				EMWriteScreen other_street, 14, 24
			Else
				other_street_array = split(other_street, " ")	'If the street information is too long, we are going to create an array of all the words and enter it word by word.
				col = 24
				row = 14
				for each word in other_street_array
					If col + len(word) + 1 > 47 Then			'If the word will run over the line, we go to the next line
						row = row + 1
						col = 24
						If row = 16 then Exit for				'If we move to a thrid line - the street information will just cut off
					End If
					If col <> 24 Then word = " " & word			'Adding a space before the word if we are not at the first column
					EMWriteScreen word, row, col				'Entering the word in the correct place on the correct street line in MAXIS
					col = col + len(word) + 1					'moving to the next space for the next word
				next
			End If
			EMWriteScreen other_city, 16, 24					'writing in the city
			EMWriteScreen other_state, 17, 24					'writing in the state
			EMWriteScreen other_zip, 17, 32						'writing in the zip code

			transmit											'saving the OTHER address
			EMReadScreen post_office_warning, 7, 3, 6			'Reading if MAXIS indicates this is not a 'VALID ADDRESS' -- we just transmit past this
			If UCASE(post_office_warning) = "WARNING" Then transmit
		End If
		EMReadScreen memo_input_screen, 17, 2, 37				'Checking to see if we made it to the 'MEMO INPUT SCREEN' as there are sometimes warning messages
		If memo_input_screen <> "Memo Input Screen" Then transmit 'moving past any warning message

		EMReadScreen memo_input_screen, 17, 2, 37				'Checking avain to ensure the memo was opened'
		If memo_input_screen = "Memo Input Screen" Then memo_opened = True		'setting the output to be sure the script knows it is ready to write a MEMO
	End If
	If memo_opened = False AND end_script = True Then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
end function

function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
'--- This function checks a fooder month and year variable provided in a dialog loop and ensures that is numeric and 2 digits long.
'~~~~~ footer_month: This is whatever variable you are using for the footer month
'~~~~~ footer_year: This is whatever variable you are using for the footer year
'~~~~~ err_msg_var: This is the variable you are using for the error message handling in the dialog loop
'~~~~~ bullet_char: This is the character(s) that you are using to identify each line in the error message
'===== Keywords: MAXIS, dialog, footer month
    If IsNumeric(footer_month) = FALSE Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be a number, review and reenter the footer month information."
    Else
        footer_month = footer_month * 1
        If footer_month > 12 OR footer_month < 1 Then err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be between 1 and 12. Review and reenter the footer month information."
        footer_month = right("00" & footer_month, 2)
    End If

    If len(footer_year) < 2 Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be at least 2 characters long, review and reenter the footer year information."
    Else
        If IsNumeric(footer_year) = FALSE Then
            err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be a number, review and reenter the footer year information."
        Else
            footer_year = right("00" & footer_year, 2)
        End If
    End If
end function

function validate_MAXIS_case_number(err_msg_variable, list_delimiter)
'--- This function checks the MAXIS_case_number variable to ensure it is present and valid
'~~~~~ err_msg_variable: the variable used in error handling within a dialog's do - loop
'~~~~~ list_delimiter: a single character to put in front of any error message to add to the err_msg_variable
'===== Keywords: MAXIS, PRISM, MMIS, dialogs
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" Then
        err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " A case number is required, enter a case number."
    Else
        If IsNumeric(MAXIS_case_number) = FALSE Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " The case number appears invalid, review the case number and update."
        If len(MAXIS_case_number) > 7 Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " The case number entered is too long, review the case number and update."
    End If
end function

function view_poli_temp(temp_one, temp_two, temp_three, temp_four)
'--- This function enters a specific POLI TEMP reference and navigates to it.
'~~~~~ temp one:    1st portion of TE reference. Do not add period. If blank, leave blank string ("").
'~~~~~ temp two:    2nd portion of TE reference. Do not add period. If blank, leave blank string ("").
'~~~~~ temp three:  3rd portion of TE reference. Do not add period. If blank, leave blank string ("").
'~~~~~ temp four:   4th portion of TE reference. Do not add period. If blank, leave blank string ("").
'===== Keywords: MAXIS, dialogs, procedure, POLI TEMP
	Call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
	EMWriteScreen "TEMP", 5, 40     'Writes TEMP

	'Writes the panel_title selection
	Call write_value_and_transmit("TABLE", 21, 71)

    'updates the length of POLI TEMP reference to minimum of 2
	If temp_one <> "" Then temp_one = right("00" & temp_one, 2)
	If len(temp_two) = 1 Then temp_two = right("00" & temp_two, 2)
	If len(temp_three) = 1 Then temp_three = right("00" & temp_three, 2)
	If len(temp_four) = 1 Then temp_four = right("00" & temp_four, 2)

    'creating the temp reference including TE at begining and periods as delimeter
	total_code = "TE" & temp_one & "." & temp_two
	If temp_three <> "" Then total_code = total_code & "." & temp_three
	If temp_four <> "" Then total_code = total_code & "." & temp_four

    'Enters the full TE reference and transmit to complete navigation
	Call write_value_and_transmit(total_code, 3, 21) '
    Call write_value_and_transmit("X", 6, 4)
end function

function word_doc_open(doc_location, objWord, objDoc)
'--- This function opens a specific word document.
'~~~~~ doc_location: location of word document
'~~~~~ ObjWord: leave as 'ObjWord'
'~~~~~ objDoc: leave as 'objDoc'
'===== Keywords: MAXIS, PRISM, MMIS, Word
	'Opens Word object
	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True		'We want to see it

	'Opens the specific Word doc
	set objDoc = objWord.Documents.Add(doc_location)
end function

function word_doc_update_field(field_name, variable_for_field, objDoc)
'--- This function updates specific fields on a word document
'~~~~~ field_name: name of the field to update
'~~~~~ variable_for_field: information to be updated
'~~~~~ objDoc: leave as 'objDoc'
'===== Keywords: MAXIS, PRISM, MMIS, Word
	objDoc.FormFields(field_name).Result = variable_for_field	'Simply enters the Word document field based on these three criteria
end function

Function write_editbox_in_person_note(x, y) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array
    EMGetCursor row, col
    If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col
    If (row < 18 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(5)
    If (row = 5 and col = 3) then EMSendKey space(5)
    EMSendKey x & " "
    If right(x, 1) = ";" then
      EMSendKey "<backspace>" & "<backspace>"
      EMGetCursor row, col
      If row = 18 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(5)
      Else
        EMSendKey "<newline>" & space(5)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

function write_bullet_and_variable_in_CASE_NOTE(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CASE notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CASE note
'===== Keywords: MAXIS, bullet, CASE note
	If trim(variable) <> "" then
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
			If character_test <> "" or noting_row >= 18 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
				End if
			End if
		Loop until character_test = ""

		'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
		If len(bullet) >= 14 then
			indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
		Else
			indent_length = len(bullet) + 4 'It's four more for the reason explained above.
		End if

		'Writes the bullet
		EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

		'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + (len(bullet) + 4)

		'Splits the contents of the variable into an array of words
        variable = trim(variable)
        If right(variable, 1) = ";" Then variable = left(variable, len(variable) - 1)
        If left(variable, 1) = ";" Then variable = right(variable, len(variable) - 1)
		variable_array = split(variable, " ")

		For each word in variable_array
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

                'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
                EMReadScreen end_of_case_note_check, 1, 24, 2
                If end_of_case_note_check = "A" then
                    EMSendKey "<PF3>"												'PF3s
                    EMWaitReady 0, 0
                    EMSendKey "<PF9>"												'PF9s (opens new note)
                    EMWaitReady 0, 0
                    EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
                    EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
                    noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
                Else
                    noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                End If
			End if

			'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
			If noting_col = 3 then
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
			If right(word, 1) = ";" then
				noting_row = noting_row + 1
				noting_col = 3
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

function write_bullet_and_variable_in_CCOL_note(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CCOL notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CCOL note
'===== Keywords: MAXIS, bullet, CCOL note
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        'msgbox varible & vbcr & "noting_row " & noting_row
        noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
        Do
            EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
            If character_test <> "" or noting_row >= 19 then
                noting_row = noting_row + 1
                'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
                If noting_row >= 19 then
                    PF8
                    'msgbox "sent PF8"
                    EMReadScreen next_page_confirmation, 4, 19, 3
                    'msgbox "next_page_confirmation " & next_page_confirmation
                    IF next_page_confirmation = "More" THEN
                        next_page = TRUE
                        noting_row = 5
                    Else
                        next_page = FALSE
                    End If
                    'msgbox "next_page " & next_page
                Else
                    noting_row = noting_row + 1
                End if
            End if
        Loop until character_test = ""

        'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
        If len(bullet) >= 14 then
            indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
        Else
            indent_length = len(bullet) + 4 'It's four more for the reason explained above.
        End if

        'Writes the bullet
        EMWriteScreen "* " & bullet & ": ", noting_row, noting_col
        'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
        noting_col = noting_col + (len(bullet) + 4)
        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array
            'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
            If len(word) + noting_col > 80 then
                noting_row = noting_row + 1
                noting_col = 3
            End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
            If noting_row >= 19 then
                PF8
                noting_row = 5
                'Msgbox "what's Happening? Noting row: " & noting_row
            End if

            'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
            If noting_col = 3 then
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Writes the word and a space using EMWriteScreen
            EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

            'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
            If right(word, 1) = ";" then
                noting_row = noting_row + 1
                noting_col = 3
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if
            'Increases noting_col the length of the word + 1 (for the space)
            noting_col = noting_col + (len(word) + 1)
        Next
        'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
    	EMSetCursor noting_row + 1, 3
    End if
end function

Function write_bullet_and_variable_in_MMIS_NOTE(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style MMIS notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into MMIS note
'===== Keywords: MMIS, bullet, CASE note
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
        Do
            EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            If character_test <> " " or noting_row >= 20 then
                noting_row = noting_row + 1

                'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
                If noting_row >= 20 then
                    PF11
                    noting_row = 5
                End if
            End if
        Loop until character_test = " "

        'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
        If len(bullet) >= 14 then
            indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
        Else
            indent_length = len(bullet) + 4 'It's four more for the reason explained above.
        End if

        'Writes the bullet
        EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

        'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
        noting_col = noting_col + (len(bullet) + 4)

        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array

            'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
            If len(word) + noting_col > 80 then
                noting_row = noting_row + 1
                noting_col = 8
            End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
            If noting_row >= 20 then
                PF11
                noting_row = 5
            End if

            'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
            If noting_col = 8 then
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Writes the word and a space using EMWriteScreen
            EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

            'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
            If right(word, 1) = ";" then
                noting_row = noting_row + 1
                noting_col = 8
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Increases noting_col the length of the word + 1 (for the space)
            noting_col = noting_col + (len(word) + 1)
        Next

        'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
        EMSetCursor noting_row + 1, 3
    End if
End Function

function write_date(date_variable, date_format_variable, screen_row, screen_col)
'--- This function will write a date in any format desired.
'~~~~~ date_variable: date to write
'~~~~~ date_format_variable: format of date. this should be a string with the correct spaces between month/day/year examples: MM DD YY, MM YY, MM  DD  YYYY
'~~~~~ screen_row: row to write date
'~~~~~ screen_col: column to write date
'===== Keywords: MAXIS, MMIS, PRISM, date, format
    'Figures out the format of the month. If it was "MM", "M", or not present.
    If instr(ucase(date_format_variable), "MM") <> 0 then
        month_format = "MM"
        month_position = instr(ucase(date_format_variable), "MM")
    Elseif instr(ucase(date_format_variable), "M") <> 0 then
        month_format = "M"
        month_position = instr(ucase(date_format_variable), "M")
    Else
        month_format = ""
        month_position = 0
    End if

    'Figures out the format of the day. If it was "DD", "D", or not present.
    If instr(ucase(date_format_variable), "DD") <> 0 then
        day_format = "DD"
        day_position = instr(ucase(date_format_variable), "DD")
    Elseif instr(ucase(date_format_variable), "D") <> 0 then
        day_format = "D"
        day_position = instr(ucase(date_format_variable), "D")
    Else
        day_format = ""
        day_position = 0
    End if

    'Figures out the format of the year. If it was "YYYY", "YY", or not present.
    If instr(ucase(date_format_variable), "YYYY") <> 0 then
        year_format = "YYYY"
        year_position = instr(ucase(date_format_variable), "YYYY")
    Elseif instr(ucase(date_format_variable), "YY") <> 0 then
        year_format = "YY"
        year_position = instr(ucase(date_format_variable), "YY")
    Else
        year_format = ""
        year_position = 0
    End if

    'Formats the month. Separates the month into its own variable and adds a zero if needed.
    var_month = datepart("m", date_variable)
    IF len(var_month) = 1 and month_format = "MM" THEN var_month = "0" & var_month

    'Formats the day. Separates the day into its own variable and adds a zero if needed.
    var_day = datepart("d", date_variable)
    IF len(var_day) = 1 and day_format = "DD" THEN var_day = "0" & var_day

    'Formats the year based on "YY" or "YYYY" formatting.
    If year_format = "YY" then
        var_year = right(datepart("yyyy", date_variable), 2)
    ElseIf year_format = "YYYY" then
        var_year = datepart("yyyy", date_variable)
    END IF

    If month_position <> 0 Then EMWriteScreen var_month, screen_row, screen_col + month_position - 1
    If day_position <> 0 Then EMWriteScreen var_day, screen_row, screen_col + day_position - 1
    If year_position <> 0 Then EMWriteScreen var_year, screen_row, screen_col + year_position - 1
end function

Function write_new_line_in_person_note(x)
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80 + 1 ) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

function write_three_columns_in_CASE_NOTE(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
'--- This function writes variables into three seperate columns into case notes
'~~~~~ col_01_start_point: column where to write the 1st variable
'~~~~~ col_01_variable: name of 1st variable to write
'~~~~~ col_02_start_point: column where to write the 2nd variable
'~~~~~ col_02_variable: name of 2nd variable to write
'~~~~~ col_03_start_point: column where to write the 3rd variable
'~~~~~ col_03_variable: name of 3rd variable to write
'===== Keywords: MAXIS, case note, three columns, format
  EMGetCursor row, col
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMGetCursor row, col
  EMWriteScreen "                                                                              ", row, 3
  EMSetCursor row, col_01_start_point
  EMSendKey col_01_variable
  EMSetCursor row, col_02_start_point
  EMSendKey col_02_variable
  EMSetCursor row, col_03_start_point
  EMSendKey col_03_variable
  EMSendKey "<newline>"
  EMGetCursor row, col
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
end function

function write_value_and_transmit(input_value, row, col)
'--- This function writes a specific value and transmits.
'~~~~~ input_value: information to be entered
'~~~~~ row: row to write the input_value
'~~~~~ col: column to write the input_value
'===== Keywords: MAXIS, PRISM, case note, three columns, format
	EMWriteScreen input_value, row, col
	transmit
end function

function write_variable_in_CASE_NOTE(variable)
'--- This function writes a variable in CASE note
'~~~~~ variable: information to be entered into CASE note from script/edit box
'===== Keywords: MAXIS, CASE note
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
            If character_test <> "" or noting_row >= 18 then

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
                Else
                    noting_row = noting_row + 1
				End if
			End if
		Loop until character_test = ""

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		For each word in variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End if
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

Function write_variable_in_CCOL_note(variable)
    ''--- This function writes a variable in CCOL note
    '~~~~~ variable: information to be entered into CASE note from script/edit box
    '===== Keywords: MAXIS, CASE note
    If trim(variable) <> "" THEN
    	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    	'msgbox varible & vbcr & "noting_row " & noting_row
        noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    	'The following figures out if we need a new page, or if we need a new case note entirely as well.
    	Do
    		EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
    		character_test = trim(character_test)
    		If character_test <> "" or noting_row >= 19 then
                noting_row = noting_row + 1
    		    'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
    			If noting_row >= 19 then
    				PF8
                    'msgbox "sent PF8"
    				EMReadScreen next_page_confirmation, 4, 19, 3
                    'msgbox "next_page_confirmation " & next_page_confirmation
    				IF next_page_confirmation = "More" THEN
    					next_page = TRUE
                        noting_row = 5
    				Else
						next_page = FALSE
    				End If
                    'msgbox "next_page " & next_page
    			Else
    				noting_row = noting_row + 1
    			End if
    		End if
    	Loop until character_test = ""

    	'Splits the contents of the variable into an array of words
    	variable_array = split(variable, " ")

        For each word in variable_array
            'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
            If len(word) + noting_col > 80 then
                noting_row = noting_row + 1
                noting_col = 3
            End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
            If noting_row >= 19 then
                PF8
                noting_row = 5
                'Msgbox "what's Happening? Noting row: " & noting_row
            End if

            'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
            If noting_col = 3 then
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Writes the word and a space using EMWriteScreen
            EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

            'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
            If right(word, 1) = ";" then
                noting_row = noting_row + 1
                noting_col = 3
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if
            'Increases noting_col the length of the word + 1 (for the space)
            noting_col = noting_col + (len(word) + 1)
        Next
        'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
    	EMSetCursor noting_row + 1, 3
    End if
End Function

Function write_variable_in_MMIS_NOTE(variable)
''--- This function writes a variable in MMIS case note
'~~~~~ variable: information to be entered into CASE note from script/edit box
'===== Keywords: MMIS, CASE note
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
			If character_test <> " " or noting_row >= 20 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 20 then
                    PF11
                    noting_row = 5
				End if
			End if
		Loop until character_test = " "

        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array
            word = trim(word)
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 8
			End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 20 then
                PF11
                noting_row = 5
			End if

            'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
    End if
End Function

function write_variable_in_SPEC_MEMO(variable)
'--- This function writes a variable in SPEC/MEMO
'~~~~~ variable: information to be entered into SPEC/MEMO
'===== Keywords: MAXIS, SPEC, MEMO
    EMGetCursor memo_row, memo_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    memo_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    'The following figures out if we need a new page
    Do
        EMReadScreen line_test, 60, memo_row, memo_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
        line_test = trim(line_test)
        'MsgBox line_test
        If line_test <> "" OR memo_row >= 18 Then
            memo_row = memo_row + 1

            'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
            If memo_row >= 18 then
                PF8
                memo_row = 3					'Resets this variable to 3
            End if
        End If

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then script_end_procedure("The WCOM/MEMO area is already full and no additional informaion can be added. This script should be run prior to adding manual wording.")

    Loop until line_test = ""

    'Each word becomes its own member of the array called variable_array.
    variable_array = split(variable, " ")

    For each word in variable_array
        'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
        If len(word) + memo_col > 74 then
            memo_row = memo_row + 1
            memo_col = 15
        End if

        'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
        If memo_row >= 18 then
            PF8
            memo_row = 3					'Resets this variable to 3
        End if

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then
            PF10
            end_msg = "The WCOM/MEMO area is already full and no additional informaion can be added. The wording that was not added and the script ended on is:" & vbNewLine & vbNewLine & variable & vbNewLine & vbNewLine & "**This script should be run prior to adding manual wording.**"
            script_end_procedure(end_msg)
        End If
        'Writes the word and a space using EMWriteScreen
        EMWriteScreen word & " ", memo_row, memo_col

        'Increases memo_col the length of the word + 1 (for the space)
        memo_col = memo_col + (len(word) + 1)
    Next

    'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand).
    EMSetCursor memo_row + 1, 15
end function

function write_variable_in_TIKL(variable)
'--- This function writes a variable in TIKL
'~~~~~ variable: information to be entered into TIKL
'===== Keywords: MAXIS, TIKL
	IF len(variable) <= 60 THEN
		tikl_line_one = variable
	ELSE
		tikl_line_one_len = 61
		tikl_line_one = left(variable, tikl_line_one_len)
		IF right(tikl_line_one, 1) = " " THEN
			whats_left_after_one = right(variable, (len(variable) - tikl_line_one_len))
		ELSE
			DO
				tikl_line_one = left(variable, (tikl_line_one_len - 1))
				IF right(tikl_line_one, 1) <> " " THEN tikl_line_one_len = tikl_line_one_len - 1
			LOOP UNTIL right(tikl_line_one, 1) = " "
			whats_left_after_one = right(variable, (len(variable) - (tikl_line_one_len - 1)))
		END IF
	END IF

	IF (whats_left_after_one <> "" AND len(whats_left_after_one) <= 60) THEN
		tikl_line_two = whats_left_after_one
	ELSEIF (whats_left_after_one <> "" AND len(whats_left_after_one) > 60) THEN
		tikl_line_two_len = 61
		tikl_line_two = left(whats_left_after_one, tikl_line_two_len)
		IF right(tikl_line_two, 1) = " " THEN
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - tikl_line_two_len))
		ELSE
			DO
				tikl_line_two = left(whats_left_after_one, (tikl_line_two_len - 1))
				IF right(tikl_line_two, 1) <> " " THEN tikl_line_two_len = tikl_line_two_len - 1
			LOOP UNTIL right(tikl_line_two, 1) = " "
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - (tikl_line_two_len - 1)))
		END IF
	END IF

	IF (whats_left_after_two <> "" AND len(whats_left_after_two) <= 60) THEN
		tikl_line_three = whats_left_after_two
	ELSEIF (whats_left_after_two <> "" AND len(whats_left_after_two) > 60) THEN
		tikl_line_three_len = 61
		tikl_line_three = right(whats_left_after_two, tikl_line_three_len)
		IF right(tikl_line_three, 1) = " " THEN
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - tikl_line_three_len))
		ELSE
			DO
				tikl_line_three = left(whats_left_after_two, (tikl_line_three_len - 1))
				IF right(tikl_line_three, 1) <> " " THEN tikl_line_three_len = tikl_line_three_len - 1
			LOOP UNTIL right(tikl_line_three, 1) = " "
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - (tikl_line_three_len - 1)))
		END IF
	END IF

	IF (whats_left_after_three <> "" AND len(whats_left_after_three) <= 60) THEN
		tikl_line_four = whats_left_after_three
	ELSEIF (whats_left_after_three <> "" AND len(whats_left_after_three) > 60) THEN
		tikl_line_four_len = 61
		tikl_line_four = left(whats_left_after_three, tikl_line_four_len)
		IF right(tikl_line_four, 1) = " " THEN
			tikl_line_five = right(whats_left_after_three, (len(whats_left_after_three) - tikl_line_four_len))
		ELSE
			DO
				tikl_line_four = left(whats_left_after_three, (tikl_line_four_len - 1))
				IF right(tikl_line_four, 1) <> " " THEN tikl_line_four_len = tikl_line_four_len - 1
			LOOP UNTIL right(tikl_line_four, 1) = " "
			tikl_line_five = right(whats_left_after_three, (tikl_line_four_len - 1))
		END IF
	END IF

	EMWriteScreen tikl_line_one, 9, 3
	IF tikl_line_two <> "" THEN EMWriteScreen tikl_line_two, 10, 3
	IF tikl_line_three <> "" THEN EMWriteScreen tikl_line_three, 11, 3
	IF tikl_line_four <> "" THEN EMWriteScreen tikl_line_four, 12, 3
	IF tikl_line_five <> "" THEN EMWriteScreen tikl_line_five, 13, 3
	transmit
end function

function write_variable_with_indent_in_CASE_NOTE(variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CASE notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CASE note
'===== Keywords: MAXIS, bullet, CASE note
    variable = trim(variable)
	If variable <> "" then
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
			If character_test <> "" or noting_row >= 18 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
				End if
			End if
		Loop until character_test = ""

        indent_length = 5

		'Writes the bullet
        If IsNumeric(left(variable, 1)) = False Then
            EMWriteScreen "  - ", noting_row, noting_col
        Else
            EMWriteScreen "  ", noting_row, noting_col
        End If
		'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + 4

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		For each word in variable_array
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

                'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
                EMReadScreen end_of_case_note_check, 1, 24, 2
                If end_of_case_note_check = "A" then
                    EMSendKey "<PF3>"												'PF3s
                    EMWaitReady 0, 0
                    EMSendKey "<PF9>"												'PF9s (opens new note)
                    EMWaitReady 0, 0
                    EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
                    EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
                    noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
                Else
                    noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                End If
			End if

			'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
			If noting_col = 3 then
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
			If right(word, 1) = ";" then
				noting_row = noting_row + 1
				noting_col = 3
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

'END OF FUNCTIONS LIBRARY========================================================================================================================================================================================
