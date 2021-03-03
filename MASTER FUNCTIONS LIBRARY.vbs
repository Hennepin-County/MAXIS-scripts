'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script "library" contains functions and variables that the other BlueZone scripts use very commonly. The other BlueZone scripts contain a few lines of code that run
'this script and get the functions. This saves time in writing and copy/pasting the same functions in many different places. Only add functions to this script if they've
'been tested in other scripts first. This document is actively used by live scripts, so it needs to be functionally complete at all times.
'
'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============

'CHANGELOG BLOCK ===========================================================================================================
actual_script_name = name_of_script
name_of_script = "Functions Library"
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
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

'Global function to actually RUN'
Call confirm_tester_information

'Time arrays which can be used to fill an editbox with the convert_array_to_droplist_items function
time_array_15_min = array("7:00 AM", "7:15 AM", "7:30 AM", "7:45 AM", "8:00 AM", "8:15 AM", "8:30 AM", "8:45 AM", "9:00 AM", "9:15 AM", "9:30 AM", "9:45 AM", "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "1:00 PM", "1:15 PM", "1:30 PM", "1:45 PM", "2:00 PM", "2:15 PM", "2:30 PM", "2:45 PM", "3:00 PM", "3:15 PM", "3:30 PM", "3:45 PM", "4:00 PM", "4:15 PM", "4:30 PM", "4:45 PM", "5:00 PM", "5:15 PM", "5:30 PM", "5:45 PM", "6:00 PM")
time_array_30_min = array("7:00 AM", "7:30 AM", "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM")

'Array of all the upcoming holidays
HOLIDAYS_ARRAY = Array(#11/11/20#, #11/26/20#, #11/27/20#, #12/25/20#, #1/1/21#, #1/18/21#, #2/15/21#, #5/31/21#, #7/5/21#, #9/6/21#, #11/11/21#, #11/25/21#, #11/26/21#, #12/24/21#, #12/31/21#)

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
IF CM_mo = "01" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #01/21/2021#
ELSEIF CM_mo = "02" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #02/18/2021#
ELSEIF CM_mo = "03" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #03/19/2021#
ELSEIF CM_mo = "04" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #04/20/2021#
ELSEIF CM_mo = "05" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #05/20/2021#
ELSEIF CM_mo = "06" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #06/18/2021#
ELSEIF CM_mo = "07" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #07/21/2021#
ELSEIF CM_mo = "08" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #08/19/2021#
ELSEIF CM_mo = "09" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #09/20/2021#
ELSEIF CM_mo = "10" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #10/21/2021#
ELSEIF CM_mo = "11" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #11/18/2021#
ELSEIF CM_mo = "12" AND CM_yr = "21" THEN
    ten_day_cutoff_date = #12/21/2021#
ELSEIF CM_mo = "12" AND CM_yr = "20" THEN
    ten_day_cutoff_date = #12/21/2020#                                'last month of current year
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
county_list = county_list+chr(9)+"06 - Big Stone"
county_list = county_list+chr(9)+"07 - Blue Earth"
county_list = county_list+chr(9)+"08 - Brown"
county_list = county_list+chr(9)+"09 - Carlton"
county_list = county_list+chr(9)+"10 - Carver"
county_list = county_list+chr(9)+"11 - Cass"
county_list = county_list+chr(9)+"12 - Chippewa"
county_list = county_list+chr(9)+"13 - Chisago"
county_list = county_list+chr(9)+"14 - Clay"
county_list = county_list+chr(9)+"15 - Clearwater"
county_list = county_list+chr(9)+"16 - Cook"
county_list = county_list+chr(9)+"17 - Cottonwood"
county_list = county_list+chr(9)+"18 - Crow Wing"
county_list = county_list+chr(9)+"19 - Dakota"
county_list = county_list+chr(9)+"20 - Dodge"
county_list = county_list+chr(9)+"21 - Douglas"
county_list = county_list+chr(9)+"22 - Faribault"
county_list = county_list+chr(9)+"23 - Fillmore"
county_list = county_list+chr(9)+"24 - Freeborn"
county_list = county_list+chr(9)+"25 - Goodhue"
county_list = county_list+chr(9)+"26 - Grant"
county_list = county_list+chr(9)+"27 - Hennepin"
county_list = county_list+chr(9)+"28 - Houston"
county_list = county_list+chr(9)+"29 - Hubbard"
county_list = county_list+chr(9)+"30 - Isanti"
county_list = county_list+chr(9)+"31 - Itasca"
county_list = county_list+chr(9)+"32 - Jackson"
county_list = county_list+chr(9)+"33 - Kanabec"
county_list = county_list+chr(9)+"34 - Kandiyohi"
county_list = county_list+chr(9)+"35 - Kittson"
county_list = county_list+chr(9)+"36 - Koochiching"
county_list = county_list+chr(9)+"37 - Lac Qui Parle"
county_list = county_list+chr(9)+"38 - Lake"
county_list = county_list+chr(9)+"39 - Lake Of Woods"
county_list = county_list+chr(9)+"40 - Le Sueur"
county_list = county_list+chr(9)+"41 - Lincoln"
county_list = county_list+chr(9)+"42 - Lyon"
county_list = county_list+chr(9)+"43 - Mcleod"
county_list = county_list+chr(9)+"44 - Mahnomen"
county_list = county_list+chr(9)+"45 - Marshall"
county_list = county_list+chr(9)+"46 - Martin"
county_list = county_list+chr(9)+"47 - Meeker"
county_list = county_list+chr(9)+"48 - Mille Lacs"
county_list = county_list+chr(9)+"49 - Morrison"
county_list = county_list+chr(9)+"50 - Mower"
county_list = county_list+chr(9)+"51 - Murray"
county_list = county_list+chr(9)+"52 - Nicollet"
county_list = county_list+chr(9)+"53 - Nobles"
county_list = county_list+chr(9)+"54 - Norman"
county_list = county_list+chr(9)+"55 - Olmsted"
county_list = county_list+chr(9)+"56 - Otter Tail"
county_list = county_list+chr(9)+"57 - Pennington"
county_list = county_list+chr(9)+"58 - Pine"
county_list = county_list+chr(9)+"59 - Pipestone"
county_list = county_list+chr(9)+"60 - Polk"
county_list = county_list+chr(9)+"61 - Pope"
county_list = county_list+chr(9)+"62 - Ramsey"
county_list = county_list+chr(9)+"63 - Red Lake"
county_list = county_list+chr(9)+"64 - Redwood"
county_list = county_list+chr(9)+"65 - Renville"
county_list = county_list+chr(9)+"66 - Rice"
county_list = county_list+chr(9)+"67 - Rock"
county_list = county_list+chr(9)+"68 - Roseau"
county_list = county_list+chr(9)+"69 - St. Louis"
county_list = county_list+chr(9)+"70 - Scott"
county_list = county_list+chr(9)+"71 - Sherburne"
county_list = county_list+chr(9)+"72 - Sibley"
county_list = county_list+chr(9)+"73 - Stearns"
county_list = county_list+chr(9)+"74 - Steele"
county_list = county_list+chr(9)+"75 - Stevens"
county_list = county_list+chr(9)+"76 - Swift"
county_list = county_list+chr(9)+"77 - Todd"
county_list = county_list+chr(9)+"78 - Traverse"
county_list = county_list+chr(9)+"79 - Wabasha"
county_list = county_list+chr(9)+"80 - Wadena"
county_list = county_list+chr(9)+"81 - Waseca"
county_list = county_list+chr(9)+"82 - Washington"
county_list = county_list+chr(9)+"83 - Watonwan"
county_list = county_list+chr(9)+"84 - Wilkin"
county_list = county_list+chr(9)+"85 - Winona"
county_list = county_list+chr(9)+"86 - Wright"
county_list = county_list+chr(9)+"87 - Yellow Medicine"
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

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

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

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName

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

'=========================================================================================================================================================================== FUNCTIONS RELATED TO GLOBAL CONSTANTS
FUNCTION income_test_SNAP_categorically_elig(household_size, income_limit) '165% FPG
	'See Combined Manual 0019.06
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "20") OR (MAXIS_footer_year = "21") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1755										'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 2371										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 2987
		IF household_size = 4 THEN income_limit = 3603
		IF household_size = 5 THEN income_limit = 4219
		IF household_size = 6 THEN income_limit = 4835
		IF household_size = 7 THEN income_limit = 5451
		IF household_size = 8 THEN income_limit = 6067
		IF household_size > 8 THEN income_limit = 6067 + (616 * (household_size- 8))
	ELSE
        '2019 Amounts
        IF household_size = 1 THEN income_limit = 1718
		IF household_size = 2 THEN income_limit = 2236
		IF household_size = 3 THEN income_limit = 2933
		IF household_size = 4 THEN income_limit = 3541
		IF household_size = 5 THEN income_limit = 4149
		IF household_size = 6 THEN income_limit = 4757
		IF household_size = 7 THEN income_limit = 5364
		IF household_size = 8 THEN income_limit = 5972
		IF household_size > 8 THEN income_limit = 5972 + (608 * (household_size- 8))
	END IF

	valid_through_date = #10/01/2021#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

FUNCTION income_test_SNAP_gross(household_size, income_limit) '130% FPG
	'See Combined Manual 0019.06
	'Also used for sponsor income
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "20") OR (MAXIS_footer_year = "21") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1383								'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 1868										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 2353
		IF household_size = 4 THEN income_limit = 2839
		IF household_size = 5 THEN income_limit = 3324
		IF household_size = 6 THEN income_limit = 3809
		IF household_size = 7 THEN income_limit = 4295
		IF household_size = 8 THEN income_limit = 4705
		IF household_size > 8 THEN income_limit = 4705 + (486 * (household_size- 8))
	ELSE
        '2019 Amounts
        IF household_size = 1 THEN income_limit = 1354
		IF household_size = 2 THEN income_limit = 1832
		IF household_size = 3 THEN income_limit = 2311
		IF household_size = 4 THEN income_limit = 2790
		IF household_size = 5 THEN income_limit = 3269
		IF household_size = 6 THEN income_limit = 3748
		IF household_size = 7 THEN income_limit = 4227
		IF household_size = 8 THEN income_limit = 4705
		IF household_size > 8 THEN income_limit = 4705 + (479 * (household_size- 8))
	END IF

	valid_through_date = #10/01/2021#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

FUNCTION income_test_SNAP_net(household_size, income_limit)
	'See Combined Manual 0020.12 - Net income standard 100% FPG
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "20") OR (MAXIS_footer_year = "21") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1064										'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 1437										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 1810
		IF household_size = 4 THEN income_limit = 2184
		IF household_size = 5 THEN income_limit = 2557
		IF household_size = 6 THEN income_limit = 2930
		IF household_size = 7 THEN income_limit = 3304
		IF household_size = 8 THEN income_limit = 3677
		IF household_size > 8 THEN income_limit = 3677 + (374 * (household_size- 8))
	ELSE
        '2019 Amounts
        IF household_size = 1 THEN income_limit = 1041
        IF household_size = 2 THEN income_limit = 1410
        IF household_size = 3 THEN income_limit = 1778
        IF household_size = 4 THEN income_limit = 2146
        IF household_size = 5 THEN income_limit = 2515
        IF household_size = 6 THEN income_limit = 2883
        IF household_size = 7 THEN income_limit = 3251
        IF household_size = 8 THEN income_limit = 3620
        IF household_size > 8 THEN income_limit = 3620 + (369 * (household_size- 8))
	END IF

	valid_through_date = #10/01/2021#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

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
		EMWriteScreen "x", 7, 26
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
			EMWriteScreen "x", 17, 29
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
		EMWriteScreen "x", 6, 26
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
    EMWriteScreen "x", 19, 38
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
		EMWriteScreen "x", 19, 71
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
		EMWriteScreen "x", 19, 54
	ELSE								'this is the new position
		EMWriteScreen "x", 19, 48
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
      EMWriteScreen "x", 10, 26
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
      EMWriteScreen "x", 6, 56
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
  call navigate_to_MAXIS_screen("stat", left(panel_read_from, 4))

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
  				EMWriteScreen "x", 12, 39
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
	EmWriteScreen "x", 13, 57
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
	If read_abawd_status = "13" THEN  abawd_status = "ABAWD = ABAWD banked months."
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

function change_date_to_soonest_working_day(date_to_change)
'--- This function will change a date that is on a weekend or Hennepin County holiday to the next working date before the date provided, the date will remain the same if it is not a holiday or weekend.
'~~~~~ date_to_change: variable in the form of a date - this will change once the function is called
'===== Keywords: MAXIS, date, change
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

		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8

		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")

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
				warning_box = MsgBox("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
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
	script_list_URL = t_drive & "\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script

    Set objNet = CreateObject("WScript.NetWork")            'Getting the script user's windows ID
    windows_user_ID = objNet.UserName
    user_ID_for_validation = ucase(windows_user_ID)

    'Now the script will look to see if this user is a tester that needs their information confirmed.
    Leave_Confirmation = FALSE                              'This will allow the user to cancel the update if they desire
    For each tester in tester_array                         'looping through all of the testers
        If user_ID_for_validation = tester.tester_id_number Then            'If the person who is running the script is a tester
            If tester.tester_confirmed = FALSE Then                         'If the information is not saved as confirmed.
                'Script user is asked if they can confirm their information now.
                confirm_testing_now = MsgBox("Hello " & tester.tester_first_name & "! Thank you for agreeing to test BlueZone Scripts. You are an invaluable part of our process and the development of new tools and scripts." & vbNewLine & vbNewLine & "To be sure the testing functionality works correctly, we need to be sure we have the correct information about you. We need you to confrim a few details, this will take less than 5 minutes." & vbNewLine & vbNewLine & "Do you have time to confirm your information now?", vbQuestion + vbYesNo, "Confirm Tester Deail")

                If confirm_testing_now = vbYes Then     'If they select 'Yes' the script will run the dialogs to confirm information
                    show_initial_dialog = TRUE          'this is set to show the initial dialog because there are 2 dialogs that loop together and once we pass the first, we don't want to see it again
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
                                    BeginDialog the_dialog, 0, 0, 370, 215, "Detailed Tester Information"      'first dialog just lists the properties we already know
                                      ButtonGroup ButtonPressed
                                        OkButton 310, 195, 50, 15
                                      Text 60, 15, 40, 10, "First Name:"
                                      Text 60, 35, 40, 10, "Last Name:"
                                      Text 50, 55, 50, 10, "Email Address:"
                                      Text 10, 75, 90, 10, "Hennepin County ID (WF#):"
                                      Text 40, 95, 60, 10, "MAXIS X-Number:"
                                      Text 40, 115, 60, 10, "Supervisor Name:"
                                      Text 40, 135, 60, 10, "Population/Team:"
                                      Text 70, 155, 25, 10, "Region:"
                                      Text 70, 175, 25, 10, "Groups:"
                                      Text 110, 15, 105, 10, tester.tester_first_name
                                      Text 110, 35, 105, 10, tester.tester_last_name
                                      Text 110, 55, 140, 10, tester.tester_email
                                      Text 110, 75, 60, 10, tester.tester_id_number
                                      Text 110, 95, 60, 10, tester.tester_x_number
                                      Text 110, 115, 150, 10, tester.tester_supervisor_name
                                      Text 110, 135, 60, 10, tester.tester_population
                                      Text 110, 155, 60, 10, tester.tester_region
                                      Text 110, 175, 150, 10, Join(tester.tester_groups, ",")
                                      DropListBox 280, 10, 80, 45, "Correct"+chr(9)+"Incorrect - Change", first_name_action
                                      DropListBox 280, 30, 80, 45, "Correct"+chr(9)+"Incorrect - Change", last_name_action
                                      DropListBox 280, 50, 80, 45, "Correct"+chr(9)+"Incorrect - Change", email_action
                                      DropListBox 280, 70, 80, 45, "Correct"+chr(9)+"Incorrect - Change", id_number_action
                                      DropListBox 280, 90, 80, 45, "Correct"+chr(9)+"Incorrect - Change", x_number_action
                                      DropListBox 280, 110, 80, 45, "Correct"+chr(9)+"Incorrect - Change", supervisor_action
                                      DropListBox 280, 130, 80, 45, "Correct"+chr(9)+"Incorrect - Change", population_action
                                      DropListBox 280, 150, 80, 45, "Correct"+chr(9)+"Incorrect - Change", region_action
                                      DropListBox 280, 170, 80, 45, "Correct"+chr(9)+"Incorrect - Change", groups_action
                                      Text 10, 195, 130, 15, "Please reach out to the BlueZone Script team with any questions."
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
                                BeginDialog the_dialog, 0, 0, 265, 215, "Detailed Tester Information"
                                  ButtonGroup ButtonPressed
                                    OkButton 210, 195, 50, 15
                                  Text 60, 15, 40, 10, "First Name:"
                                  Text 60, 35, 40, 10, "Last Name:"
                                  Text 50, 55, 50, 10, "Email Address:"
                                  Text 10, 75, 90, 10, "Hennepin County ID (WF#):"
                                  Text 40, 95, 60, 10, "MAXIS X-Number:"
                                  Text 40, 115, 60, 10, "Supervisor Name:"
                                  Text 40, 135, 60, 10, "Population/Team:"
                                  Text 70, 155, 25, 10, "Region:"
                                  Text 70, 175, 25, 10, "Groups:"
                                  Text 10, 195, 130, 15, "Please reach out to the BlueZone Script team with any questions."
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
                                  If region_action = "Incorrect - Change" Then
                                    new_region = tester.tester_region
                                    EditBox 110, 150, 60, 15, new_region
                                  Else
                                    Text 110, 155, 60, 10, tester.tester_region
                                  End If
                                  If groups_action = "Incorrect - Change" Then
                                    new_groups = join(tester.tester_groups)
                                    EditBox 110, 170, 150, 15, new_groups
                                  Else
                                    Text 110, 175, 150, 10, Join(tester.tester_groups, ",")
                                  End If
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

	'Closes script if the number gets too high (very rare circumstance, just errorproofing)
	If col_in_excel >= 105 then script_end_procedure("This script is only able to assign excel columns to 104 distinct digits. You've exceeded this number, and this script cannot continue.")
end function


function create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
'--- This function is used to grab all active X numbers according to the supervisor X number(s) inputted
'~~~~~ array_name: name of array that will contain all the supervisor's staff x numbers
'~~~~~ supervisor_array: list of supervisor's x numbers seperated by comma
'===== Keywords: MAXIS, array, supervisor, worker number, create
	'Create string with the alphabet
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")

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
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
end function

function create_array_of_all_active_x_numbers_in_county(array_name, county_code)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	call navigate_to_MAXIS_screen("rept", "user")

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
    objMail.Display                                 'To display message

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

            IF TIKL_mo = "01" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #01/21/2021#
            ELSEIF TIKL_mo = "02" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #02/18/2021#
            ELSEIF TIKL_mo = "03" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #03/19/2021#
            ELSEIF TIKL_mo = "04" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #04/20/2021#
            ELSEIF TIKL_mo = "05" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #05/20/2021#
            ELSEIF TIKL_mo = "06" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #06/18/2021#
            ELSEIF TIKL_mo = "07" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #07/21/2021#
            ELSEIF TIKL_mo = "08" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #08/19/2021#
            ELSEIF TIKL_mo = "09" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #09/20/2021#
            ELSEIF TIKL_mo = "10" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #10/21/2021#
            ELSEIF TIKL_mo = "11" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #11/18/2021#
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "21" THEN
                ten_day_cutoff = #12/21/2021#
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #12/21/2020#
            Else
                missing_date = True 'in case TIKL time spans exceed 10 day cut off calendar.
            End if

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

function determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
'--- Function used to return booleans on case and program status based on CASE CURR information. There is no input informat but MAXIS_case_number needs to be defined.
'~~~~~ case_active: Outputs BOOLEAN of if the case is active in any MAXIS program
'~~~~~ case_pending: Outputs BOOLEAN of if the case is pending for any MAXIS Program
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
'~~~~~ unknown_cash_pending: BOOLEAN of if the case has a general 'CASH' program pending but it has not been defined
'===== Keywords: MAXIS, case status, output, status
    Call navigate_to_MAXIS_screen("CASE", "CURR")           'First the function will navigate to CASE/CURR so the inofrmation discovered is based on current status
    family_cash_case = FALSE                                'defaulting all of the booleans
    adult_cash_case = FALSE
    ga_case = FALSE
    msa_case = FALSE
    mfip_case = FALSE
    dwp_case = FALSE
    grh_case = FALSE
    snap_case = FALSE
    ma_case = FALSE
    msp_case = FALSE
    case_active = FALSE
    case_pending = FALSE
    unknown_cash_pending = FALSE
    'The function will use the same functionality for each program and search CASE:CURR to find the program deader for detail about the status.
    'If 'ACTIVE', 'APP CLOSE', 'APP OPEN', or 'PENDING' is listed after the header the function will mark the boolean for that program as 'TRUE'
    'If 'ACTIVE', 'APP CLOSE', or 'APP OPEN' is listed, the function will mark case_active as TRUE
    'If 'PENDING' is listed, the function wil mark case_pending as TRUE
    row = 1                                                 'First we will look for SNAP
    col = 1
    EMSearch "FS:", row, col
    If row <> 0 Then
        EMReadScreen fs_status, 9, row, col + 4
        fs_status = trim(fs_status)
        If fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN" Then
            snap_case = TRUE
            case_active = TRUE
        End If
        If fs_status = "PENDING" Then
            snap_case = TRUE
            case_pending = TRUE
        ENd If
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
        End If
        If grh_status = "PENDING" Then
            grh_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1                                             'Looking for MSA information
    col = 1
    EMSearch "MSA:", row, col
    If row <> 0 Then
        EMReadScreen ms_status, 9, row, col + 5
        ms_status = trim(ms_status)
        If ms_status = "ACTIVE" or ms_status = "APP CLOSE" or ms_status = "APP OPEN" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
        End If
        If ms_status = "PENDING" Then
            msa_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1                                             'Looking for GA information
    col = 1
    EMSearch "GA:", row, col
    If row <> 0 Then
        EMReadScreen ga_status, 9, row, col + 4
        ga_status = trim(ga_status)
        If ga_status = "ACTIVE" or ga_status = "APP CLOSE" or ga_status = "APP OPEN" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_active = TRUE
        End If
        If ga_status = "PENDING" Then
            ga_case = TRUE
            adult_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1                                             'Looking for DWP information
    col = 1
    EMSearch "DWP:", row, col
    If row <> 0 Then
        EMReadScreen dw_status, 9, row, col + 4
        dw_status = trim(dw_status)
        If dw_status = "ACTIVE" or dw_status = "APP CLOSE" or dw_status = "APP OPEN" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
        End If
        If dw_status = "PENDING" Then
            dwp_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1                                             'Looking for MFIP information
    col = 1
    EMSearch "MFIP:", row, col
    If row <> 0 Then
        EMReadScreen mf_status, 9, row, col + 6
        mf_status = trim(mf_status)
        If mf_status = "ACTIVE" or mf_status = "APP CLOSE" or mf_status = "APP OPEN" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_active = TRUE
        End If
        If mf_status = "PENDING" Then
            mfip_case = TRUE
            family_cash_case = TRUE
            case_pending = TRUE
        ENd If
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
        End If
        If ma_status = "PENDING" Then
            ma_case = TRUE
            case_pending = TRUE
        ENd If
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
        End If
        If qm_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
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
        End If
        If sl_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
    End If
    row = 1                                             'Looking for QI information for MSA programs
    col = 1
    EMSearch "QI:", row, col
    If row <> 0 Then
        EMReadScreen qm_status, 9, row, col + 5
        qm_status = trim(qm_status)
        If qm_status = "ACTIVE" or qm_status = "APP CLOSE" or qm_status = "APP OPEN" Then
            msp_case = TRUE
            case_active = TRUE
        End If
        If qm_status = "PENDING" Then
            msp_case = TRUE
            case_pending = TRUE
        ENd If
    End If
End Function

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

function fix_read_data(search_string)
'--- This function fixes data that we are reading from PRISM that includes underscores. The function searches the variable and removes underscores. Then, the fix case function is called to format the string to the correct case & the data is trimmed to remove any excess spaces.
'~~~~~ search_string: the string for the variable to be searched
'===== Keywords: MAXIS, MMIS, PRISM, name, data, fix
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
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
			worker_county_code = "x1" & two_digit_county_code_variable
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
                EMWriteScreen "x", row, 4
                transmit
            Else
                row = 1
                col = 1
                EMSearch " C4", row, col
                If row <> 0 Then
                    EMWriteScreen "x", row, 4
                    transmit
                Else
                    script_end_procedure_with_error_report("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
                End If
            End If

            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "x", row, col - 3
            transmit
        Else
            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "x", row, col - 3
            transmit
        End If
    END IF
End Function

function HH_member_custom_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
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

	'This part takes care of remaining navigation buttons, designed to go to a single panel.
	If ButtonPressed = ADHI_button then call navigate_to_MAXIS_screen("case", "ADHI")
	If ButtonPressed = CURR_button then call navigate_to_MAXIS_screen("case", "CURR")
	If ButtonPressed = ELIG_DWP_button then call navigate_to_MAXIS_screen("elig", "DWP_")
	If ButtonPressed = ELIG_FS_button then call navigate_to_MAXIS_screen("elig", "FS__")
	If ButtonPressed = ELIG_GA_button then call navigate_to_MAXIS_screen("elig", "GA__")
	If ButtonPressed = ELIG_HC_button then call navigate_to_MAXIS_screen("elig", "HC__")
	If ButtonPressed = ELIG_MFIP_button then call navigate_to_MAXIS_screen("elig", "MFIP")
	If ButtonPressed = ELIG_MSA_button then call navigate_to_MAXIS_screen("elig", "MSA_")
	If ButtonPressed = ELIG_WB_button then call navigate_to_MAXIS_screen("elig", "WB__")
	If ButtonPressed = ELIG_GRH_button then call navigate_to_MAXIS_screen("elig", "GRH_")
	IF ButtonPressed = ELIG_SUMM_button then call navigate_to_MAXIS_screen("elig", "SUMM")
	If ButtonPressed = ABPS_button then call navigate_to_MAXIS_screen("stat", "ABPS")
	If ButtonPressed = ACCI_button then call navigate_to_MAXIS_screen("stat", "ACCI")
	If ButtonPressed = ACCT_button then call navigate_to_MAXIS_screen("stat", "ACCT")
	If ButtonPressed = ADDR_button then call navigate_to_MAXIS_screen("stat", "ADDR")
	If ButtonPressed = ADME_button then call navigate_to_MAXIS_screen("stat", "ADME")
	If ButtonPressed = ALTP_button then call navigate_to_MAXIS_screen("stat", "ALTP")
	If ButtonPressed = AREP_button then call navigate_to_MAXIS_screen("stat", "AREP")
	If ButtonPressed = BILS_button then call navigate_to_MAXIS_screen("stat", "BILS")
	If ButtonPressed = BUDG_button then call navigate_to_MAXIS_screen("stat", "BUDG")
	If ButtonPressed = BUSI_button then call navigate_to_MAXIS_screen("stat", "BUSI")
	If ButtonPressed = CARS_button then call navigate_to_MAXIS_screen("stat", "CARS")
	If ButtonPressed = CASH_button then call navigate_to_MAXIS_screen("stat", "CASH")
	If ButtonPressed = COEX_button then call navigate_to_MAXIS_screen("stat", "COEX")
	If ButtonPressed = DCEX_button then call navigate_to_MAXIS_screen("stat", "DCEX")
	If ButtonPressed = DFLN_button then call navigate_to_MAXIS_screen("stat", "DFLN")
	If ButtonPressed = DIET_button then call navigate_to_MAXIS_screen("stat", "DIET")
	If ButtonPressed = DISA_button then call navigate_to_MAXIS_screen("stat", "DISA")
	If ButtonPressed = DISQ_button then call navigate_to_MAXIS_screen("stat", "DISQ")
	If ButtonPressed = EATS_button then call navigate_to_MAXIS_screen("stat", "EATS")
	If ButtonPressed = EMMA_button then call navigate_to_MAXIS_screen("stat", "EMMA")
	If ButtonPressed = EMPS_button then call navigate_to_MAXIS_screen("stat", "EMPS")
	If ButtonPressed = FACI_button then call navigate_to_MAXIS_screen("stat", "FACI")
	If ButtonPressed = FMED_button then call navigate_to_MAXIS_screen("stat", "FMED")
	If ButtonPressed = HCMI_button then call navigate_to_MAXIS_screen("stat", "HCMI")
	If ButtonPressed = HCRE_button then call navigate_to_MAXIS_screen("stat", "HCRE")
	If ButtonPressed = HEST_button then call navigate_to_MAXIS_screen("stat", "HEST")
	If ButtonPressed = IMIG_button then call navigate_to_MAXIS_screen("stat", "IMIG")
	If ButtonPressed = INSA_button then call navigate_to_MAXIS_screen("stat", "INSA")
	If ButtonPressed = JOBS_button then call navigate_to_MAXIS_screen("stat", "JOBS")
	If ButtonPressed = MEDI_button then call navigate_to_MAXIS_screen("stat", "MEDI")
	If ButtonPressed = MEMB_button then call navigate_to_MAXIS_screen("stat", "MEMB")
	If ButtonPressed = MEMI_button then call navigate_to_MAXIS_screen("stat", "MEMI")
    If ButtonPressed = MEMO_button then call navigate_to_MAXIS_screen("spec", "MEMO")
	If ButtonPressed = MMSA_button then call navigate_to_MAXIS_screen("stat", "MMSA")
	If ButtonPressed = MONT_button then call navigate_to_MAXIS_screen("stat", "MONT")
    If ButtonPressed = NOTE_button then call navigate_to_MAXIS_screen("case", "NOTE")
	If ButtonPressed = OTHR_button then call navigate_to_MAXIS_screen("stat", "OTHR")
	If ButtonPressed = PACT_button then call navigate_to_MAXIS_screen("stat", "PACT")
	If ButtonPressed = PARE_button then call navigate_to_MAXIS_screen("stat", "PARE")
	If ButtonPressed = PBEN_button then call navigate_to_MAXIS_screen("stat", "PBEN")
	If ButtonPressed = PDED_button then call navigate_to_MAXIS_screen("stat", "PDED")
	If ButtonPressed = PREG_button then call navigate_to_MAXIS_screen("stat", "PREG")
	If ButtonPressed = PROG_button then call navigate_to_MAXIS_screen("stat", "PROG")
	If ButtonPressed = RBIC_button then call navigate_to_MAXIS_screen("stat", "RBIC")
	If ButtonPressed = REMO_button then call navigate_to_MAXIS_screen("stat", "REMO")
	If ButtonPressed = REST_button then call navigate_to_MAXIS_screen("stat", "REST")
	If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("stat", "REVW")
	If ButtonPressed = SANC_button then call navigate_to_MAXIS_screen("stat", "SANC")
	If ButtonPressed = SCHL_button then call navigate_to_MAXIS_screen("stat", "SCHL")
	If ButtonPressed = SECU_button then call navigate_to_MAXIS_screen("stat", "SECU")
	If ButtonPressed = SHEL_button then call navigate_to_MAXIS_screen("stat", "SHEL")
	If ButtonPressed = SIBL_button then call navigate_to_MAXIS_screen("stat", "SIBL")
	If ButtonPressed = SPON_button then call navigate_to_MAXIS_screen("stat", "SPON")
    If ButtonPressed = SSRT_button then call navigate_to_MAXIS_screen("stat", "SSRT")
	If ButtonPressed = STEC_button then call navigate_to_MAXIS_screen("stat", "STEC")
	If ButtonPressed = STIN_button then call navigate_to_MAXIS_screen("stat", "STIN")
	If ButtonPressed = STWK_button then call navigate_to_MAXIS_screen("stat", "STWK")
	If ButtonPressed = SWKR_button then call navigate_to_MAXIS_screen("stat", "SWKR")
	If ButtonPressed = TIME_button then call navigate_to_MAXIS_screen("stat", "TIME")
	If ButtonPressed = TRAN_button then call navigate_to_MAXIS_screen("stat", "TRAN")
	If ButtonPressed = TYPE_button then call navigate_to_MAXIS_screen("stat", "TYPE")
	If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("stat", "UNEA")
    If ButtonPressed = WCOM_button then call navigate_to_MAXIS_screen("spec", "WCOM")
    If ButtonPressed = WKEX_button then call navigate_to_MAXIS_screen("stat", "WKEX")
	If ButtonPressed = WREG_button then call navigate_to_MAXIS_screen("stat", "WREG")
end function

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
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
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
    End if
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
        EMWriteScreen "x", row, col - 3
        transmit

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		End Select
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
        instr(dail_msg, "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR") OR _
        instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
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
        instr(dail_msg, "MEMBER HAS TURNED 60 - FSET:WORK REG HAS BEEN UPDATED") OR _
        instr(dail_msg, "LAST GRADE COMPLETED") OR _
        instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
        instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
        instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
        instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
        instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
        instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
        instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
        instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
        instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
        instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
        instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
        instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
        instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
        instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
        instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
        instr(dail_msg, "CASE NOT AUTO-APPROVED HRF/SR/RECERT DUE") OR _
        instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
        instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
        instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
        instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
        instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
        instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
        instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
        instr(dail_msg, "- TRANS #") OR _
        instr(dail_msg, "PERSON/S REQD SNAP NOT IN SNAP UNIT") OR _
        instr(dail_msg, "RSDI UPDATED - (REF") OR _
        instr(dail_msg, "SSI UPDATED - (REF") OR _
        instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then
            actionable_dail = False
        '----------------------------------------------------------------------------------------------------CORRECT STAT EDITS over 5 days old
    Elseif dail_type = "STAT" or instr(dail_msg, "NEW FIAT RESULTS EXIST") then
        EmReadscreen stat_date, 8, dail_row, 39
        If isdate(stat_date) = False then
            EmReadscreen alt_stat_date, 8, dail_row, 49
            If isdate(alt_stat_date) = True then
                stat_date = alt_stat_date
            End if
        End if
        If isdate(stat_date) = True then
            five_days_ago = DateAdd("d", -5, date)
            If cdate(five_days_ago) => cdate(stat_date) then
                'messages over 5 days old are non-actionable
                actionable_dail = False
            Else
                actionable_dail = True
            End if
        else
            actionable_dail = True
        End if
    '----------------------------------------------------------------------------------------------------REMOVING PEPR messages not CM or CM + 1
    Elseif dail_type = "PEPR" then
        if dail_month = this_month or dail_month = next_month then
            actionable_dail = True
        Else
            actionable_dail = False ' delete the old messages
        End if
    '----------------------------------------------------------------------------------------------------clearing elig messages older than CM
    Elseif instr(dail_msg, "OVERPAYMENT POSSIBLE") or inStr(dail_msg, "DISBURSE EXPEDITED SERVICE") or instr(dail_msg, "NEW FIAT RESULTS EXIST") or instr(dail_msg, "NEW FS VERSION MUST BE APPROVED") or instr(dail_msg, "APPROVE NEW ELIG RESULTS RECOUPMENT HAS INCREASED") or instr(dail_msg, "PERSON/S REQD FS NOT IN FS UNIT") then
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
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
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
'--- This function sends or hits the PF13 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF13
  EMSendKey "<PF13>"
  EMWaitReady 0, 0
end function

function PF14()
'--- This function sends or hits the PF14 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF14
  EMSendKey "<PF14>"
  EMWaitReady 0, 0
end function

function PF15()
'--- This function sends or hits the PF15 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF15
  EMSendKey "<PF15>"
  EMWaitReady 0, 0
end function

function PF16()
'--- This function sends or hits the PF16 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF16
  EMSendKey "<PF16>"
  EMWaitReady 0, 0
end function

function PF17()
'--- This function sends or hits the PF17 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF17
  EMSendKey "<PF17>"
  EMWaitReady 0, 0
end function

function PF18()
'--- This function sends or hits the PF18 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF18
  EMSendKey "<PF18>"
  EMWaitReady 0, 0
end function

function PF19()
'--- This function sends or hits the PF19 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
end function

function PF20()
'--- This function sends or hits the PF20 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function PF21()
'--- This function sends or hits the PF21 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF21
  EMSendKey "<PF21>"
  EMWaitReady 0, 0
end function

function PF22()
'--- This function sends or hits the PF22 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF22
  EMSendKey "<PF22>"
  EMWaitReady 0, 0
end function

function PF23()
'--- This function sends or hits the PF23 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF23
  EMSendKey "<PF23>"
  EMWaitReady 0, 0
end function

function PF24()
'--- This function sends or hits the PF24 key.
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

        'Determining if the script was successful
        If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
            SCRIPT_success = -1
        else
            SCRIPT_success = 0
        end if

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

        'Determining if the script was successful
        If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
            SCRIPT_success = -1
        else
            SCRIPT_success = 0
        end if

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
                  CheckBox 95, 115, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
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

            Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)

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

    script_list_URL = t_drive & "\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script

    Set objNet = CreateObject("WScript.NetWork")
    windows_user_ID = objNet.UserName
    user_ID_for_validation = ucase(windows_user_ID)

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
                    body_text = body_text & vbCr & "On script - " & name_of_script & "."
                    body_text = body_text & vbCr & "The selection type of - " & selection_type & " was entered into the function call"
                    body_text = body_text & vbCr & "The only valid options are: ALL, SCRIPT, GROUP, POPULATION, or REGION"
                    body_text = body_text & vbCr & "Review the script file particularly the call for the function select_testing_file."
                    Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", "FUNCTION ERROR - select_testing_file for " & name_of_script, body_text, "", TRUE)
            End Select

            If tester.tester_population = "BZ" Then
                allow_option = TRUE
                run_testing_file = TRUE
                selection_type = "SCRIPTWRITER"
            End If

            If run_testing_file = TRUE and allow_option = TRUE Then
                continue_with_testing_file = MsgBox("You have been selected to test this script - " & name_of_script & "." & vbNewLine & vbNewLine & "At this time you can select if you would like to run the testing file or the original file." & vbNewLine & vbNewLine & "** Would you like to test this script now?", vbQuestion + vbYesNo, "Use Testing File")
                If continue_with_testing_file = vbNo Then run_testing_file = FALSE
            End If

            If run_testing_file = TRUE Then
                tester.display_testing_message selection_type, the_selection, force_error_reporting
                ' Call tester.display_testing_message(selection_type, the_selection, force_error_reporting)
                If force_error_reporting = TRUE Then testing_run = TRUE
                If run_locally = true then
                    testing_script_url = "C:\MAXIS-scripts\" & file_path
                Else
                    testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/" & file_branch & "/" & file_path

                End If
                Call run_from_GitHub(testing_script_url)
            End If

        End If
    Next
end function

function start_a_blank_CASE_NOTE()
'--- This function navigates user to a blank case note, presses PF9, and checks to make sure you're in edit mode (keeping you from writing all of the case note on an inquiry screen).
'===== Keywords: MAXIS, case note, navigate, edit
	call navigate_to_MAXIS_screen("case", "note")
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

function start_a_new_spec_memo()
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	' Do		'TODO - Maybe add functionality to keep looping if MEMO is locked.'
	' 	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO
	'
	' 	PF5															'Creates a new MEMO. If it's unable the script will stop.
	' 	EMReadScreen case_locked_check, 11, 24, 2
	' 	If case_locked_check = "CASE LOCKED" Then Call back_to_SELF
	' Loop until case_locked_check <> "CASE LOCKED"
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
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
