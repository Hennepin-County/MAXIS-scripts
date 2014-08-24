'>>>>THIS IS A DUMMY VERSION, ONLY TO BE USED IN TESTING THE CONFIG PROGRAM<<<<
'
'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = ""
'start_time = timer
'
''LOADING ROUTINE FUNCTIONS
'Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Config source\DHS BZS config\bin\Debug\FUNCTIONS FILE.vbs")
'text_from_the_other_script = fso_command.ReadAll
'fso_command.Close
'Execute text_from_the_other_script

'----------------------------------------------------------------------------------------------------

'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------

worker_county_code = "x103"
collecting_statistics = False
EDMS_choice = "DHS eDocs"
county_name = "Becker"
county_address_line_01 = "rerw"
county_address_line_02 = "werwer"
case_noting_intake_dates = True
move_verifs_needed = False

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'SHARED VARIABLES----------------------------------------------------------------------------------------------------
checked = 1		'Value for checked boxes
unchecked = 0	'Value for unchecked boxes
cancel = 0		'Value for cancel button in dialogs
OK = -1		'Value for OK button in dialogs

'Some screens require the two digit county code, and this determines what that code is
two_digit_county_code = right(worker_county_code, 2)
If two_digit_county_code = "PW" then two_digit_county_code = "91"	'For DHS purposes











