'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL DAIL"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Connects to BlueZone
EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit

'This checks to maks sure we're in MAXIS.
MAXIS_check_function

'Finds a MAXIS case number (if applicable).
call MAXIS_case_number_finder(case_number)

'Navigates to DAIL/DAIL
call navigate_to_screen("DAIL", "DAIL")

'Ends script
script_end_procedure("")
