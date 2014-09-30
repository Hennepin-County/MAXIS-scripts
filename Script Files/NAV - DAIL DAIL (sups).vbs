'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL DAIL (sups)"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog worker_dialog, 0, 0, 171, 45, "Worker dialog"
  Text 5, 10, 130, 10, "Enter the worker number (last 3 digits):"
  EditBox 135, 5, 30, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 30, 25, 50, 15
    CancelButton 90, 25, 50, 15
EndDialog

call worker_county_code_determination(worker_county_code, two_digit_county_code)

'Connects to BlueZone
EMConnect ""

dialog worker_dialog
If buttonpressed = 0 then stopscript

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit

'This checks to maks sure we're in MAXIS.
MAXIS_check_function

'Finds a MAXIS case number (if applicable).
call MAXIS_case_number_finder(case_number)

'Navigates to DAIL/DAIL
call navigate_to_screen("DAIL", "DAIL")

'Inputs worker_number variable to the DAIL screen.
If worker_number <> "" then
	EMWriteScreen worker_county_code & worker_number, 21, 6
	transmit
End if

'Ends script
script_end_procedure("")
