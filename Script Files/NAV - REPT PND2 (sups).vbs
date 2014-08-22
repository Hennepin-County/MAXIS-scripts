'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT PND2 (sups)"
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

'THE SCRIPT

EMConnect ""

dialog worker_dialog
If buttonpressed = cancel then stopscript

MAXIS_check_function

call navigate_to_screen("rept", "pnd2")

If worker_number <> "" then
  EMWriteScreen worker_county_code & worker_number, 21, 13
  transmit
End if

script_end_procedure("")






