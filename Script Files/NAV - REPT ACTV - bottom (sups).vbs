'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT ACTV - bottom (sups)"
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
If buttonpressed = 0 then stopscript

transmit

'Now it checks to make sure MAXIS is running on this screen. If both are running the script will stop.
MAXIS_check_function
back_to_self
EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "actv", 21, 70
transmit

EMWriteScreen worker_county_code & worker_number, 21, 13
transmit

do
  PF8
  EMReadScreen last_page_check, 21, 24, 2
loop until last_page_check = "THIS IS THE LAST PAGE"

script_end_procedure("")






