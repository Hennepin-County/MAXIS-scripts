'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT REVW"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog case_number_dialog, 0, 0, 161, 41, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

'FINDING THE CASE NUMBER----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

EMConnect ""

'NAVIGATING TO THE SCREEN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found on this screen.")

call navigate_to_screen("rept", "revw")

script_end_procedure("")






