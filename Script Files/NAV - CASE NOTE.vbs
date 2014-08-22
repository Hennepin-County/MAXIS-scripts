'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - CASE NOTE"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'<<<<<<DELETE OLD REDUNDANT FUNCTIONS BELOW
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

  row = 1
  col = 1

EMSearch "Case Nbr: ", row, col

EMReadScreen case_number, 8, row, col + 10
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""

BeginDialog case_number_dialog, 0, 0, 161, 42, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup ButtonPressed_case_number
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

If case_number = "" then Dialog case_number_dialog

If case_number = "" and ButtonPressed_case_number = 0 then stopscript

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0



'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"


'Checks for MAXIS
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then
  MsgBox "You do not seem to be in MAXIS. The script will now stop."
  StopScript
End if

Sub not_in_CASE
'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMReadScreen SELF_check, 27, 2, 28
    EMWaitReady 1, 1
  loop until SELF_check = "Select Function Menu (SELF)"

  EMSendKey "<home>" + "case" + "<eraseeof>" + case_number
  EMSetCursor 21, 70
  EMSendKey "NOTE" + "<enter>"
End Sub

Sub in_CASE
  EMWriteScreen "NOTE", 20, 71
  EMSendKey "<enter>"
End Sub

EMReadScreen CASE_check, 4, 20, 21
If CASE_check = "CASE" then call in_CASE
If CASE_check <> "CASE" then call not_in_CASE

script_end_procedure("")






