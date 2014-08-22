'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - returned mail"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'FUNCTIONS
EMConnect ""

call MAXIS_case_number_finder(case_number)

BeginDialog returned_mail_dialog, 0, 0, 236, 92, "Returned mail"
  EditBox 140, 0, 65, 15, case_number
  EditBox 100, 50, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 70, 50, 15
    CancelButton 120, 70, 50, 15
  Text 35, 55, 60, 10, "Worker signature:"
  Text 30, 5, 110, 10, "Case number with returned mail:"
  Text 10, 20, 220, 25, "Note: if you have mail with an allowed forwarding address, update MAXIS per policy. Do not use this script with a forwarding address. Ask a supervisor if you have questions about returned mail policy."
EndDialog

Do
  Dialog returned_mail_dialog
  If buttonpressed = 0 then stopscript
Loop until trim(case_number) <> ""
transmit 'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then stopscript


Call navigate_to_screen("case", "note")

'If there was an error after trying to go to CASE/NOTE, the script will shut down.
EMReadScreen SELF_error_check, 27, 2, 28 
If SELF_error_check = "Select Function Menu (SELF)" then stopscript

'Now the script goes into the case note and case notes the action. 

PF9

EMReadScreen mode_check, 7, 20, 3
If mode_check <> "Mode: A" and mode_check <> "Mode: E" then
  MsgBox "Unable to start a case note. Is this inquiry mode? Is this case out of county? Right case number? Check these out and try again!"
  StopScript
End if

EMSendKey "-->Returned mail received<--" + "<newline>"
EMSendKey "* No forwarding address was indicated." + "<newline>"
EMSendKey "* Sending verification request to last known address. TIKLed for 10-day return." + "<newline>"
EMSendKey "---" + "<newline>" + worker_signature

PF3

call navigate_to_screen("dail", "writ")


ten_days_from_today = dateadd("d", date, 10)
TIKL_day = datepart("d", ten_days_from_today)
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_month = datepart("m", ten_days_from_today)
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_year = (datepart("yyyy", ten_days_from_today)) - 2000

EMWriteScreen TIKL_month, 5, 18
EMWriteScreen TIKL_day, 5, 21
EMWriteScreen TIKL_year, 5, 24
EMSetCursor 9, 3
EMSendKey "Request for address sent 10 days ago. If not responded to, take appropriate action. (TIKL generated via script)" 

transmit
PF3

MsgBox "Use the appropriate returned mail paperwork in " & EDMS_choice & ". Send the completed forms to the most recent address. The script has case noted that returned mail was received and TIKLed out for 10-day return."
script_end_procedure("")






