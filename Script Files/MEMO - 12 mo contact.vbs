'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - 12 mo contact"
start_time = timer

''LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS

BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 5, 25, 70, 10, "Sign your case note:"
  EditBox 80, 20, 75, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog


'THE SCRIPT

EMConnect ""
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
dialog case_number_dialog
If ButtonPressed = 0 then stopscript
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "MAXIS cannot be found. You may be passworded out. Please try again."
  Stopscript
End if

'THE MEMO

call navigate_to_screen("spec", "memo")
PF5
EMReadScreen MEMO_edit_mode_check, 26, 2, 28
If MEMO_edit_mode_check <> "Notice Recipient Selection" then
  MsgBox "You do not appear to be able to make a MEMO for this case. Are you in inquiry? Is this case out of county? Check these items and try again."
  Stopscript
End if
EMWriteScreen "x", 5, 10
transmit
EMSendKey "************************************************************"
EMSendKey "This notice is to remind you to report changes to your county worker by the 10th of the month following the month of the change. Changes that must be reported are address, people in your household, income, shelter costs and other changes such as legal obligation to pay child support. If you don't know whether to report a change, contact your county worker." & "<newline>"
EMSendKey "************************************************************"
PF4

'THE CASE NOTE
call navigate_to_screen("case", "note")
PF9
EMSendKey "Sent 12 month contact letter via SPEC/MEMO on " & date & ". -" & worker_sig

script_end_procedure("")






