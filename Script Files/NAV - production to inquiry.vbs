'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - production to inquiry"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect "A"
row = 1
col = 1
EMSearch "Function: ", row, col
If row = 0 then
  MsgBox "Function not found."
  StopScript
End if
EMReadScreen MAXIS_function, 4, row, col + 10
If MAXIS_function = "____" then
  MsgBox "Function not found."
  StopScript
End if

row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row = 0 then
  MsgBox "Case number not found."
  StopScript
End if
EMReadScreen case_number, 8, row, col + 10

row = 1
col = 1
EMSearch "Month: ", row, col
If row = 0 then
  MsgBox "Footer month not found."
  StopScript
End if
EMReadScreen footer_month, 2, row, col + 7
EMReadScreen footer_year, 2, row, col + 10

row = 1
col = 1
EMSearch "(", row, col
If row = 0 then
  MsgBox "Command not found."
  StopScript
End if
EMReadScreen MAXIS_command, 4, row, col + 1
If MAXIS_command = "NOTE" then MAXIS_function = "CASE"

EMConnect "B"
EMFocus

attn
EMReadScreen inquiry_check, 7, 7, 15
If inquiry_check <> "RUNNING" then 
  MsgBox "Inquiry not found. The script will now stop."
  StopScript
End if

EMWriteScreen "FMPI", 2, 15
transmit

back_to_self

EMWriteScreen MAXIS_function, 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen MAXIS_command, 21, 70
transmit

script_end_procedure("")






