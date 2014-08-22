'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL WRIT"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'<<<<<<DELETE OLD REDUNDANT FUNCTIONS BELOW
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
EMConnect ""

'SECTION 01
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "This script needs to be run from MAXIS."
  StopScript
End If

row = 1
col = 1

EMSearch "Case Nbr:", row, col
If row = 0 then
  second_row = 1
  second_col = 1
  EMSearch "Case Number:", second_row, second_col
  If second_row = 0 then
    MsgBox "A case number could not be found on this screen. The script will now stop."
    StopScript
  End If
  If second_row <> 0 then EMReadScreen case_number, 8, second_row, second_col + 13
End If
If row <> 0 then EMReadScreen case_number, 8, row, col + 10
case_number = replace(case_number, "_", "")
case_number = trim(case_number)


'SECTION 02
Do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then stopscript 'This will stop the script from acting if it passwords out.
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

EMWriteScreen "DAIL", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "WRIT", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

script_end_procedure("")






