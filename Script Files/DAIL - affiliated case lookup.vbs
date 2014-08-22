EMConnect ""

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - affiliated case lookup"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

  row = 1
  col = 1
  cola = 1
EMSearch "#", 6, col
EMSearch ")", 6, cola

case_number_digits = cola - col - 1
EMReadScreen case_number, case_number_digits, 6, col + 1
If IsNumeric(case_number) = False then MsgBox "An affiliated case could not be detected on this dail message. Try another dail message."
If IsNumeric(case_number) = False then stopscript

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"


EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey "<eraseEOF>" + case_number
EMSetCursor 21, 70
EMSendKey "note" + "<enter>"

MsgBox "You are now in case/note for the affiliated case!"

script_end_procedure("")






