'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL DAIL"
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

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0



'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"



'Now it checks to make sure MAXIS is running on this screen. If both are running the script will stop.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then EMSendKey "<attn>"
If MAXIS_check <> "MAXIS" then EMWaitReady 1, 2
If MAXIS_check <> "MAXIS" then EMReadScreen training_check, 7, 8, 15
If MAXIS_check <> "MAXIS" then EMReadScreen production_check, 7, 6, 15
If MAXIS_check <> "MAXIS" then If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If MAXIS_check <> "MAXIS" then If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If MAXIS_check <> "MAXIS" then If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If MAXIS_check <> "MAXIS" then If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If MAXIS_check <> "MAXIS" then If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If MAXIS_check <> "MAXIS" then If production_check = "RUNNING" then EMSendKey "1" + "<enter>"

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMReadScreen SELF_check, 27, 2, 28
  EMWaitReady 1, 1
loop until SELF_check = "Select Function Menu (SELF)"

EMSendKey "<home>" + "dail" + "<eraseeof>" + case_number
EMSetCursor 21, 70
EMSendKey "dail" + "<enter>"

script_end_procedure("")






