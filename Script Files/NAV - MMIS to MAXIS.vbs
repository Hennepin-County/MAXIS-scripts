'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - MMIS to MAXIS"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""


EMReadScreen PMI_number, 8, 2, 2
PMI_number = replace(PMI_number, " ", "")
If len(PMI_number) <> 8 then
  MsgBox "A PMI number could not be found on this screen!"
  stopscript
End if


'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If production_check = "RUNNING" then EMSendKey "1" + "<enter>"

'This Do...loop gets back to SELF
do
  PF3
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then stopscript
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWaitReady 1, 1

EMSendKey "<home>" + "pers" + "<eraseeof>" + "<enter>"
EMWaitReady 1, 1
EMSetcursor 15, 36
EMSendKey PMI_number
transmit
EMReadScreen MTCH_check, 4, 2, 51
If MTCH_check <> "MTCH" then stopscript
EMWriteScreen "x", 8, 5
transmit
Do
  row = 1
  col = 1
  EMSearch "  Y    ", row, col
  If row = 0 then 
    PF8
  end if
  EMReadScreen page_check, 21, 24, 2
  If page_check = "THIS IS THE ONLY PAGE" or page_check = "THIS IS THE LAST PAGE" then script_end_procedure("A case could not be found for this PMI. They could be a spouse or other member on an existing case.")
Loop until row <> 0
EMWriteScreen "x", row, 4
transmit

EMWriteScreen "case", 16, 43
EMWriteScreen "note", 21, 70
transmit

script_end_procedure("")






