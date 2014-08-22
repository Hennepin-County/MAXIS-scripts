'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - FMED deduction"
start_time = timer

'LOADING ROUTINE FUNCTIONS
'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

BeginDialog worker_sig_dialog, 0, 0, 141, 46, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMWriteScreen "p", 6, 3
EMSendKey "<enter>"
EMWaitReady 0, 0

EMWriteScreen "memo", 20, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF5>"
EMWaitReady 0, 0

EMWriteScreen "x", 5, 10
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "You are turning 60 next month, so you may be eligible for a new deduction for SNAP." + "<newline>" + "<newline>"
EMSendKey "Clients who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month." + "<newline>" + "<newline>"
EMSendKey "If you have medical bills over $35 each month, please contact your worker to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices." + "<newline>" + "<newline>"
EMSendKey "Please call your worker with questions."
EMSendKey "<PF4>"
EMWaitReady 0, 0

EMWriteScreen "case", 19, 22
EMWriteScreen "note", 19, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF9>"
EMWaitReady 0, 0

EMSendKey "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "* Sent MEMO to client about FMED deductions." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig + ", using automated script."

EMSendKey "<PF3>"
EMWaitReady 0, 0

EMSendKey "<PF3>"
EMWaitReady 0, 0

MsgBox "The script has sent a MEMO to the client about the possible FMED deduction, and case noted the action."

script_end_procedure("")






