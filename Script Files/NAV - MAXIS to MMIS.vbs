'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - MAXIS to MMIS"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'First checks to make sure you're in MAXIS.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then EMReadScreen approval_confirmation_check, 21, 3, 30
If approval_confirmation_check = "Approval Confirmation" then MAXIS_check = "MAXIS" 'Simplifies the next move
If MAXIS_check <> "MAXIS" then script_end_procedure("You aren't in MAXIS! This script works by starting in MAXIS on a case.")

'Searching for the case number, using row/col variables. If not found, the script will exit.
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row = 0 then script_end_procedure("A valid case number could not be found. This script works best from a STAT, CASE, or ELIG screen.")

'Reading the case number, then removing spaces and underscores, and adding the leading zeroes for MMIS.
EMReadScreen case_number, 8, row, col + 10
case_number = replace(replace(case_number, " ", ""), "_", "0") 'Removing any underscores.
Do
	If len(case_number) < 8 then case_number = "0" & case_number
Loop until len(case_number) = 8

'Checking to see if we are on the HC/APP screen, which is not supported at this time (case number is in different place)
EMReadScreen HC_app_check, 16, 3, 33 
If HC_app_check = "Approval Package" then script_end_procedure("The script needs to be on the previous or next screen to process this.")

'Now it will look for MMIS on both screens, and enter into it.. 
attn
EMReadScreen MMIS_A_check, 7, 15, 15
If MMIS_A_check = "RUNNING" then
	EMSendKey "10"
	transmit
Else
	attn
	EMConnect "B"
	attn
	EMReadScreen MMIS_B_check, 7, 15, 15
	If MMIS_B_check <> "RUNNING" then 
		script_end_procedure("MMIS does not appear to be running. This script will now stop.")
	Else
		EMSendKey "10"
		transmit
	End if
End if
EMFocus 'Bringing window focus to the second screen if needed.

'Sending MMIS back to the beginning screen and checking for a password prompt
Do 
  PF6
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
  EMReadScreen session_start, 18, 1, 7
Loop until session_start = "SESSION TERMINATED"

'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
EMWriteScreen "mw00", 1, 2
transmit
transmit

'Finding the right MMIS, if needed, by checking the header of the screen to see if it matches the security group selector
EMReadScreen MMIS_security_group_check, 21, 1, 35 
If MMIS_security_group_check = "MMIS MAIN MENU - MAIN" then
	EMSendKey "x"
	transmit
End if

'Now it finds the recipient file application feature and selects it.
row = 1
col = 1
EMSearch "RECIPIENT FILE APPLICATION", row, col
EMWriteScreen "x", row, col - 3
transmit

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "i", 2, 19
EMWriteScreen case_number, 9, 19
transmit
EMReadscreen RKEY_check, 4, 1, 52
If RKEY_check = "RKEY" then script_end_procedure("A correct case number was not taken from MAXIS. Check your case number and try again.")

'Now it gets to RELG for member 01 of this case.
EMWriteScreen "rcin", 1, 8
transmit
EMWriteScreen "x", 11, 2
EMWriteScreen "relg", 1, 8
transmit

script_end_procedure("")






