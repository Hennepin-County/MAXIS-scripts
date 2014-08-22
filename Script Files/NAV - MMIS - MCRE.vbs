'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - MMIS - MCRE"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 01:

EMConnect ""

EMSendKey "<attn>"
Do
  EMWaitReady 1, 1
  EMReadScreen MAI_check, 3, 1, 33
  If MAI_check = "   " then EMSendKey "<attn>"
Loop until MAI_check = "MAI"

EMReadScreen MMIS_check, 7, 15, 15 '15 for production, 16 for training (row)
If MMIS_check <> "RUNNING" then script_end_procedure("MMIS is not running on this screen. The script will now stop.")

'SECTION 02:

EMWriteScreen "10", 2, 15 '10 for production, 11 for training
transmit

Do
  PF6
  EMReadScreen session_terminated_check, 18, 1, 7
  If trim(session_terminated_check) = "" then stopscript 'This should check for a password prompt, and bounce out if there is one.
Loop until session_terminated_check = "SESSION TERMINATED"

EMWriteScreen "mw00", 1, 2
transmit
transmit

row = 1
col = 1
EMSearch "EK01", row, col
If row = 0 then script_end_procedure("EK01 (MCRE MMIS) is not found. The script will now stop.")

EMWriteScreen "x", row, 4
transmit

row = 1
col = 1
EMSearch "RECIPIENT FILE APPLICATION", row, col
If row = 0 then script_end_procedure("RECIPIENT FILE APPLICATION is not found. The script will now stop.")

EMWriteScreen "x", row, 3
transmit

script_end_procedure("")






