'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT ACTV - bottom"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit

'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check = "MAXIS" and MAXIS_check = "AXIS " then script_end_procedure("MAXIS not found. Are you passworded out? Navigate to MAXIS and try again.")

'This Do...loop gets back to SELF
do
  PF3
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

'Enter keys for REPT/ACTV and transmit
EMSendKey "<home>" & "rept" & "<eraseeof>" & "<newline>" & "<newline>" & "actv" & "<enter>"
EMWaitReady 0, 0

'Presses "PF8" until the last page is found
do
  PF8
  EMReadScreen test, 21, 24, 2
loop until test = "THIS IS THE LAST PAGE"

script_end_procedure("")






