'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - TPQY response"
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Reads case number
EMReadScreen case_number, 8, 5, 73

'Navigates to INFC
EMSendKey "i"
transmit

'Navigates to SVES
EMWriteScreen "sves", 20, 71
transmit

'Navigates to TPQY
EMWriteScreen "tpqy", 20, 70
transmit

script_end_procedure("")

