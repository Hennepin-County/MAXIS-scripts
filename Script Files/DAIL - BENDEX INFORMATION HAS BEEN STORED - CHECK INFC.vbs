'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - BENDEX INFORMATION HAS BEEN STORED - CHECK INFC"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""
EMSendKey "i" + "<enter>"

EMWaitReady 0, 0
EMSetCursor 20, 71
EMSendKey "bndx" + "<enter>"

EMWaitReady 0, 0

script_end_procedure("")






