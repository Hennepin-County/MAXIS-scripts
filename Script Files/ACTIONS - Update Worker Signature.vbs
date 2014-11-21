'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - Update Worker Sig"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

worker_signature = InputBox("Please enter what you would like for your default worker signature (NOTE: this will create the signature that is autofilled as worker signature in scripts)")

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = objNet.UserName

SET update_worker_sig_fso = CreateObject("Scripting.FileSystemObject")
SET update_worker_sig_command = update_worker_sig_fso.CreateTextFile("C:\USERS\" & windows_user_ID & "\MY DOCUMENTS\workersig.txt", 2)
update_worker_sig_command.WriteLine(worker_signature)
update_worker_sig_command.Close

script_end_procedure("")