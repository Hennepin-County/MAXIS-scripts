'LOADING GLOBAL VARIABLES
Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

'REFERENCES REDIRECT ON TDRIVE
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\redirect-utilities-one-source.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

