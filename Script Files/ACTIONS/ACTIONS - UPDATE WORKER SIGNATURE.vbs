'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - UPDATE WORKER SIGNATURE.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'----------DIALOGS----------
BeginDialog worker_sig_dlg, 0, 0, 191, 105, "Update Worker Signature"
  EditBox 10, 60, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 85, 50, 15
    CancelButton 95, 85, 50, 15
  Text 10, 10, 175, 10, "Enter what you would like for your default signature."
  Text 10, 25, 170, 25, "NOTE: This will be pre-loaded in every script. Once the script has started, you can still modify your signature in the appropriate editbox."
EndDialog

'----------THE SCRIPT----------
DIALOG worker_sig_dlg
	IF ButtonPressed = 0 THEN stopscript
	IF worker_signature = "" THEN stopscript

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = objNet.UserName

SET update_worker_sig_fso = CreateObject("Scripting.FileSystemObject")
SET update_worker_sig_command = update_worker_sig_fso.CreateTextFile("C:\USERS\" & windows_user_ID & "\MY DOCUMENTS\workersig.txt", 2)
update_worker_sig_command.Write(worker_signature)
update_worker_sig_command.Close

script_end_procedure("")
