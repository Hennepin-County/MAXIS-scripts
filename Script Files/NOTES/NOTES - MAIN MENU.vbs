'If date > #10/24/2016# then MsgBox "Note: you appear to be using the old redirect files. You likely need to reinstall your scripts."			'Yells at you if you haven't updated in a while
IF run_locally = TRUE THEN
	script_repository = left(script_repository, len(script_repository) - 13)
ELSE
	IF use_master_branch = TRUE THEN
		script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/"									'Resets to release settings
	ELSE
		script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/RELEASE/"
	END IF
End If 

'LOADING SCRIPT
script_URL = script_repository & "notes/main-menu.vbs"

IF run_locally = False THEN
	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_URL
	req.open "GET", script_URL, FALSE									'Attempts to open the script_URL
	req.send													'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
		Execute req.responseText								'Executes the script code
	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
									"Script URL: " & script_URL & vbNewLine & vbNewLine &_
									"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
									vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
		StopScript
	END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF
