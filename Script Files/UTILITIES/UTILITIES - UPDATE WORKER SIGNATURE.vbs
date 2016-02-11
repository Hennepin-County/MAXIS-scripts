'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - UPDATE WORKER SIGNATURE.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'----------DIALOGS----------
BeginDialog update_worker_signature_dialog, 0, 0, 191, 105, "Update Worker Signature"
  EditBox 10, 60, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 85, 50, 15
    CancelButton 95, 85, 50, 15
  Text 10, 10, 175, 10, "Enter what you would like for your default signature."
  Text 10, 25, 170, 25, "NOTE: This will be pre-loaded in every script. Once the script has started, you can still modify your signature in the appropriate editbox."
EndDialog

'----------THE SCRIPT----------
dialog update_worker_signature_dialog				'Shows the dialog
IF ButtonPressed = cancel THEN stopscript			'Handling for if cancel is pressed
IF worker_signature = "" THEN stopscript			'If they enter nothing, it exits

'This creates an object which collects the username from the Windows logon. We need this to determine the correct location for the My Documents folder.
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName		'Saves the .UserName object as a new variable, windows_user_ID

'Opens an FSO, opens workersig.txt, writes the new signature in, and exits
SET update_worker_sig_fso = CreateObject("Scripting.FileSystemObject")
SET update_worker_sig_command = update_worker_sig_fso.CreateTextFile("C:\USERS\" & windows_user_ID & "\MY DOCUMENTS\workersig.txt", 2)
update_worker_sig_command.Write(worker_signature)
update_worker_sig_command.Close

'Script ends
script_end_procedure("")
