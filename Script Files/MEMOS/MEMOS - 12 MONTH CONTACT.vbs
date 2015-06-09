'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - 12 MO CONTACT.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'<<DELETE REDUNDANCIES!

'DIALOGS

BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 5, 25, 70, 10, "Sign your case note:"
  EditBox 80, 20, 75, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog


'THE SCRIPT

EMConnect ""
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
dialog case_number_dialog
If ButtonPressed = 0 then stopscript
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "MAXIS cannot be found. You may be passworded out. Please try again."
  Stopscript
End if

'THE MEMO

call navigate_to_screen("spec", "memo")
PF5
EMReadScreen MEMO_edit_mode_check, 26, 2, 28
If MEMO_edit_mode_check <> "Notice Recipient Selection" then
  MsgBox "You do not appear to be able to make a MEMO for this case. Are you in inquiry? Is this case out of county? Check these items and try again."
  Stopscript
End if
EMWriteScreen "x", 5, 10
transmit
EMSendKey "************************************************************"
EMSendKey "This notice is to remind you to report changes to your county worker by the 10th of the month following the month of the change. Changes that must be reported are address, people in your household, income, shelter costs and other changes such as legal obligation to pay child support. If you don't know whether to report a change, contact your county worker." & "<newline>"
EMSendKey "************************************************************"
PF4

'THE CASE NOTE
call navigate_to_screen("case", "note")
PF9
EMSendKey "Sent 12 month contact letter via SPEC/MEMO on " & date & ". -" & worker_sig

script_end_procedure("")






