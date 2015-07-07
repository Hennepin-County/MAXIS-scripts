'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - APPLYMN APPLICATION RECEIVED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog apply_MN_dialog, 0, 0, 291, 125, "Apply MN"
  EditBox 60, 5, 75, 15, case_number
  EditBox 90, 25, 75, 15, app_date
  EditBox 185, 25, 40, 15, app_time
  DropListBox 235, 25, 35, 15, "AM"+chr(9)+"PM", AM_PM
  EditBox 65, 45, 100, 15, confirmation_number
  EditBox 50, 65, 205, 15, progs_applied_for
  DropListBox 55, 85, 80, 15, "N/A"+chr(9)+"known to EBT"+chr(9)+"unknown to EBT", EBT_status
  DropListBox 180, 85, 105, 15, "SPEC/XFERed to worker."+chr(9)+"Indexed to worker.", actions_taken
  EditBox 70, 105, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 105, 50, 15
    CancelButton 235, 105, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 80, 10, "Apply MN app rec'd on"
  Text 170, 30, 10, 10, "at"
  Text 5, 50, 55, 10, "Confirmation #:"
  Text 5, 70, 45, 10, "Applying for:"
  Text 5, 90, 40, 10, "EBT status:"
  Text 145, 90, 30, 10, "Actions:"
  Text 5, 110, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Finds case number
call MAXIS_case_number_finder(case_number)

'Shows dialog and navigates to case note
Dialog apply_MN_dialog
cancel_confirmation

'checking for an active MAXIS session	
Call check_for_MAXIS(True)
	
call navigate_to_MAXIS_screen("case", "note")

'Case notes information
Call write_variable_in_CASE_NOTE("ApplyMN app rec'd on " & app_date & " at " & app_time & " " & AM_PM)
call write_variable_in_case_note("Confirmation #", confirmation_number, 6) 'x is the header, y is the variable for the edit box which will be put in the case note.
call write_variable_in_case_note("Applying for", progs_applied_for, 6) 'x is the header, y is the variable for the edit box which will be put in the case note.
If EBT_status <> "N/A" then call Call write_variable_in_CASE_NOTE("* Client is " & EBT_status & ".")
call Call write_variable_in_CASE_NOTE("* " & actions_taken)
call Call write_variable_in_CASE_NOTE("---")
call Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")