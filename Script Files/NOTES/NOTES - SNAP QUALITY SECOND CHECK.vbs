OPTION EXPLICIT

name_of_script = "NOTES - SNAP QUALITY SECOND CHECK.vbs"
start_time = timer

DIM name_of_script
DIM start_time
DIM FuncLib_URL
DIM run_locally
DIM default_directory
DIM beta_agency
DIM req
DIM fso
DIM row

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
'END OF GLOBAL VARIABLES----------------------------------------------------------------------------------------------------

'DECLARING VARIABLES--------------------------------------------------------------------------------------------------------
DIM SNAP_quality_Second_Check_dialog
DIM ButtonPressed
DIM case_number
DIM returned_to_worker_check
DIM SNAP_approved_check
DIM reviewed_by
DIM grant_amount

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog SNAP_quality_Second_Check_dialog, 0, 0, 241, 80, "SNAP quality Second check dialog"
  EditBox 55, 5, 60, 15, case_number
  CheckBox 125, 10, 115, 10, "A SNAP error exists for this case", returned_to_worker_check
  CheckBox 5, 25, 175, 10, "SNAP budget and issuance is correct and approved ", SNAP_approved_check
  EditBox 135, 40, 55, 15, grant_amount
  EditBox 55, 60, 70, 15, reviewed_by
  ButtonGroup ButtonPressed
    OkButton 135, 60, 50, 15
    CancelButton 190, 60, 50, 15
  Text 5, 45, 125, 10, "If approved, what is the grant amount:"
  Text 5, 65, 45, 10, "Reviewed by:"
  Text 5, 10, 45, 10, "Case number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Grabs case number
CALL MAXIS_case_number_finder(case_number)

DO
	DO
		DO
			DO
				Do
					Dialog SNAP_quality_Second_Check_dialog
					cancel_confirmation
					IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid case number"
				LOOP UNTIL IsNumeric(case_number) = TRUE
				If reviewed_by = "" THEN MsgBox "You must sign the case note."
			LOOP until reviewed_by <> ""
			If (returned_to_worker_check = 0 AND SNAP_approved_check = 0) THEN MsgBox "You must check either that the case is correct and has been approved, or an error exists."
		LOOP UNTIL returned_to_worker_check = 1 OR SNAP_approved_check = 1
		If (returned_to_worker_check = 1 AND SNAP_approved_check = 1) THEN MsgBox "You must check either that the case is correct and has been approved, or an error exists."
	LOOP UNTIL returned_to_worker_check = 1 OR SNAP_approved_check = 1
	If (SNAP_approved_check = 1 AND grant_amount = "") THEN Msgbox "You must enter the SNAP grant amount."
LOOP until (SNAP_approved_check = 1 AND grant_amount <> "") OR (SNAP_approved_check = 0 AND grant_amount = "") 	


'Dollar bill symbol will be added to numeric variables (in grant_amount)
IF grant_amount <> "" THEN grant_amount = "$" & grant_amount

'Checking to make sure user is still in active MAXIS session
check_for_MAXIS(TRUE)

'The CASE NOTE----------------------------------------------------------------------------------------------------
'navigates to and starts a new case note
Call start_a_blank_CASE_NOTE
'Case note if case is incorrect
If returned_to_worker_check = 1 THEN 
	Call write_variable_in_CASE_NOTE("~~~SNAP 2nd Check complete, further action required~~~")
	Call write_variable_in_CASE_NOTE("* An error exists in the SNAP budget or issuance.")  
	Call write_variable_in_CASE_NOTE("* The case has been returned to the worker and supervisor for correction.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(reviewed_by)
	'Case note if case is correct
	ELSE IF SNAP_approved_check = 1 THEN 
		Call write_variable_in_CASE_NOTE("~~~SNAP 2nd Check complete & approved " & grant_amount & " SNAP grant~~~")
		Call write_variable_in_CASE_NOTE("* SNAP budget and issuance is correct.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(reviewed_by)
	END IF
END IF	

Script_end_procedure("")

