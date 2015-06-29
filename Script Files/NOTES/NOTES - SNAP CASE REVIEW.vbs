'OPTION EXPLICIT

name_of_script = "NOTES - SNAP CASE REVIEW.vbs"
start_time = timer

'DIM name_of_script
'DIM start_time
'DIM FuncLib_URL
'DIM run_locally
'DIM default_directory
'DIM beta_agency
'DIM req
'DIM fso
'DIM row

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
'DIM SNAP_quality_case_review_dialog
'DIM ButtonPressed
'DIM case_number
'DIM MAXIS_footer_month
'DIM MAXIS_footer_year
'DIM SNAP_status
'DIM grant_amount
'DIM worker_signature
'DIM footer_month
'DIM footer_year


'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog SNAP_quality_case_review_dialog, 0, 0, 246, 85, "SNAP quality case review dialog"
  EditBox 65, 5, 65, 15, case_number
  EditBox 185, 5, 25, 15, MAXIS_footer_month
  EditBox 215, 5, 25, 15, MAXIS_footer_year
  DropListBox 135, 25, 105, 15, "Select one..."+chr(9)+"correct & approved"+chr(9)+"error exists", SNAP_status
  EditBox 135, 45, 105, 15, grant_amount
  EditBox 65, 65, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 65, 50, 15
    CancelButton 190, 65, 50, 15
  Text 5, 70, 60, 10, "Worker signature:"
  Text 5, 30, 100, 10, "SNAP budget/issuance status:"
  Text 5, 10, 45, 10, "Case number:"
  Text 135, 10, 45, 10, "Footer month:"
  Text 5, 50, 125, 10, "If approved, what is the grant amount:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'Grabs case number
CALL MAXIS_case_number_finder(case_number)

'Grabbing the footer month/year
Call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
	footer_month = MAXIS_footer_month
	call find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
	If row <> 0 then footer_year = MAXIS_footer_year
End if

DO
	DO
		DO
			Do
				DO
					Dialog SNAP_quality_case_review_dialog
					cancel_confirmation
					IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid case number"
				LOOP UNTIL IsNumeric(case_number) = TRUE
				If worker_signature = "" THEN MsgBox "You must sign the case note."
			LOOP until worker_signature <> ""
			If (MAXIS_footer_month = "" AND MAXIS_footer_year = "") OR (MAXIS_footer_month <> "" AND MAXIS_footer_year = "") OR (MAXIS_footer_month = "" AND MAXIS_footer_year <> "") THEN MsgBox "You must enter the footer year AND footer month."
		LOOP until (MAXIS_footer_month <> "" AND MAXIS_footer_year <> "")
		If SNAP_status = "Select one..." THEN MsgBox "You must check either that the case is correct and approved, or an error exists."
	LOOP UNTIL SNAP_status <> "Select one..."
	If (SNAP_status = "correct & approved" AND grant_amount = "") OR (SNAP_status = "error exists" AND grant_amount <> "") THEN Msgbox "You must either select 'error exists', and leave the grant amount blank OR select 'correct & approved', and enter the grant amount. "
LOOP until (SNAP_status = "correct & approved" AND grant_amount <> "") OR (SNAP_status = "error exists" AND grant_amount = "") 	


'Dollar bill symbol will be added to numeric variables (in grant_amount)
IF grant_amount <> "" THEN grant_amount = "$" & grant_amount

'Checking to make sure user is still in active MAXIS session
check_for_MAXIS(TRUE)

'The CASE NOTE----------------------------------------------------------------------------------------------------
'navigates to case note and creates a new one
Call start_a_blank_CASE_NOTE
'Case note if case is incorrect
If SNAP_status = "error exists" THEN
	Call write_variable_in_CASE_NOTE("~~~SNAP case review complete, further action required~~~")
	Call write_variable_in_CASE_NOTE("* An error exists in the SNAP budget or issuance.")  
	Call write_variable_in_CASE_NOTE("* The case has been returned to the worker and supervisor for correction.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	'Case note if case is correct
	ELSEIF SNAP_status = "correct & approved" THEN 
		Call write_variable_in_CASE_NOTE("~~~SNAP case review complete & app'd for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " of " & grant_amount & " SNAP grant~~~")
		Call write_variable_in_CASE_NOTE("* SNAP case has been reviewed, and the budget and issuance is correct.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)	
END If

Script_end_procedure("")

