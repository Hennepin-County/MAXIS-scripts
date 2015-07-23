'STATS GATHERING----------------------------------------------------------------------------------------------------

'IMPORTANT!!! change the name part ..." NOTES - CAF.vbs "...to the file you want this to open.
name_of_script = "NOTES - MNSURE HC RETRO APPL.vbs"
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



'dialogs--------------------------------------------------------------------------------------------------------------------

'this dialog is the core case notes of the application processing case notes.


BeginDialog MNSure_HC_Appl_dialog, 0, 0, 346, 245, "MNSure HC Appl"
  EditBox 80, 30, 50, 15, case_number
  EditBox 215, 30, 75, 15, curam_case_number
  EditBox 115, 50, 60, 15, HC_Appl_date_Recvd
  EditBox 230, 50, 90, 15, time_gap_between
  DropListBox 115, 70, 60, 15, "Select One"+chr(9)+"1 month"+chr(9)+"2 months"+chr(9)+"3 months", retro_coverage_months
  EditBox 165, 90, 65, 15, hc_closed_120days
  EditBox 90, 105, 245, 15, HH_members_requesting
  DropListBox 90, 125, 60, 15, "Select One"+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Pending", HC_Appl_status
  EditBox 110, 150, 220, 15, missing_documents
  EditBox 65, 170, 270, 15, action_done_taken
  CheckBox 5, 195, 80, 10, "TIKL for 10 day return", tikl_return_date
  EditBox 80, 225, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 210, 225, 50, 15
    CancelButton 265, 225, 50, 15
  Text 5, 15, 220, 10, "MA application in Curam, Retro eligibility determination requested "
  Text 5, 230, 70, 10, "Worker's Signature:"
  Text 140, 35, 70, 10, "Curam case number:"
  Text 5, 55, 105, 10, "Curam Application Rec'vd Date:"
  Text 5, 110, 80, 10, "HH Membs requesting:"
  Text 5, 130, 70, 10, "HC Appl Status?"
  Text 5, 35, 70, 10, "Maxis Case Number:"
  Text 5, 90, 155, 10, " If HC closed within 120 days, date of closure:"
  Text 185, 55, 40, 10, "Gap Month:"
  Text 5, 155, 100, 10, "Verification Forms Needed:"
  Text 10, 75, 100, 10, "# of Retro Months requested:"
  Text 5, 175, 50, 10, "Action Taken:"
EndDialog


'build script-----------------------------------------------------------------------------------------------------------------------------

EMConnect ""
EMFocus

call MAXIS_case_number_finder (case_number)

'do a dialog loop for the worker to enter the condition for the case notes
			

	DO
		DO
			DO
				Do

					Dialog MNSure_HC_Appl_dialog
					cancel_confirmation

					'Looping these conditions for the dialog boxes
					IF retro_coverage_months = "Select One" THEN MsgBox "If Retro months requested, please select how many months requested"
				Loop until retro_coverage_months <> "Select One"
		

				IF HH_members_requesting = "" THEN MsgBox "Enter HH members requesting for this medical"
			Loop until HH_members_requesting <> ""

			IF HC_Appl_status = "" THEN MsgBox "Select a status for this application"
		Loop until HC_Appl_status <> ""

		IF missing_documents = "" THEN MsgBox "Your list of documents are not entered, or Not Applicable"
	Loop until missing_documents <> ""

					

call start_a_blank_CASE_NOTE


'call write_variable_in_CASE_NOTE("MNSure-HC-Retro-Appl")
'Case Noting via Script....

CALL write_variable_in_CASE_NOTE("HC Application for Curam & Retro Coverage")
Call write_bullet_and_variable_in_CASE_NOTE("Case Number", case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Curam Case #", curam_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Curam Appl Rec'vd", HC_Appl_date_Recvd)
Call write_bullet_and_variable_in_CASE_NOTE("Gap Month", time_gap_between)
Call write_bullet_and_variable_in_CASE_NOTE("# of Retro months HC request?", retro_coverage_months)
Call write_bullet_and_variable_in_CASE_NOTE("HC Closed within 120 days?", hc_closed_120days)
Call write_bullet_and_variable_in_CASE_NOTE("HH Membs Requesting", HH_members_requesting)
Call write_bullet_and_variable_in_CASE_NOTE("HC Appl Status", HC_Appl_status)
Call write_bullet_and_variable_in_CASE_NOTE("Verification Forms Needed", missing_documents)
Call write_bullet_and_variable_in_CASE_NOTE("Missing Documents", missing_documents)
Call write_bullet_and_variable_in_CASE_NOTE("Requests due back", hc_request_due_date)
Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_done_taken)
Call write_bullet_and_variable_in_CASE_NOTE("Request due back", request_due_back)
Call write_variable_in_CASE_NOTE(worker_signature)



script_end_procedure ("")




	
