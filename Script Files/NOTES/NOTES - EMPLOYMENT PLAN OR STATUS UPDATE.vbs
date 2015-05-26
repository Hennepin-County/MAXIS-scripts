'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs"
start_time = timer

'Option Explicit

DIM beta_agency
DIM FuncLib_URL, req, fso

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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 196, 130, "Status Update / Employment Plan"
  ButtonGroup ButtonPressed
    OkButton 55, 95, 50, 15
    CancelButton 120, 95, 50, 15
  DropListBox 15, 40, 85, 15, "Status Update"+chr(9)+"Employment Plan", update_type
  DropListBox 110, 40, 60, 15, "Received"+chr(9)+"Sent", received_sent
  EditBox 110, 10, 65, 15, case_number
  EditBox 110, 60, 65, 15, document_date
  Text 20, 15, 55, 15, "Case Number:"
  Text 20, 65, 65, 15, "Document Date:"
EndDialog


BeginDialog employment_plan_dialog, 0, 0, 311, 265, "Employment Plan Received"
   ButtonGroup ButtonPressed
    OkButton 185, 245, 50, 15
    CancelButton 245, 245, 50, 15
  DropListBox 65, 10, 45, 15, "DWP"+chr(9)+"MFIP", program_list
  EditBox 210, 10, 30, 15, hh_member
  EditBox 65, 35, 75, 15, ES_provider
  EditBox 210, 35, 85, 15, ES_counselor
  DropListBox 65, 60, 90, 15, "1. Employment Search"+chr(9)+"2. Employment"+chr(9)+"3. High School / GED"+chr(9)+"4. Higher Ed"+chr(9)+"5. Health / Medical", primary_activity
  EditBox 210, 60, 40, 15, activity_hours
  CheckBox 65, 80, 45, 15, "FSS", FSS_check
  CheckBox 115, 80, 35, 15, "UP", UP_check
  CheckBox 160, 80, 40, 15, "Other", other_check
  EditBox 65, 100, 135, 15, job_info
  CheckBox 215, 100, 75, 15, "Verif on file.", job_verif_check
  EditBox 65, 120, 135, 15, school_info
  CheckBox 215, 120, 70, 15, "Verif on file.", school_verif_check
  EditBox 65, 140, 85, 15, disa_end_date
  CheckBox 215, 140, 70, 15, "MOF on file.", MOF_check
  EditBox 65, 160, 235, 15, actions_taken
  EditBox 65, 180, 235, 15, other_notes
  EditBox 205, 220, 95, 15, worker_signature
  Text 5, 15, 55, 10, "Program:"
  Text 5, 35, 50, 15, "ES Provider:"
  Text 155, 35, 45, 15, "Counselor:"
  Text 5, 60, 55, 15, "Primary Activity:"
  Text 160, 60, 45, 10, "Hours:"
  Text 5, 160, 55, 10, "Actions Taken:"
  Text 10, 180, 50, 10, "Other Notes:"
  Text 130, 220, 65, 15, "Worker Signature:"
  Text 130, 10, 70, 15, "HH Member number:"
  Text 5, 85, 40, 10, "Status:"
  Text 5, 100, 30, 15, "Job:"
  Text 5, 120, 35, 15, "School:"
  Text 5, 140, 50, 15, "Disa end date:"
  
EndDialog

EndDialog

BeginDialog status_update_dialog, 0, 0, 246, 195, "Status Update"
  ButtonGroup ButtonPressed
    OkButton 125, 175, 50, 15
    CancelButton 190, 175, 50, 15
  CheckBox 20, 5, 195, 15, "Sanction Imposed (Use MFIP sanction script to note.)", sanction_imposed_check
  CheckBox 20, 25, 65, 15, "Sanction Cured", sanction_cured_check
  EditBox 160, 25, 70, 15, compliance_date
  GroupBox 5, 50, 230, 50, "ES Status Change"
  EditBox 45, 70, 20, 15, hh_member
  EditBox 110, 70, 20, 15, ES_status
  EditBox 185, 70, 45, 15, effective_date
  EditBox 60, 105, 175, 15, actions_taken
  EditBox 60, 125, 175, 15, other_notes
  EditBox 75, 145, 80, 15, worker_signature
  Text 90, 30, 65, 10, "Compliance Date:"
  Text 10, 70, 30, 15, "Member:"
  Text 70, 70, 40, 20, "New ES Status:"
  Text 5, 105, 35, 15, "Notes:"
  Text 5, 125, 50, 15, "Actions Taken:"
  Text 5, 145, 60, 15, "Worker Signature:"  
  Text 145, 70, 35, 20, "Effective Date:"
EndDialog

'-grabbing case number
EMConnect ""

call MAXIS_case_number_finder(case_number)

'---------------Calling the case number dialog
DO
	DO
		DO
			Dialog case_number_dialog
			IF ButtonPressed = 0 THEN stopscript
		LOOP UNTIL ButtonPressed = OK
		IF isnumeric(case_number) = FALSE THEN MsgBox "You must enter a case number. Please try again."
	LOOP UNTIL isnumeric(case_number) = True
	IF isdate(document_date) = FALSE THEN  MsgBox "Please enter a valid document date."
LOOP UNTIL isdate(document_date) = True

'-------Employment plan dialog
IF update_type = "Employment Plan" THEN 
	DO
		DO	
			DO
				DO
					DO 
						Dialog employment_plan_dialog
						IF ButtonPressed = 0 THEN stopscript
					LOOP UNTIL ButtonPressed = OK
					IF actions_taken = "" THEN MsgBox "Please complete the actions taken field."
				LOOP UNTIL actions_taken <> ""
				IF worker_signature = "" THEN MsgBox "Please sign your case note."
			LOOP UNTIL worker_signature <> ""
			IF primary_activity = "2. Employment" and job_info = "" THEN MsgBox "You have entered the primary activity as employment but did not enter any job info.  Please complete the job information field."
		LOOP UNTIL job_info <> "" or primary_activity <> "2. Employment"
		IF primary_activity = "3. High School / GED" OR primary_activity = "4. Higher Ed" AND school_info = "" THEN MsgBox "You have entered the primary activity as education but did not complete the school information field.  Please complete the school info field or enter a different primary activity."
	LOOP UNTIL school_info <> "" OR primary_activity = "1. Employment Search" OR primary_activity = "2. Employment" OR primary_activity = "5. Health / Medical"
END IF

'Status update Dialog
IF update_type = "Status Update" THEN
	DO
		DO
			DO
				Dialog status_update_dialog
				IF ButtonPressed = 0 THEN stopscript
			LOOP UNTIL ButtonPressed = OK
			IF sanction_imposed_check = unchecked and actions_taken <> "" and received_sent = "Received" THEN MsgBox "Please complete the actions taken field."
		LOOP until sanction_imposed_check = checked OR actions_taken <> "" OR received_sent = "Sent"
		IF worker_signature = "" THEN MsgBox "Please sign your case note."
	LOOP until worker_signature <> ""
END IF

'----Writing the note
call check_for_MAXIS(true)

call start_a_blank_CASE_NOTE
'Writing the employment plan note
IF update_type = "Employment Plan" THEN 
	call write_variable_in_CASE_NOTE("***Employment Plan Received for member " & hh_member & "***")
	call write_bullet_and_variable_in_CASE_NOTE("Plan Date", document_date)
	call write_variable_in_CASE_NOTE("ES Provider: " & ES_provider & " Counselor: " & ES_counselor)
	call write_variable_in_CASE_NOTE("Primary Activity: " & primary_activity & " " & activity_hours & " hours per week.")
	IF FSS_check = checked THEN call write_variable_in_CASE_NOTE("EMPS Status: FSS")
	IF UP_check = checked THEN call write_variable_in_CASE_NOTE("EMPS Status: Universal Participation.")
	IF job_info <> "" and job_verif_check = checked THEN call write_variable_in_CASE_NOTE("Job information reported: " & job_info & " Verif on file.")
	IF job_info <> "" and job_verif_check = unchecked THEN call write_variable_in_CASE_NOTE("Job information reported: " & job_info & " NO verif on file.")
	IF school_info <> "" and school_verif_check = checked THEN call write_variable_in_CASE_NOTE("School information reported: " & school_info & " Verif on file.")
	IF school_info <> "" and school_verif_check = unchecked THEN call write_variable_in_CASE_NOTE("School information reported: " & school_info & " NO Verif on file.")
	IF disa_end_date <> "" and MOF_check = checked THEN call write_variable_in_CASE_NOTE("DISA end date: " & disa_end_date & " MOF on file.")
	IF disa_end_date <> "" and MOF_check = unchecked THEN call write_variable_in_CASE_NOTE("DISA end date: " & disa_end_date)
END IF
'Writing the status update note.
IF update_type = "Status Update" THEN
	IF received_sent = "Sent" THEN call write_variable_in_CASE_NOTE("Status update sent to ES on " & document_date)
	IF sanction_imposed_check = checked THEN call write_variable_in_CASE_NOTE("Status update to impose sanction received on: " & document_date)
	IF sanction_cured_check = checked THEN 
		call write_variable_in_CASE_NOTE("Status update to cure sanction received on: " & document_date)
		call write_variable_in_CASE_NOTE("Compliance date: " & compliance_date)
	END IF
	IF hh_member <> "" THEN 
		call write_variable_in_CASE_NOTE("Status update received to change ES status of member: " & hh_member & " on " & document_date)
		call write_variable_in_CASE_NOTE("New ES Status: " & ES_status & " Effective: " & effective_date)
	END IF
END IF
IF actions_taken <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)


script_end_procedure("")

	
