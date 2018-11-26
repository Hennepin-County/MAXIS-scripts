'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'For agencies using the access es database, we need to collect the ES providers before dialog
IF collecting_ES_statistics = TRUE THEN
'Collecting ES agencies from database
		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3
		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")
		'Opening DB
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
		'This looks for an existing case number and edits it if needed
		set rs = objConnection.Execute("SELECT SiteName FROM ESSitesTbl")' Grabbing all ES agency site names
		Dim ES_agencies
		ES_agencies = rs.GetRows()
	objConnection.Close
	set rs = nothing

	call convert_array_to_droplist_items (ES_agencies, ES_agency_list) 'this function turns the names into a droplist format'
END IF

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog agency_dropdown_case_number_dialog, 0, 0, 196, 160, "Status Update / Employment Plan"
  ButtonGroup ButtonPressed
    OkButton 65, 130, 50, 15
    CancelButton 125, 130, 50, 15
  EditBox 110, 10, 65, 15, MAXIS_case_number
  DropListBox 15, 35, 85, 15, "Status Update"+chr(9)+"Employment Plan", update_type
  DropListBox 115, 35, 60, 15, "Received"+chr(9)+"Sent", received_sent
  DropListBox 75, 60, 100, 15, ES_agency_list, agency
  EditBox 110, 10, 65, 15, MAXIS_case_number
  EditBox 110, 85, 65, 15, document_date
  Text 20, 85, 65, 15, "Document Date:"
  Text 20, 60, 55, 15, "Agency:"
EndDialog

BeginDialog case_number_dialog, 0, 0, 196, 130, "Status Update / Employment Plan"
  ButtonGroup ButtonPressed
    OkButton 55, 95, 50, 15
    CancelButton 120, 95, 50, 15
  DropListBox 15, 40, 85, 15, "Status Update"+chr(9)+"Employment Plan", update_type
  DropListBox 110, 40, 60, 15, "Received"+chr(9)+"Sent", received_sent
  EditBox 110, 10, 65, 15, MAXIS_case_number
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
  EditBox 60, 105, 175, 15, other_notes
  EditBox 60, 125, 175, 15, actions_taken
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

call MAXIS_case_number_finder(MAXIS_case_number)

'---------------Calling the case number dialog
DO
	DO
		DO
			IF collecting_ES_statistics = TRUE THEN
				Dialog agency_dropdown_case_number_dialog
			ELSE
				Dialog case_number_dialog
			END IF
			IF ButtonPressed = 0 THEN stopscript
		LOOP UNTIL ButtonPressed = OK
		IF isnumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must enter a case number. Please try again."
	LOOP UNTIL isnumeric(MAXIS_case_number) = True
	IF isdate(document_date) = FALSE THEN  MsgBox "Please enter a valid document date."
LOOP UNTIL isdate(document_date) = True

IF collecting_ES_statistics = TRUE THEN
	'Getting the counselor list based on chosen sitename
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
		'looking up the provider value from provider table due to field problems
		set counselor_rs = objConnection.Execute("SELECT CounselorName FROM ESCounselorsTbl WHERE CounselorSiteText = '" & agency & "'")'Grabbing all counselors that match site name
		Dim counselor_list
		counselor_list = counselor_rs.GetRows() 'converts counselors to an array
		objConnection.Close
		set counselor_rs = nothing
		call convert_array_to_droplist_items(counselor_list, ES_counselors)
		'Creating the dialog here so it will populate with counselor names correctly
		BeginDialog employment_plan_dialog, 0, 0, 311, 265, "Employment Plan Received"
		   ButtonGroup ButtonPressed
		    OkButton 185, 245, 50, 15
		    CancelButton 245, 245, 50, 15
		  DropListBox 65, 10, 45, 15, "DWP"+chr(9)+"MFIP", program_list
		  EditBox 210, 10, 30, 15, hh_member
		  Text 65, 35, 75, 15, agency
		  DropListBox 210, 35, 85, 15, ES_counselors, ES_counselor
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
END IF

''------Employment plan dialog
IF update_type = "Employment Plan" THEN
	DO
	Dialog employment_plan_dialog
		err_msg = ""
		IF ButtonPressed = 0 THEN stopscript
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "Please complete the actions taken field."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
		IF primary_activity = "2. Employment" and job_info = "" THEN err_msg = err_msg & vbCr & "You have entered the primary activity as employment but did not enter any job info.  Please complete the job information field."
		IF primary_activity = "3. High School / GED" AND school_info = "" THEN err_msg = err_msg & vbCr & "You have entered the primary activity as education but did not complete the school information field."
		IF primary_activity = "4. Higher Ed" AND school_info = "" THEN err_msg = err_msg & vbCr & "You have entered the primary activity as education but did not complete the school information field."
		IF err_msg <> "" THEN msgbox "The following errors must be resolved before continuing: " & err_msg
	LOOP UNTIL err_msg = ""
END IF

'Status update Dialog
IF update_type = "Status Update" THEN
	DO
		err_msg = ""
		Dialog status_update_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF sanction_imposed_check = unchecked and actions_taken = "" THEN err_msg = err_msg & vbCr & "Please indicate what actions were taken."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
		IF err_msg <> "" THEN msgbox "Please resolve the following errors to continue:" & err_msg
	LOOP until err_msg = ""
END IF


'----Writing the note
call check_for_MAXIS(False)

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
	IF received_sent = "Received" and sanction_cured_check = unchecked and sanction_imposed_check = unchecked THEN call write_variable_in_CASE_NOTE("Status update received on " & document_date)
END IF
IF actions_taken <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

IF collecting_ES_statistics = True THEN
	'Updating the database
	call write_MAXIS_info_to_ES_database(MAXIS_case_number, hh_member, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, disa_end_date, primary_activity, ESDate, agency, ES_Counselor, ES_active, insert_string)
END IF

script_end_procedure("")
