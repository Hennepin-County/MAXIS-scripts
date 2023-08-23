'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-PARTNER CALLS.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'--------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 306, 240, "Partner Calls"
  EditBox 230, 160, 65, 15, maxis_case_number
  EditBox 230, 180, 65, 15, when_contact_was_made
  EditBox 230, 200, 65, 15, worker_signature
  EditBox 80, 20, 65, 15, ESP_Name
  EditBox 80, 40, 65, 15, ESP_Phone_number
  EditBox 80, 60, 65, 15, ESP_FSS_comments
  EditBox 80, 95, 65, 15, CP_name
  EditBox 80, 115, 65, 15, CP_phone_number
  EditBox 80, 135, 65, 15, CP_Comments
  EditBox 80, 170, 65, 15, OTHER_name
  EditBox 80, 190, 65, 15, Other_phone_number
  EditBox 80, 210, 65, 15, Other_comments
  EditBox 230, 20, 65, 15, PO_name
  EditBox 230, 40, 65, 15, PO_phone_number
  EditBox 230, 60, 65, 15, PO_comments
  EditBox 230, 95, 65, 15, RRH_name
  EditBox 230, 115, 65, 15, RRH_phone_number
  EditBox 230, 135, 65, 15, RRH_comments
  ButtonGroup ButtonPressed
    OkButton 190, 220, 50, 15
    CancelButton 245, 220, 50, 15
  GroupBox 10, 160, 140, 70, "OTHER:"
  Text 15, 215, 45, 10, "Comments:"
  GroupBox 10, 10, 140, 70, "ESP/FSS:"
  GroupBox 10, 85, 140, 70, "CP:"
  Text 15, 100, 65, 10, "Organization/Name: "
  Text 165, 45, 50, 10, "Phone number: "
  Text 165, 25, 65, 10, "Organization/Name: "
  GroupBox 160, 10, 140, 70, "PO:"
  Text 165, 65, 45, 10, "Comments:"
  Text 15, 25, 65, 10, "Organization/Name: "
  Text 15, 45, 50, 10, "Phone number: "
  Text 15, 140, 45, 10, "Comments:"
  Text 165, 120, 50, 10, "Phone number: "
  Text 165, 100, 65, 10, "Organization/Name: "
  GroupBox 160, 85, 140, 70, "RRH:"
  Text 165, 140, 45, 10, "Comments:"
  Text 15, 65, 45, 10, "Comments:"
  Text 175, 165, 50, 10, "Case Number:"
  Text 15, 120, 50, 10, "Phone number: "
  Text 205, 185, 20, 10, "Date:"
  Text 15, 195, 50, 10, "Phone number: "
  Text 15, 175, 65, 10, "Organization/Name: "
  Text 165, 205, 60, 10, "Worker Signature:"
EndDialog
'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF when_contact_was_made = "" then err_msg = err_msg & vbNewLine & "* Please enter the date contact was made"
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Partner Calls ###")
CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
CALL write_bullet_and_variable_in_case_note("ESP/FSS Organization/Name", ESP_name)
CALL write_bullet_and_variable_in_case_note("ESP/FSS Phone number", ESP_Phone_number)
CALL write_bullet_and_variable_in_case_note("Comments", ESP_FSS_comments)
IF ESP_name <> "" THEN Call write_variable_in_CASE_NOTE("---")
CALL write_bullet_and_variable_in_case_note("CP Organization/Name", CP_name)
CALL write_bullet_and_variable_in_case_note("CP Phone number", CP_Phone_number)
CALL write_bullet_and_variable_in_case_note("Comments", CP_comments)
IF CP_name <> "" THEN Call write_variable_in_CASE_NOTE("---")
CALL write_bullet_and_variable_in_case_note("PO Organization/Name", PO_name)
CALL write_bullet_and_variable_in_case_note("PO Phone number", PO_Phone_number)
CALL write_bullet_and_variable_in_case_note("Comments", PO_comments)
IF PO_name <> "" THEN Call write_variable_in_CASE_NOTE("---")
CALL write_bullet_and_variable_in_case_note("Rapid Re-Housing Organization/Name", RRH_name)
CALL write_bullet_and_variable_in_case_note("RRH Phone number", RRH_Phone_number)
CALL write_bullet_and_variable_in_case_note("Comments", RRH_comments)
IF RRH_name <> "" THEN Call write_variable_in_CASE_NOTE("---")
CALL write_bullet_and_variable_in_case_note("Organization/Name", OTHER_name)
CALL write_bullet_and_variable_in_case_note("Phone number", Other_Phone_number)
CALL write_bullet_and_variable_in_case_note("Comments",  Other_comments)
IF OTHER_name <> "" THEN Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure_with_error_report("Case note has been created. Please take any additional action needed for your case.")
