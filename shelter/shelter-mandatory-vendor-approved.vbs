'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-MANDATORY VENDOR APPROVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
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
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("03/01/2020", "Updated TIKL functionality and updated functionality to create SPEC/MEMO.", "Ilse Ferris")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/01/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'--------------------------------------------------------------------------------------------------SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 296, 120, "Mandatory vendors approved"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  DropListBox 205, 5, 85, 15, "Select one..."+chr(9)+"Using ACF"+chr(9)+"Using County shelters", vendor_reason
  CheckBox 10, 25, 280, 10, "Referred to Luther Social Services for budgeting classes at 1 (888) 577-2227.", client_referred
  CheckBox 10, 50, 140, 10, "Send mandatory vendor MEMO to client.", send_MEMO_checkbox
  CheckBox 10, 60, 215, 10, "Set TIKL for 11 months to re-evaluate mandatory vendor status.", Set_TIKL
  EditBox 55, 80, 135, 15, other_notes
  EditBox 65, 100, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 100, 50, 15
    CancelButton 240, 100, 50, 15
  Text 5, 85, 40, 10, "Other notes: "
  Text 115, 10, 90, 10, "Mandatory vendor reason:"
  GroupBox 5, 40, 285, 35, "Actions:"
  Text 5, 105, 60, 10, "Worker signature:"
  Text 5, 10, 45, 10, "Case number:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF vendor_reason = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a vendor reason."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

IF send_MEMO_checkbox = 1 then
    Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("Call 1 (888) 577-2227 to get more information or enroll. You will be placed on mandatory vendor because you have used shelter or have requested assistance for housing issues.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("You will remain on mandatory vendor for 12 months. If you move, or your rent changes you must let your team know at least 15 days before the end of the month to make this change.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("Call your Human Service Representative Team at the end of this 12 month period if you want them to stop vendoring your rent at that time. Budgeting classes are free to you and available through the Lutheran Social Services. If you have any questions call the Shelter Team at (612)-348-9410.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("Sincerely, Shelter Team")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4
END IF

'----------------------------------------------------------------------------------------------------Creates a TIKL for the revoucher date
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If set_TIKL = 1 then Call create_TIKL("Mandatory vendor status started 11 months ago. Please re-evaluate vendor status.", 335, date, False, TIKL_note_text)

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### MANDATORY VENDORS APPROVED FOR 12 MONTHS ###")
Call write_variable_in_CASE_NOTE("* Mandatory vendor approved by Program Manager. Removal of vendor requires Program Manager approval.")
Call write_bullet_and_variable_in_CASE_NOTE("Client put on mandatory vendor due to", vendor_reason)
If client_referred = 1 then call write_variable_in_CASE_NOTE("* Referred client to Lutheran Social Services for budgeting classes at: 1(888) 577-2227. ")
If send_MEMO_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Sent SPEC/MEMO to client re: mandatory vendoring for the next 12 months.")
If set_TIKL = 1 then call write_variable_in_CASE_NOTE("* Set TIKL for 11 months to re-evaluate mandatory vendor status.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
