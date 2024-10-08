'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IV-E-ADOPTION ASSISTANCE.vbs"
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
call changelog_update("09/09/2024", "Changed the phrasing from Cancelled to Placement Ended, removed the TIKL for the 18th Birthday. Additionally removed the phrasing around updating the Case File.", "Casey Love, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("11/25/2019", "Updated backend functionality, changlog, and removed cancel confirmation option.", "Ilse Ferris, Hennepin County")
call changelog_update("11/25/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 181, 75, "Select an Adoption Assistance option"
  EditBox 95, 10, 60, 15, MAXIS_case_number
  DropListBox 95, 30, 75, 12, "Select one..."+chr(9)+"Placement Ended"+chr(9)+"Child in placement"+chr(9)+"Closed"+chr(9)+"Opened", action_option
  ButtonGroup ButtonPressed
    OkButton 65, 50, 50, 15
    CancelButton 120, 50, 50, 15
  Text 10, 35, 80, 10, "Select the action to take:"
  Text 45, 15, 45, 10, "Case number:"
EndDialog
'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_without_confirmation

		call validate_MAXIS_case_number(err_msg, "*")
		IF action_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an Adoption Assistance option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)

If action_option = "Closed" then
    dialog1 = ""
	BeginDialog dialog1, 0, 0, 316, 85, "Adoption Assistance Closed"
		EditBox 95, 5, 215, 15, cancel_reason
		EditBox 95, 25, 75, 15, effective_date
		CheckBox 185, 30, 90, 10, "Transferred case to 4EC", transferred_checkbox
		EditBox 95, 45, 215, 15, other_notes
		EditBox 95, 65, 105, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 205, 65, 50, 15
			CancelButton 260, 65, 50, 15
		Text 45, 50, 40, 10, "Other notes:"
		Text 30, 10, 60, 10, "AA Closed reason:"
		Text 30, 70, 60, 10, "Worker signature:"
		Text 40, 30, 50, 10, "Effective date:"
	EndDialog


	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_without_confirmation
			IF cancel_reason = "" then err_msg = err_msg & vbNewLine & "* Enter the reason the placement has ended.."
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("--AA canceled effective " & effective_date & "--")
    Call write_bullet_and_variable_in_CASE_NOTE("Reason(s) Closed", cancel_reason)
    If transferred_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Transferred case to 4EC.")
END IF

If action_option = "Child in placement" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 311, 195, "Adoption Assistance child in placement"
      EditBox 70, 10, 60, 15, placement_begins
      EditBox 70, 30, 235, 15, Rule_five
      EditBox 70, 50, 235, 15, AREP
      EditBox 70, 70, 235, 15, placed
      EditBox 70, 90, 235, 15, MMIS
      EditBox 70, 110, 235, 15, results
      EditBox 70, 130, 235, 15, other_notes
      CheckBox 70, 150, 95, 10, "Transferred case to EW4", transfer_checkbox
      EditBox 70, 165, 125, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 200, 165, 50, 15
        CancelButton 255, 165, 50, 15
      Text 40, 35, 25, 10, "Rule 5:"
      Text 40, 75, 25, 10, "Placed:"
      Text 45, 55, 20, 10, "AREP:"
      Text 10, 170, 60, 10, "Worker signature:"
      Text 30, 135, 40, 10, "Other notes:"
      Text 40, 115, 30, 10, "Results:"
      Text 45, 95, 25, 10, "MMIS:"
      Text 15, 15, 55, 10, "Elig begin date:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_without_confirmation

			If isDate(placement_begins) = False then err_msg = err_msg & vbNewLine & "* Enter a valid eligibility date."
			If Rule_five = "" then err_msg = err_msg & vbNewLine & "* Enter the Rule 5 information."
			If placed = "" then err_msg = err_msg & vbNewLine & "* Enter the placement information."
			If MMIS = "" then err_msg = err_msg & vbNewLine & "* Enter the MMIS information."
			If results = "" then err_msg = err_msg & vbNewLine & "* Enter the results information."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("--AA child in placement effective " & placement_begins & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("Rule 5", Rule_five)
	Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
	Call write_bullet_and_variable_in_CASE_NOTE("Placed", placed)
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS", MMIS)
	Call write_bullet_and_variable_in_CASE_NOTE("Results", Results)
    Call write_variable_in_CASE_NOTE("* When placement ends, transfer case to FG1, update AREP, ADDR and case file.")
	If transfer_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Transferred case to EW4.")
END IF

If action_option = "Placement Ended" then
    dialog1 = ""
	BeginDialog dialog1, 0, 0, 291, 120, "Adoption Assistance Placement Ended"
		EditBox 115, 5, 65, 15, placement_ended
		EditBox 70, 25, 215, 15, other_notes
		CheckBox 70, 60, 90, 10, "Transferred case to FG1", transferred_checkbox
		CheckBox 180, 60, 70, 10, "Deleted FC panels", deleted_FC_checkbox
		CheckBox 70, 75, 95, 10, "Updated AREP and ADDR", deleted_panels_checkbox
		EditBox 70, 100, 105, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 180, 100, 50, 15
			CancelButton 235, 100, 50, 15
		Text 25, 30, 40, 10, "Other notes:"
		Text 5, 10, 110, 10, "AREP reported placement ended:"
		Text 5, 105, 60, 10, "Worker signature:"
		GroupBox 5, 45, 280, 50, "Check each action that was completed:"
		Text 185, 10, 50, 10, "(date)"
	EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_without_confirmation

			If isDate(placement_ended) = False then err_msg = err_msg & vbNewLine & "* Enter a valid placement ending date."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("--AA Placement Ended effective " & placement_ended & "--")
	If transferred_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Transferred case to FG1.")
	If deleted_FC_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Deleted FC panels.")
	If deleted_panels_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Updated AREP and ADDR panels.")
END IF

If action_option = "Opened" then
    dialog1 = ""
	BeginDialog dialog1, 0, 0, 316, 205, "Adoption Assistance opened"
		EditBox 80, 10, 60, 15, app_date
		EditBox 250, 10, 60, 15, case_closed
		EditBox 80, 30, 60, 15, finalization_date
		EditBox 250, 30, 60, 15, effective_date
		EditBox 80, 50, 230, 15, AREP
		EditBox 80, 70, 230, 15, insa
		EditBox 80, 90, 230, 15, spec_adpt
		DropListBox 80, 110, 60, 15, "Select one..."+chr(9)+"No"+chr(9)+"Yes", PMI_merge
		DropListBox 250, 110, 60, 15, "Select one..."+chr(9)+"No"+chr(9)+"Yes", priv_request
		EditBox 80, 135, 60, 15, AA_monies
		CheckBox 150, 130, 160, 10, "Request sent to verify Social Security Number", SSN_verif_checkbox
		CheckBox 150, 140, 160, 10, "Request sent to verify other health insurance", OHI_checkbox
		EditBox 80, 160, 230, 15, other_notes
		EditBox 80, 185, 120, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 205, 185, 50, 15
			CancelButton 260, 185, 50, 15
		Text 20, 35, 60, 10, "Finalization date:"
		Text 50, 55, 25, 10, "AREP:"
		Text 5, 140, 70, 10, "AA rec'd per month: $"
		Text 20, 15, 60, 10, "Application date:"
		Text 40, 165, 40, 10, "Other notes:"
		Text 155, 15, 90, 10, "IV-E/Non IV-E case closed:"
		Text 160, 115, 85, 10, "PRIV/Block request done: "
		Text 5, 95, 75, 10, "SPEC/ADPT function:"
		Text 55, 75, 20, 10, "INSA:"
		Text 40, 115, 40, 10, "PMI merge:"
		Text 20, 190, 60, 10, "Worker signature:"
		Text 160, 35, 85, 10, "09-X/10-X App'd effective:"
	EndDialog

	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_without_confirmation

			If isDate(app_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid application date."
			If isDate(case_closed) = False then err_msg = err_msg & vbNewLine & "* Enter a valid date the IV-E/Non IV-E case closed."
			If isDate(finalization_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid finalization date."
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid 09-X/10-X effective date."
			If spec_adpt = "" then err_msg = err_msg & vbNewLine & "* Enter the SPEC/ADPT function information."
			If PMI_merge = "Select one..." then err_msg = err_msg & vbNewLine & "* Has there a PMI merge?"
			If priv_request = "Select one..." then err_msg = err_msg & vbNewLine & "* Has a priv/block request been done?"
			If IsNumeric(AA_monies) = False then err_msg = err_msg & vbNewLine & "* Enter a valid AA amount received."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("**AA opened effective " & effective_date & "**")
	Call write_bullet_and_variable_in_CASE_NOTE("Application date", app_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Finalization date", finalization_date)
	Call write_bullet_and_variable_in_CASE_NOTE("IV-E/Non IV-E case closed", case_closed)
    Call write_bullet_and_variable_in_CASE_NOTE("09-X/10-X effective date", effective_date)
	Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
	Call write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
	Call write_bullet_and_variable_in_CASE_NOTE("SPEC/ADPT function", spec_adpt)
	Call write_bullet_and_variable_in_CASE_NOTE("PMI merge", PMI_merge)
    call write_bullet_and_variable_in_CASE_NOTE("PRIV/blocked request", priv_request)
	Call write_bullet_and_variable_in_CASE_NOTE("Amt of AA monies rec'd per month", AA_monies)
    If SSN_verif_checkbox = 1 then call write_variable_in_CASE_NOTE("* Sent verification request to verify client's SSN.")
	If OHI_checkbox = 1 then call write_variable_in_CASE_NOTE("* Sent verification request to verify other health insurance.")
END IF

Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")
