'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - INCARCERATION.vbs"
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
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/12/2020", "Updated HSR Manual link for Facility List due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("01/27/2020", "Removed handling for the DAIL deletion.", "MiKayla Handley, Hennepin County")
call changelog_update("04/24/2019", "Update to run on DAIL.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'
'-------------------------------------------------------------------------------------------------------THE SCRIPT
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check <> "DAIL" THEN script_end_procedure("Your cursor is not set on a message type. Please select an appropriate DAIL message and try again.")
IF dail_check = "DAIL" THEN
	EMSendKey "T"
	TRANSMIT
	EMReadScreen DAIL_type, 4, 6, 6 'read the DAIL msg'
	DAIL_type = trim(DAIL_type)
	IF DAIL_type = "ISPI" THEN
		match_found = TRUE
	ELSE
		match_found = FALSE
		script_end_procedure("This is not an supported DAIL ISPI currently. Please select an ISPI DAIL, and run the script again.")
	END IF
	IF match_found = TRUE THEN
	    'do we need a date rcvd or save that for docs rcvd'
	    'The following reads the message in full for the end part (which tells the worker which message was selected)
	    EMReadScreen full_message, 59, 6, 20
		full_message = trim(full_message)
		'SVES PRISONER MATCH FOR SSN #xxx-xx-xxxx (last name,first initial)'
		full_message = replace(full_message, "SSN #           ", " ")	'need to replaces ssn'
		EmReadScreen MAXIS_case_number, 8, 5, 73
	    MAXIS_case_number = trim(MAXIS_case_number)

		EMReadScreen extra_info, 1, 06, 80
		IF extra_info = "+" or extra_info = "&" THEN
	        EMSendKey "X"
		    TRANSMIT
	        'THE ENTIRE MESSAGE TEXT IS DISPLAYED'
	        EmReadScreen error_msg, 37, 24, 02
		    row = 1
		    col = 1
		    EMSearch "Case Number", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.
		    'If row = 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
		    EMReadScreen first_line, 61, row + 3, col - 40 'SVES PRISONER MATCH FOR SSN # Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format.
		    	'first_line = replace(first_line, "SSN #           ", " ")	'need to replaces ssn'
		    	first_line = trim(first_line)
		    EMReadScreen second_line, 61, row + 4, col - 40
		    	second_line = trim(second_line)
		    EMReadScreen third_line, 61, row + 5, col - 40 'maxis name'
		    	third_line = trim(third_line)
		    	'third_line = replace(third_line, ",", ", ")
		    EMReadScreen fourth_line, 61, row + 6, col - 40'new hire name'
		    	fourth_line = trim(fourth_line)
		    	'fourth_line = replace(fourth_line, ",", ", ")
		    EMReadScreen fifth_line, 61, row + 7, col - 40'new hire name'
		    	fifth_line = trim(fifth_line)
			EmReadScreen client_name, 15, 9, 46
			client_name = trim(client_name)
			EmReadScreen SSN_number, 11, 9, 34
			EmReadScreen confinement_start_date, 10, 10, 22
			EmReadScreen release_date, 10, 10, 22
			release_date = trim(release_date)
			EmReadScreen facility_contact, 40, 11, 22
			facility_contact = trim(facility_contact)
			EmReadScreen facility_phone, 13, 12, 27
			TRANSMIT 'exits the additonal information'
		END IF

        'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
        Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 371, 185, "Incarceration"
		  EditBox 95, 45, 110, 15, incarceration_location
		  EditBox 95, 65, 90, 15, date_out
		  EditBox 95, 85, 90, 15, po_info
		  EditBox 95, 105, 270, 15, actions_taken
		  EditBox 95, 125, 270, 15, verifs_needed
		  EditBox 95, 145, 270, 15, other_notes
		  CheckBox 215, 70, 140, 10, "Create a TIKL to check for release date", tikl_checkbox
		  CheckBox 215, 90, 60, 10, "Reviewed ECF", ECF_reviewed
		  CheckBox 280, 90, 80, 10, "Updated STAT/FACI", update_faci_checkbox
		  ButtonGroup ButtonPressed
		    PushButton 265, 5, 100, 15, "HSR Manual - FACI", HSR_manual_button
		    PushButton 265, 25, 100, 15, "Inmate Locator", inmate_locator_button
		  DropListBox 265, 45, 100, 15, "Select One:"+chr(9)+"County Correctional Facility"+chr(9)+"Non-County Adult Correctional", faci_type
		  EditBox 95, 165, 100, 15, worker_signature
		  ButtonGroup ButtonPressed
		    OkButton 260, 165, 50, 15
		    CancelButton 315, 165, 50, 15
		  Text 215, 50, 45, 10, "Facility Type:"
		  Text 5, 70, 85, 10, "Anticipated Release Date:"
		  GroupBox 5, 5, 250, 35, "DAIL Information"
		  Text 5, 50, 75, 10, "Incarceration Location:"
		  Text 5, 110, 50, 10, "Actions Taken:"
		  Text 5, 90, 75, 10, "Probation Officer Info:"
		  Text 10, 15, 240, 20, full_message
		  Text 5, 130, 75, 10, "Verification(s) Needed:"
		  Text 5, 150, 45, 10, "Other Notes:"
		  Text 5, 170, 60, 10, "Worker signature:"
		EndDialog

        when_contact_was_made = date & ", " & time

		Do
			Do
				err_msg = ""
				Do
					Dialog Dialog1
					cancel_confirmation
					If ButtonPressed = inmate_locator_button then CreateObject("WScript.Shell").Run("https://www.bop.gov/inmateloc/")
					If ButtonPressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Facility_List.aspx")
				Loop until ButtonPressed = -1
				IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
				If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
			LOOP UNTIL err_msg = ""									'loops until all errors are resolved
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in
    	Call check_for_MAXIS(False)
    END IF

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    IF tikl_checkbox = CHECKED THEN Call create_TIKL("Check status of HH member " & hh_member & "'s incarceration at " & facility_contact & ". Incarceration Start Date was " & confinement_start_date & ".", 0, date_out, False, TIKL_note_text)

    'THE CASENOTE----------------------------------------------------------------------------------------------------
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9 'edit mode
    CALL write_variable_in_CASE_NOTE("=== " & DAIL_type & " - MESSAGE PROCESSED " & "===")
    'CALL write_variable_in_case_note("* " & full_message)
    CALL write_variable_in_case_note("SVES PRISONER MATCH FOR" & client_name)
    CALL write_variable_in_case_note(second_line)
    CALL write_variable_in_case_note(third_line)
    CALL write_variable_in_case_note(fourth_line)
    CALL write_variable_in_case_note(fifth_line)
    CALL write_variable_in_case_note("---")
    IF ECF_reviewed = CHECKED THEN CALL write_variable_in_case_note("* ECF reviewed")
    IF update_faci_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Updated STAT/FACI")
    CALL write_bullet_and_variable_in_case_note("Incarceration Location", incarceration_location)
    CALL write_bullet_and_variable_in_case_note("Anticipted Release Date", date_out)
	CALL write_bullet_and_variable_in_case_note("Facility Type", faci_type)
    CALL write_bullet_and_variable_in_case_note("Probation Information", date_out)
    CALL write_bullet_and_variable_in_case_note("Actions taken" , actions_taken)
	CALL write_bullet_and_variable_in_case_note("Action taken on", when_contact_was_made)
    CALL write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)
    CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
    IF tikl_checkbox = CHECKED THEN CALL write_variable_in_case_note("* TIKL created for anticipated release date.")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

script_end_procedure_with_error_report(DAIL_type & vbcr &  first_line & vbcr & " DAIL has been case noted")
