'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TRANSFER CASE.vbs"
start_time = timer
STATS_counter = 1                  	'sets the stats counter at one
STATS_manualtime = 229              'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/28/2022", "Multiple updates made ensuring that the transfer is complete and removing the case from in-county transfers.", "MiKayla Handley, Hennepin County.")
CALL changelog_update("05/21/2021", "Updated browser to default when opening SIR from Internet Explorer to Edge.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/09/2020", "No issues with SPEC/MEMO for out-of-county cases. SIR Announcement from 11/05/20 stated an issue was identified. Hennepin County's script project is seperate from DHS's script project. We are not experiencing the reported issue. Thank you!", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/20/2020", "Updated link for out-of-county use form.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/16/2020", "Added closing message box with information if a case is identified as being transferred to the MNPrairie Bank of counties.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/20/2020", "Rewrite.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
function view_poli_temp(temp_one, temp_two, temp_three, temp_four)
	call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
	EMWriteScreen "TEMP", 5, 40     'Writes TEMP

	'Writes the panel_title selection
	Call write_value_and_transmit("TABLE", 21, 71)

	If temp_one <> "" Then temp_one = right("00" & temp_one, 2)
	If len(temp_two) = 1 Then temp_two = right("00" & temp_two, 2)
	If len(temp_three) = 1 Then temp_three = right("00" & temp_three, 2)
	If len(temp_four) = 1 Then temp_four = right("00" & temp_four, 2)

	total_code = "TE" & temp_one & "." & temp_two
	If temp_three <> "" Then total_code = total_code & "." & temp_three
	If temp_four <> "" Then total_code = total_code & "." & temp_four

	EMWriteScreen total_code, 3, 21
	transmit

	EMWriteScreen "X", 6, 4
	transmit
end function
'--------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
back_to_self                                        'added to ensure we have the time to update and send the case in the background

EMReadScreen worker_number, 7, 22, 8                'reading the current workers number '
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

closing_message = "Transfer case is complete."      'setting up closing_message variable for possible additions later based on conditions
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 201, 85, "Transfer Case"
  EditBox 80, 5, 45, 15, MAXIS_case_number
  EditBox 150, 25, 45, 15, transfer_to_worker
  EditBox 80, 45, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 100, 65, 45, 15
    CancelButton 150, 65, 45, 15
    PushButton 130, 5, 65, 15, "CASE TRANSFER", search_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 70, 10, "Servicing Worker:"
  Text 5, 50, 60, 10, "Worker Signature:"
  Text 80, 30, 50, 10, "(transferring to)"
EndDialog
'Runs the first dialog - which confirms the case number

DO
    active_worker_found = TRUE
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
            Call validate_MAXIS_case_number(err_msg, "*")
            IF len(transfer_to_worker) <> 7 THEN err_msg = err_msg & vbNewLine & "* Please enter the new servicing worker."
            IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            If ButtonPressed = search_button Then               'Pulling up the hsr page if the button was pressed.
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/transfers.aspx?web=1"
                err_msg = "LOOP"
            Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
    		    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
            End If
    	Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    transfer_to_worker = trim(transfer_to_worker)                'formatting the information entered in the dialog
    transfer_to_worker = Ucase(transfer_to_worker)               'making sure we are capital in all things'
    worker_number = trim(worker_number)
    worker_number = Ucase(worker_number)

    '----------Checks that the worker or agency is valid---------- 'must find user information before transferring to account for privileged cases.
    IF transfer_to_worker = "X127CCL" THEN script_end_procedure("This case is will be transferred via an automated script after being closed for 4 months, the script will now end.")

    call navigate_to_MAXIS_screen("REPT", "USER")
    EMWriteScreen transfer_to_worker, 21, 12
    TRANSMIT

    EMReadScreen error_message, 75, 24, 2
    error_message = trim(error_message)
    EMReadScreen inactive_worker, 8, 7, 38
    IF inactive_worker = "INACTIVE" THEN
        active_worker_found = false
        error_message = "THIS WORKER DOES NOT APPEAR TO BE ACTIVE PLEASE REVIEW THE WORKER NUMBER AND TRY AGAIN."
    END IF
    IF error_message =  "NO WORKER FOUND WITH THIS ID" THEN active_worker_found = false
    IF active_worker_found = FALSE THEN MsgBox "ATTENTION - " & error_message
LOOP UNTIL active_worker_found = TRUE

transfer_out_of_county = FALSE                                          'setting varible to false'
IF left(transfer_to_worker, 4) <> "X127" THEN transfer_out_of_county = TRUE 'setting the out of county BOOLEAN'
'read the panel'
EMWriteScreen "X", 7, 3 ' navigating to read the worker information'
TRANSMIT
EMReadScreen worker_agency_name, 43, 8, 27
worker_agency_name = trim(worker_agency_name)

IF worker_agency_name = "" THEN 						'If we are unable to find the alias for the worker we will just use the worker name as it is what is used on notices anyway
	EMReadScreen worker_agency_name, 43, 7, 27
	worker_agency_name = trim(worker_agency_name)
	name_length = len(worker_agency_name)
	comma_location = InStr(worker_agency_name, ",")
	worker_agency_name = right(worker_agency_name, (name_length - comma_location)) & " " & left(worker_agency_name, (comma_location - 1)) 'this section will reorder the name of the worker since it is stored here as last, first. the comma_location - 1 removes the comma from the "last,"
END IF
EMReadScreen mail_addr_line_one, 43, 9, 27              ' really only need for out of county but read for all '
	mail_addr_line_one = trim(mail_addr_line_one)
EMReadScreen mail_addr_line_two, 43, 10, 27
	mail_addr_line_two = trim(mail_addr_line_two)
EMReadScreen mail_addr_line_three, 43, 11, 27
	mail_addr_line_three = trim(mail_addr_line_three)
EMReadScreen mail_addr_line_four, 43, 12, 27
	mail_addr_line_four = trim(mail_addr_line_four)
EMReadScreen worker_agency_phone, 14, 13, 27
EMReadScreen worker_county_code, 2, 15, 32

cash_cfr = worker_county_code ' for updating the out of county'
hc_cfr = worker_county_code
action_completed = True 'leaving REPT USER now'

CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) ' need discovery on priv cases for xfer handling'
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
'MNPrairie Bank Support - MNPrairie Bank cases all go to Steele (county code 74)'s ICT transfer.
'Agencies in the MNPrairie Bank are Dodge (county code 20), Steele (county code 74), and Waseca (county code 81)
IF servicing_worker = "X120ICT" OR servicing_worker = "X181ICT" THEN servicing_worker = "X174ICT"
IF servicing_worker = "X162ICT" THEN worker_agency_phone = "651-266-4444" 'Ramsey County has an individuals workers phone previously'

If transfer_out_of_county = False THEN      'If a transfer_to_worker was entered - we are attempting the transfer
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER IN COUNTY
	EMWriteScreen "x", 7, 16                               'transfer within county option
    TRANSMIT
    PF9                                                    'putting the transfer in edit mode
    EMreadscreen primary_worker, 7, 21, 16                  'how does PW act differently than SW?'
	EMreadscreen servicing_worker, 7, 18, 65               'checking to see if the transfer_to_worker is the same as the primary_worker (because then it won't transfer)
    EMreadscreen second_servicing_worker, 7, 18, 74        'checking to see if the transfer_to_worker is the same as the second_servicing_worker (because then it won't transfer)
    IF second_servicing_worker <> "_______" THEN CALL clear_line_of_text(18, 74)

    IF servicing_worker = transfer_to_worker THEN          'If they match, cancel the transfer and save the information about the 'failure'
        end_msg = "This case is already in the requested worker's number."
		PF10 'backout
		PF3 'SPEC menu
		PF3 'SELF Menu'
	ELSE                                                   'otherwise we are going for the transfer
	    EMWriteScreen transfer_to_worker, 18, 61            'entering the worker information
	    TRANSMIT                                           'saving - this should then take us to the transfer menu
        EMReadScreen panel_check, 4, 2, 55                 'reading to see if we made it to the right place

        If panel_check = "XWKR" THEN                       'this is not the right place
            end_msg = "Transfer of this case to " & transfer_to_worker & " has failed."
            PF10 'backout
            PF3 'SPEC menu
            PF3 'SELF Menu'
        Else                                                 'if we are in the right place - read to see if the new worker is the transfer_to_worker
            EMReadScreen primary_worker, 7, 21, 16
            If primary_worker <> transfer_to_worker THEN     'if it is not the transfer_to_worker - the transfer failed.
                end_msg = "Transfer of this case to " & transfer_to_worker & " has failed."
            End If
        End If
	END IF
ELSE 'this means out of county is TRUE '
    CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

    hc_cfr_no_change_checkbox = CHECKED
    cash_cfr_no_change_checkbox = CHECKED
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 346, 280, "Out of County Case Transfer"
      EditBox 80, 5, 45, 15, client_move_date
      DropListBox 80, 25, 45, 15, "No"+chr(9)+"Yes", excluded_time_dropdown
      EditBox 205, 25, 45, 15, excluded_date
      EditBox 80, 45, 45, 15, METS_case_number
      DropListBox 205, 45, 45, 15, "Active"+chr(9)+"Inactive"+chr(9)+"N/A", mets_status_dropdown
      EditBox 80, 65, 260, 15, transfer_reason
      EditBox 80, 85, 260, 15, action_to_be_taken
      EditBox 140, 105, 200, 15, requested_verifs
      EditBox 200, 125, 140, 15, expected_changes
      EditBox 50, 145, 290, 15, other_notes
      CheckBox 10, 175, 145, 10, "Check here if CASH CFR is not changing.", cash_cfr_no_change_checkbox
      CheckBox 20, 185, 220, 10, "Check here to manually set the CFR and change date for CASH.", manual_cfr_cash_checkbox
      EditBox 70, 195, 20, 15, cash_cfr
      EditBox 175, 195, 20, 15, cash_cfr_month
      EditBox 200, 195, 20, 15, cash_cfr_year
      CheckBox 10, 215, 150, 10, "Check here if the HC CFR is not changing.", hc_cfr_no_change_checkbox
      CheckBox 20, 225, 205, 10, "Check here to manually set the CFR and change date for HC.", manual_cfr_hc_checkbox
      EditBox 70, 235, 20, 15, hc_cfr
      EditBox 180, 235, 20, 15, hc_cfr_month
      EditBox 205, 235, 20, 15, hc_cfr_year
      ButtonGroup ButtonPressed
        OkButton 235, 260, 50, 15
        CancelButton 290, 260, 50, 15
        PushButton 135, 5, 65, 15, "SPEC/XFER", XFER_button
        PushButton 205, 5, 65, 15, "POLI/TEMP", POLI_TEMP_button
        PushButton 275, 5, 65, 15, "USEFORM", useform_xfer_button
      Text 5, 10, 60, 10, "Client Move Date"
      Text 5, 30, 55, 10, "Excluded time?"
      Text 140, 30, 40, 10, "Begin Date:"
      Text 5, 50, 75, 10, "METS Case Number:"
      Text 140, 50, 65, 10, "METS Case Status:"
      Text 5, 70, 70, 10, "Reason For Transfer:"
      Text 5, 90, 70, 10, "Actions To Be Taken:"
      Text 5, 110, 135, 10, "List All Requested/Pending Verifications:"
      Text 5, 130, 195, 10, "Note any expected changes in household's circumstances:"
      Text 5, 150, 45, 10, "Other Notes:"
      GroupBox 5, 165, 335, 90, "Current Financial Responsibility County (CFR)"
      Text 95, 240, 75, 10, "Change Date (MM YY)"
      Text 20, 200, 45, 10, "Current CFR:"
      Text 10, 285, 45, 10, "Current CFR:"
      Text 85, 285, 75, 10, "Change Date (MM YY)"
      Text 20, 240, 45, 10, "Current CFR:"
      Text 95, 200, 75, 10, "Change Date (MM YY)"
    EndDialog

	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
            IF transfer_reason = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a reason for transfer."
            IF excluded_time_dropdown = "Yes" AND isdate(excluded_date) = False THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date for the start of excluded time or double check that the client's absense is due to excluded time."
			IF isdate(client_move_date) = False THEN  err_msg = err_msg & vbNewLine & "* Please enter a valid date for client move."
			IF (ma_case = True AND excluded_time_dropdown = "No") THEN  err_msg = err_msg & vbNewLine & "* Please select whether the client is on excluded time."
			IF manual_cfr_cash_checkbox = CHECKED AND cash_cfr_no_change_checkbox = CHECKED THEN err_msg = err_msg & vbNewLine & "* Please select whether the CFR for CASH is changing or not. Review input."
			IF manual_cfr_hc_checkbox = CHECKED AND hc_cfr_no_change_checkbox = CHECKED THEN  err_msg = err_msg & vbNewLine & "* Please select whether the CFR for HC is changing or not. Review input."
			IF (mets_status_dropdown = "Active" and METS_case_number = "") then err_msg = err_msg & vbNewLine & "* Please enter a METS case number."
            IF ButtonPressed = POLI_TEMP_button THEN CALL view_poli_temp("02", "08", "095", "") 'TE02.08.095' there is no forth variable
            IF ButtonPressed = XFER_button THEN CALL MAXIS_dialog_navigation()
            IF ButtonPressed = useform_xfer_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/transfers.aspx?web=1"

            IF ButtonPressed = useform_xfer_button or ButtonPressed = XFER_button or ButtonPressed = POLI_TEMP_button THEN
                err_msg = "LOOP"
            Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
                IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
            End If
        Loop until err_msg = ""
        call back_to_self ' this is for if the worker has used the POLI/TEMP navigation'
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = False					'loops until user passwords back in

    'SENDING a SPEC/MEMO - this happens before the case note, and transfer -  we overwrite the information
    '----------Sending the Client a SPEC/MEMO notifying them of the details of the transfer----------

    Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    		'Writes the appt letter into the MEMO.
    Call write_variable_in_SPEC_MEMO("Your case has been transferred. Your new agency/worker is: " & worker_agency_name & "")
    Call write_variable_in_SPEC_MEMO("If you have any questions, or to send in requested proofs,")
    Call write_variable_in_SPEC_MEMO("please direct all communications to the agency listed.")
    Call write_variable_in_SPEC_MEMO(worker_agency_name)
    Call write_variable_in_SPEC_MEMO(mail_addr_line_one)
    Call write_variable_in_SPEC_MEMO(mail_addr_line_two)
    Call write_variable_in_SPEC_MEMO(mail_addr_line_three)
    Call write_variable_in_SPEC_MEMO(mail_addr_line_four)
    Call write_variable_in_SPEC_MEMO(worker_agency_phone)
    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
    PF4'save and exit

    '----------The case note of the reason for the XFER----------
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("~ Case transferred to " & transfer_to_worker & " ~")
    Call write_bullet_and_variable_in_CASE_NOTE("Active programs", list_active_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", list_pending_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason case was transferred", transfer_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("Actions to be taken", action_to_be_taken)
    Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verifs)
    Call write_bullet_and_variable_in_CASE_NOTE("Expected changes in household's circumstances:", expected_changes)
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    IF mets_status_dropdown = "Active" THEN Call write_bullet_and_variable_in_CASE_NOTE("* Client is active on HC through METS case number:", METS_case_number)
    call write_bullet_and_variable_in_case_note("Client move date", client_move_date)
    call write_bullet_and_variable_in_case_note("Change report sent", date) 'defaulting information '
    call write_bullet_and_variable_in_case_note("Case file sent:", date) 'defaulting information '
    IF excluded_time_dropdown = "Yes" THEN
        call write_bullet_and_variable_in_case_note("Excluded Time" , "Yes, Begins " & excluded_date)
    ELSEIF excluded_time_dropdown = "No" THEN
        call write_bullet_and_variable_in_case_note("Excluded Time", excluded_time_dropdown)
    END IF
    IF ma_case = True THEN
        CALL write_bullet_and_variable_in_case_note("HC County of Financial Responsibility", hc_cfr)
        IF hc_cfr_no_change_checkbox = UNCHECKED THEN
            CALL write_bullet_and_variable_in_case_note("HC CFR Change Date", (hc_cfr_month & "/" & hc_cfr_year))
        ELSE
            CALL write_bullet_and_variable_in_case_note("HC CFR", "Not changing")
        END IF
    END IF
    IF ga_case = TRUE or msa_case = TRUE or mfip_case = TRUE or dwp_case = TRUE or grh_case = TRUE THEN
        CALL write_bullet_and_variable_in_case_note("CASH County of Financial Responsibility", cash_cfr)
        IF manual_cfr_cash_checkbox = CHECKED THEN
            CALL write_bullet_and_variable_in_case_note("CASH CFR Change Date", (cash_cfr_month & "/" & cash_cfr_year))
        ELSE
            CALL write_bullet_and_variable_in_case_note("CASH CFR", "Not changing")
        END IF
    END IF
    CALL write_variable_in_case_note("* SPEC/MEMO sent to client with new worker information.")
    IF forms_to_arep = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to AREP.")
    IF forms_to_swkr = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to social worker.")
    Call write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
    PF3

    '----------------------------------------------------------OUT OF COUNTY TRANSFER actually happening
    'this appears to be a duplicate but the handling is different for out of county vs in county'
    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER
    EMWriteScreen "x", 9, 16                               'Transfer County To County
    TRANSMIT
    EMReadScreen panel_check, 4, 2, 54                        'reading to see if we made it to the right place
    'EMReadScreen error_message, 75, 24, 2                   ' YOU ARE NOT AUTHORIZED TO TRANSFER THIS CASE OR 'LAST XFER ACTION WAS DONE AT 09:23 ON 09/23/21 BY X127L1S'
    IF panel_check = "XCTY" Then
        PF9                                                    'putting the transfer in edit mode
        call create_MAXIS_friendly_date(client_move_date, 0, 4, 28)    'Writing client move date
        call create_MAXIS_friendly_date(client_move_date, 0, 4, 61)    'this is the CRF date we dont need to ask because we dont do this'
        EMWriteScreen left(excluded_time_dropdown, 1), 5, 28            'Writes the excluded time info. Only need the left character (it's a dropdown)
        IF excluded_time_dropdown = "Yes" THEN                          'If there's excluded time, need to write the info
            call create_MAXIS_friendly_date(excluded_date, 0, 6, 28)
            EMWriteScreen hc_cfr, 15, 39
        END IF

        IF excluded_date = "" AND excluded_time_dropdown = "No" THEN
            EMWriteScreen "__", 6, 28
            EMWriteScreen "__", 6, 31
            EMWriteScreen "__", 6, 34
        END IF

        IF ma_case = True AND manual_cfr_hc_checkbox = CHECKED THEN
            EMWriteScreen hc_cfr, 14, 39
            EMWriteScreen hc_cfr_month, 14, 53
            EMWriteScreen hc_cfr_year, 14, 59
        END IF

        IF ga_case = TRUE or msa_case = TRUE or mfip_case = TRUE or dwp_case = TRUE or grh_case = TRUE and manual_cfr_cash_checkbox = CHECKED THEN 'previously we read PROG for cash one and cash two programs unsure if this is necessary'
            EMWriteScreen cash_cfr, 11, 39
            EMWriteScreen cash_cfr_month, 11, 53
            EMWriteScreen cash_cfr_year, 11, 59
            EMReadScreen cash_cfr_two, 2, 12, 39
            IF cash_cfr_two <> "__" THEN
               EMWriteScreen cash_cfr, 12, 39
               EMWriteScreen cash_cfr_month, 12, 53
               EMWriteScreen cash_cfr_year, 12, 59
            END IF
        END IF

        EMWriteScreen worker_number, 18, 28
        EMWriteScreen transfer_to_worker, 18, 61
        TRANSMIT                                           'saving - this should then take us to the transfer menu
        EMReadScreen panel_check, 4, 2, 49                 'reading to see if we made it to the right place we shou.d be back on   Transfer Selection (XFER)
        If panel_check <> "XFER" Then 'this is not the right place
            end_msg = "Transfer of this case to " & transfer_to_worker & " has failed."
            PF10 'backout
            PF3 'SPEC menu
            PF3 'SELF Menu'
        Else                                               'if we are in the right place - read to see if the new worker is the transfer_to_worker
            EMReadScreen new_primary_worker, 7, 21, 16
            IF new_primary_worker <> transfer_to_worker Then           'if it is not the transfer_to_worker - the transfer failed.
                end_msg = "Transfer of this case to " & transfer_to_worker & " has failed."
            End If
        END IF
    END IF ' confirming the xfer worked and we left the panel'
END IF ' the big one'

IF end_msg <> "" Then
    closing_message = closing_message & vbCr & vbCr & "Case did not appear to transfer:" & vbCr & end_msg
ELSE
    closing_message = closing_message & vbCr & vbCr & "Case transfer has been completed to: " & transfer_to_worker
END IF
Call script_end_procedure_with_error_report(closing_message)


'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step----------------------------------------------------------------Date completed-------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/14/22
'--Tab orders reviewed & confirmed----------------------------------------------04/14/22
'--Mandatory fields all present & Reviewed--------------------------------------04/14/22
'--All variables in dialog match mandatory fields-------------------------------04/14/22
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/14/22
'--CASE:NOTE Header doesn't look funky------------------------------------------04/14/22
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/14/22 cant do this
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------04/14/22
'--PRIV Case handling reviewed -------------------------------------------------04/14/22
'--Out-of-County handling reviewed----------------------------------------------04/19/22
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/14/22
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/14/22
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------04/14/22
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------04/14/22 and then some
'--comment code-----------------------------------------------------------------04/14/22
'--Update Changelog for release/update------------------------------------------04/14/22
'--Remove testing message boxes-------------------------------------------------04/14/22
'--Remove testing code/unnecessary code-----------------------------------------04/14/22
'--Review/update SharePoint instructions----------------------------------------04/14/22
'--Review Best Practices using BZS page ----------------------------------------04/14/22
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/14/22
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------n/A
