function claim_referral_tracking(action_taken, action_date)
'--- This function tracks the date a worker first suspects there may be a SNAP or MFIP claim. It also helps to track the discovery date and the established date of a claim. This will create or update the MISC panel and case note the referral.
'~~~~~ action_taken: 3 options exist for clearing claim referral "Sent Request for Additional Info", "Overpayment Exists", & "No Overpayment Exists"  each has different handling
'===== Keywords: MAXIS, Claim, MISC, CCOL, overpayment
    CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
    CALL Check_for_MAXIS(false)                         'Ensuring we are in a MAXIS session
    action_date = date & ""

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 221, 155, "Claim Referral Tracking"
        EditBox 65, 30, 45, 15, MAXIS_case_number
        EditBox 165, 30, 45, 15, action_date
        DropListBox 65, 50, 145, 15, "Select One:"+chr(9)+"Sent Request for Additional Info"+chr(9)+"Overpayment Exists"+chr(9)+"No Overpayment Exists", action_taken
        EditBox 65, 70, 145, 15, verif_requested
        EditBox 65, 90, 145, 15, other_notes
        EditBox 110, 110, 100, 15, worker_signature
            ButtonGroup ButtonPressed
            OkButton 115, 135, 45, 15
            CancelButton 165, 135, 45, 15
            PushButton 5, 135, 85, 15, "Claims Procedures", claims_procedures_btn
        Text 5, 5, 210, 20, "This script will only enter a STAT/MISC panel for a SNAP or MFIP federal food claim. "
        Text 15, 35, 50, 10, "Case Number: "
        Text 120, 35, 40, 10, "Action Date: "
        Text 15, 55, 45, 10, "Action Taken:"
        Text 5, 75, 55, 10, "Verif Requested:"
        Text 20, 95, 45, 10, "Other Notes:"
        Text 45, 115, 60, 10, "Worker Signature:"
    EndDialog
    DO
        DO
    	    err_msg = ""
    	    DO
                dialog Dialog1
                cancel_without_confirmation
                If ButtonPressed = claims_procedures_btn then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Claims_and_Overpayments.aspx")
            Loop until ButtonPressed = -1
    	    IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
    	    IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid action date."
    	    IF action_taken = "Select One:" then err_msg = err_msg & vbnewline & "* Please select the action taken for next step in overpayment."
            IF action_taken = "Sent Request for Additional Info" and verif_requested = "" then err_msg = err_msg & vbnewline & "* You selected that a request for additional information was sent, please advise what verifications were requested."
    	    IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
    	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'Checking for PRIV cases.
    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
    IF is_this_priv = True THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
    MAXIS_background_check      'Making sure we are out of background.

    'Grabbing case and program status information from MAXIS.
    'For tis script to work correctly, these must be correct BEFORE running the script.
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
    EMReadScreen case_status, 15, 8, 9                  'Now we are reading the CASE STATUS string from the panel - we want to make sure this does NOT read CAF1 PENDING
    EMReadScreen appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
    ega_status = "INACTIVE"
    ea_status = "INACTIVE"

    '\This functionality is how the above function reads for program information - just pulled out for these specific programs
    row = 1                                             'Looking for EGA information
    col = 1
    EMSearch "EGA", row, col
    If row <> 0 Then
        EMReadScreen ega_status, 9, row, col + 6
        ega_status = trim(ega_status)
        If ega_status = "ACTIVE" or ega_status = "APP CLOSE" or ega_status = "APP OPEN" Then ega_status = "ACTIVE"
        If ega_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
    End If

    row = 1                                             'Looking for EA information
    col = 1
    EMSearch "EA: ", row, col
    If row <> 0 Then
        EMReadScreen ea_status, 9, row, col + 5
        ea_status = trim(ea_status)
        If ea_status = "ACTIVE" or ea_status = "APP CLOSE" or ea_status = "APP OPEN" Then ea_status = "ACTIVE"
        If ea_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
    End If

    case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
    active_programs = ""        'Creates a variable that lists all the active programs on the case.
    If ga_status = "ACTIVE" Then active_programs = active_programs & "GA, "
    If msa_status = "ACTIVE" Then active_programs = active_programs & "MSA, "
    If mfip_status = "ACTIVE" Then active_programs = active_programs & "MFIP, "
    If dwp_status = "ACTIVE" Then active_programs = active_programs & "DWP, "
    If ive_status = "ACTIVE" Then active_programs = active_programs & "IV-E, "
    If grh_status = "ACTIVE" Then active_programs = active_programs & "GRH, "
    If snap_status = "ACTIVE" Then active_programs = active_programs & "SNAP, "
    If ega_status = "ACTIVE" Then active_programs = active_programs & "EGA, "
    If ea_status = "ACTIVE" Then active_programs = active_programs & "EA, "
    If cca_status = "ACTIVE" Then active_programs = active_programs & "CCA, "
    If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then active_programs = active_programs & "HC, "

    active_programs = trim(active_programs)  'trims excess spaces of active_programs
    If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

    pending_programs = ""        'Creates a variable that lists all the pending programs on the case.
    If unknown_cash_pending = True Then pending_programs = pending_programs & "Cash, "
    If ga_status = "PENDING" Then pending_programs = pending_programs & "GA, "
    If msa_status = "PENDING" Then pending_programs = pending_programs & "MSA, "
    If mfip_status = "PENDING" Then pending_programs = pending_programs & "MFIP, "
    If dwp_status = "PENDING" Then pending_programs = pending_programs & "DWP, "
    If ive_status = "PENDING" Then pending_programs = pending_programs & "IV-E, "
    If grh_status = "PENDING" Then pending_programs = pending_programs & "GRH, "
    If snap_status = "PENDING" Then pending_programs = pending_programs & "SNAP, "
    If ega_status = "PENDING" Then pending_programs = pending_programs & "EGA, "
    If ea_status = "PENDING" Then pending_programs = pending_programs & "EA, "
    If cca_status = "PENDING" Then pending_programs = pending_programs & "CCA, "
    If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then pending_programs = pending_programs & "HC, "

    pending_programs = trim(pending_programs)  'trims excess spaces of pending_programs
    If right(pending_programs, 1) = "," THEN pending_programs = left(pending_programs, len(pending_programs) - 1)

    msgbox pending_programs & vbcr & active_programs

    Call back_to_SELF

    claim_referral = False
    If (snap_status = "ACTIVE" or snap_status = "REIN") then claim_referral = True
    If (mfip_status = "ACTIVE" or mfip_status = "REIN") then claim_referral = True

    msgbox claim_referral

    IF claim_referral = False OR case_active = False then
        PROG_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "This case does not appear to have snap or cash active."  & vbNewLine & "Continue to case note only?" & vbNewLine, vbYesNo + vbQuestion, "No cash or snap programs")
        IF PROG_check = vbYes THEN case_note = True
        IF PROG_check = vbNo THEN
            case_note = False
            end_msg = end_msg & "Please review the case if cash or snap were active previously select yes and case note only."
        End if
    END IF

    IF claim_referral = True then
        case_note = True
        'Going to the MISC panel to add claim referral tracking information
        CALL navigate_to_MAXIS_screen ("STAT", "MISC")
        Row = 6
        EMReadScreen panel_number, 1, 02, 73
        IF panel_number = "0" THEN
            EMWriteScreen "NN", 20,79
            TRANSMIT
        ELSE
            DO
                'Checking to see if the MISC panel is empty, if not it will find a new line'
                EMReadScreen MISC_description, 25, row, 30
                MISC_description = replace(MISC_description, "_", "")
                IF trim(MISC_description) = "" THEN
                    EXIT DO
                ELSE
                  row = row + 1
                END IF
            LOOP UNTIL row = 17
            IF row = 17 THEN script_end_procedure("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
            'writing in the action taken and date to the MISC panel
            PF9
            '_________________________ 25 characters to write on MISC
            '???? - where does this variable get used? This might need to be another parameter in the function if it's used in the script files themselves
            IF claim_referral_tracking_dropdown = "Initial" THEN MISC_action_taken = "Claim Referral Initial"
            IF claim_referral_tracking_dropdown = "OP Non-Collectible (please specify)" THEN MISC_action_taken = "Determination-Non-Collect"
            IF claim_referral_tracking_dropdown = "No Savings/Overpayment" THEN MISC_action_taken = "Determination-No Savings"
            IF claim_referral_tracking_dropdown = "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
            EMWriteScreen MISC_action_taken, Row, 30
            EMWriteScreen date, Row, 66
            TRANSMIT '??? is this to save?

            'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
            IF action_taken = "Sent Request for Additional Info" THEN Call create_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.", 10, date, False, TIKL_note_text)
        END IF
    End if

        'The case note-------------------------------------------------------------------------------------------------
    If case_note = True then
        start_a_blank_CASE_NOTE
        Call write_variable_in_case_note("***Claim Referral Tracking-" & action_taken & "***")
        If claim_referral = True then Call write_bullet_and_variable_in_case_note("Action Date", action_date)
        Call write_bullet_and_variable_in_case_note("Pending Program(s)", pending_programs)
        CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
        IF action_taken = "Sent Request for Additional Info" THEN
            Call write_bullet_and_variable_in_case_note("Action taken", MISC_action_taken)
            Call write_variable_in_case_note(TIKL_note_text)
            Call write_bullet_and_variable_in_case_note("Verification requested", verif_requested)
        End if
        Call write_variable_in_case_note("---")
        Call write_variable_in_case_note(worker_signature)

        IF action_taken = "Sent Request for Additional Info" THEN
            end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that you sent a request for additional information. Please follow the agency's procedure(s) for claim entry once received.")
        ELSE
        	end_msg = end_msg & vbCr & ("Claim Referral Tracking - you have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
        END IF
    End if

    IF claim_referral = False then
        If case_note = TRUE THEN
            end_msg = end_msg & vbCr & "Claim Referral Tracking " & programs & " action " & action_taken 'we create some messaging to explain what happened in the script run.
        Else
            end_msg = "Claim Referral Tracking is for MFIP and SNAP cases only. Please let us know if there are further considerations needed."
        End if
    End if
    script_end_procedure_with_error_report(end_msg)
End FUNCTION
