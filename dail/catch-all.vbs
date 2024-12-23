'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - CATCH ALL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 195          'manual run time in seconds
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
'===============================================================================================END FUNCTIONS LIBRARY BLOCK

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/19/2024", "Improved script functionality and details included in CASE/NOTE.", "Mark Riegel, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("11/01/2019", "BUG FIX - resolved error where script was missing the case notes. Script should now case note every time the script is run to completion.", "Casey Love, Hennepin County")
call changelog_update("09/04/2019", "Reworded the TIKL.", "MiKayla Handley, Hennepin County")
call changelog_update("05/01/2019", "Removed the automated DAIL deletion. Workers must go back and delete manually once the DAIL has been acted on.", "MiKayla Handley, Hennepin County")
call changelog_update("04/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK

'-------------------------------------------------------------------------------------------------------THE SCRIPT
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""

EMReadScreen DAIL_type, 4, 6, 6 'read the DAIL msg'
DAIL_type = trim(DAIL_type)
If DAIL_type = "TIKL" Then EmReadScreen tikl_date, 8, 6, 11
EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)
EMReadScreen full_case_name_and_number_line, 76, 5, 5 
EMReadScreen full_dail_msg_line, 75, 6, 6

'Enters “X” on DAIL message to open full message. 
Call write_value_and_transmit("X", 6, 3)

'Read full message, including if message needs to be opened
EMReadScreen full_message_check, 36, 24, 2
If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
    EMReadScreen full_message, 61, 6, 20
    full_message = trim(full_message)
    
    'Remove X from dail message
    EMWriteScreen " ", 6, 3
Else
    ' Script reads the full DAIL message so that it can process, or not process, as needed.
    EMReadScreen full_dail_msg_line_1, 60, 9, 5
    EMReadScreen full_dail_msg_line_2, 60, 10, 5
    EMReadScreen full_dail_msg_line_3, 60, 11, 5
    EMReadScreen full_dail_msg_line_4, 60, 12, 5

    If trim(full_dail_msg_line_2) = "" Then full_dail_msg_line_1 = trim(full_dail_msg_line_1)

    full_message = trim(full_dail_msg_line_1 & full_dail_msg_line_2 & full_dail_msg_line_3 & full_dail_msg_line_4)

    'Transmit back to DAIL message
    transmit

End If

If instr(full_message, "VERIFICATIONS REQUESTED FOR THIS CASE. PLEASE REVIEW CASE") Then

    EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
    TRANSMIT

    'Search CASE/NOTEs to determine if there is a VERIFICATIONS REQUESTED CASE/NOTE

    'Using TIKL date to set too old date, no need to read for dates prior to 30 days before TIKL date
    too_old_date = DateAdd("D", -120, date)

    note_row = 5
    Do
        EMReadScreen note_date, 8, note_row, 6                  'reading the note date

        EMReadScreen note_title, 55, note_row, 25               'reading the note header
        note_title = trim(note_title)

        'VERIFICATIONS NOTES
        If left(note_title, 29) = ">>>Verifications Requested<<<" Then
            EMWriteScreen "X", note_row, 3                          'Opening the VERIF note to read the verifications
            transmit

            EMReadScreen in_correct_note, 29, 4, 3                  'making sure we are in the right note
            EMReadScreen note_list_header, 23, 4, 25

            'Here we find the right row to start reading
            If in_correct_note = ">>>Verifications Requested<<<" Then                     'making sure we're in the right note
                'Verify the due date created by the verifications needed script to confirm we have found the correct CASE NOTE. It should only be found on the first page
                row = 1
                col = 1
                EMSearch "* Verif due date:", row, col
                If row <> 0 and col <> 0 Then
                    EMReadScreen verif_due_date, 10, row, col + 18
                    verif_due_date = TRIM(verif_due_date)

                    verif_due_date_month = datepart("m", dateadd("d", 0, verif_due_date))
                    If len(var_month) = 1 then var_month = "0" & var_month
                    verif_due_date_day = datepart("d", dateadd("d", 0, verif_due_date))
                    If len(var_day) = 1 then var_day = "0" & var_day
                    verif_due_date_year = datepart("yyyy", dateadd("d", 0, verif_due_date))
                    verif_due_date = verif_due_date_month & verif_due_date_day & verif_due_date_year
                    
                    tikl_date_formatted = tikl_date
                    tikl_date_formatted_month = datepart("m", dateadd("d", 0, tikl_date_formatted))
                    If len(var_month) = 1 then var_month = "0" & var_month
                    tikl_date_formatted_day = datepart("d", dateadd("d", 0, tikl_date_formatted))
                    If len(var_day) = 1 then var_day = "0" & var_day
                    tikl_date_formatted_year = datepart("yyyy", dateadd("d", 0, tikl_date_formatted))
                    tikl_date_formatted = tikl_date_formatted_month & tikl_date_formatted_day & tikl_date_formatted_year
                    
                    If verif_due_date = tikl_date_formatted Then 
                        verifications_requested_case_note_found = True
                        Exit Do
                    Else
                        PF3     'If it is not a match, then it will PF3 out of this CASE/NOTE
                    End if
                Else
                    PF3         'If it is not a match, then it will PF3 out of this CASE/NOTE
                End If
            Else
                PF3           'this backs us out of the note if we ended up in the wrong note.
            End If
        End If

        'This is how we move through the notes and leave when we are done
		IF note_date = "        " then Exit Do
		note_row = note_row + 1
		IF note_row = 19 THEN
			PF8
			note_row = 5
		END IF
		EMReadScreen next_note_date, 8, note_row, 6
		IF next_note_date = "        " then Exit Do

    Loop until datevalue(next_note_date) < too_old_date 'looking ahead at the next case note kicking out the dates before app'

    If verifications_requested_case_note_found = True Then
    
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 281, 85, "Verifications Received Validation"
            Text 5, 5, 270, 10, "Please review the verifications requested as noted in the CASE/NOTE."
            Text 10, 25, 125, 10, "Have all verifications been received?"
            DropListBox 10, 35, 255, 15, "Select one..."+chr(9)+"Yes, delete TIKL and redirect to NOTES - DOCS RECEIVED"+chr(9)+"No, run case through background", verifications_status
            Text 5, 65, 60, 10, "Worker signature:"
            EditBox 65, 60, 95, 15, worker_signature
            ButtonGroup ButtonPressed
            OkButton 180, 60, 45, 15
            CancelButton 230, 60, 45, 15
        EndDialog

        Do
            Do
                err_msg = ""
                Dialog Dialog1
                cancel_confirmation
                If verifications_status = "Select one..." then err_msg = err_msg & vbcr & "* Please indicate whether the verifications have been received."
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
            LOOP UNTIL err_msg = ""									'loops until all errors are resolved
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        If verifications_status = "Yes, delete TIKL and redirect to NOTES - DOCS RECEIVED" Then
            PF3     'back out of CASE/NOTE
            PF3     'back to DAIL

            'Reset TIKL effective date
            tikl_date_reset = replace(tikl_date, "/", " ")
            EmWriteScreen left(tikl_date_reset, 2), 4, 67
            EmWriteScreen left(right(tikl_date_reset, 5), 2), 4, 70
            EmWriteScreen right(tikl_date_reset, 2), 4, 73
            transmit

            dail_row = 6
            Do
                EMReadScreen check_full_case_name_and_number_line, 76, dail_row - 1, 5 
                EMReadScreen check_full_dail_msg_line, 75, dail_row, 6

                If check_full_case_name_and_number_line = full_case_name_and_number_line and check_full_dail_msg_line = full_dail_msg_line Then
                    'Delete the message, match found
                    Call write_value_and_transmit("D", dail_row, 3)

                    'Handling for deleting message under someone else's x number
                    EMReadScreen other_worker_error, 25, 24, 2
                    other_worker_error = trim(other_worker_error)

                    If other_worker_error = "ALL MESSAGES WERE DELETED" or other_worker_error = "" Then
                        'Script deleted the final message in the DAIL

                        'Navigate back to SELF and add the case number
                        back_to_SELF
                        EMWriteScreen MAXIS_case_number, 18, 43

                        CALL run_from_GitHub(script_repository & "notes/documents-received.vbs")

                    ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
                        'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
                        transmit

                        'Navigate back to SELF and add the case number
                        back_to_SELF
                        EMWriteScreen MAXIS_case_number, 18, 43

                        CALL run_from_GitHub(script_repository & "notes/documents-received.vbs")

                    End If
                End If
                    
                dail_row = dail_row + 1

                'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                EMReadScreen new_case, 8, dail_row, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN 
                    'If there is NOT a new case number, the script will top the message
                    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    'If the script does find that there is a new case number (indicated by "CASE NBR"), it will end as it was unable to find the matching TIKL
                    script_end_procedure_with_error_report("The script was unable to return to the correct TIKL message. Please delete manually.")
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                'Resets the DAIL row since the message has now been topped
                dail_row = 6  
            Loop

        ElseIf verifications_status = "No, run case through background" Then
            PF3     'back out of CASE/NOTE
            PF3     'back to DAIL

            'Navigate to STAT
            Call write_value_and_transmit("S", 6, 3)

            'Run case through background
            Call write_value_and_transmit("BGTX", 20, 71)

            'Transmit past STAT/WRAP
            Transmit
            
            script_end_procedure_with_error_report("Success! The case has been run through background. Please review the case again.")
            
        End If
        
    Else
        script_end_procedure_with_error_report("Unable to find the corresponding 'Verifications Requested' CASE/NOTE. The script will now end.")
    End If        

End If

If instr(full_message, "SSN HAS NOT BEEN VERIFIED IN OVER 60 DAYS") Then

    'Dialog with links to policy references
    Dialog1 = "" 'blanking out dialog name
    BeginDialog Dialog1, 0, 0, 311, 155, "DAIL - SSN HAS NOT BEEN VERIFIED"
    ButtonGroup ButtonPressed
        PushButton 5, 45, 65, 15, "CM 12.0012.03", combined_manual_btn
        PushButton 5, 65, 65, 15, "TE 02.12.14", poli_temp_btn
        PushButton 5, 85, 65, 15, "HSR Manual", hsr_manual_btn
        PushButton 5, 105, 65, 15, "Script Instructions", script_instructions_btn
        OkButton 205, 135, 50, 15
        CancelButton 255, 135, 50, 15
    Text 5, 5, 55, 10, "DAIL Message - "
    Text 60, 5, 245, 10, full_message
    Text 5, 20, 300, 20, "This DAIL message is not currently supported by scripts. Please see the following policies/ procedures for information on how to process:"
    Text 75, 50, 95, 10, "Link to Combined Manual"
    Text 75, 70, 75, 10, "Link to POLI/TEMP"
    Text 75, 90, 85, 10, "Link to HSR Manual"
    Text 75, 110, 85, 10, "Link to Script Instructions"
    EndDialog

    DO
        Do
            err_msg = ""    'This is the error message handling
            Dialog Dialog1
            cancel_without_confirmation
            If ButtonPressed = combined_manual_btn Then
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_001203"
                err_msg = "LOOP"
            End If
            If ButtonPressed = poli_temp_btn Then
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.08.081%20DAIL%20MESSAGE%20%20%20SSN%20NOT%20VERIFIED.pdf"
                err_msg = "LOOP"
            End If
            If ButtonPressed = hsr_manual_btn Then 
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/PEPR.aspx#ssn-has-not-been-verified-in-over-60-days"
                err_msg = "LOOP"
            End If
            If ButtonPressed = script_instructions_btn Then 
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%20CATCH%20ALL.docx"
                err_msg = "LOOP"
            End If
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'End the script.
    script_end_procedure("Please follow the instructions provided in the Combined Manual, POLI/TEMP, and/or HSR Manual. The script will now end.")
End If

If instr(full_message, "GA HAS BEEN ACTV FOR 2 YEARS - REFER TO SSA IF APPROPRIATE") Then

    'Dialog with links to policy references
    Dialog1 = "" 'blanking out dialog name
    BeginDialog Dialog1, 0, 0, 311, 155, "DAIL - GA HAS BEEN ACTV FOR 2 YEARS - REFER TO SSA IF APPROPRIATE"
        ButtonGroup ButtonPressed
        PushButton 5, 45, 65, 15, "CM 12.0012.12", combined_manual_btn
        PushButton 5, 65, 65, 15, "HSR Manual", hsr_manual_btn
        PushButton 5, 85, 65, 15, "Script Instructions", script_instructions_btn
        OkButton 205, 135, 50, 15
        CancelButton 255, 135, 50, 15
        Text 5, 5, 55, 10, "DAIL Message - "
        Text 60, 5, 245, 10, full_message
        Text 5, 20, 300, 20, "This DAIL message is not currently supported by scripts. Please see the following policies/ procedures for information on how to process:"
        Text 75, 50, 95, 10, "Link to Combined Manual"
        Text 75, 70, 85, 10, "Link to HSR Manual"
        Text 75, 90, 85, 10, "Link to Script Instructions"
    EndDialog

    DO
        Do
            err_msg = ""    'This is the error message handling
            Dialog Dialog1
            cancel_without_confirmation
            If ButtonPressed = combined_manual_btn Then
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001212"
                err_msg = "LOOP"
            End If
            If ButtonPressed = hsr_manual_btn Then 
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/PEPR.aspx#ga-has-been-actv-for-2-years-%E2%80%93-refer-to-ssa-if-appropriate"
                err_msg = "LOOP"
            End If
            If ButtonPressed = script_instructions_btn Then 
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%20CATCH%20ALL.docx"
                err_msg = "LOOP"
            End If
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'End the script.
    script_end_procedure("Please follow the instructions provided in the Combined Manual and/or HSR Manual. The script will now end.")
End If

If instr(full_message, "DISA HAS ENDED - REVIEW DISA STATUS OR REDETERMINE ELIG") Then

    'Dialog with links to policy references
    Dialog1 = "" 'blanking out dialog name
    BeginDialog Dialog1, 0, 0, 311, 155, "DAIL - DISA HAS ENDED - REVIEW DISA STATUS OR REDETERMINE ELIG"
        ButtonGroup ButtonPressed
        PushButton 5, 45, 65, 15, "CM 12.0012.15", combined_manual_btn
        PushButton 5, 65, 65, 15, "Script Instructions", script_instructions_btn
        OkButton 205, 135, 50, 15
        CancelButton 255, 135, 50, 15
        Text 5, 5, 55, 10, "DAIL Message - "
        Text 60, 5, 245, 10, "full_message"
        Text 5, 20, 300, 20, "This DAIL message is not currently supported by scripts. Please see the following policies/ procedures for information on how to process:"
        Text 75, 50, 95, 10, "Link to Combined Manual"
        Text 75, 70, 85, 10, "Link to Script Instructions"
    EndDialog

    DO
        Do
            err_msg = ""    'This is the error message handling
            Dialog Dialog1
            cancel_without_confirmation
            If ButtonPressed = combined_manual_btn Then
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001215"
                err_msg = "LOOP"
            End If
            If ButtonPressed = script_instructions_btn Then 
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%20CATCH%20ALL.docx"
                err_msg = "LOOP"
            End If
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'End the script.
    script_end_procedure("Please follow the instructions provided in the Combined Manual, POLI/TEMP, and/or HSR Manual. The script will now end.")
End If

EMWriteScreen "S", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 225, "DAIL_type &   MESSAGE PROCESSED"
  GroupBox 5, 5, 270, 55, "DAIL for case #  &  MAXIS_case_number"
  Text 10, 20, 260, 35, full_message
  Text 10, 65, 45, 15, "Date Doc(s) Received:"
  EditBox 70, 65, 40, 15, docs_rcvd_date
  Text 115, 70, 50, 10, "(if applicable)"
  Text 10, 90, 55, 10, "MEMB Number:"
  EditBox 70, 85, 20, 15, memb_number
  Text 10, 110, 50, 10, "Actions taken:"
  EditBox 70, 105, 205, 15, actions_taken
  Text 10, 130, 50, 10, "Verifs needed:"
  EditBox 70, 125, 205, 15, verifs_needed
  Text 10, 150, 45, 10, "Other notes:"
  EditBox 70, 145, 205, 15, other_notes
  CheckBox 10, 165, 110, 10, "Check here if you want to TIKL", TIKL_checkbox
  CheckBox 10, 180, 90, 10, "ECF has been reviewed ", ECF_reviewed
  Text 5, 210, 60, 10, "Worker signature:"
  EditBox 65, 205, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 205, 45, 15
    CancelButton 230, 205, 45, 15
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
		If trim(docs_rcvd_date) <> "" THEN 
            If isdate(docs_rcvd_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid date that the forms were received."
        End If
		If (isnumeric(memb_number) = False and len(memb_number) > 2) then err_msg = err_msg & vbcr & "* Please Enter a valid member number."
    	If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

EMReadScreen are_we_in_stat, 14, 20, 11
EMReadScreen are_we_at_dail, 4, 2, 48
If are_we_in_stat = "Function: STAT" Then
    PF3
    EMReadScreen are_we_at_dail, 4, 2, 48
    If are_we_at_dail <> "DAIL" Then
        Call back_to_SELF
        EMWriteScreen "        ", 18, 43
        EMWriteScreen MAXIS_case_number, 18, 43
        Call navigate_to_MAXIS_screen("DAIL", "DAIL")
    End If
ElseIf are_we_at_dail <> "DAIL" Then
    Call back_to_SELF
    EMWriteScreen "        ", 18, 43
    EMWriteScreen MAXIS_case_number, 18, 43
    Call navigate_to_MAXIS_screen("DAIL", "DAIL")
End If

IF TIKL_checkbox = 1 then Call create_TIKL("Review case for requested verifications or actions needed: " & verifs_needed & ".", 10, date, False, TIKL_note_text)

Call start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("-" & DAIL_type & " PROCESSED - " & full_message & "-")
CALL write_variable_in_case_note("---")
IF ECF_reviewed = CHECKED THEN CALL write_variable_in_case_note("* Case file has been reviewed.")
If trim(docs_rcvd_date) <> "" Then CALL write_bullet_and_variable_in_case_note("Date Doc(s) Received", docs_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
IF TIKL_checkbox = CHECKED THEN CALL write_variable_in_case_note(TIKL_date_text)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report(DAIL_type & vbcr & full_message & vbcr & " DAIL has been case noted")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------