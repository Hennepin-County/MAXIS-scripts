'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SELECT WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("03/13/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function Create_List_Of_Notices
	Erase NOTICES_ARRAY
	no_notices = FALSE
	If notice_panel = "WCOM" Then
		wcom_row = 7
		array_counter = 0
		Do
			ReDim Preserve NOTICES_ARRAY(3, array_counter)
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE

			NOTICES_ARRAY(selected,    array_counter) = unchecked
			NOTICES_ARRAY(information, array_counter) = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
			NOTICES_ARRAY(MAXIS_row,   array_counter) = wcom_row

			array_counter = array_counter + 1
			wcom_row = wcom_row + 1

			EMReadScreen next_notice, 4, wcom_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If

	If notice_panel = "MEMO" Then
		memo_row = 7
		array_counter = 0
		Do
			ReDim Preserve NOTICES_ARRAY(3, array_counter)
			EMReadScreen notice_date, 8,  memo_row, 19
			EMReadScreen notice_info, 31, memo_row, 29
			EMReadScreen notice_stat, 8,  memo_row, 67

			notice_date = trim(notice_date)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE

			NOTICES_ARRAY(selected,    array_counter) = unchecked
			NOTICES_ARRAY(information, array_counter) = notice_info & " - " & notice_date & " - Status: " & notice_stat
			NOTICES_ARRAY(MAXIS_row,   array_counter) = memo_row

			array_counter = array_counter + 1
			memo_row = memo_row + 1

			EMReadScreen next_notice, 4, memo_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If
End Function


EMConnect ""

Dim NOTICES_ARRAY()
ReDim NOTICES_ARRAY(3,0)

Const selected = 0
Const information = 1
Const MAXIS_row = 2

Call check_for_MAXIS(False)

'Finds MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

EMReadScreen which_panel, 4, 2, 47
If which_panel <> "WCOM" then
    If MAXIS_case_number <> "" Then
        Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	    notice_panel = "WCOM"
	    at_notices = True
    Else
        at_notices = FALSE
    End If
Else
    at_notices = TRUE
    notice_panel = "WCOM"
End If


If at_notices = True then

	EMReadScreen MAXIS_footer_month, 2, 3, 46
	EMReadScreen MAXIS_footer_year,  2, 3, 51

	Create_List_Of_Notices

End If


Do
	err_msg = ""

    If NOTICES_ARRAY(0, 0) <> "" Then
        For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)
            EMReadScreen desc, 20, NOTICES_ARRAY(MAXIS_row, notices_listed), 30
            if desc = "ELIG Approval Notice" Then
                EMReadScreen print_status, 7, NOTICES_ARRAY(MAXIS_row, notices_listed), 71
                If print_status = "Waiting" Then NOTICES_ARRAY(selected, notices_listed) = checked
            End If
        Next
    End If

	dlg_y_pos = 65
	dlg_length = 125 + (UBound(NOTICES_ARRAY, 2) * 20)

	BeginDialog find_notices_dialog, 0, 0, 205, dlg_length, "Notices to add WCOM"
	  Text 5, 10, 50, 10, "Case Number"
	  EditBox 65, 5, 50, 15, MAXIS_case_number
	  Text 5, 30, 120, 10, "In which month was the notice sent?"
	  EditBox 140, 25, 20, 15, MAXIS_footer_month
	  EditBox 165, 25, 20, 15, MAXIS_footer_year
	  ButtonGroup ButtonPressed
	    PushButton 60, 50, 50, 10, "Find Notices", find_notices_button
	  If no_notices = FALSE Then
		  For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)
		  	CheckBox 10, dlg_y_pos, 185, 10, NOTICES_ARRAY(information, notices_listed), NOTICES_ARRAY(selected, notices_listed)
			dlg_y_pos = dlg_y_pos + 15
		  Next
	  Else
	  	  Text 10, dlg_y_pos, 185, 10, "**No Notices could be found here.**"
		  dlg_y_pos = dlg_y_pos + 15
	  End If
	  dlg_y_pos = dlg_y_pos + 5
	  EditBox 75, dlg_y_pos, 125, 15, worker_signature
	  dlg_y_pos = dlg_y_pos + 5
	  Text 5, dlg_y_pos, 60, 10, "Worker Signature:"
	  dlg_y_pos = dlg_y_pos + 15
	  ButtonGroup ButtonPressed
	    OkButton 100, dlg_y_pos, 50, 15
	    CancelButton 150, dlg_y_pos, 50, 15
	EndDialog

	Dialog find_notices_dialog
	cancel_confirmation

	notice_selected = FALSE
	For notice_to_print = 0 to UBound(NOTICES_ARRAY, 2)
		If NOTICES_ARRAY(selected, notice_to_print) = checked Then notice_selected = TRUE
	Next

	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
	If notice_panel = "Select One..." Then err_msg = err_msg & vbNewLine & "- Select where the notice to print is."
	If MAXIS_footer_month = "" or MAXIS_footer_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice to be copied to a Word Document."

	If ButtonPressed = find_notices_button then
		If notice_panel <> "Select One..." AND MAXIS_case_number <> "" AND MAXIS_footer_month <> "" AND MAXIS_footer_year <> "" Then
			Call navigate_to_MAXIS_screen ("SPEC", notice_panel)
			EMWriteScreen MAXIS_footer_month, 3, 46
			EMWriteScreen MAXIS_footer_year, 3, 51

			transmit
			Create_List_Of_Notices
			err_msg = "LOOP"
		Else
			err_msg = err_msg & vbNewLine & "!!! Cannot read a list of notices without a case number entered, and footer month & year entered !!!"
		End If
	End If

	If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg

Loop Until err_msg = ""

Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

EMWriteScreen MAXIS_footer_month, 3, 46
EMWriteScreen MAXIS_footer_year, 3, 51

transmit

SNAP_notice = FALSE
MFIP_notice = FALSE
GA_notice = FALSE
MSA_notice = FALSE

For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)
    If NOTICES_ARRAY(selected, notices_listed) = checked Then
        EMReadScreen notice_prog, 3, NOTICES_ARRAY(MAXIS_row, notices_listed), 25
        notice_prog = trim(notice_prog)
        If notice_prog = "FS" Then SNAP_notice = TRUE
        If notice_prog = "MF" Then MFIP_notice = TRUE
        If notice_prog = "GA" Then GA_notice = TRUE
        If notice_prog = "MS" Then MSA_notice = TRUE
    End If
Next

'DIALOG to select the WCOM to add
BeginDialog wcom_selection_dlg, 0, 0, 251, 240, "Select WCOM"
  CheckBox 15, 25, 190, 10, "WCOM for SNAP Duplicate Assistance in another state", duplicate_assistance_wcom_checkbox
  Text 10, 10, 95, 10, "Check the WCOM needed."
  CheckBox 15, 40, 145, 10, "WCOM for closing due to Returned Mail", returned_mail_wcom_checkbox
  CheckBox 15, 55, 155, 10, "WCOM for SNAP closed via PACT due to FPI", pact_fraud_wcom_checkbox
  CheckBox 15, 70, 130, 10, "WCOM for Temp disabled ABAWDs", temp_disa_abawd_wcom_checkbox
  CheckBox 15, 85, 85, 10, "WCOM for Client Death", client_death_wcom_checkbox
  CheckBox 15, 100, 125, 10, "WCOM for MFIP to SNAP transition", mfip_to_snap_wcom_checkbox
  CheckBox 15, 115, 215, 10, "WCOM for ABAWD with postponed WREG verifs for EXP SNAP", wreg_postponed_verif_wcom_checkbox
  CheckBox 15, 130, 160, 10, "WCOM for possible Banked Months available", banked_mos_avail_wcom_checkbox
  CheckBox 15, 145, 180, 10, "WCOM for Banked Months closing due to non-coop", banked_mos_non_coop_wcom_checkbox
  CheckBox 15, 160, 235, 10, "WCOM for Banked Months closing due to all available months used.", banked_mos_used_wcom_checkbox
  CheckBox 15, 175, 235, 10, "WCOM for ABAWD WREG coded for Child under 18", abawd_child_coded_wcom_checkbox
  CheckBox 15, 190, 205, 10, "WCOM for Failure to comply FSET - Good Cause Information", fset_fail_to_comply_wcom_checkbox
  CheckBox 15, 205, 150, 10, "WCOM for SNAP closed/denied with PACT", snap_pact_wcom_checkbox
  ButtonGroup ButtonPressed
    OkButton 140, 220, 50, 15
    CancelButton 195, 220, 50, 15
EndDialog

Dialog wcom_selection_dlg
cancel_confirmation

'TODO create logic to identify if the verbiage is too long (for instance if mutliple messages are selected) the script will add a MEMO to the case to give all the verbiage required.
'SPEC/WCOM  allows for 60 characters in each line
'           and 15 lines
'MEMOs have the same amount of space on a page - but have 1 additional page' Case 266334 has examples in SPEC
'Need a Function
end_of_wcom_line = 0
end_of_wcom_row = 1

end_of_memo_line = 0
end_of_memo_row = 1

using_memo = FALSE

Dim array_of_msg_lines ()
ReDim array_of_msg_lines(0)

'Use ; to denote a new line
Function add_words_to_message(message_to_add)

    If trim(message_to_add) <> "" Then
        message_array = split(message_to_add, " ")

        If using_memo = FALSE Then
            end_of_line = end_of_wcom_line
        Else
            end_of_line = end_of_memo_line
        End If

        'ERASE array_of_msg_lines
        Dim array_of_msg_lines ()
        ReDim array_of_msg_lines(0)

        message_line = ""
        lines_in_msg = 0

        For each word in message_array
            'MsgBox lines_in_msg
            If len(word) + len(message_line) > 59 Then
                ReDim Preserve array_of_msg_lines(lines_in_msg)
                array_of_msg_lines(lines_in_msg) = message_line
                lines_in_msg = lines_in_msg + 1

                message_line = ""
            End If

            message_line = message_line & replace(word, ";", "") & " "

            IF right(word, 1) = ";" Then
                ReDim Preserve array_of_msg_lines(lines_in_msg)
                array_of_msg_lines(lines_in_msg) = message_line
                lines_in_msg = lines_in_msg + 1

                message_line = ""
            End If
        Next

        ReDim Preserve array_of_msg_lines(lines_in_msg)
        array_of_msg_lines(lines_in_msg) = message_line
        lines_in_msg = lines_in_msg + 1
        lines_in_msg = lines_in_msg + 1

        MsgBox "End of WCOM Row: " & end_of_wcom_row & vbNewLine & "Lines Used:" & lines_in_msg
        if end_of_wcom_row + lines_in_msg > 15 then using_memo = TRUE

        If using_memo = FALSE Then      'Now add it to the full array with a for next using the small array'

            If UBound(WCOM_TO_WRITE_ARRAY) = 0 Then
                notice_line = 0
            Else
                notice_line = UBound(WCOM_TO_WRITE_ARRAY) + 1
                ReDim Preserve WCOM_TO_WRITE_ARRAY(notice_line)
                WCOM_TO_WRITE_ARRAY(notice_line) = "-      - - - - - - - - - - - - - - - - - - - -       -"
                notice_line = notice_line + 1
            End If

            For each entry in array_of_msg_lines
            MsgBox entry
                ReDim Preserve WCOM_TO_WRITE_ARRAY(notice_line)
                WCOM_TO_WRITE_ARRAY(notice_line) = entry
                notice_line = notice_line + 1
            Next

            end_of_wcom_row = end_of_wcom_row + lines_in_msg
        Else
            If UBound(MEMO_TO_WRITE_ARRAY) = 0 Then
                notice_line = 0
            Else
                notice_line = UBound(MEMO_TO_WRITE_ARRAY) + 1
                ReDim Preserve MEMO_TO_WRITE_ARRAY(notice_line)
                MEMO_TO_WRITE_ARRAY(notice_line) = "-      - - - - - - - - - - - - - - - - - - - -       -"
                notice_line = notice_line + 1
            End If

            For each entry in array_of_msg_lines
            MsgBox entry
                ReDim Preserve MEMO_TO_WRITE_ARRAY(notice_line)
                MEMO_TO_WRITE_ARRAY(notice_line) = entry
                notice_line = notice_line + 1
            Next
        End If

        end_of_memo_row = end_of_memo_row + lines_in_msg
    Else

    End If
    'Split all the words in the message
    'If at a ; then go to new line

    'add each line to the right array
End Function

'split the wording of each of the wcoms into words, then for each word add the length to the line
'once the length exceeds 60 then go to the next row and contine.
'At the end of each message, add a blank row.
'Add each line of the message to the WCOM array.
'Once the end of the WCOM is reached then start adding to the MEMO array

Dim WCOM_TO_WRITE_ARRAY ()
Dim MEMO_TO_WRITE_ARRAY ()

ReDim WCOM_TO_WRITE_ARRAY (0)
ReDim MEMO_TO_WRITE_ARRAY (0)

'After all of the messages are assessed then when it is written - use a for-next to write each line (which is one element in the array)
'into the wcom or memo.

If duplicate_assistance_wcom_checkbox = checked Then
    Call navigate_to_MAXIS_screen ("STAT", "MEMI")
    EMReadScreen dup_state, 2, 14, 78

    If dup_state = "__" Then dup_state = ""
    dup_month = MAXIS_footer_month
    dup_year = MAXIS_footer_year

    BeginDialog wcom_details_dlg, 0, 0, 121, 90, "WCOM Details"
      EditBox 70, 20, 25, 15, dup_state
      EditBox 70, 40, 15, 15, dup_month
      EditBox 90, 40, 15, 15, dup_year
      ButtonGroup ButtonPressed
        OkButton 60, 65, 50, 15
      Text 5, 10, 110, 10, "Duplicate SNAP in another state"
      Text 5, 25, 50, 10, "Previous State:"
      Text 5, 45, 60, 10, "In (Month/Year)"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If dup_state = "" Then err_msg = err_msg & vbNewLine & "* Enter the state in which client already received SNAP."
        If dup_month = "" or dup_year = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year for which SNAP in MN is being denied due to receipt of benefits in another state."

        If err_msg <> "" Then MsgBox "Please resolve before continuing:" & vbNewLine & err_msg
    Loop until err_msg = ""

    CALL add_words_to_message("You received SNAP benefits from the state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & ". You cannot recceive SNAP benefits from two states at the same time.")
End If

If returned_mail_wcom_checkbox = checked Then

    BeginDialog wcom_details_dlg, 0, 0, 126, 85, "WCOM Details"
      EditBox 75, 20, 45, 15, rm_sent_date
      EditBox 75, 40, 45, 15, rm_due_date
      ButtonGroup ButtonPressed
        OkButton 60, 65, 50, 15
      Text 5, 5, 110, 10, "Returned Mail"
      Text 5, 25, 65, 10, "Verif Request Sent:"
      Text 5, 45, 65, 10, "Verif Request Due:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg
        if isdate(rm_sent_date) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the request for address information was sent."
        if isdate(rm_due_date) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the response for address information was due."

        if err_msg <> "" Then msgBox "Resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""

    CALL add_words_to_message("Your mail has been returned to our agency. On " & rm_sent_date & " you were sent a Request for you to contact this agency because of this returned mail. You did not contact this agency by " & rm_due_date & " so your SNAP case has been closed.")
End If

If pact_fraud_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 281, 85, "WCOM Details"
      EditBox 75, 20, 45, 15, new_hh_memb
      EditBox 215, 20, 60, 15, SNAP_close_date
      EditBox 75, 40, 200, 15, new_memb_verifs
      ButtonGroup ButtonPressed
        OkButton 225, 65, 50, 15
      Text 5, 5, 120, 10, "New HH Member Information Failed"
      Text 5, 25, 65, 10, "New person in HH:"
      Text 140, 25, 70, 10, "SNAP close eff date:"
      Text 5, 45, 65, 10, "Verifs requested:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If new_hh_memb = "" Then err_msg = err_msg & vbNewLine & "*Enter the name of the person who has joined the household."
        If isdate(SNAP_close_date) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date on which SNAP will close."
        If new_memb_verifs = "" Then err_msg = err_msg & vbNewLine & "*Enter the verifications that were needed to add this person to the case." & vbNewLine & "If no verifications are required - this is not the correct WCOM to use."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""

    CALL add_words_to_message("This agency received a request to add " & new_hh_memb & " but the information requested to add this person was not received. The information needed was: " & new_memb_verifs & ". This person and their income is mandatory to be provided and because this informaiton has not been provided, your SNAP case will be closed on " & SNAP_close_date & " ")
End If

If temp_disa_abawd_wcom_checkbox Then
    BeginDialog wcom_details_dlg, 0, 0, 131, 60, "WCOM Details"
      EditBox 105, 20, 20, 15, numb_disa_mos
      ButtonGroup ButtonPressed
        OkButton 75, 40, 50, 15
      Text 5, 5, 120, 10, "DISA indicated on form from Doctor"
      Text 0, 25, 105, 10, "Number of months of disability"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If numb_disa_mos = "" Then err_msg = err_msg & vbNewLine & "*Enter the number of months the disability is expected to last from the doctor's information."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""

    CALL add_words_to_message("You are exempt from the ABAWD work providsion because you are unable to work for " & numb_disa_mos & " months per your Doctor statement.")
End If

If client_death_wcom_checkbox Then
    CALL add_words_to_message("This SNAP case has been closed because the only eligible unit member has died.")
End If

If mfip_to_snap_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 166, 80, "WCOM Details"
      EditBox 5, 40, 155, 15, MFIP_closing_reason
      ButtonGroup ButtonPressed
        OkButton 110, 60, 50, 15
      Text 5, 5, 155, 10, "MFIP is closing and SNAP has been assessed"
      Text 5, 25, 105, 10, "Why is MFIP closing:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If MFIP_closing_reason = "" Then err_msg = err_msg & vbNewLine & "*List all reasons why SNAP is closing."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""

    CALL add_words_to_message("You are no longer eligible for MFIP because" & MFIP_closing_reason & ".")

End If

If wreg_postponed_verif_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 281, 115, "WCOM Details"
      EditBox 60, 20, 135, 15, abawd_name
      EditBox 5, 55, 270, 15, wreg_verifs_needed
      EditBox 60, 75, 55, 15, wreg_verifs_due_date
      EditBox 60, 95, 55, 15, snap_closure_date
      ButtonGroup ButtonPressed
        OkButton 225, 95, 50, 15
      Text 5, 5, 185, 10, "Postponed Verification of WREG information/exemption"
      Text 5, 25, 50, 10, "Client Name:"
      Text 5, 40, 70, 10, "Verifications Needed:"
      Text 5, 80, 40, 10, "Verifs Due:"
      Text 5, 100, 50, 10, "SNAP Closure:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If abawd_name = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that has used 3 ABAWD months."
        If wreg_verifs_needed = "" Then err_msg = err_msg & vbNewLine & "* List all WREG verifications needed."
        If isdate(wreg_verifs_due_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verifications are due."
        If isdate(snap_closure_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the day SNAP will close if verifications are not received."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    CALL add_words_to_message(abawd_name & " has used their 3 entitled months of SNAP benefits as an Able Bodied Adult Without Dependents. Verification of " & wreg_verifs_needed & " has been postponed. You must turn in verification of " & wreg_verifs_needed & " by " & wreg_verifs_due_date & " to continue to be eligible for SNAP benefits. If you do not turn in the required verifications, your case will close on " & snap_closure_date & ".")
End If

If banked_mos_avail_wcom_checkbox Then
    CALL add_words_to_message("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
End If

If banked_mos_non_coop_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 201, 65, "WCOM Details"
      EditBox 60, 20, 135, 15, banked_abawd_name
      ButtonGroup ButtonPressed
        OkButton 145, 45, 50, 15
      Text 5, 5, 185, 10, "Client that failed Banked Months E&T requirement."
      Text 5, 25, 50, 10, "Client Name:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If banked_abawd_name = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that has not cooperated with E&T."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""

    CALL add_words_to_message("You have been receiving SNAP banked months. Your SNAP case is closing because " & banked_abawd_name & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before you SNAP clsoes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
End If

If banked_mos_used_wcom_checkbox Then
    CALL add_words_to_message("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all ofther eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
End If

If abawd_child_coded_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 201, 65, "WCOM Details"
      EditBox 60, 20, 135, 15, exempt_abawd_name
      ButtonGroup ButtonPressed
        OkButton 145, 45, 50, 15
      Text 5, 5, 185, 10, "Client exempt from ABAWD due to child in the SNAP Unit."
      Text 5, 25, 50, 10, "Client Name:"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If exempt_abawd_name = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that os using child under 18 years exemption."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""

    CALL add_words_to_message(exempt_abawd_name & " is exempt from the Able Bodied Adults Without Dependents (ABAWD) Work Requirements due to a child(ren) under the age of 18 in the SNAP unit.")
End If

If fset_fail_to_comply_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 201, 85, "WCOM Details"
      EditBox 5, 40, 190, 15, fset_fail_reason
      ButtonGroup ButtonPressed
        OkButton 145, 65, 50, 15
      Text 5, 5, 115, 10, "Client did not meet SNAP E & T rules"
      Text 5, 25, 95, 10, "Reasons client failed FSET"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If fset_fail_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reasons the client failed E&T."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""

    CALL add_words_to_message("Reasons for not meeting the rules: " & fset_fail_reason & ";" & "You can keep getting your SNAP benefits if you show you had a good reason for not meeting the SNAP E & T rules. If you had a good reason, tell us right away.;" & "What do you do next:;" &_
    "You must meet the SNAP E & T rules by the end of the month. If you want to meet the rules, contact your county worker at 612-596-1300, or your SNAP E &T provider at 612-596-7411. You can tell us why you did not meet with the rules. If you had a good reason for not meeting the SNAP E & T rules, contact your SNAP E & T provider right away.")
End If

If snap_pact_wcom_checkbox Then

    BeginDialog wcom_details_dlg, 0, 0, 301, 85, "WCOM Details"
      DropListBox 65, 5, 45, 45, "Select One..."+chr(9)+"CLOSED"+chr(9)+"DENIED", SNAP_close_or_deny
      EditBox 5, 40, 290, 15, pact_close_reason
      ButtonGroup ButtonPressed
        OkButton 245, 65, 50, 15
      Text 5, 10, 55, 10, "SNAP case was "
      Text 120, 10, 35, 10, "on PACT."
      Text 5, 25, 95, 10, "SNAP case closed reason(s):"
    EndDialog

    Do
        err_msg = ""

        Dialog wcom_details_dlg

        If SNAP_close_or_deny = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the case was closed or denied."
        If pact_close_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reasons the SNAP was denied."

        If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""

    CALL add_words_to_message("Your SNAP case was " & SNAP_close_or_deny & " because " & pact_close_reason)
End If

'Navigate to the correct SPEC screen to select the notice
Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

EMWriteScreen MAXIS_footer_month, 3, 46
EMWriteScreen MAXIS_footer_year, 3, 51

transmit

For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)

    If NOTICES_ARRAY(selected, notices_listed) = checked Then
        'Open the Notice
        EMWriteScreen "X", NOTICES_ARRAY(MAXIS_row, notices_listed), 13
        transmit

        PF9     'Put in to edit mode - the worker comment input screen
        EMSetCursor 03, 15

        For each msg_line in WCOM_TO_WRITE_ARRAY
            CALL write_variable_in_SPEC_MEMO(msg_line)
            MsgBox "Look"
        Next

        PF4
        PF3

        If UBound(MEMO_TO_WRITE_ARRAY) > 0 or MEMO_TO_WRITE_ARRAY(0) <> "" Then
            CALL navigate_to_MAXIS_screen("SPEC", "MEMO")

            start_a_new_spec_memo

            For each msg_line in MEMO_TO_WRITE_ARRAY
                CALL write_variable_in_SPEC_MEMO(msg_line)
            Next

            PF4
        End If

        back_to_self
    End If
    ' If NOTICES_ARRAY(selected, notices_listed) = checked Then
    '
    '     'Open the Notice
    '     EMWriteScreen "X", NOTICES_ARRAY(MAXIS_row, notices_listed), 13
    '     transmit
    '
    '     PF9     'Put in to edit mode - the worker comment input screen
    '     EMSetCursor 03, 15
    '     'Write the comment
    '     If duplicate_assistance_wcom_checkbox = checked
    '         CALL write_variable_in_SPEC_MEMO("You received SNAP benefits from the state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & ". You cannot recceive SNAP benefits from two states at the same time.")
    '     End If
    '     If returned_mail_wcom_checkbox = checked
    '         CALL write_variable_in_SPEC_MEMO("Your mail has been returned to our agency. On " & rm_sent_date & " you were sent a Request for you to contact this agency because of this returned mail. You did not contact this agency by " & rm_due_date & " so your SNAP case has been closed.")
    '     End If
    '     If pact_fraud_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("This agency received a request to add " & new_hh_memb & " but the information requested to add this person was not received. The information needed was: " & new_memb_verifs & ". This person and their income is mandatory to be provided and because this informaiton has not been provided, your SNAP case will be closed on " & SNAP_close_date & " ")
    '     End If
    '     If temp_disa_abawd_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("You are exempt from the ABAWD work providsion because you are unable to work for " & numb_disa_mos & " months per your Doctor statement.")
    '     End If
    '     If client_death_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("This SNAP case has been closed because the only eligible unit member has died.")
    '     End If
    '     If mfip_to_snap_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("You are no longer eligible for MFIP because" & MFIP_closing_reason & ".")
    '     End If
    '     If wreg_postponed_verif_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO(name_var & " has used their 3 entitled months of SNAP benefits as an Able Bodied Adult Without Dependents. Verification of " & wreg_varif_var & " has been postponed. You must turn in verification of " & wreg_verif_var & " by " & date_due_verif & " to continue to be eligible for SNAP benefits. If you do not turn in the required verifications, your case will close on " & date_of_closure_var & ".")
    '     End If
    '     If banked_mos_avail_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
    '     End If
    '     If banked_mos_non_coop_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because " & hh_memb_var & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before you SNAP clsoes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
    '     End If
    '     If banked_mos_used_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all ofther eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
    '     End If
    '     If abawd_child_coded_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO(abawd_name_var & " is exempt from the Able Bodied Adults Without Dependents (ABAWD) Work Requirements due to a child(ren) under the age of 18 in the SNAP unit.")
    '     End If
    '     If fset_fail_to_comply_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("Reasons for not meeting the rules: " & fset_reasons_ver)
    '         CALL write_variable_in_SPEC_MEMO("You can keep getting your SNAP benefits if you show you had a good reason for not meeting the SNAP E & T rules. If you had a good reason, tell us right away.")
    '         CALL write_variable_in_SPEC_MEMO("")
    '         CALL write_variable_in_SPEC_MEMO("What do you do next:"
    '         CALL write_variable_in_SPEC_MEMO("You must meet the SNAP E & T rules by the end of the month. If you want to meet the rules, contact your county worker at 612-596-1300, or your SNAP E &T provider at 612-596-7411. You can tell us why you did not meet with the rules. If you had a good reason for not meeting the SNAP E & T rules, contact your SNAP E & T provider right away.")
    '     End If
    '     If snap_pact_wcom_checkbox
    '         CALL write_variable_in_SPEC_MEMO("Your SNAP case was " & close_deny_var & " because " & close_deny_reason_var)
    '     End If
    '
    '     PF4
    '     PF3
    '
    ' End If
Next


CALL navigate_to_MAXIS_screen("CASE", "NOTE")

start_a_blank_case_note

If using_memo = FALSE Then CALL write_variable_in_CASE_NOTE("*** Added WCOM for to Notice to clarify action taken ***")
If using_memo = TRUE Then CALL write_variable_in_CASE_NOTE("*** Added WCOM & MEMO for to Notice to clarify action taken ***")

If duplicate_assistance_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("Advised duplicate assistance from state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & " was received.")
If returned_mail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("Explained returned mail was received. Verification request sent: " & rm_sent_date & " and Due: " & rm_due_date & " with no response caused SNAP case closure.")
If pact_fraud_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained New Household Member: " & new_hh_memb & " added. Verification needed: " & new_memb_verifs & ". Verification not received causing closure.")
If temp_disa_abawd_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Advised client meets ABAWD exemption of temporary inability to work for " & numb_disa_mos & " months per Doctor statement.")
If client_death_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Advised closure due to client death.")
If mfip_to_snap_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained MFIP closure due to: " & MFIP_closing_reason & ".")
If wreg_postponed_verif_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Advised " & name_var & " has used their 3 ABAWD months. Postponed WREG verification: " & wreg_varif_var & " is due: " & date_due_verif & " or SNAP will close on " & date_of_closure_var & ".")
If banked_mos_avail_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
If banked_mos_non_coop_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained " & hh_memb_var & " was receiving Banked Months and fail cooperation with E & T. Explained requesting Good Cause, and future banked months ineligibility.")
If banked_mos_used_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained Banked Months were being used are are now all used. Advised to review other WREG/ABAWD exemptions.")
If abawd_child_coded_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained " & abawd_name_var & " is ABAWD and WREG exemptd due to a child(ren) under the age of 18 in the SNAP unit.")
If fset_fail_to_comply_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Advised SNAP is closing due to FSET requirements not being met. Reasons for not meeting the rules: " & fset_reasons_ver & " Advised of good cause and contact information.")
If snap_pact_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("Explained SNAP case was " & close_deny_var & " because " & close_deny_reason_var)

CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
