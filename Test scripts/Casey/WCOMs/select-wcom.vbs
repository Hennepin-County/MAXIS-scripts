'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SELECT WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================
'run_locally = TRUE
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
call changelog_update("03/13/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTION===================================================================================================================
'Function created to list of all the notices in either WCOM or MEMO - with information
'IDEA - Create_List_Of_Notices function may need to be updated for this particular script
Function Create_List_Of_Notices
    'This function is fairly specific at this time to work when being called within the loop of a dynamic dialog.
    'The array filled will be used to list the notices in the dialog. (Array and constants are defined in the script - outside of the function)
	Erase NOTICES_ARRAY            'Clear the array at the beginning of the function because this can be re called on a loop for dialog display
	no_notices = FALSE             'setting this at the beginning - this will be turned to TRUE if nothing is found in on the specified panel
	If notice_panel = "WCOM" Then      'if the dialog inputs 'WCOM' then the function will go to WCOM
		wcom_row = 7                   'setting initial variables
		array_counter = 0
		Do
			ReDim Preserve NOTICES_ARRAY(3, array_counter)       'resizing the array
			EMReadScreen notice_date, 8,  wcom_row, 16           'getting all the detail from each notice information
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)                      'Formatting the notice information
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE     'This resets the notices boolean to indicate the notice type and month/year have no waiting notices

			NOTICES_ARRAY(selected,    array_counter) = unchecked                'Adding the notice information to the array
			NOTICES_ARRAY(information, array_counter) = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
			NOTICES_ARRAY(MAXIS_row,   array_counter) = wcom_row

			array_counter = array_counter + 1            'incrementing the counter and row
			wcom_row = wcom_row + 1

			EMReadScreen next_notice, 4, wcom_row, 30    'looking to see if another notice exists - loop will exit if no other notices are on the panel
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If

	If notice_panel = "MEMO" Then       'if the dialog inputs 'MEMO' then the function will go to MEMO
		memo_row = 7                    'setting initial variables
		array_counter = 0
		Do
			ReDim Preserve NOTICES_ARRAY(3, array_counter)       'resizing the array
			EMReadScreen notice_date, 8,  memo_row, 19           'getting all the detail from each notice information
			EMReadScreen notice_info, 31, memo_row, 29
			EMReadScreen notice_stat, 8,  memo_row, 67

			notice_date = trim(notice_date)                      'Formatting the notice information
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE     'This resets the notices boolean to indicate the notice type and month/year have no waiting notices

			NOTICES_ARRAY(selected,    array_counter) = unchecked                'Adding the notice information to the arra
			NOTICES_ARRAY(information, array_counter) = notice_info & " - " & notice_date & " - Status: " & notice_stat
			NOTICES_ARRAY(MAXIS_row,   array_counter) = memo_row

			array_counter = array_counter + 1            'incrementing the counter and row
			memo_row = memo_row + 1

			EMReadScreen next_notice, 4, memo_row, 30    'looking to see if another notice exists - loop will exit if no other notices are on the panel
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If
End Function

'Function to add verbiage to an array that will be used to write to a WCOM
'This function is used so that we can correctly asses the length of the message to write in to WCOM - this is vital for this script so that we don't miss out on WCOM verbiage
Function add_words_to_message(message_to_add)

    If trim(message_to_add) <> "" Then  'ensuring there is a value in the message to add
        message_array = split(message_to_add, " ")      'creating an array of all the words in the message

        'ERASE array_of_msg_lines
        ReDim array_of_msg_lines(0)         'blanks out this array each time because we don't want old messages to be duplicated

        message_line = ""                   'setting variables for a FOR...NEXT
        lines_in_msg = 0

        For each word in message_array          'This will look at each word in the message
            'MsgBox lines_in_msg
            If len(word) + len(message_line) > 59 Then              'there are only 59 characters available in each line
                ReDim Preserve array_of_msg_lines(lines_in_msg)     'increases the size of the array of lines in the message input
                array_of_msg_lines(lines_in_msg) = message_line     'adding the combined words to the array
                lines_in_msg = lines_in_msg + 1

                message_line = ""                                   'blanking out the combination of words for each line
            End If

            message_line = message_line & replace(word, ";", "") & " "      'Adding each word to the line

            IF right(word, 1) = ";" Then                                    'moving to a new line if ; is input
                ReDim Preserve array_of_msg_lines(lines_in_msg)
                array_of_msg_lines(lines_in_msg) = message_line
                lines_in_msg = lines_in_msg + 1

                message_line = ""
            End If
        Next

        ReDim Preserve array_of_msg_lines(lines_in_msg)         'adding the last line to the array of lines
        array_of_msg_lines(lines_in_msg) = message_line
        lines_in_msg = lines_in_msg + 1

        'MsgBox "End of WCOM Row: " & end_of_wcom_row & vbNewLine & "Lines Used:" & lines_in_msg
        'Adding a seperator if there is already a message in WCOM
        If UBound(WCOM_TO_WRITE_ARRAY) = 0 AND WCOM_TO_WRITE_ARRAY(0) = "" Then
            notice_line = 0
        Else
            notice_line = UBound(WCOM_TO_WRITE_ARRAY) + 1
            ReDim Preserve WCOM_TO_WRITE_ARRAY(notice_line)
            WCOM_TO_WRITE_ARRAY(notice_line) = "-      - - - - - - - - - - - - - - - - - - - -       -"
            notice_line = notice_line + 1
        End If

        'Here the lines for this message are added to the array that is storing all the messages in the script run as a worker can select multiple
        For each entry in array_of_msg_lines
            'MsgBox entry
            ReDim Preserve WCOM_TO_WRITE_ARRAY(notice_line)
            WCOM_TO_WRITE_ARRAY(notice_line) = trim(entry)
            notice_line = notice_line + 1
        Next

        end_of_wcom_row = end_of_wcom_row + lines_in_msg        'tracking how long the WCOM is already
    End If

End Function

'This function creates the HH Member dropdown for a number of different dialogs
function Generate_Client_List(list_for_dropdown)

	memb_row = 5       'setting the row to look at the list of members on the left hand side of the panel

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")         'go to MEMB
	Do                                                     'this loop transmits to each MEMB panel to read information for each member
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do           'this is the end of the list of members
		EMWriteScreen ref_numb, 20, 76            'writing the reference number in the command line to go to each MEMB panel
		transmit
		EMReadScreen first_name, 12, 6, 63        'reading the name on the panel
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")     'adding each client information to a string
		memb_row = memb_row + 1                   'going to the next member
	Loop until memb_row = 20

    If memb_row = 6 Then        'If the row is only 6, then there is only one person in the HH
        list_for_dropdown = right(client_info, len(client_info) - 1)    'taking the '~' off of the string
    Else
    	client_info = right(client_info, len(client_info) - 1)             'taking the left most '~' off
    	client_list_array = split(client_info, "~")                        'making this an array

    	For each person in client_list_array                               'creating the string to be added to the dialog code to fill the dropdown
    		list_for_dropdown = list_for_dropdown & chr(9) & person
    	Next
    End If

end function

'THE SCRIPT=====================================================================================================================
EMConnect ""            'Connect to BlueZone

Dim NOTICES_ARRAY()         'Creating an array to list all the notices displayed on the panel
ReDim NOTICES_ARRAY(3,0)

Const selected = 0          'Setting constants for easy readability of the array
Const information = 1
Const MAXIS_row = 2

Call check_for_MAXIS(False)     'Making sure that we are not passworded out

'Finds MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

EMReadScreen which_panel, 4, 2, 47          'Checking to see where the script is started from
If which_panel <> "WCOM" then               'If this is not on WCOM - and if the case number is known, the script will navigate to WCOM
    If MAXIS_case_number <> "" Then
        Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	    notice_panel = "WCOM"
	    at_notices = True                  'This boolean tells the script if we are already at one of the notices page (for this script ONLY WCOM)
    Else
        at_notices = FALSE
    End If
Else
    at_notices = TRUE
    notice_panel = "WCOM"
End If


If at_notices = True then               'generating a list of notices if we are at WCOM - so the following dialog will not be empty if we start at WCOM

	EMReadScreen MAXIS_footer_month, 2, 3, 46
	EMReadScreen MAXIS_footer_year,  2, 3, 51

	Create_List_Of_Notices

End If

'This is the DO...LOOP for the dialog to select the WCOM to add information to
Do
	err_msg = ""       'resetting the err_msg variable at the beginning of each loop for handling of correct dialogs

    If NOTICES_ARRAY(0, 0) <> "" Then           'This is looking to see if there is information in the first element of the array (indicating the array has data)
        For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)                                          'looking at all the notices
            EMReadScreen desc, 20, NOTICES_ARRAY(MAXIS_row, notices_listed), 30                     'reading the description of the notice
            if desc = "ELIG Approval Notice" Then                                                   'if the notice is an elig approval - this will check to see if the notice is waiting - these will be prechecked
                EMReadScreen print_status, 7, NOTICES_ARRAY(MAXIS_row, notices_listed), 71
                If print_status = "Waiting" Then NOTICES_ARRAY(selected, notices_listed) = checked
            End If
        Next
    End If

	dlg_y_pos = 65     'setting some lengths and positions
	dlg_length = 125 + (UBound(NOTICES_ARRAY, 2) * 20)

	BeginDialog find_notices_dialog, 0, 0, 205, dlg_length, "Notices to add WCOM"      'This is what the dialog will look like
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

	Dialog find_notices_dialog         'display the dialog
	cancel_confirmation

	notice_selected = FALSE            'this boolean and loop will identify if no notice has been selected
	For notice_to_print = 0 to UBound(NOTICES_ARRAY, 2)
		If NOTICES_ARRAY(selected, notice_to_print) = checked Then notice_selected = TRUE
	Next

    'looking for errors in the dialog entry
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
	If MAXIS_footer_month = "" or MAXIS_footer_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice to be copied to a Word Document."

    'If the button is pressed to find notices, the loop will not entry - but instead navigate to the WCOM for the specified case and month/year
	If ButtonPressed = find_notices_button then
		If MAXIS_case_number <> "" AND MAXIS_footer_month <> "" AND MAXIS_footer_year <> "" Then  'navigation only works with case number and footer month/year
			Call navigate_to_MAXIS_screen ("SPEC", notice_panel)            'for this script - this is always WCOM
			EMWriteScreen MAXIS_footer_month, 3, 46
			EMWriteScreen MAXIS_footer_year, 3, 51

			transmit
			Create_List_Of_Notices           'using the funcation to create a list of notices for the dialog
			err_msg = "LOOP"                 'this keeps the loop from exiting since err_msg will not be blank
		Else
			err_msg = err_msg & vbNewLine & "!!! Cannot read a list of notices without a case number entered, and footer month & year entered !!!"   'If case number or footer month/year are not specified - this will be the error
		End If
	End If

    'The error message will only display if it is not blank AND is not the one to keep the loop from exiting.
	If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg

Loop Until err_msg = ""

'navigating to the panel for case case and footer month/year specified.
Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

EMWriteScreen MAXIS_footer_month, 3, 46
EMWriteScreen MAXIS_footer_year, 3, 51

transmit

'This will cycle through all the notices that are on WCOM
For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)

    If NOTICES_ARRAY(selected, notices_listed) = checked Then   'If the worker selected the notice
        'Navigate to the correct SPEC screen to select the notice
        Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

        EMWriteScreen MAXIS_footer_month, 3, 46
        EMWriteScreen MAXIS_footer_year, 3, 51

        transmit

        'Open the Notice
        EMWriteScreen "X", NOTICES_ARRAY(MAXIS_row, notices_listed), 13
        transmit

        PF9     'Put in to edit mode - the worker comment input screen

        'Checking to see that the WCOM goes in to edit mode because otherwise we can't add WCOMs
        EMReadScreen edit_mode_check, 18, 24, 36
        If edit_mode_check = "UPDATE NOT ALLOWED" Then
            PF3
            end_msg = "Could not put the WCOM (" & NOTICES_ARRAY(information, notices_listed) & ") in to EDIT mode to add a WCOM. Likely this notice has already been printed. Review the notices on this case and run the script again if needed."
            script_end_procedure(end_msg)
        End If

        'Making sure there is no other text entered in the WCOM area as it needs to be open to being written in.
        For wcom_row = 3 to 17
            EMReadScreen wcom_line, 60, wcom_row, 15
            'msgBox "~" & wcom_line & "~"
            If trim(wcom_line) <> "" Then
                PF10
                PF3
                script_end_procedure("This script must be run before adding any additional WCOMs. If there is a manual WCOM to add, run the script first, then add the manual WCOM second.")
            End If
        Next

        PF10    'exiting without saving since we didn't do anything yet
        PF3

        back_to_self
        wcom_row = ""
    End If
Next

'setting these variables
'IDEA the WCOMs available will vary depending on the type of notice that was selected - since each program has different WCOM needs
SNAP_notice = FALSE
MFIP_notice = FALSE
GA_notice = FALSE
MSA_notice = FALSE

'This bit identifies which type of notice has been selected - so that in the future WCOMs listed in the next dialog can be adjusted based on the type of notice
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
BeginDialog wcom_selection_dlg, 0, 0, 241, 345, "Select WCOM"
  Text 10, 10, 95, 10, "Check the WCOM needed."
  GroupBox 10, 25, 225, 240, "SNAP"
  CheckBox 20, 40, 150, 10, "SNAP closed/denied with PACT", snap_pact_wcom_checkbox
  CheckBox 20, 55, 155, 10, "SNAP closed via PACT for new HH Member", pact_fraud_wcom_checkbox
  CheckBox 20, 70, 145, 10, "SNAP closing due to Returned Mail", snap_returned_mail_wcom_checkbox
  CheckBox 20, 85, 115, 10, "SNAP closing and MFIP opening", snap_to_mfip_wcom_checkbox
  CheckBox 20, 100, 190, 10, "SNAP Duplicate Assistance in another state", duplicate_assistance_wcom_checkbox
  CheckBox 20, 115, 190, 10, "SNAP Duplicate Assistance on another case in MN", dup_assistance_in_MN_wcom_checkbox
  CheckBox 20, 130, 85, 10, "SNAP Applicant Death", client_death_wcom_checkbox
  CheckBox 20, 145, 205, 10, "SNAP Postponed verif of CAF pg 9 Signature - for EXP SNAP", signature_postponed_verif_wcom_checkbox
  CheckBox 20, 160, 190, 10, "ABAWD - Postponed WREG verifs for EXP SNAP", wreg_postponed_verif_wcom_checkbox
  CheckBox 20, 175, 155, 10, "ABAWD WREG coded for Child under 18", abawd_child_coded_wcom_checkbox
  CheckBox 20, 190, 130, 10, "ABAWD - Temporarily disabled", temp_disa_abawd_wcom_checkbox
  CheckBox 20, 205, 160, 10, "ABAWD - Banked Months possibly available", banked_mos_avail_wcom_checkbox
  CheckBox 20, 220, 180, 10, "ABAWD - Banked Months E and T voluntary", banked_mos_vol_e_t_wcom_checkbox
  CheckBox 20, 235, 190, 10, "ABAWD - Banked Months closing for months used", banked_mos_used_wcom_checkbox
  CheckBox 20, 250, 195, 10, "FSET - Failure to comply with Good Cause Information", fset_fail_to_comply_wcom_checkbox
  GroupBox 10, 265, 225, 55, "Cash"
  CheckBox 20, 275, 55, 10, "CASH Denied", cash_denied_checkbox
  CheckBox 20, 290, 125, 10, "CASH closing due to Returned Mail", mfip_returned_mail_wcom_checkbox
  CheckBox 20, 305, 125, 10, "MFIP Closing and SNAP opening", mfip_to_snap_wcom_checkbox
  ButtonGroup ButtonPressed
    OkButton 135, 325, 50, 15
    CancelButton 185, 325, 50, 15
EndDialog

' Dim myBtn
'
' myBtn = Dialog(wcom_selection_dlg)
' MsgBox "The user pressed button " & myBtn


'Initial declaration of arrays
Dim array_of_msg_lines ()
Dim WCOM_TO_WRITE_ARRAY ()
'Eventually this checkbox dialog will be dynamic and the WCOMs available will be different based on the programs of the notices selected.
'THIS is a big loop that will be used to make sure the WCOM is not too long
Do      'Just made this  loop - this needs sever testing.
    big_err_msg = ""            'this error message is called something different because there are other err_msg variables that happen within this loop for each WCOM

    Dialog wcom_selection_dlg       'running the dialog to select which WCOMs are going to be added
    cancel_confirmation

    end_of_wcom_line = 0            'setting variables to asses length of WCOM
    end_of_wcom_row = 1

    'setting the arrays to blank for each loop - they will be refilled once the checkboxes are selected again
    ReDim array_of_msg_lines(0)
    ReDim WCOM_TO_WRITE_ARRAY (0)

    'Here there is an IF statement for each checkbox - each WCOM may have it's own dialog and the verbiage will be added to the array for the WCOM lines
    If snap_pact_wcom_checkbox = checked Then             'SNAP closed with PACT
        'code for the dialog for PACT closure (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 301, 85, "WCOM Details"
          DropListBox 65, 5, 45, 45, "Select One..."+chr(9)+"CLOSED"+chr(9)+"DENIED", SNAP_close_or_deny
          EditBox 5, 40, 290, 15, pact_close_reason
          ButtonGroup ButtonPressed
            OkButton 245, 65, 50, 15
          Text 5, 10, 55, 10, "SNAP case was "
          Text 120, 10, 35, 10, "on PACT."
          Text 5, 25, 95, 10, "SNAP case closed reason(s):"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If SNAP_close_or_deny = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the case was closed or denied."
            If pact_close_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reasons the SNAP was denied."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("Your SNAP case was " & SNAP_close_or_deny & " because " & pact_close_reason & ".")
    End If

    If pact_fraud_wcom_checkbox = checked Then        'FPI findings indicate another person
        'code for the dialog for closing for fpi result (this dialog has the same name in each IF to prevent the over 7 dialog error)
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

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If trim(new_hh_memb) = "" Then err_msg = err_msg & vbNewLine & "*Enter the name of the person who has joined the household."
            If isdate(SNAP_close_date) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date on which SNAP will close."
            If trim(new_memb_verifs) = "" Then err_msg = err_msg & vbNewLine & "*Enter the verifications that were needed to add this person to the case." & vbNewLine & "If no verifications are required - this is not the correct WCOM to use."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("This agency received a request to add " & new_hh_memb & " but the information requested to add this person was not received. The information needed was: " & new_memb_verifs & ". This person and their income is mandatory to be provided and because this information has not been provided, your SNAP case will be closed on " & SNAP_close_date & ". ")
    End If

    If snap_returned_mail_wcom_checkbox = checked Then           'Returned Mail
        'code for the dialog for returned mail (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 126, 85, "WCOM Details"
          EditBox 75, 20, 45, 15, rm_sent_date_snap
          EditBox 75, 40, 45, 15, rm_due_date_snap
          ButtonGroup ButtonPressed
            OkButton 60, 65, 50, 15
          Text 5, 5, 110, 10, "SNAP Returned Mail"
          Text 5, 25, 65, 10, "Verif Request Sent:"
          Text 5, 45, 65, 10, "Verif Request Due:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            if isdate(rm_sent_date_snap) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the request for address information was sent."
            if isdate(rm_due_date_snap) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the response for address information was due."
            if err_msg <> "" Then msgBox "Resolve to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("Your mail has been returned to our agency. On " & rm_sent_date_snap & " you were sent a request for you to contact this agency because of this returned mail. You did not contact this agency by " & rm_due_date_snap & " so your SNAP case has been closed.")
    End If

    If snap_to_mfip_wcom_checkbox = checked then        'SNAP closing for MFIP opening - no input needed
        CALL add_words_to_message("Your SNAP case is closing because food benefits will be included in the MFIP benefit.")
    End If

    If duplicate_assistance_wcom_checkbox = checked Then        'Duplicate assistance in another state
        If dup_state = "" Then                                  'If this is blank the script will look for it on MEMI - but if we are looping and the worker has already filled it in, the script will let that value stand
            Call navigate_to_MAXIS_screen ("STAT", "MEMI")
            EMReadScreen dup_state, 2, 14, 78
        End If
        If dup_state = "__" Then dup_state = ""                 'formatting state to not have underscores

        If dup_month = "" Then dup_month = MAXIS_footer_month   'setting the month and year as a default
        If dup_year = "" Then dup_year = MAXIS_footer_year

        'code for the dialog for dup assistance (this dialog has the same name in each IF to prevent the over 7 dialog error)
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

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If trim(dup_state) = "" Then err_msg = err_msg & vbNewLine & "* Enter the state in which client already received SNAP."
            If trim(dup_month) = "" or trim(dup_year) = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year for which SNAP in MN is being denied due to receipt of benefits in another state."
            If err_msg <> "" Then MsgBox "Please resolve before continuing:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You received SNAP benefits from the state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & ". You cannot recceive SNAP benefits from two states at the same time.")
    End If

    If dup_assistance_in_MN_wcom_checkbox = checked Then        'Duplicate assistance in MN

        BeginDialog wcom_details_dlg, 0, 0, 116, 70, "WCOM Details"
          EditBox 70, 25, 15, 15, mn_dup_month
          EditBox 90, 25, 15, 15, mn_dup_year
          ButtonGroup ButtonPressed
            OkButton 60, 50, 50, 15
          Text 5, 10, 110, 10, "Duplicate SNAP in this state"
          Text 5, 30, 60, 10, "In (Month/Year)"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If trim(mn_dup_month) = "" or trim(mn_dup_year) = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year for which SNAP in MN is being denied due to receipt of benefits in another state."
            If err_msg <> "" Then MsgBox "Please resolve before continuing:" & vbNewLine & err_msg
        Loop until err_msg = ""

        CALL add_words_to_message("You will not be eligible for SNAP benefits this month since you have received SNAP benefits on another case in the month of " & mn_dup_month & "/" & mn_dup_year & ".; Per program rules SNAP participants are not eligible for duplicate benefits in the same benefit month.")
    End If

    If client_death_wcom_checkbox = checked Then      'Client death
        single_person = FALSE

        Call navigate_to_MAXIS_screen("STAT", "MEMB")

        EmWriteScreen "01", 20, 76
        transmit

        EMReadScreen second_member, 2, 6, 3
        If second_member = "  " Then single_person = TRUE


        Do
            EMReadScreen date_of_death, 10, 19, 42
            date_of_death = replace(date_of_death, " ", "/")

            If date_of_death <> "__/__/____" Then
                EMReadScreen first_name, 12, 6, 63
                EMReadScreen last_name, 25, 6, 30
                EMReadScreen middle_initial, 1, 6, 79

                first_name = replace(first_name, "_", "")
                last_name = replace(last_name, "_", "")
                middle_initial = replace(middle_initial, "_", "")

                If middle_initial = "" Then
                    deceased_client = first_name & " " & last_name
                Else
                    deceased_client = first_name & " " & middle_initial & ". " & last_name
                End If

                EMReadScreen client_ref_nbr, 2, 4, 33

                Exit Do
            Else
                date_of_death = ""
                transmit
            ENd If

            EMReadScreen check_for_last_memb, 13, 24, 2
        Loop Until check_for_last_memb = "ENTER A VALID"

        others_on_eats = FALSE

        If single_person = TRUE then
            only_elig_checkbox = checked
        Else
            Call navigate_to_MAXIS_screen("STAT", "EATS")
            EMReadScreen check_for_eats, 14, 24, 7
            If check_for_eats <> "DOES NOT EXIST" Then
                EMReadScreen all_eat_together, 1, 4, 72
                If all_eat_together = "Y" Then
                    others_on_eats = TRUE
                Else
                    row = 1
                    col = 1
                    EMSearch client_ref_nbr, row, col

                    If col <> 39 Then
                        others_on_eats = TRUE
                    Else
                        EMReadScreen next_in_group, row, 43
                        If next_in_group <> "__" Then others_on_eats = TRUE
                    End If
                ENd If
            End If
            If others_on_eats = FALSE Then only_elig_checkbox = checked
        End If
        'MsgBox "single person - " & single_person & vbNewLine & "others on eats - " & others_on_eats

        BeginDialog wcom_details_dlg, 0, 0, 236, 60, "WCOM Details"
          EditBox 95, 5, 135, 15, deceased_client
          CheckBox 5, 25, 220, 10, "Check here if client was only eligible SNAP member on this case.", only_elig_checkbox
          EditBox 95, 40, 55, 15, date_of_death
          ButtonGroup ButtonPressed
            OkButton 180, 40, 50, 15
          Text 5, 10, 85, 10, "Name of deceased client:"
          Text 25, 45, 60, 10, "Date of death:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If trim(deceased_client) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client who died."
            If IsDate(date_of_death) = "" Then err_msg = err_msg & vbNewLine & "* Enter a valid date for date of death."

            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""

        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        If only_elig_checkbox = unchecked Then
            CALL add_words_to_message("This SNAP case has closed because the applicant " & deceased_client & " has died.; To continue to receive SNAP benefits, you must reapply for a new case.")
        Else
            CALL add_words_to_message("This SNAP case has closed because the applicant " & deceased_client & " has died.")
        End If
    End If

    If signature_postponed_verif_wcom_checkbox = checked Then       'postponed signature of CAF for XFS

        BeginDialog wcom_details_dlg, 0, 0, 201, 40, "WCOM Details"
          EditBox 60, 20, 55, 15, snap_closure_date_sig
          ButtonGroup ButtonPressed
            OkButton 145, 20, 50, 15
          Text 5, 5, 185, 10, "Postponed Verification of WREG information/exemption"
          Text 5, 25, 50, 10, "SNAP Closure:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If isdate(snap_closure_date_sig) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the day SNAP will close if verifications are not received."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""

        CALL add_words_to_message("Receipt of signature on application (last page signature required) has been poostponed. Return the application page with signature and date to continue to be eligible for SNAP benefits. If not received, SNAP will close on " & snap_closure_date_sig & ".")
    End If

    If wreg_postponed_verif_wcom_checkbox = checked Then      'XFS Postponed verifs are in WREG
        'code for the dialog for postponed verifs from WREG (this dialog has the same name in each IF to prevent the over 7 dialog error)
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

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If abawd_name = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that has used 3 ABAWD months."
            If wreg_verifs_needed = "" Then err_msg = err_msg & vbNewLine & "* List all WREG verifications needed."
            If isdate(wreg_verifs_due_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verifications are due."
            If isdate(snap_closure_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the day SNAP will close if verifications are not received."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message(abawd_name & " has used their 3 entitled months of SNAP benefits as an Able Bodied Adult Without Dependents. Verification of " & wreg_verifs_needed & " has been postponed. You must turn in verification of " & wreg_verifs_needed & " by " & wreg_verifs_due_date & " to continue to be eligible for SNAP benefits. If you do not turn in the required verifications, your case will close on " & snap_closure_date & ".")
    End If

    If abawd_child_coded_wcom_checkbox = checked Then         'ABAWD exemption for care of child
        'code for the dialog for ABAWD child exemption (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 201, 65, "WCOM Details"
          EditBox 60, 20, 135, 15, exempt_abawd_name
          ButtonGroup ButtonPressed
            OkButton 145, 45, 50, 15
          Text 5, 5, 185, 10, "Client exempt from ABAWD due to child in the SNAP Unit."
          Text 5, 25, 50, 10, "Client Name:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If exempt_abawd_name = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that os using child under 18 years exemption."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message(exempt_abawd_name & " is exempt from the Able Bodied Adults Without Dependents (ABAWD) Work Requirements due to a child(ren) under the age of 18 in the SNAP unit.")
    End If

    If temp_disa_abawd_wcom_checkbox = checked Then       'Verified temporary disa for ABAWD exemption
        'code for the dialog for temporary disa for ABAWD (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 131, 60, "WCOM Details"
          EditBox 105, 20, 20, 15, numb_disa_mos
          ButtonGroup ButtonPressed
            OkButton 75, 40, 50, 15
          Text 5, 5, 120, 10, "DISA indicated on form from Doctor"
          Text 0, 25, 105, 10, "Number of months of disability"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If trim(numb_disa_mos) = "" Then err_msg = err_msg & vbNewLine & "*Enter the number of months the disability is expected to last from the doctor's information."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You are exempt from the ABAWD work provision because you are unable to work for " & numb_disa_mos & " months per your Doctor statement.")
    End If

    If banked_mos_avail_wcom_checkbox = checked Then      'ABAWD expired - Banked Months available - NO WORKER INPUT NEEDED
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You have used all of your available ABAWD months. You may be eligible for SNAP banked months. Please contact your team if you have questions.")
    End If

    If banked_mos_vol_e_t_wcom_checkbox = checked Then       'Banked Months closing due to FSET non-coop - NO WORKER INPUT NEEDED
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You have been approved to receive additional SNAP benefits under the SNAP Banked Months program. Working with Employment and Training is voluntary under this program. If you'd like work with Employment and Training, please contact your team.")
    End If

    If banked_mos_used_wcom_checkbox = checked Then       'Banked Months expired - NO WORKER INPUT NEEDED
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your team if you have questions.")
    End If

    If fset_fail_to_comply_wcom_checkbox = checked Then           'Fail to comply with FSET
        'code for the dialog for fail to comply with FSET (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 201, 85, "WCOM Details"
          EditBox 5, 40, 190, 15, fset_fail_reason
          ButtonGroup ButtonPressed
            OkButton 145, 65, 50, 15
          Text 5, 5, 115, 10, "Client did not meet SNAP E & T rules"
          Text 5, 25, 95, 10, "Reasons client failed FSET"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If fset_fail_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reasons the client failed E&T."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        ' CALL add_words_to_message("Reasons for not meeting the rules: " & fset_fail_reason & ". You can keep getting your SNAP benefits if you show you had a good reason for not meeting the SNAP E & T rules. If you had a good reason, tell us right away.;" & "What do you do next:;" &_
        ' "You must meet the SNAP E & T rules by the end of the month. If you want to meet the rules, contact your county worker at 612-596-1300, or your SNAP E &T provider at 612-596-7411. You can tell us why you did not meet with the rules. If you had a good reason for not meeting the SNAP E & T rules, contact your SNAP E & T provider right away.")

        CALL add_words_to_message("What to do next:; * You must meet the SNAP E&T rules by the end of the month. If you want to meet the rules, contact your team at 612-596-1300, or your SNAP E&T provider at 612-596-7411.; * You can tell us why you did not meet the rules. If you had a good reason for not meeting the SNAP E&T rules, contact your SNAP E&T provider right away.")
    End If

    If cash_denied_checkbox = checked Then              'Information to add to CASH denial

        BeginDialog wcom_details_dlg, 0, 0, 321, 105, "WCOM Details"
          DropListBox 5, 25, 50, 45, "Select one..."+chr(9)+"MFIP"+chr(9)+"DWP"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"RCA", cash_one_program
          EditBox 115, 25, 200, 15, cash_one_reason
          DropListBox 5, 65, 50, 45, "Select one..."+chr(9)+"MFIP"+chr(9)+"DWP"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"RCA", cash_two_program
          EditBox 115, 65, 200, 15, cash_two_reason
          ButtonGroup ButtonPressed
            OkButton 265, 85, 50, 15
          Text 5, 10, 90, 10, "First cash program denied:"
          Text 60, 30, 55, 10, "denied because:"
          Text 5, 50, 100, 10, "Second cash program denied:"
          Text 60, 70, 55, 10, "denied because:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If cash_one_program = "Select one..." Then err_msg = err_msg & vbNewLine & "* Select at least one cash program that is actually being denied."
            If trim(cash_one_reason) = "" Then err_msg = err_msg & vbNewLine & "* Explain why the first cash progam is being denied."
            If cash_two_program = "Select one..." AND trim(cash_two_reason) = "" Then err_msg = err_msg & vbNewLine & "* Explain why the second cash program is being denied."

            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""

        If cash_two_program <> "Select one..." Then
            CALL add_words_to_message(cash_one_program & " is being denied because " & cash_one_reason & ".; " & cash_two_program & " is being denied because " & cash_two_reason & ".")
        Else
            CALL add_words_to_message(cash_one_program & " is being denied because " & cash_one_reason & ".")
        End If
    End If

    If mfip_returned_mail_wcom_checkbox = checked Then      'Returned mail for cash
        'code for the dialog for returned mail (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 126, 85, "WCOM Details"
          EditBox 75, 20, 45, 15, rm_sent_date_cash
          EditBox 75, 40, 45, 15, rm_due_date_cash
          ButtonGroup ButtonPressed
            OkButton 60, 65, 50, 15
          Text 5, 5, 110, 10, "CASH Returned Mail"
          Text 5, 25, 65, 10, "Verif Request Sent:"
          Text 5, 45, 65, 10, "Verif Request Due:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            if isdate(rm_sent_date_cash) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the request for address information was sent."
            if isdate(rm_due_date_cash) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the response for address information was due."
            if err_msg <> "" Then msgBox "Resolve to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""

        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("Your mail has been returned to our agency. On " & rm_sent_date_cash & " you were sent a request for you to contact this agency because of this returned mail. You can avoid having your case closed if you contact this agency by " & rm_due_date_cash & ".")
    End If

    If mfip_to_snap_wcom_checkbox = checked Then      'MFIP is closing and SNAP is opening
        'code for the dialog for MFIP closure reason when SNAP is reassessed (this dialog has the same name in each IF to prevent the over 7 dialog error)
        BeginDialog wcom_details_dlg, 0, 0, 166, 80, "WCOM Details"
          EditBox 5, 40, 155, 15, MFIP_closing_reason
          ButtonGroup ButtonPressed
            OkButton 110, 60, 50, 15
          Text 5, 5, 155, 10, "MFIP is closing and SNAP has been assessed"
          Text 5, 25, 105, 10, "Why is MFIP closing:"
        EndDialog

        Do                          'displaying the dialog and ensuring that all required information is entered
            err_msg = ""

            Dialog wcom_details_dlg

            If MFIP_closing_reason = "" Then err_msg = err_msg & vbNewLine & "*List all reasons why SNAP is closing."
            If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
        CALL add_words_to_message("You are no longer eligible for MFIP because " & MFIP_closing_reason & ".")
    End If


    'This assesses if the message generated is too long for WCOM. If so then the checklist will reappear along with each selected WCOM dialog so it can be changed
    If UBOUND(WCOM_TO_WRITE_ARRAY) > 14 Then big_err_msg = big_err_msg & vbNewLine & "The amount of text/information that is being added to WCOM will exceed the 15 lines available on MAXIS WCOMs. Please reduce the number of WCOMs that have been selected or reduce the amount of text in the selected WCOM."
    'MsgBox "UBOUND of array is " & UBOUND(WCOM_TO_WRITE_ARRAY)

    ' If end_of_wcom_row > 14 Then big_err_msg = big_err_msg & vbNewLine & "The amount of text/information that is being added to WCOM will exceed the 15 lines available on MAXIS WCOMs. Please reduce the number of WCOMs that have been selected or reduce the amount of text in the selected WCOM."
    ' MsgBox "End of WCOM ROW is " & end_of_wcom_row
    'Leave this here - testing purposes
    wcom_to_display = ""
    For each msg_line in WCOM_TO_WRITE_ARRAY
        if wcom_to_display = "" Then
            wcom_to_display = msg_line
        else
            wcom_to_display = wcom_to_display & vbNewLine & msg_line
        end if
    Next
    'MsgBox wcom_to_display

    If big_err_msg <> "" Then MsgBox "*** Please resolved the following to continue ***" & vbNewLine & big_err_msg
Loop until big_err_msg = ""

'This will cycle through all the notices that are on WCOM
For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)

    If NOTICES_ARRAY(selected, notices_listed) = checked Then   'If the worker selected the notice
        'Navigate to the correct SPEC screen to select the notice
        Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

        EMWriteScreen MAXIS_footer_month, 3, 46
        EMWriteScreen MAXIS_footer_year, 3, 51

        transmit

        'Open the Notice
        EMWriteScreen "X", NOTICES_ARRAY(MAXIS_row, notices_listed), 13
        transmit

        PF9     'Put in to edit mode - the worker comment input screen
        EMSetCursor 03, 15

        For each msg_line in WCOM_TO_WRITE_ARRAY        'each line in this array will be written to the WCOM
            CALL write_variable_in_SPEC_MEMO(msg_line)
        Next

        'MsgBox "Look"
        PF4     'Save the WCOM
        PF3     'Exit the WCOM

        back_to_self
    End If
Next

'Now the action will be case noted
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

start_a_blank_case_note

CALL write_variable_in_CASE_NOTE("*** Added WCOM for to Notice to clarify action taken ***")
CALL write_variable_in_CASE_NOTE("Inormation added to the following WCOM notices in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ":")
For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)
    If NOTICES_ARRAY(selected, notices_listed) = checked Then
        CALL write_variable_in_CASE_NOTE("* " & NOTICES_ARRAY(information, notices_listed))
    End If
Next
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE("Detail added to each notice:")

If snap_pact_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("* SNAP case was " & SNAP_close_or_deny & " because " & pact_close_reason & ".")
If pact_fraud_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("* Request to add: " & new_hh_memb & ". Verification needed: " & new_memb_verifs & ". Verification not received causing closure.")
If snap_returned_mail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Returned mail was received. Verification request sent: " & rm_sent_date_snap & " and Due: " & rm_due_date_snap & " with no response caused SNAP case closure.")
If snap_to_mfip_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* SNAP is closing but food will be included in MFIP.")
If duplicate_assistance_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Duplicate assistance from state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & " was received.")
If dup_assistance_in_MN_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Duplicate assistance on a case in MN during the month of " & mn_dup_month & "/" & mn_dup_year & ".")
If client_death_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Closure due to client death.")
If signature_postponed_verif_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Postponed signature on CAF page 9.")
If wreg_postponed_verif_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & abawd_name & " has used their 3 ABAWD months. Postponed WREG verification: " & wreg_verifs_needed & " is due: " & wreg_verifs_due_date & ".")
If abawd_child_coded_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & exempt_abawd_name & " is ABAWD and WREG exempt due to a child(ren) under the age of 18 in the SNAP unit.")
If temp_disa_abawd_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Client meets ABAWD exemption of temporary inability to work for " & numb_disa_mos & " months per Doctor statement.")
If banked_mos_avail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* ABAWD months have been used, explained Banked Months may be available.")
If banked_mos_vol_e_t_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* E&T is voluntary with Banked Months.")
If banked_mos_non_coop_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & banked_abawd_name & " was receiving Banked Months and fail cooperation with E & T. Explained requesting Good Cause, and future banked months ineligibility.")
If banked_mos_used_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Banked Months were being used are are now all used. Advised to review other WREG/ABAWD exemptions.")
If fset_fail_to_comply_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* SNAP is closing due to FSET requirements not being met. Reasons for not meeting the rules: " & fset_fail_reason & ". Advised of good cause and contact information.")
If cash_denied_checkbox = checked Then
    CALL write_variable_in_CASE_NOTE("* " & cash_one_program & " denied because " & cash_one_reason & ".")
    If cash_two_program <> "Select one..." Then CALL write_variable_in_CASE_NOTE("* " & cash_two_program & " denied because " & cash_two_reason & ".")
End If
If mfip_returned_mail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Returned mail was received. Verification request sent: " & rm_sent_date_cash & ", cash denied and can be reopened if verif received by: " & rm_due_date_cash & ".")
If mfip_to_snap_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* MFIP closure due to: " & MFIP_closing_reason & ".")

CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
