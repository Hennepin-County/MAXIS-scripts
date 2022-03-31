'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - ADD WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                               'manual run time in seconds
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
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
Call changelog_update("10/20/2021", "Updated online document submission option to include MNBenefits. Added Health Care PARIS match WCOM.", "Ilse Ferris, Hennepin County")
Call changelog_update("09/02/2021", "Added functionality to support sending a WCOM about any Expedited SNAP Postponed Verification.", "Casey Love, Hennepin County")
Call changelog_update("04/09/2020", "Multiple updates to available WCOMs:##~## - ALL Banked Months WCOMs are removed as no Banked Months are currently being issued.##~## - Added client name information to Temporary Disabled ABAWD WCOM##~## - Added WCOM for Care of Child under 6 Exemption.##~## - Added WCOM for close/deny due to no verifications for when the notice reads 'No Eligible Members'.##~## - Added WCOM for Ineligible Student. ##~## - Added WCOM for Voluntary Quit.##~## - Added WCOM for Future Eligibility Begin Date for SNAP.##~##", "Casey Love, Hennepin County")
Call changelog_update("03/22/2019", "Reformatted the Select WCOM Dialog. The layout is now clearer as to what WCOMs are for ABAWD. Additional DHS mandated WCOMs are indicated with an asterisk (*).", "Casey Love, Hennepin County")
Call changelog_update("03/07/2019", "Removed WCOMs for duplicate assistance (in MN and Out of State) and client death as notices have been updated to include details of these changes.", "Casey Love, Hennepin County")
call changelog_update("02/20/2019", "Adjusted wording to fit ABAWD Voluntary E&T with Homeless Exemption on a single WCOM.", "Casey Love, Hennepin County")
call changelog_update("01/18/2019", "Reorganized and renamed WCOM options in user dialog for ease of use.", "Ilse Ferris, Hennepin County")
call changelog_update("01/16/2019", "Updated Banked months homeless WCOM to be used for all ABAWD cases. This option allows users to notify a client of the potential for the homeless ABAWD exemption.", "Ilse Ferris, Hennepin County")
call changelog_update("11/01/2018", "Removed 'Failure to Comply' WCOM and added WCOM for Voluntary SNAP E&T..", "Casey Love, Hennepin County")
call changelog_update("09/30/2018", "Initial version.", "Casey Love, Hennepin County")

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
            trimmed_word = trim(word)
            trimmed_word = replace(word, ";", "")
            If len(trimmed_word) + len(message_line) > 59 Then              'there are only 59 characters available in each line
                'MsgBox "On the word ~" & trimmed_word & "~ the line was too long." & vbNewLine & "Line is currenlty ~" & message_line & "~" & vbNewLine & "The position is " & len(trimmed_word) + len(message_line)
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
            'MsgBox "~" & entry & "~" & vbNewLine & vbNewLine & "Which is " & len(entry) & " characters long."
            ReDim Preserve WCOM_TO_WRITE_ARRAY(notice_line)
            WCOM_TO_WRITE_ARRAY(notice_line) = trim(entry)
            notice_line = notice_line + 1
        Next

        end_of_wcom_row = end_of_wcom_row + lines_in_msg        'tracking how long the WCOM is already
    End If

End Function

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
notice_panel = "WCOM"

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

        Dialog1 = ""
    	BeginDialog Dialog1, 0, 0, 205, dlg_length, "Notices to add WCOM"      'This is what the dialog will look like
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

    	Dialog Dialog1         'display the dialog
    	cancel_confirmation

    	notice_selected = FALSE            'this boolean and loop will identify if no notice has been selected
    	For notice_to_print = 0 to UBound(NOTICES_ARRAY, 2)
    		If NOTICES_ARRAY(selected, notice_to_print) = checked Then notice_selected = TRUE
    	Next

        'looking for errors in the dialog entry
    	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
    	If MAXIS_footer_month = "" or MAXIS_footer_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
    	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice that needs a WCOM added."

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
    call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

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

CALL Generate_Client_List(client_dropdown, "Select One...")

'Initial declaration of arrays
Dim array_of_msg_lines ()
Dim WCOM_TO_WRITE_ARRAY ()
'Eventually this checkbox dialog will be dynamic and the WCOMs available will be different based on the programs of the notices selected.
'THIS is a big loop that will be used to make sure the WCOM is not too long
Do
    Do      'Just made this  loop - this needs sever testing.
        big_err_msg = ""            'this error message is called something different because there are other err_msg variables that happen within this loop for each WCOM

        'DIALOG to select the WCOM to add
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 241, 395, "Check the WCOM needed"
            CheckBox 10, 35, 195, 10, "Online Document Submission Options", clt_virtual_dropbox_checkbox
            CheckBox 20, 70, 195, 10, "E and T Voluntary *", voluntary_e_t_wcom_checkbox
            CheckBox 20, 85, 195, 10, "Homeless exemption information", abawd_homeless_wcom_checkbox
            CheckBox 20, 100, 195, 10, "WREG Exemption coded - Temporarily disabled *", temp_disa_abawd_wcom_checkbox
            CheckBox 20, 115, 195, 10, "WREG Exemption coded - Care of Child under 18 *", abawd_child_18_coded_wcom_checkbox
            CheckBox 20, 130, 195, 10, "WREG Exemption coded - Care of Child under 6 *", abawd_child_6_coded_wcom_checkbox
            CheckBox 20, 145, 195, 10, "Voluntary Quit WCOM - non-PWE", voluntary_quit_wcom_checkbox
            CheckBox 20, 175, 195, 10, "No Eligible Members and verifs missing or unclear *", additional_verif_wcom_checkbox
            CheckBox 20, 190, 195, 10, "Closed/denied with PACT *", snap_pact_wcom_checkbox
            CheckBox 20, 205, 195, 10, "Closed via PACT for new HH Member *", pact_fraud_wcom_checkbox
            CheckBox 20, 220, 195, 10, "Closing due to Returned Mail *", snap_returned_mail_wcom_checkbox
            CheckBox 20, 235, 195, 10, "Closing SNAP and MFIP opening *", snap_to_mfip_wcom_checkbox
            CheckBox 20, 250, 195, 10, "EXP SNAP - Postponed verifs *", postponed_verif_wcom_checkbox
            CheckBox 20, 265, 195, 10, "EXP SNAP - Postponed verif of CAF page 9 Signature *", signature_postponed_verif_wcom_checkbox
            CheckBox 20, 280, 195, 10, "Ineligible Student WCOMs", inelig_student_wcoms_checkbox
            CheckBox 20, 295, 195, 10, "Future Eligibility Begin Date WCOM", future_elig_wcom_checkbox
            CheckBox 20, 325, 60, 10, "CASH Denied *", cash_denied_checkbox
            CheckBox 20, 340, 130, 10, "CASH closing due to Returned Mail*", mfip_returned_mail_wcom_checkbox
            CheckBox 20, 355, 125, 10, "MFIP Closing and SNAP opening *", mfip_to_snap_wcom_checkbox
            CheckBox 10, 375, 100, 10, "PARIS Match - Health Care", paris_match_HC_checkbox
            ButtonGroup ButtonPressed
              OkButton 135, 375, 50, 15
              CancelButton 185, 375, 50, 15
            GroupBox 15, 60, 215, 100, "ABAWD's"
            GroupBox 15, 165, 215, 145, "Other SNAP"
            GroupBox 5, 315, 230, 55, "Cash"
            Text 20, 5, 210, 25, "Select WCOM(s) to add to the notice. Reminder: you can select more than one as required for the case, use multiple categories if necessary. "
            GroupBox 5, 50, 230, 265, "SNAP"
        EndDialog

		' CheckBox 10, 35, 220, 10, "HC - July COLA Income Change Explanation", july_cola_wcom          'this is a TEMP WCOM - need to redesign based on notice type and adding HC WCOMs.
		' CheckBox 25, 150, 140, 10, "Banked Months - E and T voluntary *", banked_mos_vol_e_t_wcom_checkbox
		' CheckBox 25, 165, 175, 10, "Banked Months - Closing for all 9 months used", banked_mos_used_wcom_checkbox
		' CheckBox 25, 180, 145, 10, "Banked Months -  Possibly available", banked_mos_avail_wcom_checkbox
		' GroupBox 20, 135, 195, 60, "Banked Months"

        Dialog Dialog1       'running the dialog to select which WCOMs are going to be added
        cancel_confirmation

        end_of_wcom_line = 0            'setting variables to asses length of WCOM
        end_of_wcom_row = 1

        'setting the arrays to blank for each loop - they will be refilled once the checkboxes are selected again
        ReDim array_of_msg_lines(0)
        ReDim WCOM_TO_WRITE_ARRAY (0)

		If clt_virtual_dropbox_checkbox = checked Then CALL add_words_to_message("You can submit documents Online at www.MNbenefits.mn.gov or Email with documents attachment. EMAIL: hhsews@hennepin.us (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types).")

        If july_cola_wcom = checked Then
            'code for the dialog for PACT closure (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 206, 75, "WCOM Details"
              DropListBox 125, 35, 75, 45, "Select One..."+chr(9)+"RSDI"+chr(9)+"SSI"+chr(9)+"RSDI & SSI", HC_Income_with_COLA
              ButtonGroup ButtonPressed
                OkButton 150, 55, 50, 15
              Text 5, 10, 195, 20, "This WCOM explains that income was increased in January but that changes was disregarded until July."
              Text 5, 35, 115, 10, "Income that increased in January:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If HC_Income_with_COLA = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select which income has a COLA that is now being counted."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("EXPLANATION OF CHANGE OF INCOME:; The change of income listed in this notice is due to " & HC_Income_with_COLA & " increase that went into effect in January. This income increase is disregarded for 6 months and only counted starting in July. The change is only in how income is counted for Health Care and not in the income you receive from " & HC_Income_with_COLA & ".; If you have additional questions please contact us at 612-596-1300.")
        End If

        If snap_pact_wcom_checkbox = checked Then             'SNAP closed with PACT
            'code for the dialog for PACT closure (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 301, 85, "WCOM Details"
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

                Dialog Dialog1
                cancel_confirmation

                If SNAP_close_or_deny = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the case was closed or denied."
                If pact_close_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reasons the SNAP was denied."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("Your SNAP case was " & SNAP_close_or_deny & " because " & pact_close_reason & ".")
        End If

        'Here there is an IF statement for each checkbox - each WCOM may have it's own dialog and the verbiage will be added to the array for the WCOM lines
        If additional_verif_wcom_checkbox = checked Then             'SNAP closed with PACT
            'code for the dialog for PACT closure (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 156, 45, "WCOM Details"
              DropListBox 65, 5, 45, 45, "Select One..."+chr(9)+"CLOSED"+chr(9)+"DENIED", SNAP_close_or_deny
              ButtonGroup ButtonPressed
                OkButton 100, 25, 50, 15
              Text 5, 10, 55, 10, "SNAP case was "
              Text 120, 10, 35, 10, "on PACT."
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If SNAP_close_or_deny = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the case was closed or denied."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 301, 105, "WCOM Details"
              EditBox 5, 25, 290, 15, add_verifs_missing
              EditBox 105, 45, 60, 15, add_verifs_due_date
              If SNAP_close_or_deny = "DENIED" Then
                  EditBox 75, 65, 60, 15, verifs_app_date
                  Text 15, 70, 60, 10, "Application Date:"
              End If
              If SNAP_close_or_deny = "CLOSED" Then
                  EditBox 95, 65, 60, 15, verifs_closure_date
                  Text 15, 70, 80, 10, "Closure Effective Date:"
              End If
              ButtonGroup ButtonPressed
                OkButton 245, 85, 50, 15
              Text 5, 10, 210, 10, "Missing Veifications that caused the SNAP case to be CLOSED:"
              Text 15, 50, 90, 10, "Verifications were due on:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                add_verifs_missing = trim(add_verifs_missing)

                If add_verifs_missing = "" Then err_msg = err_msg & vbNewLine & "* List the verifications that were not received and caused the SNAP to be " & SNAP_close_or_deny & ", or indicate this closure is due to an Ineligible Student."
                If IsDate(add_verifs_due_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verifications were due."
                If SNAP_close_or_deny = "CLOSED" Then
                    If IsDate(verifs_closure_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the effective date of closure."
                End If
                If SNAP_close_or_deny = "DENIED" Then
                    If IsDate(verifs_app_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date of application."
                End If

                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            If SNAP_close_or_deny = "DENIED" Then
                day_60_after_app = DateAdd("d", 60, verifs_app_date)
                CALL add_words_to_message("Your SNAP application has been denied because you did not provide: " & add_verifs_missing & ". This proof was needed by " & add_verifs_due_date & ".  If you need assistance getting this proof please contact your worker at the number listed on this notice by " & day_60_after_app & ".")
            End If

            If SNAP_close_or_deny = "CLOSED" Then
                If DatePart("d", verifs_closure_date) = 1 Then verifs_closure_date = DateAdd("d", -1, verifs_closure_date)
                CALL add_words_to_message("Your SNAP case will close because you did not provide: " & add_verifs_missing & ".  This proof was needed by " & add_verifs_due_date & ".  If you need assistance getting this proof please contact your worker at the number listed on this notice by " & verifs_closure_date & ".")
            End If
        End If

        If pact_fraud_wcom_checkbox = checked Then        'FPI findings indicate another person
            'code for the dialog for closing for fpi result (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 281, 85, "WCOM Details"
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

                Dialog Dialog1
                cancel_confirmation

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
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 126, 85, "WCOM Details"
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

                Dialog Dialog1
                cancel_confirmation

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

        If signature_postponed_verif_wcom_checkbox = checked Then       'postponed signature of CAF for XFS

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 201, 40, "WCOM Details"
              EditBox 60, 20, 55, 15, snap_closure_date_sig
              ButtonGroup ButtonPressed
                OkButton 145, 20, 50, 15
              Text 5, 5, 185, 10, "Postponed Verification of WREG information/exemption"
              Text 5, 25, 50, 10, "SNAP Closure:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If isdate(snap_closure_date_sig) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the day SNAP will close if verifications are not received."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            CALL add_words_to_message("Receipt of signature on application (last page signature required) has been poostponed. Return the application page with signature and date to continue to be eligible for SNAP benefits. If not received, SNAP will close on " & snap_closure_date_sig & ".")
        End If

        If inelig_student_wcoms_checkbox = checked Then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 261, 110, "WCOM Details"
              DropListBox 115, 20, 140, 45, client_dropdown, inelig_student_name
              DropListBox 60, 40, 135, 50, "Select One..."+chr(9)+"PART of a SNAP HH"+chr(9)+"the only member of SNAP HH", inelig_student_HH_detail
              DropListBox 130, 60, 125, 45, "Select One..."+chr(9)+"SNAP E&T education plan"+chr(9)+"Federal or State Work Study", inelig_student_proof
              EditBox 90, 80, 55, 15, verifs_due_date
              ButtonGroup ButtonPressed
                OkButton 205, 80, 50, 15
              Text 10, 5, 215, 10, "WCOMs available for if an ineligible student is on the SNAP case."
              Text 20, 25, 85, 10, "Ineligible Student Name:"
              Text 20, 45, 35, 10, "Student is"
              Text 20, 65, 105, 10, "Student did not provide proof of"
              Text 20, 85, 65, 10, "Verifs were due on"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If inelig_student_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the name of the person who is not eligible for SNAP due to student status."
                If inelig_student_HH_detail = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the student is the only person in the SNAP unit or if the student is with others also in the SNAP unit."
                If inelig_student_proof = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the verification the student should have provided."
                If IsDate(verifs_due_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date the verifications were due on."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            inelig_student_name = right(inelig_student_name, len(inelig_student_name)-5)

            inelig_student_message = ""
            If inelig_student_HH_detail = "PART of a SNAP HH" Then inelig_student_message = inelig_student_name & " is not included in your SNAP unit as an eligible student "
            If inelig_student_HH_detail = "the only member of SNAP HH" Then inelig_student_message = "SNAP is denied because " & inelig_student_name & " is not an eligible student "
            If inelig_student_proof = "SNAP E&T education plan" Then inelig_student_message = inelig_student_message & "and no proof your education plan meets the student requirment of the SNAP Employment & Training (E&T) Program has been received, it was due on " & verifs_due_date & "."
            If inelig_student_proof = "Federal or State Work Study" Then inelig_student_message = inelig_student_message & "and no proof of your work with a Federal or State Work Study program has been received, it was due on  " & verifs_due_date & "."
            inelig_student_message = inelig_student_message & " If you need help getting this proof, please contact your worker at the number listed below."

            CALL add_words_to_message(inelig_student_message)

        End If

        If future_elig_wcom_checkbox = checked Then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 201, 60, "WCOM Details"
              EditBox 85, 20, 50, 15, future_elig_request_date
              EditBox 85, 40, 50, 15, future_elig_begin_date
              ButtonGroup ButtonPressed
                OkButton 145, 40, 50, 15
              Text 5, 5, 185, 10, "Future Eiligibility Begin Date Request Information"
              Text 25, 25, 55, 10, "Date of request:"
              Text 5, 45, 80, 10, "Requested Begin Date:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If IsDate(future_elig_request_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the client requested a future SNAP elig date."
                If IsDate(future_elig_begin_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date SNAP eligibility should begin."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            Call convert_date_into_MAXIS_footer_month(future_elig_begin_date, future_elig_mo, future_elig_yr)

            CALL add_words_to_message("On " & future_elig_request_date & " you verbally requested eligibility of SNAP to begin " & future_elig_begin_date & ".  Because of this, your benefits were denied in the month of application and were determined for " & future_elig_mo & "/" & future_elig_yr & ".")

        End If

        If postponed_verif_wcom_checkbox = checked Then      'XFS Postponed verifs are in WREG
            'code for the dialog for postponed verifs from WREG (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 281, 170, "WCOM Details"
			  EditBox 5, 30, 270, 15, postponed_verifs_needed
			  EditBox 140, 60, 135, 15, abawd_name
			  EditBox 5, 90, 270, 15, wreg_verifs_needed
			  EditBox 60, 130, 55, 15, wreg_verifs_due_date
			  EditBox 60, 150, 55, 15, snap_closure_date
			  ButtonGroup ButtonPressed
			    OkButton 225, 150, 50, 15
			  Text 5, 5, 185, 10, "Postponed Verification details"
			  Text 5, 20, 70, 10, "Verifications Needed:"
			  GroupBox 0, 50, 285, 75, "Postponed WREG Verifications"
			  Text 70, 65, 70, 10, "ABAWD Client Name:"
			  Text 5, 80, 105, 10, "WREG Verifications Needed:"
			  Text 100, 110, 175, 10, "(verifications needed to confirm WREG status ONLY)"
			  Text 5, 135, 40, 10, "Verifs Due:"
			  Text 5, 155, 50, 10, "SNAP Closure:"
			EndDialog


            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

				postponed_verifs_needed = trim(postponed_verifs_needed)
				wreg_verifs_needed = trim(wreg_verifs_needed)
				abawd_name = trim(abawd_name)
                If abawd_name = "" AND wreg_verifs_needed <> "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that has used 3 ABAWD months."
                If wreg_verifs_needed = "" AND postponed_verifs_needed = "" Then err_msg = err_msg & vbNewLine & "* List all WREG verifications needed."
                If isdate(wreg_verifs_due_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the verifications are due."
                If isdate(snap_closure_date) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the day SNAP will close if verifications are not received."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY

			msg_wording = "You are getting SNAP right away because you meet the criteria for expedited SNAP.  You still need to provide the following postponed verification(s) that is needed to continue your SNAP eligibility.  We will mail out a separate verification request form for the requested verification(s). "
			If postponed_verifs_needed <> "" Then msg_wording = msg_wording & "Verifications: " & postponed_verifs_needed & ". "
            If wreg_verifs_needed <> "" Then msg_wording = msg_wording & abawd_name & " has used their 3 entitled months of SNAP benefits as an Able Bodied Adult Without Dependents. Verification of " & wreg_verifs_needed & " has been postponed. "
			msg_wording = msg_wording & "You must turn in all verifications by " & wreg_verifs_due_date & " to continue to be eligible for SNAP benefits. If you do not turn in the required verifications, your case will close on " & snap_closure_date & "."
			Call add_words_to_message(msg_wording)
        End If

        If abawd_child_18_coded_wcom_checkbox = checked Then         'ABAWD exemption for care of child
            'code for the dialog for ABAWD child exemption (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 201, 65, "WCOM Details"
              DropListBox 60, 20, 135, 15, client_dropdown, abawd_exempt_child_18_name
              ButtonGroup ButtonPressed
                OkButton 145, 45, 50, 15
              Text 5, 5, 185, 10, "Client exempt from ABAWD due to child in the SNAP Unit."
              Text 5, 25, 50, 10, "Client Name:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If abawd_exempt_child_18_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that is using child under 18 years exemption."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            abawd_exempt_child_18_name = right(abawd_exempt_child_18_name, len(abawd_exempt_child_18_name)-5)
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message(abawd_exempt_child_18_name & " is exempt from the Able Bodied Adults Without Dependents (ABAWD) Work Requirements due to a child(ren) under the age of 18 in the SNAP unit.")
        End If

        If abawd_child_6_coded_wcom_checkbox = checked Then         'ABAWD exemption for care of child
            'code for the dialog for ABAWD child exemption (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 235, 65, "WCOM Details"
              DropListBox 60, 20, 135, 15, client_dropdown, abawd_exempt_child_6_name
              ButtonGroup ButtonPressed
                OkButton 145, 45, 50, 15
              Text 5, 5, 220, 10, "Client exempt from ABAWD due to child 6 or under in the SNAP Unit."
              Text 5, 25, 50, 10, "Client Name:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If abawd_exempt_child_6_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the name of the client that is using child under 6 years exemption."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            abawd_exempt_child_6_name = right(abawd_exempt_child_6_name, len(abawd_exempt_child_6_name)-5)
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message(abawd_exempt_child_6_name & " is exempt from the Able Bodied Adults Without Dependents (ABAWD) Work Requirements due to caring for a child under the age of 6.")
        End If

        If voluntary_quit_wcom_checkbox = checked Then

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 246, 80, "WCOM Details"
              DropListBox 85, 20, 150, 45, client_dropdown, vol_quit_name
              DropListBox 85, 40, 150, 45, "Select One..."+chr(9)+"voluntarily quit"+chr(9)+"reduced work hours below 30 per week"+chr(9)+"refused suitable employment", vol_quit_sanction_reason
              ButtonGroup ButtonPressed
                OkButton 185, 60, 50, 15
              Text 5, 5, 120, 10, "Select voluntary quit details:"
              Text 10, 25, 70, 10, "Name of the member:"
              Text 10, 45, 70, 10, "Reason for sanction:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If vol_quit_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Choose the name of the client whole coluntarily quit."
                If vol_quit_sanction_reason = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the reason for the voluntary quit sanction."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            vol_quit_name = right(vol_quit_name, len(vol_quit_name)-5)

            CALL add_words_to_message(vol_quit_name & " is sanctioned from SNAP because they have " & vol_quit_sanction_reason & ". They are sanctioned until they return to the same job, they accept similar employment or they become exempt from work registration for a reason other than receiving Unemployment Compensation.")

        End If

        If temp_disa_abawd_wcom_checkbox = checked Then       'Verified temporary disa for ABAWD exemption
            'code for the dialog for temporary disa for ABAWD (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 211, 80, "WCOM Details"
              DropListBox 75, 20, 130, 45, client_dropdown, temp_disa_memb_name
              EditBox 185, 40, 20, 15, numb_disa_mos
              ButtonGroup ButtonPressed
                OkButton 155, 60, 50, 15
              Text 5, 5, 120, 10, "DISA indicated on form from Doctor"
              Text 10, 25, 60, 10, "Disabled Member"
              Text 80, 45, 105, 10, "Number of months of disability"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If temp_disa_memb_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Choose the ABAWD Client."
                If trim(numb_disa_mos) = "" Then err_msg = err_msg & vbNewLine & "* Enter the number of months the disability is expected to last from the doctor's information."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            temp_disa_memb_name = right(temp_disa_memb_name, len(temp_disa_memb_name)-5)
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message(temp_disa_memb_name & " is exempt from the ABAWD work provision because you are unable to work for " & numb_disa_mos & " months per your Doctor statement.")
        End If

        If voluntary_e_t_wcom_checkbox = checked Then

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 171, 60, "WCOM Details"
              DropListBox 30, 20, 130, 45, client_dropdown, abawd_memb_name
              ButtonGroup ButtonPressed
                OkButton 115, 40, 50, 15
              Text 5, 5, 160, 10, "ABAWD Client to advise about Voluntary E and T"
              Text 5, 25, 25, 10, "Client:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If abawd_memb_name = "Select One..." Then err_msg = err_msg & vbNewLine & "* Choose the ABAWD Client."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            abawd_memb_name = right(abawd_memb_name, len(abawd_memb_name)-5)

            CALL add_words_to_message("Minnesota has changed the rules for time-limited SNAP recipients. " & abawd_memb_name & " is not required to participate in SNAP Employment and Training (SNAP E&T), but may choose to. Participation in SNAP E&T may extend your SNAP benefits and offer you support as you seek employment. Ask your worker about SNAP E&T.")
        End If

        If (abawd_homeless_wcom_checkbox = checked OR banked_mos_avail_wcom_checkbox = checked OR banked_mos_vol_e_t_wcom_checkbox = checked OR banked_mos_used_wcom_checkbox = checked) AND voluntary_e_t_wcom_checkbox = unchecked THen
            CALL add_words_to_message("You receive time-limited SNAP as you are an ABAWD (Able-bodied adult without dependents).")
        End If

        If abawd_homeless_wcom_checkbox = checked Then
            CALL add_words_to_message("You previously reported that you are homeless, specifically defined for this purpose as lacking both:; *Fixed/regular nighttime residence (inc. temporary housing); *Access to work-related necessities (shower/laundry/etc.); Based on this information, you may qualify for SNAP benefits that are not time-limited. If you believe you meet the homeless and unfit for employment exemption (or any other exemption), please contact your team.")
        End If

        'currently no checkbox for this one - we should never be using it as client's no longer need to request banked months
        If banked_mos_avail_wcom_checkbox = checked Then      'ABAWD expired - Banked Months available - NO WORKER INPUT NEEDED
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("You have used all of your available ABAWD months. You may be eligible for SNAP banked months. Please contact your team if you have questions.")
        End If

        If banked_mos_vol_e_t_wcom_checkbox = checked Then       'Banked Months closing due to FSET non-coop - NO WORKER INPUT NEEDED
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("Working with Employment and Training is voluntary as a banked months recipient, if you would like to work with Employment and Training, please contact your team.")
            'CALL add_words_to_message("You have been approved to receive additional SNAP benefits under the SNAP Banked Months program. Working with Employment and Training is voluntary under this program. If you'd like work with Employment and Training, please contact your team.")
        End If

        If banked_mos_used_wcom_checkbox = checked Then       'Banked Months expired - NO WORKER INPUT NEEDED
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your team if you have questions.")
        End If

        If cash_denied_checkbox = checked Then              'Information to add to CASH denial

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 321, 105, "WCOM Details"
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

                Dialog Dialog1
                cancel_confirmation

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
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 126, 85, "WCOM Details"
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

                Dialog Dialog1
                cancel_confirmation

                if isdate(rm_sent_date_cash) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the request for address information was sent."
                if isdate(rm_due_date_cash) = False Then err_msg = err_msg & vbNewLine & "*Enter a valid date for when the response for address information was due."
                if err_msg <> "" Then msgBox "Resolve to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("Your mail has been returned to our agency. On " & rm_sent_date_cash & " you were sent a request for you to contact this agency because of this returned mail. You can avoid having your case closed if you contact this agency by " & rm_due_date_cash & ".")
        End If

        If mfip_to_snap_wcom_checkbox = checked Then      'MFIP is closing and SNAP is opening
            'code for the dialog for MFIP closure reason when SNAP is reassessed (this dialog has the same name in each IF to prevent the over 7 dialog error)
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 166, 80, "WCOM Details"
              EditBox 5, 40, 155, 15, MFIP_closing_reason
              ButtonGroup ButtonPressed
                OkButton 110, 60, 50, 15
              Text 5, 5, 155, 10, "MFIP is closing and SNAP has been assessed"
              Text 5, 25, 105, 10, "Why is MFIP closing:"
            EndDialog

            Do                          'displaying the dialog and ensuring that all required information is entered
                err_msg = ""

                Dialog Dialog1
                cancel_confirmation

                If MFIP_closing_reason = "" Then err_msg = err_msg & vbNewLine & "*List all reasons why SNAP is closing."
                If err_msg <> "" Then MsgBox "Resolve the following to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            'Adding the verbiage to the WCOM_TO_WRITE_ARRAY
            CALL add_words_to_message("You are no longer eligible for MFIP because " & MFIP_closing_reason & ".")
        End If

        If paris_match_HC_checkbox = checked then CALL add_words_to_message("You do not qualify for Medical Assistance because you are not a Minnesota resident. (Code of Federal Regulations, title 42, section 435.403)")

        'This assesses if the message generated is too long for WCOM. If so then the checklist will reappear along with each selected WCOM dialog so it can be changed
        If UBOUND(WCOM_TO_WRITE_ARRAY) > 14 Then big_err_msg = big_err_msg & vbNewLine & "The amount of text/information that is being added to WCOM will exceed the 15 lines available on MAXIS WCOMs. Please reduce the number of WCOMs that have been selected or reduce the amount of text in the selected WCOM."

        If big_err_msg <> "" Then
            MsgBox "*** Please resolved the following to continue ***" & vbNewLine & big_err_msg
        Else
            'Leave this here - testing purposes
            wcom_to_display = ""
            For each msg_line in WCOM_TO_WRITE_ARRAY
                if wcom_to_display = "" Then
                    wcom_to_display = msg_line
                else
                    wcom_to_display = wcom_to_display & vbNewLine & msg_line
                end if
            Next
            review_wcom_text = MsgBox("The WCOM will read:" & vbNewLine & vbNewLine & wcom_to_display, vbOkCancel, "Review the WCOM Text")
            If review_wcom_text = vbCancel then script_end_procedure("")
        End If
    Loop until big_err_msg = ""
    call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

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

        PF4     'Save the WCOM
        PF3     'Exit the WCOM

        back_to_self
    End If
Next

'Now the action will be case noted
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

start_a_blank_case_note

CALL write_variable_in_CASE_NOTE("*** Added WCOM for to Notice to clarify action taken ***")
CALL write_variable_in_CASE_NOTE("Information added to the following WCOM notices in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ":")
For notices_listed = 0 to UBound(NOTICES_ARRAY, 2)
    If NOTICES_ARRAY(selected, notices_listed) = checked Then
        CALL write_variable_in_CASE_NOTE("* " & NOTICES_ARRAY(information, notices_listed))
    End If
Next
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE("Detail added to each notice:")

If clt_virtual_dropbox_checkbox Then CALL write_variable_in_CASE_NOTE("* Information about online document submission options.")
If july_cola_wcom = checked Then CALL write_variable_in_CASE_NOTE("* HC income budgeted increase as COLA disregards has ended for " & HC_Income_with_COLA & ".")
If snap_pact_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("* SNAP case was " & SNAP_close_or_deny & " because " & pact_close_reason & ".")
If pact_fraud_wcom_checkbox Then CALL write_variable_in_CASE_NOTE("* Request to add: " & new_hh_memb & ". Verification needed: " & new_memb_verifs & ". Verification not received causing closure.")
If snap_returned_mail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Returned mail was received. Verification request sent: " & rm_sent_date_snap & " and Due: " & rm_due_date_snap & " with no response caused SNAP case closure.")
If snap_to_mfip_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* SNAP is closing but food will be included in MFIP.")
If duplicate_assistance_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Duplicate assistance from state of: " & dup_state & " during the month of " & dup_month & "/" & dup_year & " was received.")
If dup_assistance_in_MN_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Duplicate assistance on a case in MN during the month of " & mn_dup_month & "/" & mn_dup_year & ".")
If client_death_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Closure due to client death.")
If signature_postponed_verif_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Postponed signature on CAF page 9.")
If inelig_student_wcoms_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Ineligible Student Information about " & inelig_student_name & " - proof needed: " & inelig_student_proof & ".")
If future_elig_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* SNAP eligibility begin date reqeusted to be changed to: " & future_elig_begin_date & ". Request made on " & future_elig_request_date & ".")
If postponed_verif_wcom_checkbox = checked Then
	CALL write_variable_in_CASE_NOTE("* Postponed verifications requested, due: " & wreg_verifs_due_date)
	If postponed_verifs_needed <> "" Then CALL write_variable_in_CASE_NOTE("   -Verifications: " & postponed_verifs_needed & ".")
	If wreg_verifs_needed <> "" Then CALL write_variable_in_CASE_NOTE("   -" & abawd_name & " has used their 3 ABAWD months. Postponed WREG verification: " & wreg_verifs_needed & ".")
End If
If abawd_child_18_coded_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & abawd_exempt_child_18_name & " is ABAWD and WREG exempt due to a child(ren) under the age of 18 in the SNAP unit.")
If abawd_child_6_coded_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & abawd_exempt_child_6_name & " is ABAWD and WREG exempt due to care of a child under 6.")
If voluntary_quit_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & vol_quit_name & " is sanctioned from SNAP due to: " & vol_quit_sanction_reason & ".")
If additional_verif_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Verifs not provided: " & add_verifs_missing & ", which were due on " & add_verifs_due_date & ".")
If temp_disa_abawd_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & temp_disa_memb_name & " meets ABAWD exemption of temporary inability to work for " & numb_disa_mos & " months per Doctor statement.")
If voluntary_e_t_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Voluntary SNAP E&T offered to " & abawd_memb_name & ".")
If abawd_homeless_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Information about ABAWD Exemption for homelessness.")
If banked_mos_avail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* ABAWD months have been used, explained Banked Months may be available.")
If banked_mos_vol_e_t_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* E&T is voluntary with Banked Months.")
If banked_mos_non_coop_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* " & banked_abawd_name & " was receiving Banked Months and fail cooperation with E & T. Explained requesting Good Cause, and future banked months ineligibility.")
If banked_mos_used_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Client has exhausted all available Banked Months. Advised to review other WREG/ABAWD exemptions.")
If cash_denied_checkbox = checked Then
    CALL write_variable_in_CASE_NOTE("* " & cash_one_program & " denied because " & cash_one_reason & ".")
    If cash_two_program <> "Select one..." Then CALL write_variable_in_CASE_NOTE("* " & cash_two_program & " denied because " & cash_two_reason & ".")
End If
If mfip_returned_mail_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* Returned mail was received. Verification request sent: " & rm_sent_date_cash & ", cash denied and can be reopened if verif received by: " & rm_due_date_cash & ".")
If mfip_to_snap_wcom_checkbox = checked Then CALL write_variable_in_CASE_NOTE("* MFIP closure due to: " & MFIP_closing_reason & ".")
If paris_match_HC_checkbox = checked then Call write_variable_in_CASE_NOTE("* Information about health care closure due to lack of state residency.")

CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("WCOMs added to Notice and case note created.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/20/2021
'--Tab orders reviewed & confirmed----------------------------------------------10/20/2021
'--Mandatory fields all present & Reviewed--------------------------------------10/20/2021
'--All variables in dialog match mandatory fields-------------------------------09/02/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/20/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------10/20/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/20/2021
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/02/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------09/02/2021
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/02/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------09/02/2021
'--Script name reviewed---------------------------------------------------------09/02/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/02/2021
'--comment Code-----------------------------------------------------------------09/02/2021
'--Update Changelog for release/update------------------------------------------09/02/2021
'--Remove testing message boxes-------------------------------------------------09/02/2021
'--Remove testing code/unnecessary code-----------------------------------------09/02/2021
'--Review/update SharePoint instructions----------------------------------------10/20/2021
'--Review Best Practices using BZS page ----------------------------------------09/02/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/02/2021
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/20/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/02/2021
