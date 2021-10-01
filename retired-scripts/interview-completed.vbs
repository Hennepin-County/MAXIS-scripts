'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INTERVIEW COMPLETED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/30/2019", "Updated CAF type options to 'Application', 'Recertification' and 'Addendum'", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2018", "Added Outlook reminder for the 1st of the next month if case is unable to be updated for recerts due to renewal month being CM plus one.", "Ilse Ferris, Hennepin County")
call changelog_update("03/02/2018", "Interview header updated for on demand waiver handling.", "MiKayla Handley, Hennepin County")
call changelog_update("09/25/2017", "Updated to allow for cases that do not have CAF date listed on STAT/REVW.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = datepart("m", date)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = "" & datepart("yyyy", date) - 2000

'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 95, 25, 20, 15, MAXIS_footer_month
  EditBox 120, 25, 20, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_checkbox
  CheckBox 50, 60, 30, 10, "HC", HC_checkbox
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 135, 60, 35, 10, "EMER", EMER_checkbox
  DropListBox 70, 80, 75, 15, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"Addendum", CAF_type
  ButtonGroup ButtonPressed
    OkButton 35, 100, 50, 15
    CancelButton 95, 100, 50, 15
  Text 35, 10, 45, 10, "Case number:"
  Text 30, 30, 65, 10, "Footer month/year:"
  GroupBox 5, 45, 170, 30, "Programs applied for"
  Text 30, 85, 35, 10, "CAF type:"
EndDialog

Do
	Do
		err_msg = ""
  		Dialog Dialog1 'Runs the first dialog that gathers program information and case number
  		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine &  "* You need to type a valid case number."
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If cash_checkbox = 0 AND HC_checkbox = 0 AND SNAP_checkbox = 0 AND EMER_checkbox = 0 then err_msg = err_msg & vbNewLine & "* Select at least one program."
  		If CAF_type = "Select One..." then err_msg = err_msg & vbNewLine &  "* You must select the type of CAF you interviewed"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false							'Loops until we affirm that we're ready to case note.

'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")

'Because some cases don't have HCRE dates listed, so when you try to go past PROG the script gets caught up. Do...loop handles this instance.
PF3		'exits PROG to prompt HCRE if HCRE isn't complete
Do
	EMReadscreen HCRE_panel_check, 4, 2, 50
	If HCRE_panel_check = "HCRE" then
		PF10	'exists edit mode in cases where HCRE isn't complete for a member
		PF3
	END IF
Loop until HCRE_panel_check <> "HCRE"		'repeats until case is not in the HCRE panel

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
If CAF_type = "Recertification" then
  Call navigate_to_MAXIS_screen("STAT", "REVW")
	EMReadScreen recd_date, 8, 13, 37
	If recd_date = "__ __ __" then
	  CAF_datestamp = ""
	Else
		call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
	End if
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
	IF DateDiff ("d", CAF_datestamp, date) > 60 THEN CAF_datestamp = ""							'This will disregard Application Dates that are older than 60 days. IF and old dste is pulled, the next dialog will require the worker to enter the correct date
End if

If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "Cash, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "Emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

interview_date = date & ""		'Defaults the date of the interview to today's date.

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 451, 280, "Interview Dialog"
  EditBox 65, 5, 75, 15, caf_datestamp
  EditBox 205, 5, 75, 15, interview_date
  ComboBox 345, 5, 100, 15, "Office"+chr(9)+"Phone", interview_type
  EditBox 45, 25, 400, 15, HH_comp
  EditBox 60, 45, 385, 15, earned_income
  EditBox 75, 65, 370, 15, unearned_income
  EditBox 45, 85, 400, 15, expenses
  EditBox 35, 105, 410, 15, assets
  CheckBox 15, 140, 135, 10, "Check here if this case is expedited.", expedited_checkbox
  EditBox 140, 155, 300, 15, why_xfs
  EditBox 185, 175, 255, 15, reason_expedited_wasnt_processed
  EditBox 50, 200, 395, 15, other_notes
  EditBox 60, 220, 385, 15, verifs_needed
  EditBox 230, 240, 105, 15, worker_signature
  CheckBox 10, 260, 425, 10, "Check here to set an Outlook reminder to update STAT/REVW after 1st of the next month (if client completes the recert interview).", outlook_check
  ButtonGroup ButtonPressed
    OkButton 340, 240, 50, 15
    CancelButton 395, 240, 50, 15
  Text 150, 10, 50, 10, "Interview Date:"
  Text 290, 10, 50, 10, "Interview Type:"
  Text 5, 30, 35, 10, "HH Comp:"
  Text 5, 50, 55, 10, "Earned Income:"
  Text 5, 70, 60, 10, "Unearned Income:"
  Text 5, 90, 35, 10, "Expenses:"
  Text 5, 110, 25, 10, "Assets:"
  GroupBox 5, 125, 440, 70, "Expedited SNAP"
  Text 15, 160, 125, 10, "Explain why case is expedited or not:"
  Text 15, 180, 165, 10, "Reason expedited wasn't processed (if applicable) "
  Text 5, 205, 45, 10, "Other Notes:"
  Text 170, 245, 60, 10, "Worker Signature"
  Text 5, 225, 50, 10, "Verifs Needed:"
  Text 5, 10, 55, 10, "CAF Datestamp:"
EndDialog
DO
	Do
		err_msg = ""
		Dialog Dialog1			'Displays the Interview Dialog
		cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
		If CAF_datestamp = "" or isDate(CAF_datestamp) = False THEN err_msg = err_msg & vbcr & "* Enter a valid application datestamp."
        If interview_date = "" or isDate(interview_date) = False THEN err_msg = err_msg & vbcr & "* Enter a valid interview date."
		IF (SNAP_checkbox = checked) AND (why_xfs = "") THEN err_msg = err_msg & vbcr & "* SNAP is pending, you must explain your Expedited Determination"
        IF worker_signature = "" THEN err_msg = err_msg & vbcr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false							'Loops until we affirm that we're ready to case note.

'This code will update the interview date in PROG.
If CAF_type <> "Recertification" AND CAF_type <> "Addendum" Then        'Interview date is not on PROG for recertifications or addendums
    If SNAP_checkbox = checked OR cash_checkbox = checked Then          'Interviews are only required for Cash and SNAP
        intv_date_needed = FALSE
        Call navigate_to_MAXIS_screen("STAT", "PROG")                   'Going to STAT to check to see if there is already an interview indicated.

        If SNAP_checkbox = checked Then                                 'If the script is being run for a SNAP interview
            EMReadScreen entered_intv_date, 8, 10, 55                   'REading what is entered in the SNAP interview
            'MsgBox "SNAP interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE  'If this is blank - the script needs to prompt worker to update it
        End If

        If cash_checkbox = checked THen                             'If the script is bring run for a Cash interview
            EMReadScreen cash_one_app, 8, 6, 33                     'First the script needs to identify if it is cash 1 or cash 2 that has the application information
            EMReadScreen cash_two_app, 8, 7, 33

            cash_one_app = replace(cash_one_app, " ", "/")          'Turning this in to a date format
            cash_two_app = replace(cash_two_app, " ", "/")

            If cash_one_app <> "__/__/__" Then      'Error handling - VB doesn't like date comparisons with non-dates
                if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6     'If date of application on PROG matches script date of applicaton
            End If
            If cash_two_app <> "__/__/__" Then
                if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
            End If

            EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
            'MsgBox "Cash interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
        End If

        If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
            If SNAP_checkbox = checked Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
            If cash_checkbox = checked Then prog_update_cash_checkbox = checked

			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
            BeginDialog Dialog1, 0, 0, 231, 130, "Update PROG?"
              OptionGroup RadioGroup1
                RadioButton 10, 10, 155, 10, "YES! Update PROG with the Interview Date", confirm_update_prog
                RadioButton 10, 60, 90, 10, "No, do not update PROG", do_not_update_prog
              EditBox 165, 5, 50, 15, interview_date
              CheckBox 25, 25, 30, 10, "SNAP", prog_update_SNAP_checkbox
              CheckBox 25, 40, 30, 10, "CASH", prog_update_cash_checkbox
              Text 20, 75, 200, 10, "Reason PROG should not be updated with the Interview Date:"
              EditBox 20, 90, 195, 15, no_update_reason
              ButtonGroup ButtonPressed
                OkButton 175, 110, 50, 15
            EndDialog

            DO    'Running the dialog
                Do
                    err_msg = ""
                    Dialog Dialog1
                    'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                    If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                    IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."
		    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		    	LOOP UNTIL err_msg = ""
		    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		    Loop until are_we_passworded_out = false					'loops until user passwords back in

            If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
                CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

                keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
                keep_footer_year = MAXIS_footer_year

                app_month = DatePart("m", CAF_datestamp)    'Setting the footer month and year to the app month.
                app_year = DatePart("yyyy", CAF_datestamp)

                MAXIS_footer_month = right("00" & app_month, 2)
                MAXIS_footer_year = right(app_year, 2)

                CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
                PF9                                             'Edit

                intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
                intv_day = DatePart("d", interview_date)
                intv_yr = DatePart("yyyy", interview_date)

                intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
                intv_day = right("00"&intv_day, 2)
                intv_yr = right(intv_yr, 2)

                If prog_update_SNAP_checkbox = checked Then     'If it was selected to SNAP interview to be updated
                    programs_w_interview = "SNAP"               'Setting a variable for case noting

                    EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
                    EMWriteScreen intv_day, 10, 58
                    EMWriteScreen intv_yr, 10, 61
                End If

                If prog_update_cash_checkbox = checked Then     'If it was selected to update for Cash
                    If programs_w_interview = "" Then programs_w_interview = "CASH"     'variable for the case note
                    If programs_w_interview <> "" Then programs_w_interview = "SNAP and CASH"
                    EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
                    EMReadScreen cash_two_app, 8, 7, 33

                    cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
                    cash_two_app = replace(cash_two_app, " ", "/")

                    If cash_one_app <> "__/__/__" Then              'Comparing them to the date of application to determine which row to use
                        if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6
                    End If
                    If cash_two_app <> "__/__/__" Then
                        if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                    End If

                    EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                    EMWriteScreen intv_day, prog_row, 58
                    EMWriteScreen intv_yr, prog_row, 61
                End If

                transmit                                    'Saving the panel

                MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
                MAXIS_footer_year = keep_footer_year
            End If
        ENd If
    End If
End If
'MsgBox confirm_update_prog

If outlook_check = 1 then
    reminder_date = CM_plus_1_mo & "/01/" & CM_plus_1_yr
    'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
    Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "Update received date and interview date on STAT/REVW for: " & MAXIS_case_number, "", "", TRUE, 5, "")
    Outlook_remider = True
End if

'Navigates to case note, and checks to make sure we aren't in inquiry.
start_a_blank_CASE_NOTE

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " Recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("~ Interview Completed ~")
CALL write_variable_in_CASE_NOTE ("** Case note for Interview only - full case note of CAF processing to follow.")
CALL write_bullet_and_variable_in_CASE_NOTE("Application Type", CAF_type)
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
CALL write_bullet_and_variable_in_CASE_NOTE("CAF Datestamp", CAF_datestamp)
CALL write_variable_in_CASE_NOTE("* Interview type: " & interview_type & " - Interview date: " & interview_date)
If confirm_update_prog = 1 Then CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG for " & programs_w_interview)
If do_not_update_prog = 1 Then CALL write_bullet_and_variable_in_CASE_NOTE("PROG WAS NOT UPDATED WITH INTERVIEW DATE, because", no_update_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned Income", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Unearned Income", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Expenses", expenses)
CALL write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
IF (expedited_checkbox = unchecked) AND (SNAP_checkbox = checked) THEN CALL write_variable_in_CASE_NOTE ("* NOT Expedited SNAP")
CALL write_bullet_and_variable_in_CASE_NOTE ("Explanation of Expedited Determination", why_xfs)		'Worker can detail how they arrived at if client is expedited or not - particularly useful if different from screening
CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
If Outlook_remider = True then call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
end_msg = "Success! Interview has been successfully noted. Once processing is completed remember to run the CAF Script for detailed case note."
If do_not_update_prog = 1 Then end_msg = end_msg & vbNewLine & vbNewLine & "It was selected that PROG would NOT be updated because " & no_update_reason
script_end_procedure(end_msg)
