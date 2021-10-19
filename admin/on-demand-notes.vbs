'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ON DEMAND.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("10/01/2021", "GitHub #189 Updated script to remove correction email process.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
CALL Check_for_MAXIS(false)                         'Ensuring we are not passworded out
back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

closing_message = "On Demand Application Waiver process has been case noted." 'setting up closing_message variable for possible additions later based on conditions

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 226, 160, "Notes On Demand"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 175, 25, 50, 15, application_date
  EditBox 175, 45, 50, 15, notice_sent_date
  EditBox 175, 65, 50, 15, case_note_date
  EditBox 175, 85, 50, 15, interview_date
  CheckBox 5, 90, 75, 10, "Update STAT/PROG", update_PROG_CHECKBOX
  DropListBox 55, 105, 170, 15, "Select One:"+chr(9)+"Case was not pended timely"+chr(9)+"Client completed application interview"+chr(9)+"Client has not completed application interview"+chr(9)+"Denied programs for no interview"+chr(9)+"Interview not needed for MFIP to SNAP transition"+chr(9)+"Other(please describe)", case_status_dropdown
  EditBox 55, 120, 170, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 120, 140, 50, 15
    CancelButton 175, 140, 50, 15
    PushButton 140, 5, 60, 15, "CASE/NOTE", CASE_NOTE_button
  Text 5, 50, 60, 10, "Date notice sent:"
  Text 5, 30, 65, 10, "Date of application:"
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 125, 45, 10, "Other notes:"
  Text 5, 70, 60, 10, "Date of case note:"
  Text 105, 90, 60, 10, "Date of interview:"
  Text 5, 110, 45, 10, "Case status:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	'Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF case_status_dropdown = "Case was not pended timely" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid NOMI date."
		IF case_status_dropdown = "Denied programs for no interview" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid NOMI date."
		IF case_status_dropdown = "Client completed application interview" and update_PROG_CHECKBOX <> CHECKED THEN err_msg = err_msg & vbNewLine & "* Please update PROG with interview date."
		IF case_status_dropdown = "Client completed application interview" and application_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case note date."
		IF case_status_dropdown = "Client completed application interview" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid interview date."
		IF case_status_dropdown = "Client has not completed application interview" THEN
		 	IF interview_date = "" or isDate(interview_date) = False THEN err_msg = err_msg & vbcr & "* Pleasse enter a valid interview date."
		END IF
		'IF update_PROG_CHECKBOX = CHECKED THEN
		IF case_status_dropdown = "Other(please describe)" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a description of what occurred."
        IF ButtonPressed = CASE_NOTE_button then call navigate_to_MAXIS_screen("CASE", "NOTE")
	'Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Checking for PRIV cases.
'If update_PROG_CHECKBOX = CHECKED THEN          'Interviews are only required for Cash and SNAP
    intv_date_needed = FALSE
	'Checking for PRIV cases.
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)  'Going to STAT to check to see if there is already an interview indicated.
	IF is_this_priv = TRUE THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
    EMReadScreen entered_intv_date, 8, 10, 55                   'Reading what is entered in the SNAP interview
    MsgBox "SNAP interview date - " & entered_intv_date
    If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE  'If this is blank - the script needs to prompt worker to update it
    EMReadScreen cash_one_app, 8, 6, 33                     'First the script needs to identify if it is cash 1 or cash 2 that has the application information
    EMReadScreen cash_two_app, 8, 7, 33
    EMReadScreen grh_cash_app, 8, 9, 33
    cash_one_app = replace(cash_one_app, " ", "/")          'Turning this in to a date format
    cash_two_app = replace(cash_two_app, " ", "/")
    grh_cash_app = replace(grh_cash_app, " ", "/")
    If cash_one_app <> "__/__/__" Then      'Error handling - VB doesn't like date comparisons with non-dates
        If IsDate(cash_one_app) = TRUE Then
            if DateDiff("d", cash_one_app, application_date) = 0 then prog_row = 6     'If date of application on PROG matches script date of application
        End If
    End If
    If cash_two_app <> "__/__/__" THEN
        If IsDate(cash_two_app) = TRUE Then
            if DateDiff("d", cash_two_app, application_date) = 0 then prog_row = 7
        End If
    End If
    If grh_cash_app <> "__/__/__" THEN
        If IsDate(grh_cash_app) = TRUE THen
            if DateDiff("d", grh_cash_app, application_date) = 0 then prog_row = 9
        End If
    End If
    EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
    'MsgBox "Cash interview date - " & entered_intv_date
    If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
	If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
        If the_process_for_snap = "Application" Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
        If the_process_for_cash = "Application" Then prog_update_cash_checkbox = checked


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
        'Running the dialog
        Do
            Do
                err_msg = ""
                Dialog Dialog1
                'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."
                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE

        If confirm_update_prog = 1 THEN     'If the dialog selects to have PROG updated
            CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month
            keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
            keep_footer_year = MAXIS_footer_year
            app_month = DatePart("m", application_date)    'Setting the footer month and year to the app month.
            app_year = DatePart("yyyy", application_date)
            MAXIS_footer_month = right("00" & app_month, 2)
            MAXIS_footer_year = right(app_year, 2)
            Call back_to_SELF
            CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
            PF9                                             'Edit
            intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
            intv_day = DatePart("d", interview_date)
            intv_yr = DatePart("yyyy", interview_date)
            intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
            intv_day = right("00"&intv_day, 2)
            intv_yr = right(intv_yr, 2)
            intv_date_to_check = intv_mo & " " & intv_day & " " & intv_yr
            If prog_update_SNAP_checkbox = checked THEN     'If it was selected to SNAP interview to be updated
                programs_w_interview = "SNAP"               'Setting a variable for case noting
                EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
                EMWriteScreen intv_day, 10, 58
                EMWriteScreen intv_yr, 10, 61
            End If
            If prog_update_cash_checkbox = checked THEN     'If it was selected to update for Cash
                If programs_w_interview = "" THEN programs_w_interview = "CASH"     'variable for the case note
                If programs_w_interview <> "" THEN programs_w_interview = "SNAP and CASH"
                EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
                EMReadScreen cash_two_app, 8, 7, 33
                EMReadScreen grh_cash_app, 8, 9, 33
                cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
                cash_two_app = replace(cash_two_app, " ", "/")
                grh_cash_app = replace(grh_cash_app, " ", "/")
                If cash_one_app <> "__/__/__" THEN              'Comparing them to the date of application to determine which row to use
                    If IsDate(cash_one_app) = TRUE THEN
                        if DateDiff("d", cash_one_app, application_date) = 0 then prog_row = 6
                    End If
                End If
                If cash_two_app <> "__/__/__" THEN
                    If IsDate(cash_two_app) = TRUE THEN
                        if DateDiff("d", cash_two_app, application_date) = 0 then prog_row = 7
                    End If
                End If
                If grh_cash_app <> "__/__/__" THEN
                    If IsDate(grh_cash_app) = TRUE THEN
                        if DateDiff("d", grh_cash_app, application_date) = 0 then prog_row = 9
                    End If
                End If
                EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                EMWriteScreen intv_day, prog_row, 58
                EMWriteScreen intv_yr, prog_row, 61
            End If
            transmit                                    'Saving the panel
            Call HCRE_panel_bypass
            Call back_to_SELF
            Call MAXIS_background_check
            MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
            MAXIS_footer_year = keep_footer_year
        End If
    END If
    If intv_date_needed = TRUE and confirm_update_prog = 1 THEN         'If previous code has determined that PROG needs to be updated
        snap_intv_date_updated = FALSE
        cash_intv_date_updated = FALSE
        show_prog_update_failure = FALSE
        Call back_to_SELF
        CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
        If prog_update_SNAP_checkbox = checked Then
            EMReadScreen new_snap_intv_date, 8, 10, 55
            If new_snap_intv_date = intv_date_to_check Then snap_intv_date_updated = TRUE
            If snap_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
        End If
        If prog_update_cash_checkbox = checked THEN
            EMReadScreen new_cash_intv_date, 8, prog_row, 55
            If new_cash_intv_date = intv_date_to_check Then cash_intv_date_updated = TRUE
            If cash_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
        End If
        If show_prog_update_failure = TRUE THEN
            fail_msg = "You have requested the script update PROG for "
            If prog_update_SNAP_checkbox = checked AND prog_update_cash_checkbox = checked THEN
                fail_msg = fail_msg & "Cash and SNAP "
            ElseIf prog_update_SNAP_checkbox = checked THEN
                fail_msg = fail_msg & "SNAP "
            ElseIf prog_update_cash_checkbox = checked THEN
                fail_msg = fail_msg & "Cash "
            End If
            fail_msg = fail_msg & "to enter the interview date on PROG." & vbCr & vbCr & "The script was unable to update PROG completely." & vbCr
            If prog_update_SNAP_checkbox = checked THEN
                fail_msg = fail_msg & " - The SNAP Interview Date was not entered." & vbCr
            ElseIf prog_update_cash_checkbox = checked THEN
                fail_msg = fail_msg & " - The Cash Interview Date was not entered." & vbCr
            End If
            fail_msg = fail_msg & vbCr & "The PROG panel will need to be updated manually with the interview information."
            MsgBox fail_msg
        End If
    End If
End If


denial_date = dateadd("d", 10, denial_date)

'NOW WE START CASE NOTING - there are a few
start_a_blank_case_note
IF case_status_dropdown = "Client completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " on "  & interview_date & " PROG updated ~")
	CALL write_variable_in_CASE_NOTE("* Completed by previous worker per case note dated: " & case_note_date)
ELSEIF case_status_dropdown = "Client has not completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " ~")
	CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
	CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & interview_date)
	CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview.")
    Call write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it, because the case was not pended timely a NOMI still needs to be sent and adequate time provided to the client to comply. Denial can be done after " & denial_date)
ELSEIF case_status_dropdown = "Case was not pended timely" THEN
    CALL write_variable_in_CASE_NOTE("~ Client has not completed application interview ~")
    CALL write_variable_in_CASE_NOTE("* Application date:" & application_date)
    CALL write_variable_in_CASE_NOTE("* NOMI sent to client on:" & interview_date)
    CALL write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it. Because the case was not pended timely a NOMI still needs to be sent and adequate time provided to the client to comply.")
ELSEIF case_status_dropdown = "Denied programs for no interview" THEN
    CALL write_variable_in_CASE_NOTE("* Application date:" & application_date)
    CALL write_variable_in_CASE_NOTE("* Reason for denial: interview was not completed timely")
    CALL write_variable_in_CASE_NOTE("* NOMI sent to client on:" & interview_date)
   	CALL write_variable_in_CASE_NOTE("* Confirmed client was provided sufficient 10 day notice.")
ELSEIF case_status_dropdown = "Interview not needed for MFIP to SNAP transition" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " ~")
	CALL write_variable_in_CASE_NOTE("* MFIP to SNAP transition no interview required updated PROG to reflect this")
END IF
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE (worker_signature)
PF3 'to save the case note'

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/01/2021
'--Tab orders reviewed & confirmed----------------------------------------------10/01/2021
'--Mandatory fields all present & Reviewed--------------------------------------10/01/2021
'--All variables in dialog match mandatory fields-------------------------------10/01/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/01/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/01/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------10/01/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures(w/ or w/o error messaging)-----------------------------10/01/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/01/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------10/01/20211
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/01/2021
'--Comment Code-----------------------------------------------------------------10/01/2021
'--Update Changelog for release/update------------------------------------------10/01/2021
'--Remove testing message boxes-------------------------------------------------10/01/2021
'--Remove testing code/unnecessary code-----------------------------------------10/01/2021
'--Review/update SharePoint instructions----------------------------------------10/01/2021
'--Review Best Practices using BZS page ----------------------------------------10/01/2021
'--Review script information on SharePoint BZ Script List-----------------------10/01/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/01/2021
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/01/2021
'--Complete misc. documentation (if applicable)---------------------------------10/01/2021
'--Update project team/issue contact (if applicable)----------------------------10/01/2021
