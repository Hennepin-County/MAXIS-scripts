'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SUBSEQUENT APPLICATION.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("02/01/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone

If skip_start_of_subsequent_apps <> True Then
    CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
    call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
    back_to_self                                        'added to ensure we have the time to update and send the case in the background
    EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

    'Initial Dialog - Case number
    Dialog1 = ""                                        'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 191, 135, "Application Received"
      EditBox 60, 35, 45, 15, MAXIS_case_number
      ButtonGroup ButtonPressed
        PushButton 90, 95, 95, 15, "Script Instructions", script_instructions_btn
        OkButton 80, 115, 50, 15
        CancelButton 135, 115, 50, 15
      Text 5, 10, 185, 20, "Multiple CASE:NOTEs will be entered with this script run to document the actions for pending new applications."
      Text 5, 40, 50, 10, "Case Number:"
      Text 5, 55, 185, 10, "This case should be in PND2 status for this script to run."
      Text 5, 65, 185, 30, "If the programs requested on the application are not yet pending in MAXIS, cancel this script run, pend the case to PND2 status and run the script again."
    EndDialog

    'Runs the first dialog - which confirms the case number
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
          	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
            If ButtonPressed = script_instructions_btn Then             'Pulling up the instructions if the instruction button was pressed.
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx"
                err_msg = "LOOP"
            Else                                                        'If the instructions button was NOT pressed, we want to display the error message if it exists.
    		    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
            End If
    	Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'Checking for PRIV cases.
    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
    IF is_this_priv = True THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
    MAXIS_background_check      'Making sure we are out of background.
    EMReadScreen initial_pw_for_data_table, 7, 21, 17
    EMReadScreen case_name_for_data_table, 20, 21, 46

    'Grabbing case and program status information from MAXIS.
    'For tis script to work correctly, these must be correct BEFORE running the script.
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, active_programs, programs_applied_for)
    EMReadScreen pnd2_appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2

    case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
    script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "CASE IS PENDING - " & case_pending        'Adding details about CASE/CURR information to a script report out to BZST
    If case_status = "CAF1 PENDING" OR case_pending = False Then                    'The case MUST be pending and NOT in PND1 to continue.
        call script_end_procedure_with_error_report("This case is not in PND2 status. Current case status in MAXIS is " & case_status & ". Update MAXIS to put this case in PND2 status and then run the script again.")
    End If

    call back_to_SELF           'resetting
    EMReadScreen mx_region, 10, 22, 48

    If mx_region = "INQUIRY DB" Then
        ' continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
        ' If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
    End If

    multiple_app_dates = False                          'defaulting the boolean about multiple application dates to FALSE
    EMWriteScreen MAXIS_case_number, 18, 43             'now we are going to try to get to REPT/PND2 for the case to read the application date.
    Call navigate_to_MAXIS_screen("REPT", "PND2")
    EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
    If pnd2_disp_limit = "Display Limit" Then transmit
    row = 1                                             'searching for the CASE NUMBER to read from the right row
    col = 1
    EMSearch MAXIS_case_number, row, col
    If row <> 24 and row <> 0 Then pnd2_row = row
    EMReadScreen application_date, 8, pnd2_row, 38                                  'reading and formatting the application date
    application_date = replace(application_date, " ", "/")
    EMReadScreen additional_application_check, 14, pnd2_row + 1, 17                 'looking to see if this case has a secondary application date entered
    If additional_application_check = "ADDITIONAL APP" THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.
        multiple_app_dates = True           'identifying that this case has multiple application dates - this is not used specifically yet but is in place so we can output information for managment of case handling in the future.

        EMReadScreen additional_application_date, 8, pnd2_row + 1, 38               'reading the app date from the other application line
        additional_application_date = replace(additional_application_date, " ", "/")

        'There is a specific dialog that will display if there is more than one application date so we can select the right one for this script run
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 166, 160, "Application Received"
          DropListBox 15, 70, 100, 45, application_date+chr(9)+additional_application_date, app_date_to_use
          ButtonGroup ButtonPressed
            PushButton 65, 120, 95, 15, "Open CM 05.09.06", cm_05_09_06_btn
            OkButton 55, 140, 50, 15
            CancelButton 110, 140, 50, 15
          Text 5, 10, 135, 10, "This case has a second application date."
          Text 5, 25, 165, 25, "Per CM 0005.09.06 - if a case is pending and a new app is received you should use the original application date."
          Text 5, 55, 115, 10, "Select which date you need to use:"
          Text 5, 90, 145, 30, "Please contact Knowledge Now or your Supervisor if you have questions about dates to enter in MAXIS for applications."
        EndDialog

        Do
        	Do
        		Dialog Dialog1
        		cancel_without_confirmation

                'referncing the CM policy about application dates.
                If ButtonPressed = cm_05_09_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00050906"
        	Loop until ButtonPressed = -1
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
        application_date = app_date_to_use                          'setting the application date selected to the application_date variable
    End If

    IF IsDate(application_date) = False THEN                   'If we could NOT find the application date - then it will use the PND2 application date.
        application_date = pnd2_appl_date
    End if


End If

If previously_pended_progs = "" Then
    MEMO_found = False
    screening_found = False

    Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
    too_old_date = DateAdd("D", -1, application_date)              'We don't need to read notes from before the CAF date

    note_row = 5
    previously_pended_progs = ""
    Do
        EMReadScreen note_date, 8, note_row, 6                  'reading the note date

        EMReadScreen note_title, 55, note_row, 25               'reading the note header
        note_title = trim(note_title)

        If left(note_title, 22) = "~ Application Received" Then
            app_recvd_note_found = True
            Call write_value_and_transmit("X", note_row, 3)
            in_note_row = 4
            Do
                EMReadScreen note_line, 78, in_note_row, 3
                note_line = trim(note_line)

                If left(note_line, 25) = "* Application Requesting:" Then
                    previously_pended_progs = right(note_line, len(note_line)-25)
                    previously_pended_progs = trim(previously_pended_progs)
                End If

                If left(note_line, 18) = "* Case Population:" Then
                    population_of_case = right(note_line, len(note_line)-18)
                    population_of_case = trim(population_of_case)
                End If

                in_note_row = in_note_row + 1
                If in_note_row = 18 Then
                    PF8
                    in_note_row = 4
                    EMReadScreen end_of_note, 9, 24, 14
                    If end_of_note = "LAST PAGE" Then Exit Do
                End If
            Loop until note_line = ""
            PF3
        end If

        If left(note_title, 23) = "~ Appointment letter sent in MEMO" Then MEMO_found = True
        If left(note_title, 21) = "~ Received Application for SNAP" Then screening_found = True

        if note_date = "        " then Exit Do
        note_row = note_row + 1
        if note_row = 19 then
            note_row = 5
            PF8
            EMReadScreen check_for_last_page, 9, 24, 14
            If check_for_last_page = "LAST PAGE" Then Exit Do
        End If
        EMReadScreen next_note_date, 8, note_row, 6
        if next_note_date = "        " then Exit Do
    Loop until DateDiff("d", too_old_date, next_note_date) <= 0

End If

new_programs_pended = programs_applied_for
prev_prog_list = split(previously_pended_progs, ",")
For each the_prog in prev_prog_list
    the_prog = trim(the_prog)
    new_programs_pended = trim(replace(new_programs_pended, the_prog, ""))
Next
Do
    new_programs_pended = trim(replace(new_programs_pended, ", ,", ""))
Loop until InStr(new_programs_pended, ", ,") = 0
If new_programs_pended = "" Then new_programs_pended = "None"

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 206, 310, "Subsequent Application Detail"
  DropListBox 90, 140, 95, 45, "Select One:"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Populations"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer"+chr(9)+"MNbenefits"+chr(9)+"N/A"+chr(9)+"Verbal Request", application_type
  EditBox 90, 160, 50, 15, form_date
  EditBox 90, 180, 95, 15, confirmation_number
  DropListBox 90, 200, 95, 45, "Select One:"+chr(9)+"Adults"+chr(9)+"Families"+chr(9)+"Specialty", population_of_case
  DropListBox 90, 220, 95, 45, "No - Only ES Programs"+chr(9)+"Yes - Child Care Requested", was_ccap_requested
  DropListBox 90, 255, 95, 45, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Mystery Doc Queue"+chr(9)+"Online"+chr(9)+"Phone-Verbal Request"+chr(9)+"Request to APPL Form"+chr(9)+"Virtual Drop Box", how_application_rcvd
  ButtonGroup ButtonPressed
    OkButton 90, 285, 50, 15
    CancelButton 145, 285, 50, 15
    PushButton 72, 65, 115, 13, "Change Pending Program Details", change_pending_progs_btn
    PushButton 72, 125, 115, 13, "Change Pending Program Details", change_pending_progs_btn
  GroupBox 5, 5, 190, 75, "Case Information"
  Text 10, 20, 60, 10, "Application Date: "
  Text 70, 20, 50, 10, application_date
  Text 10, 35, 60, 10, "Already Pending: "
  Text 70, 35, 120, 10, previously_pended_progs
  Text 10, 50, 60, 10, "Active Programs:"
  Text 70, 50, 120, 10, active_programs
  GroupBox 5, 85, 190, 155, "Application Information"
  Text 15, 100, 165, 10, "Programs Requested on Subsequent Application:"
  Text 25, 110, 160, 10, new_programs_pended
  Text 15, 145, 70, 10, "Form Type Received:"
  Text 15, 165, 70, 10, "Date Form Received:"
  Text 35, 185, 50, 10, "Confirmation #:"
  Text 15, 205, 70, 10, "Population/Specialty"
  Text 10, 225, 80, 10, "Was CCAP Requested?"
  GroupBox 5, 245, 190, 30, "Agency Information"
  Text 15, 260, 70, 10, "Application Received:"
EndDialog

'Displaying the dialog
Do
	Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation

        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

MsgBox "STOP NOW"

call start_a_blank_CASE_NOTE

call write_variable_in_CASE_NOTE("Subsequent Application Requesting: " & new_programs_pended)
call write_bullet_and_variable_in_CASE_NOTE("Subsequent Application Requesting", new_programs_pended)
call write_bullet_and_variable_in_CASE_NOTE("Previously Pending", previously_pended_progs)
call write_bullet_and_variable_in_CASE_NOTE("Application Date", application_date)
call write_bullet_and_variable_in_CASE_NOTE("Case Population", population_of_case)
call write_bullet_and_variable_in_CASE_NOTE("Active Programs", active_programs)
call write_variable_in_CASE_NOTE("----------------- Subsequent Application Form Info -------------------------")
call write_bullet_and_variable_in_CASE_NOTE("Form Type Received", application_type)
call write_bullet_and_variable_in_CASE_NOTE("Confirmation #", confirmation_number)
call write_bullet_and_variable_in_CASE_NOTE("Date Form Received", form_date)
call write_variable_in_CASE_NOTE("Form date is not used as the program application date for subsequent applications as it is treated as client reporting. (CM 05.09.12)")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

Call script_end_procedure_with_error_report()
