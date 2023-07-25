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
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
CALL changelog_update("01/30/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function determine_expedited_screening()
    'Now we will use the entries in the Application information to determine if this case is screened as expedited
    IF heat_AC_check = CHECKED THEN
        utilities = heat_AC_amt
    ELSEIF electric_check = CHECKED and phone_check = CHECKED THEN
        utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
    ELSEIF phone_check = CHECKED and electric_check = UNCHECKED THEN
        utilities = phone_amt
    ELSEIF electric_check = CHECKED and phone_check = UNCHECKED THEN
        utilities = electric_amt
    END IF

    'in case no options are clicked, utilities are set to zero.
    IF phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked THEN utilities = 0
    'If nothing is written for income/assets/rent info, we set to zero.
    IF income = "" THEN income = 0
    IF assets = "" THEN assets = 0
    IF rent   = "" THEN rent   = 0

    IF (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) THEN
        If population_of_case = "Families" Then transfer_to_worker = "EZ1"      'cases that screen as expedited are defaulted to expedited specific baskets based on population
        If population_of_case = "Adults" Then
            'making sure that Adults EXP baskets are not at limit


            If (EX1_basket_available = True and EX2_basket_available = False) then
                transfer_to_worker = "EX1"
            ElseIf (EX1_basket_available = False and EX2_basket_available = True) then
                transfer_to_worker = "EX2"
            Else
                'Do all the randomization here
                Randomize       'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
                random_number = Int(100*Rnd) 'rnd function returns a value greater or equal 0 and less than 1.
                If random_number MOD 2 = 1 then transfer_to_worker = "EX1"		'odd Number
                If random_number MOD 2 = 0 then transfer_to_worker = "EX2"		'even Number
            End if
        End If
        expedited_status = "Client Appears Expedited"                           'setting a variable with expedited information
        no_transfer_checkbox = unchecked
    End If
    IF (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) THEN
        expedited_status = "Client Does Not Appear Expedited"
        no_transfer_checkbox = checked
        transfer_to_worker = ""
    End If

    'families cases that have cash pending need to to to these specific baskets
    If population_of_case = "Families" and InStr(new_programs_pended, "CASH") <> 0 Then transfer_to_worker = "EY9"

    'The familiy cash basket has a backup if it has hit the display limit.
    If transfer_to_worker = "EY9" Then
        no_transfer_checkbox = unchecked
        If EY9_basket_available = False Then transfer_to_worker = "EY8"
    End If
end function

function update_programs_pending_detail(previously_pended_progs, new_programs_pended)
    prev_pend_snap_checkbox = unchecked         'default the checkboxed to unchecked
    prev_pend_cash_checkbox = unchecked
    prev_pend_grh_checkbox = unchecked
    prev_pend_emer_checkbox = unchecked
    new_pend_snap_checkbox = unchecked
    new_pend_cash_checkbox = unchecked
    new_pend_grh_checkbox = unchecked
    new_pend_emer_checkbox = unchecked

    If InStr(previously_pended_progs, "SNAP") <> 0 Then prev_pend_snap_checkbox = checked       'checking the boxes based on what is listed in the program string
    If InStr(previously_pended_progs, "CASH") <> 0 Then prev_pend_cash_checkbox = checked
    If InStr(previously_pended_progs, "GRH") <> 0 Then prev_pend_grh_checkbox = checked
    If InStr(previously_pended_progs, "EGA") <> 0 Then prev_pend_emer_checkbox = checked
    If InStr(previously_pended_progs, "EA") <> 0 Then prev_pend_emer_checkbox = checked

    If InStr(new_programs_pended, "SNAP") <> 0 Then new_pend_snap_checkbox = checked
    If InStr(new_programs_pended, "CASH") <> 0 Then new_pend_cash_checkbox = checked
    If InStr(new_programs_pended, "GRH") <> 0 Then new_pend_grh_checkbox = checked
    If InStr(new_programs_pended, "EGA") <> 0 Then new_pend_emer_checkbox = checked
    If InStr(new_programs_pended, "EA") <> 0 Then new_pend_emer_checkbox = checked

    'dialog of all of the checkboxes
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 186, 120, "Select Programs"
        Text 10, 10, 60, 10, "Already Pending: "
        CheckBox 10, 25, 35, 10, "SNAP", prev_pend_snap_checkbox
        CheckBox 50, 25, 35, 10, "Cash", prev_pend_cash_checkbox
        CheckBox 95, 25, 35, 10, "GRH", prev_pend_grh_checkbox
        CheckBox 140, 25, 35, 10, "EMER", prev_pend_emer_checkbox
        Text 10, 50, 165, 10, "Programs Requested on Subsequent Application:"
        CheckBox 10, 65, 35, 10, "SNAP", new_pend_snap_checkbox
        CheckBox 50, 65, 35, 10, "Cash", new_pend_cash_checkbox
        CheckBox 95, 65, 35, 10, "GRH", new_pend_grh_checkbox
        CheckBox 140, 65, 35, 10, "EMER", new_pend_emer_checkbox
        Text 10, 90, 60, 10, "Active Programs:"
        Text 75, 90, 100, 10, "None"
        ButtonGroup ButtonPressed
            PushButton 95, 105, 85, 10, "Return to Main Dialog", returnbtn
    EndDialog

    'displaying the dialog - no looping needed since this is embeded in another dialog and there is no way to validate the entries
    dialog Dialog1

    'resetting the program list strings
    previously_pended_progs = ""
    If prev_pend_snap_checkbox = checked Then previously_pended_progs = previously_pended_progs & ", SNAP"
    If prev_pend_cash_checkbox = checked Then previously_pended_progs = previously_pended_progs & ", CASH"
    If prev_pend_grh_checkbox = checked Then previously_pended_progs = previously_pended_progs & ", GRH"
    If prev_pend_emer_checkbox = checked Then
        If population_of_case = "Adults" Then previously_pended_progs = previously_pended_progs & ", EGA"
        If population_of_case = "Families" Then previously_pended_progs = previously_pended_progs & ", EA"
        If population_of_case = "Specialty" Then previously_pended_progs = previously_pended_progs & ", EGA"
    End If

    new_programs_pended = ""
    If new_pend_snap_checkbox = checked Then new_programs_pended = new_programs_pended & ", SNAP"
    If new_pend_cash_checkbox = checked Then new_programs_pended = new_programs_pended & ", CASH"
    If new_pend_grh_checkbox = checked Then new_programs_pended = new_programs_pended & ", GRH"
    If new_pend_emer_checkbox = checked Then
        If population_of_case = "Adults" Then new_programs_pended = new_programs_pended & ", EGA"
        If population_of_case = "Families" Then new_programs_pended = new_programs_pended & ", EA"
        If population_of_case = "Specialty" Then new_programs_pended = new_programs_pended & ", EGA"
    End If
    If new_programs_pended = "" Then new_programs_pended = "None"

    If left(previously_pended_progs, 1) = "," Then previously_pended_progs = right(previously_pended_progs, len(previously_pended_progs)-1)
    previously_pended_progs = trim(previously_pended_progs)
    If left(new_programs_pended, 1) = "," Then new_programs_pended = right(new_programs_pended, len(new_programs_pended)-1)
    new_programs_pended = trim(new_programs_pended)
end function

'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone

If skip_start_of_subsequent_apps <> True Then
    CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
    call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
    back_to_self                                        'added to ensure we have the time to update and send the case in the background
    EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

    'Initial Dialog - Case number
    Dialog1 = ""                                        'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 191, 135, "Subsequent Application Received"
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
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20SUBSEQUENT%20APPLICATION.docx"
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
    MEMO_NOTE_found = False                                     'defaulting these booleans
    screening_found = False

    Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
    too_old_date = DateAdd("D", -1, application_date)              'We don't need to read notes from before the CAF date

    note_row = 5
    previously_pended_progs = ""
    Do
        EMReadScreen note_date, 8, note_row, 6                  'reading the note date

        EMReadScreen note_title, 55, note_row, 25               'reading the note header
        note_title = trim(note_title)

        If left(note_title, 22) = "~ Application Received" Then 'Application received case note
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
        ' MsgBox "33 - " & left(note_title, 33)
        ' MsgBox "31 - " & left(note_title, 31)

        If left(note_title, 33) = "~ Appointment letter sent in MEMO" Then MEMO_NOTE_found = True       'MEMO case note
        If left(note_title, 31) = "~ Received Application for SNAP" Then screening_found = True         'Exp screening case note

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

'reading pnd2 to see if any of the baskets are over the limit and saving the findings for later functionality
EX1_basket_available = True
Call navigate_to_MAXIS_screen("REPT", "PND2")
Call write_value_and_transmit("EX1", 21, 17)
EMReadScreen pnd2_disp_limit, 13, 6, 35
If pnd2_disp_limit = "Display Limit" Then EX1_basket_available = False

EX2_basket_available = True
Call navigate_to_MAXIS_screen("REPT", "PND2")
Call write_value_and_transmit("EX2", 21, 17)
EMReadScreen pnd2_disp_limit, 13, 6, 35
If pnd2_disp_limit = "Display Limit" Then EX2_basket_available = False

EY9_basket_available = True
Call navigate_to_MAXIS_screen("REPT", "PND2")
EMWriteScreen "EY9", 21, 17
transmit
EMReadScreen pnd2_disp_limit, 13, 6, 35
If pnd2_disp_limit = "Display Limit" Then EY9_basket_available = False

back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

new_programs_pended = programs_applied_for
prev_prog_list = split(previously_pended_progs, ",")
For each the_prog in prev_prog_list
    the_prog = trim(the_prog)
    new_programs_pended = trim(replace(new_programs_pended, the_prog, ""))
Next
Do
    new_programs_pended = trim(replace(new_programs_pended, ", ,", ""))
Loop until InStr(new_programs_pended, ", ,") = 0
If left(new_programs_pended, 1) = "," Then new_programs_pended = right(new_programs_pended, len(new_programs_pended)-1)
new_programs_pended = trim(new_programs_pended)
If new_programs_pended = "" Then new_programs_pended = "None"

'Displaying the dialog
Do
	Do
        err_msg = ""

        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 275, 290, "Subsequent Application Detail"
  		  DropListBox 90, 145, 95, 45, "Select One:"+chr(9)+"ECF"+chr(9)+"Online"+chr(9)+"Request to APPL Form"+chr(9)+"In Person", how_application_rcvd
		  DropListBox 90, 165, 95, 45, "Select One:"+chr(9)+"Adults"+chr(9)+"Families"+chr(9)+"Specialty", population_of_case
          DropListBox 90, 185, 170, 45, "Select One:"+chr(9)+"CAF - 5223"+chr(9)+"MNbenefits CAF - 5223"+chr(9)+"SNAP App for Seniors - 5223F"+chr(9)+"MNsure App for HC - 6696"+chr(9)+"MHCP App for Certain Populations - 3876"+chr(9)+"App for MA for LTC - 3531"+chr(9)+"MHCP App for B/C Cancer - 3523"+chr(9)+"No Application Required", application_type
          EditBox 90, 205, 95, 15, confirmation_number
          EditBox 90, 225, 50, 15, form_date
  		  EditBox 55, 250, 210, 15, other_notes
          ButtonGroup ButtonPressed
            OkButton 160, 270, 50, 15
            CancelButton 215, 270, 50, 15
            PushButton 72, 65, 115, 13, "Change Pending Program Details", change_pending_progs_btn
            PushButton 72, 125, 115, 13, "Change Pending Program Details", change_pending_progs_btn
          GroupBox 5, 5, 190, 80, "Case Information"
          Text 10, 20, 60, 10, "Application Date: "
          Text 70, 20, 50, 10, application_date
          Text 10, 35, 60, 10, "Already Pending: "
          Text 70, 35, 120, 10, previously_pended_progs
          Text 10, 50, 60, 10, "Active Programs:"
          Text 70, 50, 120, 10, active_programs
          GroupBox 5, 85, 260, 160, "Application Information"
          Text 15, 100, 165, 10, "Programs Requested on Subsequent Application:"
          Text 25, 110, 160, 10, new_programs_pended
          Text 15, 150, 70, 10, "Application Received:"
          Text 15, 170, 70, 10, "Population/Specialty"
          Text 15, 190, 70, 10, "Form Type Received:"
          Text 35, 210, 50, 10, "Confirmation #:"
          Text 15, 230, 70, 10, "Date Form Received:"
		  Text 5, 255, 45, 10, "Other Notes:"
        EndDialog

        Dialog Dialog1
        cancel_confirmation

        If application_type = "MNbenefits CAF - 5223" AND how_application_rcvd = "Select One:" Then how_application_rcvd = "Online"
	    IF how_application_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
	    IF application_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
        IF application_type = "MNbenefits CAF - 5223" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If a MNbenefits app was received, you must enter the confirmation number and time received."
        If population_of_case = "Select One:" then err_msg = err_msg & vbNewLine & "* Please indicate the population or specialty of the case."
        If IsDate(form_date) = False Then err_msg = err_msg & vbNewLine & "* Please enter the date the subsequent application form was received."

        If ButtonPressed = change_pending_progs_btn Then
            If population_of_case = "Select One:" Then MsgBox "The EMER program option cannot be determined if checked as the EMER program is population specific." & vbCr & vbCr & "You can still check the programs, but if EMER is checked, EA/EGA will not show up on the list of programs since the population has not been selected."
            call update_programs_pending_detail(previously_pended_progs, new_programs_pended)
            err_msg = "LOOP"
        End If
        IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

Call hest_standards(heat_AC_amt, electric_amt, phone_amt, application_date) 'function to determine the hest standards depending on the application date.
no_transfer_checkbox = checked

'families cases that have cash pending need to to to these specific baskets
If population_of_case = "Families" and InStr(new_programs_pended, "CASH") <> 0 Then transfer_to_worker = "EY9"

'The familiy cash basket has a backup if it has hit the display limit.
If transfer_to_worker = "EY9" Then
    no_transfer_checkbox = unchecked
    If EY9_basket_available = False Then transfer_to_worker = "EY8"
End If

back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

send_appt_ltr = FALSE                                           'Now we need to determine if this case needs an appointment letter based on the program(s) pending
If unknown_cash_pending = True Then send_appt_ltr = TRUE
If ga_status = "PENDING" Then send_appt_ltr = TRUE
If msa_status = "PENDING" Then send_appt_ltr = TRUE
If mfip_status = "PENDING" Then send_appt_ltr = TRUE
If dwp_status = "PENDING" Then send_appt_ltr = TRUE
If grh_status = "PENDING" Then send_appt_ltr = TRUE
If snap_status = "PENDING" Then send_appt_ltr = TRUE
If emer_status = "PENDING" and emer_type = "EGA" Then send_appt_ltr = TRUE
' If emer_status = "PENDING" and emer_type = "EA" Then send_appt_ltr = TRUE

'if the case is determined to need an appointment letter the script will default the interview date
IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN
    interview_date = dateadd("d", 5, application_date)
    If interview_date <= date then interview_date = dateadd("d", 5, date)
    Call change_date_to_soonest_working_day(interview_date, "FORWARD")

    application_date = application_date & ""
    interview_date = interview_date & ""                                        'turns interview date into string for variable
End If

'displaying the dialog
Do
    Do
        err_msg = ""
        last_income = income
        last_assets = assets
        last_rent = rent
        last_heat_AC_check = heat_AC_check
        last_electric_check = electric_check
        last_phone_check = phone_check

        dlg_len = 75                'this is another dynamic dialog that needs different sizes based on what it has to display.
        If snap_status = "PENDING" and screening_found = False Then dlg_len = dlg_len + 110
        IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN dlg_len = dlg_len + 95
        IF how_application_rcvd = "Request to APPL Form" THEN dlg_len = dlg_len + 80
        If expedited_status = "" and snap_status = "PENDING" and screening_found = False Then dlg_len = 135

        'defining the actions dialog
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 266, dlg_len, "Actions in MAXIS"

            y_pos = 5
            If snap_status = "PENDING" and screening_found = False Then
                GroupBox 5, y_pos, 255, 105, "Expedited Screening"
                GroupBox 185, y_pos + 15, 70, 65, "**IMPORTANT**"
                EditBox 130, y_pos + 15, 50, 15, income
                Text 25, y_pos + 20, 95, 10, "Income received this month:"
                Text 190, y_pos + 30, 60, 45, "The income, assets and shelter costs fields will default to $0 if left blank. "
                EditBox 130, y_pos + 35, 50, 15, assets
                Text 30, y_pos + 40, 95, 10, "Cash, checking, or savings: "
                EditBox 130, y_pos + 55, 50, 15, rent
                Text 30, y_pos+ + 60, 90, 10, "AMT paid for rent/mortgage:"
                GroupBox 10, y_pos + 75, 170, 25, "Utilities claimed (check below):"
                CheckBox 15, y_pos + 85, 55, 10, "Heat (or AC)", heat_AC_check
                CheckBox 85, y_pos + 85, 45, 10, "Electricity", electric_check
                CheckBox 140, y_pos + 85, 35, 10, "Phone", phone_check
                y_pos = y_pos + 110
            End If
            If expedited_status <> "" OR snap_status <> "PENDING" OR screening_found = True Then
                GroupBox 5, y_pos, 255, 45, "Transfer Information"
                EditBox 95, y_pos + 10, 30, 15, transfer_to_worker
                If expedited_status = "Client Appears Expedited" Then Text 130, y_pos + 15, 130, 10, "This case screened as EXPEDITED."
                If expedited_status = "Client Does Not Appear Expedited" Then Text 130, y_pos + 15, 130, 10, "Case screened as NOT EXPEDITED."
                Text 10, y_pos + 15, 85, 10, "Transfer the case to x127"
                CheckBox 20, y_pos + 30, 185, 10, "Check here if this case does not require a transfer.", no_transfer_checkbox
                y_pos = y_pos + 55
                IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN
                    GroupBox 5, y_pos, 255, 90, "Appointment Notice"
                    y_pos = y_pos + 15
                    Text 15, y_pos, 35, 10, "CAF date:"
                    Text 50, y_pos, 55, 15, application_date
                    Text 120, y_pos, 60, 10, "Appointment date:"
                    Text 185, y_pos, 55, 15, interview_date
                    y_pos = y_pos + 15
                    Text 50, y_pos, 195, 10, "The NOTICE cannot be cancelled or changed from this script."
                    y_pos = y_pos + 10
                    Text 50, y_pos, 190, 20, "An Eligibility Worker can make changes/cancellations to the notice in MAXIS."
                    y_pos = y_pos + 20
                    Text 50, y_pos, 200, 10, "This script follows the requirements for the On Demand Waiver."
                    y_pos = y_pos + 10
                    odw_btn_y_pos = y_pos
                    y_pos = y_pos + 25
                End If
                IF how_application_rcvd = "Request to APPL Form" THEN
                    GroupBox 5, y_pos, 255, 75, "Request to APPL Information"
                    y_pos = y_pos + 10
                    reset_y = y_pos
                    EditBox 85, y_pos, 45, 15, request_date
                    Text 15, y_pos + 5, 60, 10, "Submission Date:"
                    y_pos = y_pos + 20
                    ' EditBox 85, y_pos, 45, 15, request_worker_number
                    ' Text 15, y_pos + 5, 60, 10, "Requested By X#:"
                    ' y_pos = y_pos + 20
                    EditBox 85, y_pos, 45, 15, METS_case_number
                    Text 15, y_pos + 5, 55, 10, "METS Case #:"
                    y_pos = reset_y
                    CheckBox 150, y_pos, 55, 10, "MA Transition", MA_transition_request_checkbox
                    y_pos = y_pos + 15
                    CheckBox 150, y_pos, 60, 10, "Auto Newborn", Auto_Newborn_checkbox
                    y_pos = y_pos + 15
                    CheckBox 150, y_pos, 85, 10, "METS Retro Coverage", METS_retro_checkbox
                    y_pos = y_pos + 15
                    CheckBox 150, y_pos, 85, 10, "Team 603 will process", team_603_email_checkbox
                    y_pos = y_pos + 25
                End If
            End If
            ButtonGroup ButtonPressed
                OkButton 155, y_pos, 50, 15
                CancelButton 210, y_pos, 50, 15
                If expedited_status <> "" OR snap_status <> "PENDING" OR screening_found = True Then
                    IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN PushButton 50, odw_btn_y_pos, 125, 13, "HSR Manual - On Demand Waiver", on_demand_waiver_button
                End If
        EndDialog


        Dialog Dialog1
        cancel_confirmation

        If expedited_status = "" and snap_status = "PENDING" and screening_found = False Then err_msg = "LOOP"
        If last_income <> income Then err_msg = "LOOP"
        If last_assets <> assets Then err_msg = "LOOP"
        If last_rent <> rent Then err_msg = "LOOP"
        If last_heat_AC_check <> heat_AC_check Then err_msg = "LOOP"
        If last_electric_check <> electric_check Then err_msg = "LOOP"
        If last_phone_check <> phone_check Then err_msg = "LOOP"

        If snap_status = "PENDING" and screening_found = False Then
            IF income = "" THEN income = "0"
            IF assets = "" THEN assets = "0"
            IF rent   = "" THEN rent   = "0"
            If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) THEN
                err_msg = err_msg & vbnewline & "* The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
            Else
                Call determine_expedited_screening
            End If
        End If

        If expedited_status <> "" OR snap_status <> "PENDING" OR screening_found = True Then
            IF no_transfer_checkbox = UNCHECKED AND transfer_to_worker = "" then err_msg = err_msg & vbNewLine & "* You must enter the basket number the case to be transferred by the script or check that no transfer is needed."
            IF no_transfer_checkbox = CHECKED and transfer_to_worker <> "" then err_msg = err_msg & vbNewLine & "* You have checked that no transfer is needed, please remove basket number from transfer field."
            IF no_transfer_checkbox = UNCHECKED AND len(transfer_to_worker) > 3 AND isnumeric(transfer_to_worker) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the last 3 digits of the worker number for transfer."
            IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN
                If IsDate(interview_date) = False Then err_msg = err_msg & vbNewLine & "* The Interview Date needs to be entered as a valid date."
            End If
            IF how_application_rcvd = "Request to APPL Form" THEN
                IF request_date = "" THEN err_msg = err_msg & vbNewLine & "* If a request to APPL was received, you must enter the date the form was submitted."
                IF METS_retro_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg & vbNewLine & "* You have checked that this is a METS Retro Request, please enter a METS IC #."
                IF MA_transition_request_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg &  vbNewLine & "* You have checked that this is a METS Transition Request, please enter a METS IC #."
            End If
        End If

        If ButtonPressed = on_demand_waiver_button Then
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/On_Demand_Waiver.aspx"
            err_msg = "LOOP"
        End If

        IF err_msg <> "" and left(err_msg, 4) <> "LOOP"THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine

        LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has     not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE

application_type = replace(application_type, " - ", " (DHS-") & ")"

call start_a_blank_CASE_NOTE

If new_programs_pended = "None" then Call write_variable_in_CASE_NOTE("Subsequent Application with No Additional Programs Requested")
If new_programs_pended <> "None" then call write_variable_in_CASE_NOTE("Subsequent Application Requesting: " & new_programs_pended)
call write_bullet_and_variable_in_CASE_NOTE("Subsequent Application Requesting", new_programs_pended)
call write_bullet_and_variable_in_CASE_NOTE("Previously Pending", previously_pended_progs)
call write_bullet_and_variable_in_CASE_NOTE("Application Date", application_date)
call write_bullet_and_variable_in_CASE_NOTE("Case Population", population_of_case)
call write_bullet_and_variable_in_CASE_NOTE("Active Programs", active_programs)
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)
call write_variable_in_CASE_NOTE("----------------- Subsequent Application Form Info -------------------------")
call write_bullet_and_variable_in_CASE_NOTE("Form Type Received", application_type)
call write_bullet_and_variable_in_CASE_NOTE("Confirmation #", confirmation_number)
call write_bullet_and_variable_in_CASE_NOTE("Date Form Received", form_date)
call write_variable_in_CASE_NOTE("Form date is not used as the program application date for subsequent applications as it is treated as client reporting. (CM 05.09.12)")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

PF3

If how_application_rcvd = "Request to APPL Form" THEN                           'specific functionality if the application was pended from a request to APPL form
    If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then        'HC cases - we need to add the persons pending HC to the CNOTE
        Call navigate_to_MAXIS_screen("STAT", "HCRE")                           'we are going to read this information from the HCRE panel.

        hcre_row = 10                   'top row
        household_persons = ""          'starting with a blank string
        Do                              'we are going to look at each row
            EMReadScreen hcre_app_date, 8, hcre_row, 51             'read the app_date
            EMReadScreen hcre_ref_nbr, 2, hcre_row, 24              'read the reference number
            'if the app date matches the app date we are processing, we will save the reference number to the list of all that match
            If hcre_app_date = app_date_with_banks Then household_persons = household_persons & hcre_ref_nbr & ", "

            hcre_row = hcre_row + 1         'go to the next row.
            If hcre_row = 18 Then           'go to the next page IF we are at the last row
                PF20
                hcre_row = 10
                EMReadScreen last_page_check, 9, 24, 14
                If last_page_check = "LAST PAGE" Then Exit Do   'leave the loop once we have reached the last page of persons on HCRE
            End If
        Loop
        household_persons = trim(household_persons)         'formatting the list of persons requesting HC
        If right(household_persons, 1) = "," THEN household_persons = left(household_persons, len(household_persons) - 1)
    End If
End If

'Functionality to send emails if the case was pended from a 'Request to APPL'
IF how_application_rcvd = "Request to APPL Form" Then
	send_email_to = ""
	cc_email_to = ""
	If team_603_email_checkbox = CHECKED Then send_email_to = "HSPH.EWS.TEAM.603@hennepin.us"

	email_subject = "Request to APPL Form has been processed for MAXIS Case # " & MAXIS_case_number
	email_body = "Request to APPL form has been received and processed."
	email_body = email_body & vbCr & vbCr & "MAXIS Case # " & MAXIS_case_number & " has been pended and is ready to be processed."
	If METS_case_number <> "" Then email_body = email_body & vbCr & "This case is associated with METS Case # " & METS_case_number & "."

	If METS_retro_checkbox = CHECKED and MA_transition_request_checkbox = CHECKED and Auto_Newborn_checkbox = CHECKED THEN
		email_body = email_body & vbCr & vbCr & "Request to APPL was received for:"
		If METS_retro_checkbox = CHECKED Then email_body = email_body & vbCr & "- METS Retro Request"
		If MA_transition_request_checkbox = CHECKED Then email_body = email_body & vbCr & "- MA Transition"
		If Auto_Newborn_checkbox = CHECKED Then email_body = email_body & vbCr & "- Auto Newborn"
	End If
	IF send_appt_ltr = TRUE THEN email_body = email_body & vbCr & vbCr & "A SPEC/MEMO has been created. If the client has completed the interview, please cancel the notice and update STAT/PROG with the interview information. Case Assignment is not tasked with cancelling or preventing this notice from being generated."
	email_body = email_body & vbCr & vbCr & "Case is ready to be processed."

    Call create_outlook_email("", send_email_to, cc_email_to, "", email_subject, 1, False, "", "", False, "", email_body, False, "", FALSE)
    'Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
End If

'Expedited Screening CNOTE for cases where SNAP is pending
If snap_status = "PENDING" and screening_found = False Then
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
    CALL write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
    CALL write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
    CALL write_variable_in_CASE_NOTE("        Utilities (AMT/HEST claimed): $" & utilities)
    CALL write_variable_in_CASE_NOTE("---")
    If has_DISQ = True then CALL write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
    If has_DISQ = False then CALL write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
    If in_office_card = False then CALL write_variable_in_CASE_NOTE("Recipient will NOT be able to get an EBT card in an agency office. An EBT card has previously been provided to the household.")
    CALL write_variable_in_CASE_NOTE("---")
    IF expedited_status = "Client Does Not Appear Expedited" THEN CALL write_variable_in_CASE_NOTE("Client does not appear expedited. Application sent to ECF.")
    IF expedited_status = "Client Appears Expedited" THEN CALL write_variable_in_CASE_NOTE("Client appears expedited. Application sent to ECF.")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)
    PF3
End If

'IF a transfer is needed (by entry of a transfer_to_worker in the Action dialog) the script will transfer it here
tansfer_message = ""            'some defaults
transfer_case = False
action_completed = TRUE

If transfer_to_worker <> "" Then        'If a transfer_to_worker was entered - we are attempting the transfer
	transfer_case = True
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER
	EMWriteScreen "x", 7, 16                               'transfer within county option
	transmit
	PF9                                                    'putting the transfer in edit mode
	EMreadscreen servicing_worker, 3, 18, 65               'checking to see if the transfer_to_worker is the same as the current_worker (because then it won't transfer)
	servicing_worker = trim(servicing_worker)
	IF servicing_worker = transfer_to_worker THEN          'If they match, cancel the transfer and save the information about the 'failure'
		action_completed = False
        transfer_message = "This case is already in the requested worker's number."
		PF10 'backout
		PF3 'SPEC menu
		PF3 'SELF Menu'
	ELSE                                                   'otherwise we are going for the tranfer
	    EMWriteScreen "X127" & transfer_to_worker, 18, 61  'entering the worker ifnormation
	    transmit                                           'saving - this should then take us to the transfer menu
        EMReadScreen panel_check, 4, 2, 55                 'reading to see if we made it to the right place
        If panel_check = "XWKR" Then
            action_completed = False                       'this is not the right place
            transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            PF10 'backout
            PF3 'SPEC menu
            PF3 'SELF Menu'
        Else                                               'if we are in the right place - read to see if the new worker is the transfer_to_worker
            EMReadScreen new_pw, 3, 21, 20
            If new_pw <> transfer_to_worker Then           'if it is not the transfer_tow_worker - the transfer failed.
                action_completed = False
                transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            End If
        End If
	END IF
END IF

'SENDING a SPEC/MEMO - this happens AFTER the transfer so that the correct team information is on the notice.
'there should not be an issue with PRIV cases because going directly here we shouldn't lose the 'connection/access'
IF send_appt_ltr = TRUE and MEMO_NOTE_found = False THEN        'If we are supposed to be sending an appointment letter - it will do it here - this matches the information in ON DEMAND functionality
	last_contact_day = DateAdd("d", 30, application_date)
	If DateDiff("d", interview_date, last_contact_day) < 0 Then last_contact_day = interview_date

	'Navigating to SPEC/MEMO and opening a new MEMO
	Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    		'Writes the appt letter into the MEMO.
    Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
    Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & interview_date & ". **")
    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    Call write_variable_in_SPEC_MEMO(" ")
    CALL write_variable_in_SPEC_MEMO("All interviews are completed via phone. If you do not have a phone, go to one of our Digital Access Spaces at any Hennepin County Library or Service Center. No processing, no interviews are completed at these sites. Some Options:")
    CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
    CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
    CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
    CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
    CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
    CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
    CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
    CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
    CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at service centers(listed above)")
    Call write_variable_in_SPEC_MEMO(" ")
    CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can always request a paper copy via phone.")

	PF4

    'now we are going to read if a MEMO was created.
    spec_row = 7
    memo_found = False
    Do
        EMReadScreen print_status, 7, spec_row, 67          'we are looking for a WAITING memo - if one is found -d we are going to assume it is the right one.
        If print_status = "Waiting" Then memo_found = True
        spec_row = spec_row + 1
    Loop until print_status = "       "

    If memo_found = True Then                               'CASE NOTING the MEMO sent if it was successful
        start_a_blank_CASE_NOTE
    	Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
        Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
        Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
        Call write_variable_in_CASE_NOTE("* A link to the Domestic Violence Brochure sent to client in SPEC/MEMO as part of notice.")
        Call write_variable_in_CASE_NOTE("---")
        CALL write_variable_in_CASE_NOTE (worker_signature)
    	PF3
    End If
END IF

revw_pending_table = False                                           'Determining if we should be adding this case to the CasesPending SQL Table
If unknown_cash_pending = True Then revw_pending_table = True       'case should be pending cash or snap and NOT have SNAP active
If ga_status = "PENDING" Then revw_pending_table = True
If msa_status = "PENDING" Then revw_pending_table = True
If mfip_status = "PENDING" Then revw_pending_table = True
If dwp_status = "PENDING" Then revw_pending_table = True
If grh_status = "PENDING" Then revw_pending_table = True
If snap_status = "PENDING" Then revw_pending_table = True
If snap_status = "ACTIVE" Then revw_pending_table = False
If trim(mx_region) = "TRAINING" Then revw_pending_table = False     'we do NOT want TRAINING cases in the SQL Table.

If revw_pending_table = True Then
	MAXIS_case_number = trim(MAXIS_case_number)
    eight_digit_case_number = right("00000000"&MAXIS_case_number, 8)            'The SQL table functionality needs the leading 0s added to the Case Number

    If unknown_cash_pending = True Then cash_stat_code = "P"                    'determining the program codes for the table entry

    If ma_status = "INACTIVE" Or ma_status = "APP CLOSE" Then hc_stat_code = "I"
    If ma_status = "ACTIVE" Or ma_status = "APP OPEN" Then hc_stat_code = "A"
    If ma_status = "REIN" Then hc_stat_code = "R"
    If ma_status = "PENDING" Then hc_stat_code = "P"
    If msp_status = "INACTIVE" Or msp_status = "APP CLOSE" Then hc_stat_code = "I"
    If msp_status = "ACTIVE" Or msp_status = "APP OPEN" Then hc_stat_code = "A"
    If msp_status = "REIN" Then hc_stat_code = "R"
    If msp_status = "PENDING" Then hc_stat_code = "P"
    If unknown_hc_pending = True Then hc_stat_code = "P"

    If ga_status = "PENDING" Then ga_stat_code = "P"
    If ga_status = "REIN" Then ga_stat_code = "R"
    If ga_status = "ACTIVE" Or ga_status = "APP OPEN" Then ga_stat_code = "A"
    If ga_status = "INACTIVE" Or ga_status = "APP CLOSE" Then ga_stat_code = "I"

    If grh_status = "PENDING" Then grh_stat_code = "P"
    If grh_status = "REIN" Then grh_stat_code = "R"
    If grh_status = "ACTIVE" Or grh_status = "APP OPEN" Then grh_stat_code = "A"
    If grh_status = "INACTIVE" Or grh_status = "APP CLOSE" Then grh_stat_code = "I"

    If emer_status = "PENDING" Then emer_stat_code = "P"
    If emer_status = "REIN" Then emer_stat_code = "R"
    If emer_status = "ACTIVE" Or emer_status = "APP OPEN" Then emer_stat_code = "A"
    If emer_status = "INACTIVE" Or emer_status = "APP CLOSE" Then emer_stat_code = "I"

    If mfip_status = "PENDING" Then mfip_stat_code = "P"
    If mfip_status = "REIN" Then mfip_stat_code = "R"
    If mfip_status = "ACTIVE" Or mfip_status = "APP OPEN" Then mfip_stat_code = "A"
    If mfip_status = "INACTIVE" Or mfip_status = "APP CLOSE" Then mfip_stat_code = "I"

    If snap_status = "PENDING" Then snap_stat_code = "P"
    If snap_status = "REIN" Then snap_stat_code = "R"
    If snap_status = "ACTIVE" Or snap_status = "APP OPEN" Then snap_stat_code = "A"
    If snap_status = "INACTIVE" Or snap_status = "APP CLOSE" Then snap_stat_code = "I"

    If no_transfer_checkbox = checked Then worker_id_for_data_table = initial_pw_for_data_table     'determining the X-Number for table entry
    If no_transfer_checkbox = unchecked Then worker_id_for_data_table = transfer_to_worker
    If len(worker_id_for_data_table) = 3 Then worker_id_for_data_table = "X127" & worker_id_for_data_table

    'Setting constants
    Const adOpenStatic = 3
    Const adLockOptimistic = 3

	case_number_found_in_SQL = False
	'Read the whole table to see if this case number exists on the list
	objSQL = "SELECT * FROM ES.ES_CasesPending"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof
		sql_case_number = case_info_notes = objRecordSet("CaseNumber")
		If eight_digit_case_number = sql_case_number Then
			case_number_found_in_SQL = True
			Exit Do
		End If
		objRecordSet.MoveNext
	Loop
	'close the connection and recordset objects to free up resources
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	If case_number_found_in_SQL = True Then
		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the BZST connection to SQL Database'
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

		objRecordSet.Open "SELECT * FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection
		current_exp_code = objRecordSet("IsExpSnap")
		If snap_status = "PENDING" and screening_found = False Then current_exp_code = 1

		'close the connection and recordset objects to free up resources
		objConnection.Close

		'This is the BZST connection to SQL Database'
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

		'delete a record if the case number matches
		objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection
	Else
		current_exp_code = 1

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the BZST connection to SQL Database'
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	End If

    'delete a record if the case number matches
    ' objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_nuMEMO_foundmber & "'", objConnection
    'Add a new record with this case information'
    objRecordSet.Open "INSERT INTO ES.ES_CasesPending (WorkerID, CaseNumber, CaseName, ApplDate, FSStatusCode, CashStatusCode, HCStatusCode, GAStatusCode, GRStatusCode, EAStatusCode, MFStatusCode, IsExpSnap, UpdateDate)" &  _
                      "VALUES ('" & worker_id_for_data_table & "', '" & eight_digit_case_number & "', '" & case_name_for_data_table & "', '" & application_date & "', '" & snap_stat_code & "', '" & cash_stat_code & "', '" & hc_stat_code & "', '" & ga_stat_code & "', '" & grh_stat_code & "', '" & emer_stat_code & "', '" & mfip_stat_code & "', '" & current_exp_code & "', '" & date & "')", objConnection, adOpenStatic, adLockOptimistic

    'close the connection and recordset objects to free up resources
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
End If


'Now we create some messaging to explain what happened in the script run.
end_msg = "Subsequent Application Received has been noted."
end_msg = end_msg & vbCr & "Additional Programs requested: " & new_programs_pended & ", form received: " & form_date

If snap_status = "PENDING" and screening_found = False Then end_msg = end_msg & vbCr & vbCr & "Since SNAP is pending, an Expedtied SNAP screening has been completed and noted based on resident reported information from CAF1."

If MEMO_NOTE_found = True Then
    end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice was previously sent to the resident to alert them to the need for an interview for their requested programs and the script did not send one during this run."
Else
    IF send_appt_ltr = TRUE AND memo_found = True THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice has been sent to the resident to alert them to the need for an interview for their requested programs."
    IF send_appt_ltr = TRUE AND memo_found = False THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice about the Interview appears to have failed. Contact QI Knowledge Now to have one sent manually."
End If
If transfer_message = "" Then
    If transfer_case = True Then end_msg = end_msg & vbCr & vbCr & "Case transfer has been completed to x127" & transfer_to_worker
Else
    end_msg = end_msg & vbCr & vbCr & "FAILED CASE TRANSFER:" & vbCr & transfer_message
End If
If transfer_case = False Then end_msg = end_msg & vbCr & vbCr & "NO TRANSFER HAS BEEN REQUESTED."
IF how_application_rcvd = "Request to APPL Form" Then end_msg = end_msg & vbCr & vbCr & "CASE PENDED from a REQUEST TO APPL FORM"
script_run_lowdown = script_run_lowdown & vbCr & "END Message: " & vbCr & end_msg

Call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------01/26/2023
'--Tab orders reviewed & confirmed----------------------------------------------01/26/2023
'--Mandatory fields all present & Reviewed--------------------------------------01/26/2023
'--All variables in dialog match mandatory fields-------------------------------01/26/2023
'Review dialog names for content and content fit in dialog----------------------01/26/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------01/26/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------01/26/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function:
'--confirm that proper punctuation is used -------------------------------------01/26/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------01/26/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------01/26/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------01/26/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------01/26/2023 - keeping matched to App Received
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------01/26/2023
'--Script name reviewed---------------------------------------------------------01/26/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------01/26/2023
'--comment Code-----------------------------------------------------------------01/26/2023
'--Update Changelog for release/update------------------------------------------01/26/2023
'--Remove testing message boxes-------------------------------------------------01/26/2023
'--Remove testing code/unnecessary code-----------------------------------------01/26/2023
'--Review/update SharePoint instructions----------------------------------------In Progress
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------In Progress
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------In Progress
'--Complete misc. documentation (if applicable)---------------------------------In Progress
'--Update project team/issue contact (if applicable)----------------------------In Progress