'GATHERING STATS===========================================================================================
name_of_script = "NOTES - APPLICATION CHECK.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
CALL changelog_update("09/20/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("03/29/2022", "Removed ApplyMN as application option.", "Ilse Ferris")
call changelog_update("01/15/2021", "Added support for the pending Health Care application screening requirements.", "Ilse Ferris")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("12/17/2019", "Enhanced PND2 case pending case search.", "Ilse Ferris, Hennepin County")
call changelog_update("10/03/2019", "Updated TIKL functionality to suppress when the case is identified to either approve or deny.", "Ilse Ferris")
call changelog_update("09/06/2019", "Updated TIKL veribage", "Ilse Ferris")
call changelog_update("08/27/2019", "Updated dialog and case note to address requested enhancements (TIKL & interview still needed).", "MiKayla Handley")
call changelog_update("08/20/2019", "Bug on the script when a large PND2 list is accessed.", "Casey Love, Hennepin County")
call changelog_update("06/14/2018", "Updated dialog and case note to address requested enhancements.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "Application check completed, a case note made, and a TIKL has been set." 'setting up closing_message variable for possible additions later based on conditions

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 176, 65, "Application Check"
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 70, 25, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 45, 50, 15
    CancelButton 125, 45, 45, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 60, 10, "Worker signature:"
EndDialog

Do
	DO
		err_msg = ""
	    dialog dialog1
      	cancel_without_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("REPT", "PND2") 'information gathering to auto-populate the application date
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROD
IF priv_check = "PRIV" then script_end_procedure("This case is privileged, and cannot be accessed.")

'Ensuring that the user is in REPT/PND2
Do
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check <> "PND2" then
		back_to_SELF
		Call navigate_to_MAXIS_screen("REPT", "PND2")
	End if
LOOP until PND2_check = "PND2"

'checking the case to make sure there is a pending case.  If not script will end & inform the user no pending case exists in PND2
EMReadScreen not_pending_check, 5, 24, 2
If not_pending_check = "CASE " THEN script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

'Because inquiry and training region are stupid, when you nav to REPT/PND2 the cursor resets to row 20, col 13 or the REPT field. Just to ruin my life.
'Now code will work in all regions - production, inquiry and training.
row = 7
Do
    EmReadscreen pending_case_num, 8, row, 6
    If trim(pending_case_num) = trim(MAXIS_case_number) then
        found_case = True
        exit do
    Else
        row = row + 1
        found_case = False
    End if

LOOP until row = 19
If found_case = False then script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

HC_pending = False  'setting variable to false. This will be used to determine if HC is penfing or not to support HC screening process.

EMReadScreen app_month, 2, row, 38
EMReadScreen app_day, 2, row, 41
EMReadScreen app_year, 2, row, 44
EMReadScreen days_pending, 3, row, 50
EMReadScreen additional_application_check, 14, row + 1, 17
EMReadScreen add_app_month, 2, row + 1, 38
EMReadScreen add_app_day, 2, row + 1, 41
EMReadScreen add_app_year, 2, row + 1, 44

'Creating new variable for application check date and additional application date.
application_date = app_month & "/" & app_day & "/" & app_year
additional_application_date = add_app_month & "/" & add_app_day & "/" & add_app_year

'checking for multiple application dates.  Creates message boxes giving the user an option of which app date to choose
If additional_application_check = "ADDITIONAL APP" THEN multiple_apps = MsgBox("Do you want this application date: " & application_date, VbYesNoCancel)
If multiple_apps = vbCancel then stopscript
If multiple_apps = vbYes then application_date = application_date
IF multiple_apps = vbNo then
	additional_apps = Msgbox("Per CM 0005.09.06 - if a case is pending and a new app is received you should use the original application date." & vbcr & "Do you want this application date: " & additional_application_date, VbYesNoCancel)
	application_date = ""
	If additional_apps = vbCancel then stopscript
	If additional_apps = vbNo then script_end_procedure_with_error_report("No more application dates exist. Please review the case, and start the script again if applicable.")
	If additional_apps = vbYes then
		additional_date_found = TRUE
		application_date = additional_application_date
		row = row + 1
	END IF
End if

EMReadScreen PEND_CASH_check,	1, row, 54
EMReadScreen PEND_SNAP_check, 1, row, 62
EMReadScreen PEND_HC_check, 1, row, 65
EMReadScreen PEND_EMER_check,	1, row, 68
EMReadScreen PEND_GRH_check, 1, row, 72

If PEND_HC_check = "P" then HC_pending = True   'This will search case notes to ensure that a HC Application screening has been conducted.

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
EMReadScreen err_msg, 7, 24, 02
IF err_msg = "BENEFIT" THEN	script_end_procedure_with_error_report ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

'Reading the app date from PROG
EMReadScreen cash1_app_date, 8, 6, 33
cash1_app_date = replace(cash1_app_date, " ", "/")
EMReadScreen cash2_app_date, 8, 7, 33
cash2_app_date = replace(cash2_app_date, " ", "/")
EMReadScreen emer_app_date, 8, 8, 33
emer_app_date = replace(emer_app_date, " ", "/")
EMReadScreen grh_app_date, 8, 9, 33
grh_app_date = replace(grh_app_date, " ", "/")
EMReadScreen snap_app_date, 8, 10, 33
snap_app_date = replace(snap_app_date, " ", "/")
EMReadScreen ive_app_date, 8, 11, 33
ive_app_date = replace(ive_app_date, " ", "/")
EMReadScreen hc_app_date, 8, 12, 33
hc_app_date = replace(hc_app_date, " ", "/")
EMReadScreen cca_app_date, 8, 14, 33
cca_app_date = replace(cca_app_date, " ", "/")

'Reading the program status
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen snap_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74

'----------------------------------------------------------------------------------------------------ACTIVE program coding
EMReadScreen cash1_prog_check, 2, 6, 67     'Reading cash 1
EMReadScreen cash2_prog_check, 2, 7, 67     'Reading cash 2
EMReadScreen emer_prog_check, 2, 8, 67      'EMER Program

'Logic to determine if MFIP is active
IF cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "MS" THEN
	IF cash1_status_check = "ACTV" THEN cash_active = TRUE
END IF
IF cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "MS" THEN
	IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
END IF
IF emer_prog_check = "EG" and emer_status_check = "ACTV" THEN emer_active = TRUE
IF emer_prog_check = "EA" and emer_status_check = "ACTV" THEN emer_active = TRUE

IF cash1_status_check = "ACTV" THEN cash_active  = TRUE
IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
IF snap_status_check  = "ACTV" THEN SNAP_active  = TRUE
IF grh_status_check   = "ACTV" THEN grh_active   = TRUE
IF ive_status_check   = "ACTV" THEN IVE_active   = TRUE
IF hc_status_check    = "ACTV" THEN hc_active    = TRUE
IF cca_status_check   = "ACTV" THEN cca_active   = TRUE

active_programs = ""        'Creates a variable that lists all the active.
IF cash_active = TRUE or cash2_active = TRUE THEN active_programs = active_programs & "CASH, "
IF emer_active = TRUE THEN active_programs = active_programs & "Emergency, "
IF grh_active  = TRUE THEN active_programs = active_programs & "GRH, "
IF snap_active = TRUE THEN active_programs = active_programs & "SNAP, "
IF ive_active  = TRUE THEN active_programs = active_programs & "IV-E, "
IF hc_active   = TRUE THEN active_programs = active_programs & "HC, "
IF cca_active  = TRUE THEN active_programs = active_programs & "CCA"

active_programs = trim(active_programs)  'trims excess spaces of active_programs
If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

'----------------------------------------------------------------------------------------------------Pending programs
programs_applied_for = ""   'Creates a variable that lists all pening cases.
additional_programs_applied_for = ""
'cash I
IF cash1_status_check = "PEND" then
    If cash1_app_date = application_date THEN
        cash_pends = TRUE
		CASH_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "CASH, "
    Else
        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
	End if
End if
'cash II
IF cash2_status_check = "PEND" then
    if cash2_app_date = application_date THEN
        cash2_pends = TRUE
		CASH_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "CASH, "
    Else
        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
    End if
End if
'SNAP
IF snap_status_check  = "PEND" then
    If snap_app_date  = application_date THEN
        SNAP_pends = TRUE
		FS_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "SNAP, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "SNAP, "
    end if
End if
'GRH
IF grh_status_check = "PEND" then
    If grh_app_date = application_date THEN
        grh_pends = TRUE
		GRH_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "GRH, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "GRH, "
    End if
End if
'I-VE
IF ive_status_check = "PEND" then
    if ive_app_date = application_date THEN
        IVE_pends = TRUE
        programs_applied_for = programs_applied_for & "IV-E, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "IV-E, "
    End if
End if
'HC
IF hc_status_check = "PEND" then
    If hc_app_date = application_date THEN
        hc_pends = TRUE
		HC_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "HC, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "HC, "
    End if
End if
'CCA
IF cca_status_check = "PEND" then
    If cca_app_date = application_date THEN
        cca_pends = TRUE
        programs_applied_for = programs_applied_for & "CCA, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "CCA, "
    End if
End if
'EMER
If emer_status_check = "PEND" then
    If emer_app_date = application_date then
        emer_pends = TRUE
		EA_CHECKBOX = CHECKED
        IF emer_prog_check = "EG" THEN programs_applied_for = programs_applied_for & "EGA, "
        IF emer_prog_check = "EA" THEN programs_applied_for = programs_applied_for & "EA, "
    else
		EA_CHECKBOX = CHECKED
        IF emer_prog_check = "EG" THEN additional_programs_applied_for = additional_programs_applied_for & "EGA, "
        IF emer_prog_check = "EA" THEN additional_programs_applied_for = additional_programs_applied_for & "EA, "
    End if
End if

programs_applied_for = trim(programs_applied_for)       'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

additional_programs_applied_for = trim(additional_programs_applied_for)       'trims excess spaces of programs_applied_for
If right(additional_programs_applied_for, 1) = "," THEN additional_programs_applied_for = left(additional_programs_applied_for, len(additional_programs_applied_for) - 1)

pending_progs = trim(pending_progs) 'trims excess spaces of pending_progs
'takes the last comma off of pending_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(pending_progs, 1) = "," THEN pending_progs = left(pending_progs, len(pending_progs) - 1)

'Determines which application check the user is at----------------------------------------------------------------------------------------------------
If DateDiff("d", application_date, date) = 0 then
	application_check = "Day 1"
	reminder_date = dateadd("d", 5, application_date)
	reminder_text = "Day 5"
Elseif DateDiff("d", application_date, date) = 1 then
	application_check = "Day 1"
	reminder_date = dateadd("d", 5, application_date)
	reminder_text = "Day 5"
Elseif (DateDiff("d", application_date, date) > 1 AND DateDiff("d", application_date, date) < 10) then
	application_check = "Day 5"
	reminder_date = dateadd("d", 10, application_date)
	reminder_text = "Day 10"
Elseif (DateDiff("d", application_date, date) => 10 AND DateDiff("d", application_date, date) < 20) then
	application_check = "Day 10"
	reminder_date = dateadd("d", 20, application_date)
	reminder_text = "Day 20"
Elseif (DateDiff("d", application_date, date) => 20 AND DateDiff("d", application_date, date) < 30) then
	application_check = "Day 20"
	reminder_date = dateadd("d", 30, application_date)
	reminder_text = "Day 30"
Elseif (DateDiff("d", application_date, date) => 30 AND DateDiff("d", application_date, date) < 45) then
	application_check = "Day 30"
	reminder_date = dateadd("d", 45, application_date)
	reminder_text = "Day 45"
Elseif (DateDiff("d", application_date, date) => 45 AND DateDiff("d", application_date, date) < 60) then
	application_check = "Day 45"
	reminder_date = dateadd("d", 60, application_date)
	reminder_text = "Day 60"
Elseif DateDiff("d", application_date, date) = 60 then
	application_check = "Day 60"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
Elseif DateDiff("d", application_date, date) > 60 then
	application_check = "Over 60 days"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
END IF

'--------------------------------------------------------------------------------------------------------------------------------------Health Care Screening Portion
IF HC_pending = True then
    hc_days_pending = datediff("D", hc_app_date, date)
    'Checking case note to see if a HC screening has been completed to date
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    'starting at the 1st case note, checking the headers for the HC screening
    case_note_found = False         'defaulting to false if not able to find an expedited care note
    row = 5
    Do
        EMReadScreen first_case_note_date, 8, 5, 6 'static reading of the case note date to determine if no case notes acutually exist.
        If trim(first_case_note_date) = "" then
            case_note_found = False
            exit do
        Else
            EMReadScreen case_note_date, 8, row, 6    'incremented row - reading the case note date
            EMReadScreen case_note_header, 55, row, 25
            case_note_header = lcase(trim(case_note_header))
            If trim(case_note_date) = "" then
                case_note_found = False             'The end of the case notes has been found
                exit do
            ElseIf instr(case_note_header, "Health Care Application Screening Completed") then
                case_note_found = True     'no need for screening.
                exit do
            Else
                row = row + 1
                IF row = 19 then
                    PF8                         'moving to next case note page if at the end of the page
                    row = 5
                End if
            END IF
        END IF
    LOOP until cdate(case_note_date) < cdate(hc_app_date) 'repeats until the case note date is less than the HC application date

    interview_status = ""   'blanking out variable
    If case_note_found = False then
        'Asking the user if they wish to contact the client/arep. If yes, they will go to the screening dialog
        'If no - the user needs to provide a reason for not screening which then will be added to the Application check case note.
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 226, 65, "Health Care Application Screening Not Found"
            DropListBox 160, 5, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", interview_confirmation
            EditBox 75, 25, 145, 15, no_call_reason
            ButtonGroup ButtonPressed
            OkButton 130, 45, 45, 15
            CancelButton 175, 45, 45, 15
            Text 5, 30, 70, 10, "If no, provide reason:"
            Text 5, 10, 150, 10, "Would you like to call the resident/AREP now?"
            Text 5, 45, 125, 10, "(Reason will be captured in case note)"
        EndDialog

        Do
        	DO
        		err_msg = ""
        		Dialog Dialog1
                cancel_without_confirmation
        		If interview_confirmation = "Select one..." then err_msg = error_msg & ("Confirm if you will call the resident/AREP.")
                If interview_confirmation = "No" and trim(no_call_reason) = "" then err_msg = error_msg & ("Provide a reason for not calling the resident/AREP.")
                If interview_confirmation = "Yes" and trim(no_call_reason) <> "" then err_msg = error_msg & ("Either select Yes and clear the reason field, or select No and provide a reason for not conducting a screening.")
                If err_msg <> "" then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        If interview_confirmation = "Yes" then
            'Gathing information for the next dialog to be auto-filled
            when_contact_was_made = date & ", " & time 'updates the "when contact was made" variable to show the current date & time
            'Gathering the phone numbers for dialog from STAT/ADDR
            Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

            phone_number_list = "Select or Type|"
            If phone_number_one <> "" Then phone_number_list = phone_number_list & phone_number_one & "|"
            If phone_number_two <> "" Then phone_number_list = phone_number_list & phone_number_two & "|"
            If phone_number_three <> "" Then phone_number_list = phone_number_list & phone_number_three & "|"
            phone_number_array = split(phone_number_list, "|")

            Call convert_array_to_droplist_items(phone_number_array, phone_numbers)

            'HC Application Screening Dialog
            Do
            	Do
                    err_msg = ""
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 341, 325, "Health Care Contact"
                      DropListBox 10, 65, 65, 15, "Select one..."+chr(9)+"Phone Call"+chr(9)+"Unable to Reach"+chr(9)+"Voicemail", contact_type
                      DropListBox 80, 65, 45, 10, "to"+chr(9)+"from", contact_direction
                      ComboBox 130, 65, 85, 15, "Select or Type"+chr(9)+"Memb 01"+chr(9)+"Memb 02"+chr(9)+"AREP"+chr(9)+"SWKR"+chr(9)+who_contacted, who_contacted
                      EditBox 260, 65, 65, 15, METS_IC_number
                      ComboBox 70, 85, 75, 15, phone_numbers+chr(9)+phone_number, phone_number
                      EditBox 225, 85, 100, 15, when_contact_was_made
                      CheckBox 70, 100, 65, 10, "Used Interpreter", used_interpreter_checkbox
                      EditBox 75, 135, 250, 15, verifs_needed
                      DropListBox 265, 155, 60, 15, "Select"+chr(9)+"Yes"+chr(9)+"No", barrier_droplist
                      ButtonGroup ButtonPressed
                        PushButton 250, 170, 10, 15, "!", help_button
                      DropListBox 265, 170, 60, 15, "Select"+chr(9)+"Yes"+chr(9)+"No", reasonable_droplist
                      EditBox 75, 190, 250, 15, attested_verifs
                      DropListBox 100, 230, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", verif_confirm
                      DropListBox 100, 245, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", request_confirm
                      DropListBox 100, 260, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", atr_confirm
                      DropListBox 195, 230, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", avs_form_confirm
                      DropListBox 195, 245, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", avs_confirm
                      DropListBox 195, 260, 30, 15, "Pick"+chr(9)+"Yes"+chr(9)+"No", solq_confirm
                      CheckBox 235, 230, 80, 10, "Sent Work Number", work_number_checkbox
                      CheckBox 235, 240, 50, 10, "Sent TPQY", TPQY_checkbox
                      CheckBox 235, 250, 95, 10, "Sent VA benefit request.", VA_request_checkbox
                      ButtonGroup ButtonPressed
                        PushButton 230, 265, 90, 10, "Create VA Request Email", VA_button
                      EditBox 70, 285, 255, 15, other_notes
                      EditBox 70, 305, 150, 15, worker_signature
                      ButtonGroup ButtonPressed
                        OkButton 225, 305, 50, 15
                        CancelButton 275, 305, 50, 15
                        PushButton 150, 15, 60, 15, "Application Guide", app_guide_button
                        PushButton 210, 15, 55, 15, "COVID-19 FAQ", faq_button
                        PushButton 265, 15, 65, 15, "MAXIS Information", info_button
                      GroupBox 145, 5, 190, 35, "HC Policy/ Procedural Help"
                      Text 140, 55, 65, 10, "Who was contacted"
                      Text 10, 90, 50, 10, "Phone Number:"
                      Text 10, 15, 105, 10, "HC Application Date: " & hc_app_date
                      Text 65, 160, 195, 10, "Does the resident have a barrier to providing verifications?"
                      Text 65, 175, 185, 10, "If yes, is there a reasonable explanation for the barrier?"
                      Text 220, 70, 40, 10, "METS IC#:"
                      GroupBox 5, 5, 125, 35, "Health Care Information"
                      Text 15, 140, 60, 10, "Mandatory Verifs:"
                      GroupBox 10, 215, 325, 65, "Confirm your case actions below:"
                      Text 15, 235, 80, 10, "Reviewed Verifs on File:"
                      GroupBox 5, 45, 330, 70, "Contact Information:"
                      Text 60, 265, 35, 10, "Sent ATR:"
                      Text 25, 290, 40, 10, "Other notes:"
                      Text 30, 250, 65, 10, "Sent Verif Request:"
                      Text 20, 55, 40, 10, "Contact type"
                      Text 140, 230, 55, 10, "Sent AVS Form:"
                      Text 90, 55, 30, 10, "From/To"
                      Text 140, 245, 55, 10, "Submitted AVS:"
                      Text 150, 90, 75, 10, "Date/Time of Contact:"
                      Text 140, 260, 55, 10, "Checked SOLQ:"
                      Text 5, 310, 60, 10, "Worker signature:"
                      Text 10, 195, 65, 10, "Self-Attested Verifs:"
                      GroupBox 5, 120, 330, 90, "If you've connected with the resident/AREP review the following information:"
                      Text 10, 25, 105, 10, "Days HC is Pending: " & hc_days_pending
                    EndDialog
                    Dialog Dialog1
            		cancel_confirmation
            		If ButtonPressed = app_guide_button then
                        CreateObject("WScript.Shell").Run("https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-COVID9") 'COVID 19 Application Guide - OneSource
                        err_msg = "LOOP" & err_msg
                    End if
                    If ButtonPressed = faq_button then
                        CreateObject("WScript.Shell").Run("https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-COVID12")     'COVID 19 FAQ - OneSource
                        err_msg = "LOOP" & err_msg
                    End if
                    If ButtonPressed = info_button then
                        CreateObject("WScript.Shell").Run("https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-COVID7")     'MAXIS Information - OneSource
                        err_msg = "LOOP" & err_msg
                    End if
                    If ButtonPressed = VA_button then
                        VA_email = true
                        VA_info = "Name of Veteran: " & vbcr & "SSN of Veteran: " & vbcr & "VA File # (if known): " & vbcr & "Name of Spouse/Child receiving VA benefit (if applicable): " & vbcr & "SSN of Spouse/Child receiving VA benefit (if applicable): " & vbcr & "Relationship to Veteran (if applicable): "
                        'Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
                        Call create_outlook_email("Vetservices@Hennepin.us", "", "VA Request for Case #" & MAXIS_case_number, VA_info, "", False)   'will create email, will not send.
                        err_msg = "LOOP" & err_msg
                    End if
                    If ButtonPressed = help_button then
                        tips_tricks_msg = MsgBox("*** Tips and Tricks ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Existing policy allows enrollees or their authorized representative to provide a written statement or verbal conversation that is documented in case notes if they have a reasonable explanation for not being able to provide proofs or a signed release of information form for the worker to obtain the proofs. " & vbcr & vbcr & _
                        "For example, if a resident's workplace has been closed due to COVID-19 and they are unable to obtain verifications at this time. Reasonable explanations can include but are not limited to:" & vbNewLine & "- An employer not being available." & vbNewLine & "- The person is under quarantine." & vbNewLine & "- The person does not have access to photocopies or a fax machine.", vbInformation, "Tips and Tricks")
                        err_msg = "LOOP" & err_msg
                    End if

            	    'Mandatory fields
                    If contact_type = "Select one..." then err_msg = err_msg & vbcr & "* Enter the contact type."
                    If trim(who_contacted) = "" or who_contacted = "Select or Type" then err_msg = err_msg & vbcr & "* Enter who was contacted."
                    If trim(phone_number) = "" or trim(phone_number) = "Select or Type" then err_msg = err_msg & vbcr & "* Enter the phone number called."
                    If trim(when_contact_was_made) = "" then err_msg = err_msg & vbcr & "* Enter the date and time of contact."
                    If contact_type = "Phone Call" then
                        'mandotory fields for completing the health care application screening
                        If trim(verifs_needed) = "" then err_msg = err_msg & vbcr & "* Enter the mandatory verifications needed for this application."
                        If barrier_droplist = "Select" then err_msg = err_msg & vbcr &"* Provide an answer re: barrier to providing verifs."
                        IF reasonable_droplist = "Select" then err_msg = err_msg & vbcr &"* Provide an answer re: reason explanation for barrier"
                        If (reasonable_droplist = "Yes" and trim(attested_verifs) = "") then err_msg = err_msg & vbcr & "* Enter the self-attested verifications."
                        IF verif_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you review the verifs on file?"
                        IF request_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you send a verification request?"
                        IF atr_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you send an ATR?"
                        IF avs_form_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you send AVS forms for applicant?"
                        IF avs_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you submit request in the AVS system?"
                        IF solq_confirm = "Pick" then err_msg = err_msg & vbcr & "* Did you check SOLQ?"
                    End if
                    If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
                    IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
                LOOP UNTIL err_msg = ""									'loops until all errors are resolved
                CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
            Loop until are_we_passworded_out = false					'loops until user passwords back in

            'additions to the closing_message if these conditions apply
            If VA_email = True then closing_message = closing_message & vbcr & "Complete the rest of the VA email, and send to Vet services."
            If trim(attested_verifs) <> "" then closing_message = closing_message & vbcr & vbcr & "Use the POSTPONED CASE ACTIONS script on the power pad to track attested verifications."

            'THE CASENOTE----------------------------------------------------------------------------------------------------
            'case note header info
            If contact_type = "Phone Call" then
                interview_status = "Completed"
            Else
                interview_status = "Attempted"  'for either UNABLE TO REACH or VOICEMAIL options
            End if

            start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("Health Care Application Screening " & interview_status & " " & date)
            Call write_variable_in_CASE_NOTE("* " & contact_type & " " & contact_direction & " " & who_contacted & " completed at " & when_contact_was_made)
            Call write_bullet_and_variable_in_CASE_NOTE("Phone Number", phone_number)
            CALL write_bullet_and_variable_in_CASE_NOTE("METS/IC number", METS_IC_number)
            If interview_status = "Completed" then
                Call write_variable_in_CASE_NOTE("===Screening Inforamtion===")
                Call write_bullet_and_variable_in_CASE_NOTE("Mandatory Verifs", verifs_needed)
                Call write_bullet_and_variable_in_CASE_NOTE("Resident has a barrier to providing verifs", barrier_droplist)
                If barrier_droplist = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Resident has resonable explaination for the barrier", reasonable_droplist)
                CALL write_bullet_and_variable_in_CASE_NOTE("Self-Attested Verifs", attested_verifs)
            End if
            'Case actions
            CALL write_variable_in_CASE_NOTE("---")
            If verif_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Reviewed Verifs on File", verif_confirm)
            If request_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Sent Verification Request", request_confirm)
            If atr_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Sent ATR (Auth to Release Info)", atr_confirm)
            If avs_form_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Sent AVS Auth Form", avs_form_confirm)
            If avs_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Submitted AVS Request", avs_confirm)
            If solq_confirm <> "Pick" then Call write_bullet_and_variable_in_CASE_NOTE("Checked SOLQ system", solq_confirm)
            If work_number_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Sent Work Number request.")
            If TPQY_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Sent TPQY request.")
            If VA_request_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Sent VA request via VetServices.")
            CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
            CALL write_variable_in_CASE_NOTE("---")
            CALL write_variable_in_CASE_NOTE (worker_signature)
            PF3

            'Updating the PROG panel with the screening date
            If interview_status = "Completed" then
                no_call_reason = "" 'ensuring variable in case note is blank
                Call convert_date_into_MAXIS_footer_month(hc_app_date, MAXIS_footer_month, MAXIS_footer_year)   'converting application month into MAXIS footer month/year
                Call navigate_to_MAXIS_screen("STAT", "PROG")
                MAXIS_footer_month_confirmation                 'confirming we got to the correct footer month/year
                EMReadScreen hc_status_check, 4, 12, 74         'double checking for pending status in the application month, will update if pending
                If hc_status_check = "PEND" then
                    PF9
                    Call create_MAXIS_friendly_date(date, 0, 12, 55)    'adding in today's date to HC screening date field.
                    Transmit 'to save and exit
                    PF3     'to wrap screen
                    PF3     'to exit wrap screen
                    MAXIS_background_check
                    Call navigate_to_MAXIS_screen("STAT", "PROG")  'brings user back to STAT/PROG for APPLICATION CHECK DIALOG
                End if
            End if
        End if
    End if
End if

'----------------------------------------------------------------------------------------------------dialogs
If trim(attested_verifs) <> "" then verifs_rcvd = attested_verifs & "(**These are attested verifs.)" 'Carrying over the information from the attested verification field to be entered intot he verif rec'd field

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 386, 185, "Application Check: "  & application_check
  DropListBox 75, 15, 80, 15, "Select One:"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Pop"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer", app_type
  EditBox 175, 20, 50, 15, application_date
  DropListBox 75, 45, 155, 15, "Select One:"+chr(9)+"Interview still needed"+chr(9)+"Requested verifications not received"+chr(9)+"Partial verfications received, more are needed"+chr(9)+"Case is ready to approve or deny"+chr(9)+"Other", application_status_droplist
  CheckBox 245, 45, 135, 15, "Check to have an outlook reminder set", Outlook_reminder_checkbox
  CheckBox 275, 65, 95, 10, "Check to confirm ECF has", ECF_checkbox
  EditBox 95, 65, 170, 15, other_app_notes
  EditBox 95, 85, 170, 15, verifs_rcvd
  EditBox 95, 105, 170, 15, verifs_needed
  EditBox 95, 125, 170, 15, actions_taken
  EditBox 95, 145, 165, 15, other_notes
  EditBox 95, 165, 125, 15, worker_signature
  CheckBox 285, 125, 30, 10, "CASH", CASH_CHECKBOX
  CheckBox 330, 125, 25, 10, "EA", EA_CHECKBOX
  CheckBox 285, 135, 25, 10, "FS", FS_CHECKBOX
  CheckBox 330, 135, 30, 10, "GRH", GRH_CHECKBOX
  CheckBox 285, 145, 20, 10, "HC", HC_CHECKBOX
  ButtonGroup ButtonPressed
    OkButton 280, 165, 45, 15
    CancelButton 330, 165, 45, 15
  ButtonGroup ButtonPressed
    PushButton 240, 15, 30, 10, "AREP", AREP_button
    PushButton 345, 15, 30, 10, "JOBS", JOBS_button
    PushButton 240, 25, 30, 10, "PROG", PROG_button
    PushButton 275, 25, 30, 10, "REVW", REVW_button
    PushButton 310, 25, 30, 10, "SHEL", SHEL_button
    PushButton 345, 25, 30, 10, "UNEA", UNEA_button
    PushButton 275, 15, 30, 10, "DISA", DISA_button
    PushButton 310, 15, 30, 10, "HCRE", HCRE_button
  Text 10, 20, 55, 10, "Application type:"
  Text 175, 10, 55, 10, "Application date"
  Text 5, 50, 60, 10, "Application status:"
  Text 5, 70, 85, 10, "If status is 'other' explain:"
  Text 5, 90, 75, 10, "Verifications received:"
  Text 5, 110, 75, 10, "Pending verifications:"
  Text 5, 130, 50, 10, "Actions taken:"
  Text 5, 150, 45, 10, "Other notes:"
  Text 5, 170, 60, 10, "Worker signature:"
  GroupBox 5, 5, 160, 30, "Day 1 application check only"
  GroupBox 235, 5, 145, 35, "MAXIS navigation"
  GroupBox 280, 115, 95, 45, "Pending Programs"
  Text 275, 90, 100, 20, "**A TIKL to review the case will be created by the script"
  Text 285, 75, 90, 10, " been reviewed"
EndDialog
'main dialog
Do
	Do
		err_msg = ""
		dialog dialog1
		cancel_confirmation
		MAXIS_dialog_navigation
		If application_status_droplist = "Select One:" then err_msg = err_msg & vbNewLine & "* You must choose the application status."
		If application_status_droplist <> "Interview still needed" and actions_taken = ""  then err_msg = err_msg & vbNewLine & "* You must enter your case actions."
		If application_status_droplist = "Other" AND other_app_notes = ""  then err_msg = err_msg & vbNewLine & "* You must enter more information about the 'other' application status."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If other_app_notes <> "" Then application_status_droplist = application_status_droplist & ", " & other_app_notes

If Outlook_reminder_checkbox = CHECKED THEN
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	CALL create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "Application check: " & reminder_text & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
End if

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If application_status_droplist <> "Case is ready to approve or deny" then Call create_TIKL("Application check: " & reminder_text & ". Review case and ECF including verification requests, and take appropriate action.", 0, reminder_date, False, TIKL_note_text)

'THE CASENOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("-------------------------" & application_check & " Application Check")
IF application_check = "Day 1" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Type of application rec'd", app_type)
Call write_bullet_and_variable_in_CASE_NOTE("Application Status", application_status_droplist)
CALL write_bullet_and_variable_in_CASE_NOTE("Program(s) Applied For", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("Application Date", application_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Pending Programs", additional_programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", active_programs)
IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* ECF reviewed and verifications have been sent")
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Received", verifs_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("Pending Verifications", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
Call write_bullet_and_variable_in_case_note("Reason Health Care Interview Not Attempted", no_call_reason)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure_with_error_report(closing_message)
