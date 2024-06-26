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
Call changelog_update("05/23/2024", "This script, NOTES - APPLICATION CHECK will be retiring on 06/03/2024. This script was built for the CARL application and does not support ES Workflow's assignments.", "Ilse Ferris, Hennepin County")
Call changelog_update("04/16/2023", "Removed Health Care Screening functionality due to return to regular HC application rules.", "Ilse Ferris, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/10/2022", "Added new functionality to specifically support the review of cases with Cash, SNAP, GRH, or EMER pending that are at or after Day 30. These cases have reached the end of the Application Processing Period and should be reviewed for determination action, which may include denial. These updates better support the actions required for cases at this point.##~##", "Casey Love, Hennepin County") ''#1042
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
closing_message = "Application check is complete." 'setting up closing_message variable for possible additions later based on conditions

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
      	Call validate_MAXIS_case_number(display_ben_err_msg, "*")
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("REPT", "PND2") 'information gathering to auto-populate the application date
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROD
IF priv_check = "PRIV" then script_end_procedure("This case is privileged, and cannot be accessed.")

'Ensuring that the user is in REPT/PND2
Do
    basket_at_display_limit = False
    EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
    If pnd2_disp_limit = "Display Limit" Then
        transmit
        basket_at_display_limit = True
        EMReadScreen basket_at_limit, 7, 21, 13
    End If
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

If found_case = False then
    If basket_at_display_limit = True Then
        Call back_to_SELF
        EMWriteScreen MAXIS_case_number, 18, 43
        script_end_procedure_with_error_report("The script could not confirm that this case is pending on PND2 or gather additional necessary information from PND2 because this basket (" & basket_at_limit & ") is at the MAXIS PND2 Display limit. This is not a script issue, but a MAXIS limitation.")
    End If
    script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")
End If

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

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
msa_pending = False
grh_pending = False
If msa_status = "PENDING" Then msa_pending = True
If grh_status = "PENDING" Then grh_pending = True

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
EMReadScreen err_msg, 7, 24, 02
IF err_msg = "BENEFIT" THEN	script_end_procedure_with_error_report ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

'Reading the app date from PROG
EMReadScreen cash1_app_date, 8, 6, 33
EMReadScreen cash1_intvw_date, 8, 6, 55
cash1_app_date = replace(cash1_app_date, " ", "/")
cash1_intvw_date = replace(cash1_intvw_date, " ", "/")
EMReadScreen cash2_app_date, 8, 7, 33
EMReadScreen cash2_intvw_date, 8, 7, 55
cash2_app_date = replace(cash2_app_date, " ", "/")
cash2_intvw_date = replace(cash2_intvw_date, " ", "/")
EMReadScreen emer_app_date, 8, 8, 33
EMReadScreen emer_intvw_date, 8, 8, 55
emer_app_date = replace(emer_app_date, " ", "/")
emer_intvw_date = replace(emer_intvw_date, " ", "/")
EMReadScreen grh_app_date, 8, 9, 33
EMReadScreen grh_intvw_date, 8, 9, 55
grh_app_date = replace(grh_app_date, " ", "/")
grh_intvw_date = replace(grh_intvw_date, " ", "/")
EMReadScreen snap_app_date, 8, 10, 33
EMReadScreen snap_intvw_date, 8, 10, 55
snap_app_date = replace(snap_app_date, " ", "/")
snap_intvw_date = replace(snap_intvw_date, " ", "/")
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

'TODO: Remove and use determine_program_and_case_status_from_CASE_CURR functionality
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
interview_completed = False     'reading for interview information in case we are at Day 30
'cash I
IF cash1_status_check = "PEND" then
    If cash1_app_date = application_date THEN
        cash_pends = TRUE
		CASH_CHECKBOX = CHECKED
        programs_applied_for = programs_applied_for & "CASH, "
    Else
        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
	End if
    If cash1_intvw_date <> "__/__/__" Then
        interview_completed = True
        interview_date = cash1_intvw_date
    End If
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
    If cash2_intvw_date <> "__/__/__" Then
        interview_completed = True
        interview_date = cash2_intvw_date
    End If
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
    If snap_intvw_date <> "__/__/__" Then
        interview_completed = True
        interview_date = snap_intvw_date
    End If
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
    If grh_intvw_date <> "__/__/__" Then
        interview_completed = True
        interview_date = grh_intvw_date
    End If
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
    If emer_intvw_date <> "__/__/__" Then
        interview_completed = True
        interview_date = emer_intvw_date
    End If
End if

programs_applied_for = trim(programs_applied_for)       'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

additional_programs_applied_for = trim(additional_programs_applied_for)       'trims excess spaces of programs_applied_for
If right(additional_programs_applied_for, 1) = "," THEN additional_programs_applied_for = left(additional_programs_applied_for, len(additional_programs_applied_for) - 1)

pending_progs = trim(pending_progs) 'trims excess spaces of pending_progs
'takes the last comma off of pending_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(pending_progs, 1) = "," THEN pending_progs = left(pending_progs, len(pending_progs) - 1)
caf_programs_denial = False

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
    caf_programs_denial = True
Elseif (DateDiff("d", application_date, date) => 45 AND DateDiff("d", application_date, date) < 60) then
	application_check = "Day 45"
	reminder_date = dateadd("d", 60, application_date)
	reminder_text = "Day 60"
    caf_programs_denial = True
Elseif DateDiff("d", application_date, date) = 60 then
	application_check = "Day 60"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
    caf_programs_denial = True
Elseif DateDiff("d", application_date, date) > 60 then
	application_check = "Over 60 days"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
    caf_programs_denial = True
END IF

'Determining if a CAF based program is pending
IF cash1_status_check <> "PEND" and cash2_status_check <> "PEND" and snap_status_check  <> "PEND" and grh_status_check <> "PEND" and emer_status_check <> "PEND" Then caf_programs_denial = False

'If a CAF Based program is pending and the case is at or past Day 30, the script will pull this special Functionality
'The standard functionality for Application Checkk will NOT run if this case is at CAF programs denial
If caf_programs_denial = True Then
    allow_60_days = False
    If interview_completed = True Then interview_completed = "Yes"      'defaulting based on what is listed in PROG.'
    If interview_completed = False Then interview_completed = "No"
    day_30 = dateadd("d", 30, application_date)
    day_before_app = DateAdd("d", -1, application_date) 'will set the date one day prior to app date'

    Call navigate_to_MAXIS_screen("CASE", "NOTE")       'Reading CASE NOTE for notice dates.

    note_row = 5            'resetting the variables on the loop
    note_date = ""
    note_title = ""
    appt_date = ""
    appt_notc_date = ""
    nomi_date = ""

    Do
        EMReadScreen note_date, 8, note_row, 6      'reading the note date
        EMReadScreen note_title, 55, note_row, 25   'reading the note header
        note_title = trim(note_title)

        'These are Appointment Notice Headers
        IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
            appt_notc_date = note_date
        ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
            appt_notc_date = note_date
        ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
            EMReadScreen appt_date, 10, note_row, 63
            appt_date = replace(appt_date, "~", "")
            appt_date = trim(appt_date)
            appt_notc_date = note_date
            'MsgBox WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)
        END IF

        'These are NOMI Headers'
        IF note_title = "~ Client missed application interview, NOMI sent via sc" then nomi_date = note_date
        IF left(note_title, 32) = "**Client missed SNAP interview**" then nomi_date = note_date
        IF left(note_title, 32) = "**Client missed CASH interview**" then nomi_date = note_date
        IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then nomi_date = note_date
        IF note_title = "~ Client has not completed application interview, NOMI" then nomi_date = note_date
        IF note_title = "~ Client has not completed CASH APP interview, NOMI sen" then nomi_date = note_date
        IF note_title = "* A notice was previously sent to client with detail ab" then nomi_date = note_date

        note_row = note_row + 1
        IF note_row = 19 THEN
            PF8
            note_row = 5
        END IF
        EMReadScreen next_note_date, 8, note_row, 6
        IF next_note_date = "        " then Exit Do
    Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

    'Now we ask about the application steps. First is the interivew.
    Do
        Do
            err_msg = ""
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 346, 175, "Case May Be Ready for Denial"
              DropListBox 170, 110, 40, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", interview_completed
              EditBox 170, 130, 50, 15, interview_date
              ButtonGroup ButtonPressed
                OkButton 230, 155, 50, 15
                CancelButton 285, 155, 50, 15
                PushButton 160, 65, 165, 15, "CM 05.12.15 - Application Processing Standards", cm_05_12_15_btn
              GroupBox 5, 10, 330, 80, "Application Check: " & application_check
              Text 15, 25, 290, 10, "Programs Applied for: " & programs_applied_for
              Text 15, 40, 285, 10, "Cases at Day 30 for CAF based programs should be reviewed for possible denial."
              Text 15, 50, 315, 10, "A determination for eligibility should be made no later than 30 days from the Date of Application."
              Text 25, 65, 115, 10, "Application Date: " & application_date
              Text 25, 75, 115, 10, "Day 30: " & day_30
              Text 50, 115, 115, 10, "Has an Interview been completed?"
              Text 15, 135, 150, 10, "If so, what date was the interview completed?"
              GroupBox 5, 95, 330, 55, "SNAP, CASH, GRH, and EMER Require an interivew."
            EndDialog

            dialog Dialog1
            cancel_confirmation

            If interview_completed = "Select..."Then err_msg = err_msg & vbCr & "* Indicate if there was an interview completed."
            If interview_completed = "Yes" Then
                If IsDate(interview_date) = False Then
                    err_msg = err_msg & vbCr & "* Indicate the date that the interview was completed."
                ElseIf DateDiff("d", interview_date, application_date) > 0 Then
                    err_msg = err_msg & vbCr & "* The date entered is before the date of application, the interview can only be on or after the initial date of application."
                End If
            End If
            If ButtonPressed = cm_05_12_15_btn Then
                err_msg = "LOOP"
                run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051215"
            End If
            If err_msg <> "" and err_msg <> "LOOP" Then MsgBox "****  PLEASE RESOLVE TO CONTINUE ****" & vbCr & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = False

    'If the interview was not completed, we ask about the notices'
    If interview_completed = "No" Then
        Do
            Do
                err_msg = ""
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 346, 260, "Case May Be Ready for Denial"
                  EditBox 125, 175, 50, 15, appt_notc_date
                  EditBox 125, 195, 50, 15, nomi_date
                  ButtonGroup ButtonPressed
                    OkButton 230, 235, 50, 15
                    CancelButton 285, 235, 50, 15
                    PushButton 160, 65, 165, 15, "CM 05.12.15 - Application Processing Standards", cm_05_12_15_btn
                  GroupBox 5, 10, 330, 80, "Application Check: " & application_check
                  Text 15, 25, 290, 10, "Programs Applied for: " & programs_applied_for
                  Text 15, 40, 285, 10, "Cases at Day 30 for CAF based programs should be reviewed for possible denial."
                  Text 15, 50, 315, 10, "A determination for eligibility should be made no later than 30 days from the Date of Application."
                  Text 25, 65, 115, 10, "Application Date: " & application_date
                  Text 25, 75, 115, 10, "Day 30: " & day_30
                  GroupBox 5, 100, 330, 30, "SNAP, CASH, GRH, and EMER Require an interivew."
                  Text 20, 115, 155, 10, "INTERVIEW NOT COMPLETED"
                  GroupBox 5, 135, 330, 95, "Interview Notification Required"
                  Text 15, 150, 300, 20, "Two seperate NOTICES should have been sent to the resident for the interview notification. The resident should have received an Appointment Letter and a NOMI."
                  Text 25, 180, 100, 10, "Date Appointment Notice Sent:"
                  Text 70, 200, 55, 10, "Date NOMI Sent:"
                  Text 15, 215, 200, 10, "If these notices were not sent, leave these date fields blank."
                EndDialog

                dialog Dialog1
                cancel_confirmation

                If trim(appt_notc_date) <> "" and IsDate(appt_notc_date) = False Then err_msg = err_msg & vbCr & "* The date of the appointment notice does not appear to be a valid date. Please review and update."
                If trim(nomi_date) <> "" and IsDate(nomi_date) = False Then err_msg = err_msg & vbCr & "* The date of the NOMI does not appear to be a valid date. Please review and update."

                If ButtonPressed = cm_05_12_15_btn Then
                    err_msg = "LOOP"
                    run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051215"
                End If
                If err_msg <> "" and err_msg <> "LOOP" Then MsgBox "****  PLEASE RESOLVE TO CONTINUE ****" & vbCr & err_msg

            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = False
    End If

    'If the interview was completed, we ask about verifications.
    If interview_completed = "Yes" Then
        If (msa_pending = True or grh_pending = True) and reminder_text <> "Post day 60" Then allow_60_days = True
        Do
            Do
                err_msg = ""
                Dialog1 = ""
                If allow_60_days = True Then dlg_len = 305
                If allow_60_days = False Then dlg_len = 275
                BeginDialog Dialog1, 0, 0, 346, dlg_len, "Case May Be Ready for Denial"
                  DropListBox 110, 145, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", require_verifs
                  DropListBox 175, 175, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", verifs_sent
                  EditBox 175, 195, 50, 15, verifs_sent_date
                  DropListBox 175, 215, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", verifs_received
                  If allow_60_days = True Then
                      GroupBox 5, 245, 330, 35, "GRH or MSA Pending 30 days"
                      Text 15, 265, 180, 10, "Does this case have a disabled Household Member?"
                      DropListBox 190, 260, 45, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", disa_member_exists
                      Text 265, 265, 65, 10, "Pending Day " &  DateDiff("d", application_date, date)
                  End If
                  ButtonGroup ButtonPressed
                    OkButton 230, dlg_len-20, 50, 15
                    CancelButton 285, dlg_len-20, 50, 15
                    PushButton 160, 65, 165, 15, "CM 05.12.15 - Application Processing Standards", cm_05_12_15_btn
                  GroupBox 5, 10, 330, 80, "Application Check: " & application_check
                  Text 15, 25, 290, 10, "Programs Applied for: " & programs_applied_for
                  Text 15, 40, 285, 10, "Cases at Day 30 for CAF based programs should be reviewed for possible denial."
                  Text 15, 50, 315, 10, "A determination for eligibility should be made no later than 30 days from the Date of Application."
                  Text 25, 65, 115, 10, "Application Date: " & application_date
                  Text 25, 75, 115, 10, "Day 30: " & day_30
                  GroupBox 5, 100, 330, 30, "SNAP, CASH, GRH, and EMER Require an interivew."
                  Text 20, 115, 155, 10, "Interview Completed on: " & interview_date
                  GroupBox 5, 135, 330, 105, "Verifications may be Needed"
                  Text 15, 150, 95, 10, "Were Verifications Needed?"
                  GroupBox 20, 165, 240, 70, "If Verifications were Needed:"
                  Text 35, 180, 140, 10, "Was a Verification Request sent via ECF?"
                  Text 40, 200, 135, 10, "What Date was the request sent in ECF?"
                  Text 30, 220, 140, 10, "Were all mandatory verifications received?"
                  If (msa_pending = True or grh_pending = True) and reminder_text = "Post day 60" Then
                    Text 5, 245, 200, 20, "MSA or GRH are pending and it has been at least 60 days. (Pending Day " & DateDiff("d", application_date, date)  & ")"
                  End If
                EndDialog

                dialog Dialog1
                cancel_confirmation

                If require_verifs = "Select..."Then err_msg = err_msg & vbCr & "* Indicate if verifications were needed for this case."
                If require_verifs = "Yes" Then
                    If verifs_sent = "Select..."Then err_msg = err_msg & vbCr & "* Was the verification request form sent via ECF?."
                    If verifs_sent = "Yes" Then
                        If IsDate(verifs_sent_date) = False Then
                            err_msg = err_msg & vbCr & "* Enter the date the verification request form was sent from ECF."
                        ElseIf DateDiff("d", verifs_sent_date, application_date) > 0 Then
                            err_msg = err_msg & vbCr & "* The date entered for the verification sent date is prior to the date of application. The verifications cannot have been requested prior to the application receive date.."
                        End If
                    End If
                    If verifs_received = "Select..."Then err_msg = err_msg & vbCr & "* Indicate if all mandatory verifications have been received."
                    If allow_60_days = True Then
                        If disa_member_exists ="Select..." Then err_msg = err_msg & vbCr & " Indicate if any resident in this household is disabled."
                    End If
                End If
                If verifs_received = "Yes" Then err_msg = ""

                If ButtonPressed = cm_05_12_15_btn Then
                    err_msg = "LOOP"
                    run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051215"
                End If
                If err_msg <> "" and err_msg <> "LOOP" Then MsgBox "****  PLEASE RESOLVE TO CONTINUE ****" & vbCr & err_msg

            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = False
        If disa_member_exists <> "Yes" Then allow_60_days = False
    End If

    If allow_60_days = True Then
        If verifs_received = "Yes" Then allow_60_days = False
        If require_verifs = "No" Then allow_60_days = False
    End If

    'This is the logic bit to see what should happen to this case.
    should_deny_today = False
    complete_determination_today = False
    reason_cannot_deny = ""
    If allow_60_days = False or snap_status_check  = "PEND" Then
        If interview_completed = "No" Then
            If IsDate(appt_notc_date) = True and IsDate(nomi_date) = True Then
                If DateDiff("d", nomi_date, day_30) > 0 Then
                    should_deny_today = True                                        'DENY - No Interview, all notices sent and NOMI sent before Day 30'
                ElseIf DateDiff("d", nomi_date, date) >= 10 Then
                    should_deny_today = True                                        'DENY - No Interview, all notices sent and it has been 10 days since NOMI sent
                Else
                    reason_cannot_deny = "Interview Notices Not Timely"             'Delay denial - NOMI sent after Day 30 and has not been 10 days'
                End if
            End If
            If trim(appt_notc_date) = "" or trim(nomi_date) = "" Then
                reason_cannot_deny = "Missing Interview Notices"                    'Delay denial - No Interview - Notices not all sent
            End If
        Else
            If verifs_received = "Yes" Then
                complete_determination_today = True                                 'PROCESS - Interview done and Verifs Received'
            ElseIf require_verifs = "Yes" Then
                If verifs_sent = "No" Then
                    reason_cannot_deny = "Missing Verification Request"             'Delay Denial - Verifs Needed, no request sent'
                ElseIf IsDate(verifs_sent_date) = True Then
                    If DateDiff("d", verifs_sent_date, date) >= 10 Then
                        should_deny_today = True                                    'DENY - Interview and Verif request sent at least 10 days ago'
                    Else
                        reason_cannot_deny = "Verification Request Not Timely"      'Delay Denial - It has not been 10 days since mandaorty verifs requested'
                    End If
                End If
            ElseIf require_verifs = "No" Then
                complete_determination_today = True                                 'PROCESS - Interview done and No Verifs Needed'
            End If
        End If
    End If
    If complete_determination_today = True Then allow_60_days = False

    'saving information for error reporting'
    script_run_lowdown = script_run_lowdown & "application_check - " & application_check & vbCr
    script_run_lowdown = script_run_lowdown & "application_date - " & application_date & vbCr
    script_run_lowdown = script_run_lowdown & "day_30 - " & day_30 & vbCr
    script_run_lowdown = script_run_lowdown & "pending_progs - " & pending_progs & vbCr
    script_run_lowdown = script_run_lowdown & "programs_applied_for - " & programs_applied_for & vbCr
    script_run_lowdown = script_run_lowdown & "additional_programs_applied_for - " & additional_programs_applied_for & vbCr
    script_run_lowdown = script_run_lowdown & "interview_completed - " & interview_completed & vbCr
    script_run_lowdown = script_run_lowdown & "interview_date - " & interview_date & vbCr
    script_run_lowdown = script_run_lowdown & "appt_notc_date - " & appt_notc_date & vbCr
    script_run_lowdown = script_run_lowdown & "nomi_date - " & nomi_date & vbCr
    script_run_lowdown = script_run_lowdown & "require_verifs - " & require_verifs & vbCr
    script_run_lowdown = script_run_lowdown & "verifs_sent - " & verifs_sent & vbCr
    script_run_lowdown = script_run_lowdown & "verifs_sent_date - " & verifs_sent_date & vbCr
    script_run_lowdown = script_run_lowdown & "verifs_received - " & verifs_received & vbCr
    script_run_lowdown = script_run_lowdown & "should_deny_today - " & should_deny_today & vbCr
    script_run_lowdown = script_run_lowdown & "complete_determination_today - " & complete_determination_today & vbCr
    script_run_lowdown = script_run_lowdown & "reason_cannot_deny - " & reason_cannot_deny & vbCr
    script_run_lowdown = script_run_lowdown & "msa_pending - " & msa_pending
    script_run_lowdown = script_run_lowdown & "grh_pending - " & grh_pending
    script_run_lowdown = script_run_lowdown & "reminder_text - " & reminder_text
    script_run_lowdown = script_run_lowdown & "allow_60_days - " & allow_60_days
    script_run_lowdown = script_run_lowdown & "disa_member_exists - " & disa_member_exists

    'Creating the messages for the script_en_procedure to explain DENY or PROCESS discovery
    next_step_msg = ""
    If should_deny_today = True Then
        If allow_60_days = False Then next_step_msg = "*** This case should be denied today. ***" & vbCr & vbCr
        If allow_60_days = True Then next_step_msg = "*** This programs other than MSA/GRH should be denied today. ***" & vbCr & vbCr
        next_step_msg = next_step_msg & "Application Processing standard is to complete determining within 30 days." & vbCr & vbCr
        next_step_msg = next_step_msg & " - Day 30 for this case: " & day_30 & "." & vbCr & vbCr
        If interview_completed = "No" Then
            next_step_msg = next_step_msg & "The interview has not been completed at this time." & vbCr
            next_step_msg = next_step_msg & "We do not need to give more time for the completion of the interview. Notification of the interview requrement was sent to the resident per the requirements." & vbCr
            next_step_msg = next_step_msg & "Appointment Notice on " & appt_notc_date & "." & vbCr
            next_step_msg = next_step_msg & "NOMI on " & nomi_date & "." & vbCr
            next_step_msg = next_step_msg & "Denial should be completed via REPT/PND2, denying for no interview." & vbCr & vbCr
            If (msa_pending = True or grh_pending = True) and reminder_text <> "Post day 60" Then next_step_msg = next_step_msg & "MSA or GRH Pending - since there has been no interview, we cannot determine disability status and case should be denied at Day 30 if possible." & vbCr & vbCr
        End if
        If verifs_received = "No" and require_verifs = "Yes" and verifs_sent = "Yes" Then
            next_step_msg = next_step_msg & "The interview has been completed for this case." & vbCr
            next_step_msg = next_step_msg & "Verifications were required and requested on " & verifs_sent_date & "." & vbCr
            next_step_msg = next_step_msg & "Verifications have not been received." & vbCr
            next_step_msg = next_step_msg & "We do not need to povide more time for the submission of verifications. 10 days have been allowed for the return of verifications." & vbCr
            next_step_msg = next_step_msg & "Denial should be completed in STAT and ELIG." & vbCr & vbCr
        End If
        next_step_msg = next_step_msg & "Once the denial is processed the script NOTES - Eligibility Summary will support the CASE/NOTE completion." & vbCr & vbCr
        next_step_msg = next_step_msg & "Contact QI Knowledge Now for assistance with processing this denial if needed."
    End If

    If complete_determination_today = True Then
        next_step_msg = "*** Complete a Determination for this case today. ***" & vbCr & vbCr
        next_step_msg = next_step_msg & "This case:" & vbCr
        next_step_msg = next_step_msg & " - Has completed an interview." & vbCr
        If verifs_received = "Yes" Then
            next_step_msg = next_step_msg & " - All verifications have been received." & vbCr
        ElseIf require_verifs = "No" Then
            next_step_msg = next_step_msg & " - Does not require any verifications." & vbCr
        End If
        next_step_msg = next_step_msg & vbCr
        next_step_msg = next_step_msg & "Review all case information and complete a determination by updating STAT and approving in ELIG." & vbCr & vbCr
        next_step_msg = next_step_msg & "Once the case is processed the script NOTES - Eligibility Summary will support the CASE/NOTE completion." & vbCr & vbCr
        next_step_msg = next_step_msg & "Contact QI Knowledge Now for assistance with processing this denial if needed."
    End If
    If next_step_msg <> "" Then
        If allow_60_days = True Then MsgBox next_step_msg & vbCr & vbCr & "The script will now continue to the rest of Application Check to document the steps taken and additional needs for GRH or MSA."
        If allow_60_days = False Then
            closing_message = replace(closing_message, "Application check completed, a case note made, and a TIKL has been set.", "")

            If closing_message <> "" Then closing_message = next_step_msg & "----------------------------------------" & vbCr & closing_message
            If closing_message = "" Then closing_message = next_step_msg

            script_end_procedure_with_error_report(closing_message)     'end script -  NO CASE/NOTE
        End If
    End If

    'If there is a Delay Denial, the Application Check dialog has been mosified to explain details of the delay and allow for other information in a CASE NOTE
    If reason_cannot_deny <> "" Then
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 386, 250, "Application Check: "  & application_check
          Text 15, 60, 300, 10, "Denials should happen on Day 30. This denial must be delayed."
          If Instr(reason_cannot_deny, "Interview") <> 0 Then
              GroupBox 5, 45, 375, 95, "Denial Delay"
              Text 25, 85, 135, 10, "Appointment Notice Date: " & appt_notc_date
              Text 70, 95, 120, 10, "NOMI Date: " & nomi_date
              If reason_cannot_deny = "Missing Interview Notices" Then
                  Text 15, 75, 280, 10, "Denial cannot be completed as the correct notices have not been sent to the resident."
                  Text 15, 110, 355, 10, "No interview has been completed but the resident has not been fully informed of the interview requirement."
                  Text 15, 125, 220, 10, "CONTACT QI KNOWLEDGE NOW TO HAVE THE NOTICES SENT."
              End if
              If reason_cannot_deny = "Interview Notices Not Timely" Then
                  Text 15, 75, 280, 10, "Denial cannot be completed as the NOMI was not sent before day 30."
                  Text 15, 110, 360, 10, "We must waint until 10 days from the date the NOMI is sent to deny this case., since it was not sent before Day 30."
                  Text 15, 125, 310, 10, "Contact QI Knowledge Now if you have questions or concerns about the notices on this case."
              End If
          End If
          If Instr(reason_cannot_deny, "Verification") <> 0 Then
              GroupBox 5, 45, 375, 95, "Denial Delay"
              If reason_cannot_deny = "Missing Verification Request" Then
                  Text 15, 75, 345, 10, "Denial cannot be completed as the verification request was not sent for mandatory verifications."
                  ' Text 25, 85, 165, 10, "Verification Request Sent: " & verifs_sent_date
                  Text 15, 90, 355, 10, "SEND THE VERIFICATION REQUEST THROUGH ECF."
                  CheckBox 30, 100, 275, 10, "Check here to confirm you have sent the correct verification request in ECF.", verif_request_sent_via_ecf_checkkbox
                  Text 15, 115, 310, 10, "We must give the resident 10 days to provide mandatory verifications before denying."
              End If
              If reason_cannot_deny = "Verification Request Not Timely" Then
                  Text 15, 75, 345, 10, "Denial cannot be completed as it has not been 10 days since the Verification Request Form was sent."
                  Text 25, 85, 165, 10, "Verification Request Sent: " & verifs_sent_date
                  Text 15, 105, 310, 10, "We must give the resident 10 days to provide mandatory verifications before denying."
              End If
          End If
          EditBox 95, 150, 285, 15, verifs_rcvd
          EditBox 95, 170, 285, 15, verifs_missing
          EditBox 95, 190, 285, 15, actions_taken
          EditBox 95, 210, 285, 15, other_notes
          EditBox 95, 230, 125, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 280, 230, 45, 15
            CancelButton 330, 230, 45, 15
            PushButton 240, 15, 30, 10, "AREP", AREP_button
            PushButton 345, 15, 30, 10, "JOBS", JOBS_button
            PushButton 240, 25, 30, 10, "PROG", PROG_button
            PushButton 275, 25, 30, 10, "REVW", REVW_button
            PushButton 310, 25, 30, 10, "SHEL", SHEL_button
            PushButton 345, 25, 30, 10, "UNEA", UNEA_button
            PushButton 275, 15, 30, 10, "DISA", DISA_button
            PushButton 310, 15, 30, 10, "HCRE", HCRE_button
          Text 5, 155, 75, 10, "Verifications received:"
          Text 5, 175, 75, 10, "Pending verifications:"
          Text 5, 195, 50, 10, "Actions taken:"
          Text 5, 215, 45, 10, "Other notes:"
          Text 5, 235, 60, 10, "Worker signature:"
          GroupBox 235, 5, 145, 35, "MAXIS navigation"
          Text 5, 10, 205, 10, "Application Date: " & application_date
          Text 5, 20, 205, 10, "Day 30: " & day_30
          Text 5, 30, 220, 10, "Programs Applied for: " & programs_applied_for

        EndDialog

        Do
        	Do
        		err_msg = ""
        		dialog dialog1
        		cancel_confirmation
        		MAXIS_dialog_navigation
                If require_verifs = "Yes" and verifs_received = "No" and trim(verifs_missing) = "" Then err_msg = err_msg & vbCr & "* List the pending verirications that have not been returned."
                IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        	LOOP UNTIL err_msg = ""
        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        'THE CASENOTE----------------------------------------------------------------------------------------------------
        start_a_blank_CASE_NOTE
        CALL write_variable_in_CASE_NOTE("Application Denial Delayed")
        If reason_cannot_deny = "Missing Interview Notices" Then
            CALL write_variable_in_CASE_NOTE("No Interview has been completed.")
            CALL write_variable_in_CASE_NOTE("Resident has not received all notices ")
            CALL write_bullet_and_variable_in_CASE_NOTE("Appointment Notice Date", appt_notc_date)
            If trim(appt_notc_date) = "" Then CALL write_variable_in_CASE_NOTE("* NO Appointment Notice Sent.")
            CALL write_bullet_and_variable_in_CASE_NOTE("NOMI Date", nomi_date)
            If trim(nomi_date) = "" Then CALL write_variable_in_CASE_NOTE("* NO NOMI Sent.")
        End If
        If reason_cannot_deny = "Interview Notices Not Timely" Then
            CALL write_variable_in_CASE_NOTE("No Interview has been completed.")
            CALL write_variable_in_CASE_NOTE("NOMI was not sent before Day 30 so Denial cannot be completed until 10 days from the NOMI.")
            CALL write_bullet_and_variable_in_CASE_NOTE("Appointment Notice Date", appt_notc_date)
            CALL write_bullet_and_variable_in_CASE_NOTE("NOMI Date", nomi_date)
        End If
        If reason_cannot_deny = "Missing Verification Request" Then
            CALL write_variable_in_CASE_NOTE("Verifications were required for this case to complete a determination.")
            CALL write_variable_in_CASE_NOTE("Verification Request Form not found in case file.")
            If verif_request_sent_via_ecf_checkkbox = checked Then
                Call write_variable_in_CASE_NOTE("* Verification Request form sent today, " & date & ".")
                Call write_variable_in_CASE_NOTE("  10 days provided for return of verifications.")
                Call write_variable_in_CASE_NOTE("  Verifications due by: " & DateAdd("d", 10, date))
            End If

        End if
        If reason_cannot_deny = "Verification Request Not Timely" Then
            CALL write_variable_in_CASE_NOTE("Verifications were required for this case to complete a determination.")
            CALL write_variable_in_CASE_NOTE("We must provide 10 days for the return of Verifications.")
            CALL write_bullet_and_variable_in_CASE_NOTE("Verification Request Sent", verifs_sent_date)
            CALL write_variable_in_CASE_NOTE("  Verifications due by " & DateAdd("d", 10, verifs_sent_date))
        End if
        CALL write_bullet_and_variable_in_CASE_NOTE("Application Date", application_date)
        CALL write_bullet_and_variable_in_CASE_NOTE("Day 30", day_30)
        CALL write_bullet_and_variable_in_CASE_NOTE("Program(s) Applied For", programs_applied_for)
        CALL write_bullet_and_variable_in_CASE_NOTE("Other Pending Programs", additional_programs_applied_for)
        CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", active_programs)
        CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Received", verifs_rcvd)
        CALL write_bullet_and_variable_in_CASE_NOTE("Pending Verifications", verifs_missing)
        CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
        CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
        Call write_bullet_and_variable_in_case_note("Reason Health Care Interview Not Attempted", no_call_reason)
        CALL write_variable_in_CASE_NOTE("---")
        CALL write_variable_in_CASE_NOTE (worker_signature)

        If allow_60_days = True Then
            PF3
            closing_message = "CASE/NOTE for Delay of Denial for non-MSA/GRH programs has been made." & vbCr & vbCr &  "The script has continued to address GRH or MSA with: " & vbCr & closing_message
        End If
        If allow_60_days = False Then
            closing_message = replace(closing_message, ", a case note made, and a TIKL has been set.", " and a case note made.")
            script_end_procedure_with_error_report(closing_message)
        End If
    End If

    If allow_60_days = True Then
        other_notes = "Case has a disabled household member and "
        If msa_pending = True Then other_notes = other_notes & "MSA is Pending and "
        If grh_pending = True Then other_notes = other_notes & "GRH is Pending and "
        other_notes = other_notes & "cannot be denied until 60 days from the date of application."
    End If
End If

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
IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* Case file reviewed and verifications have been sent")
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Received", verifs_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("Pending Verifications", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
Call write_bullet_and_variable_in_case_note("Reason Health Care Interview Not Attempted", no_call_reason)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/10/2022
'--Tab orders reviewed & confirmed----------------------------------------------11/10/2022
'--Mandatory fields all present & Reviewed--------------------------------------11/10/2022
'--All variables in dialog match mandatory fields-------------------------------11/10/2022
'Review dialog names for content and content fit in dialog----------------------04/16/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------The program checkboxes are not doing anything BUT will review in rewrite, issue #803
'--CASE:NOTE Header doesn't look funky------------------------------------------11/10/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------11/10/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------11/10/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/16/2023----------------------N/A     
'--MAXIS_background_check reviewed (if applicable)------------------------------04/16/2023----------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------04/16/2023
'--Out-of-County handling reviewed----------------------------------------------04/16/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/10/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------11/10/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/16/2023----------------------N/A: Will review in rewrite, issue #803
'--Incrementors reviewed (if necessary)-----------------------------------------04/16/2023----------------------N/A: Will review in rewrite, issue #803
'--Denomination reviewed -------------------------------------------------------11/10/2022
'--Script name reviewed---------------------------------------------------------11/10/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/10/2022
'--comment Code-----------------------------------------------------------------11/10/2022
'--Update Changelog for release/update------------------------------------------04/16/2023
'--Remove testing message boxes-------------------------------------------------11/10/2022
'--Remove testing code/unnecessary code-----------------------------------------11/10/2022
'--Review/update SharePoint instructions----------------------------------------04/16/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/16/2023----------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/16/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------04/16/2023----------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------04/16/2023
'--Update project team/issue contact (if applicable)----------------------------04/16/2023
