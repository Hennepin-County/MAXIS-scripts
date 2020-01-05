'GATHERING STATS===========================================================================================
name_of_script = "NOTES - APPLICATION CHECK.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 116, 45, "Application Check"
  EditBox 65, 5, 45, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 5, 25, 50, 15
    CancelButton 60, 25, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

Do
	DO
		err_msg = ""
	    dialog dialog1
      	cancel_without_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'information gathering to auto-populate the application date
Call navigate_to_MAXIS_screen("REPT", "PND2")

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
MAXIS_row = 7
Do
    EmReadscreen pending_case_num, 8, MAXIS_row, 6
    If trim(pending_case_num) = trim(MAXIS_case_number) then
        found_case = True
        exit do
    Else
        MAXIS_row = MAXIS_row + 1
        found_case = False
    End if

LOOP until row = 19
If found_case = False then script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

EMReadScreen app_month, 2, MAXIS_row, 38
EMReadScreen app_day, 2, MAXIS_row, 41
EMReadScreen app_year, 2, MAXIS_row, 44
EMReadScreen days_pending, 3, MAXIS_row, 50
EMReadScreen additional_application_check, 14, MAXIS_row + 1, 17
EMReadScreen add_app_month, 2, MAXIS_row + 1, 38
EMReadScreen add_app_day, 2, MAXIS_row + 1, 41
EMReadScreen add_app_year, 2, MAXIS_row + 1, 44

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
		MAXIS_row = MAXIS_row + 1
	END IF
End if

EMReadScreen PEND_CASH_check,	1, MAXIS_row, 54
EMReadScreen PEND_SNAP_check, 1, MAXIS_row, 62
EMReadScreen PEND_HC_check, 1, MAXIS_row, 65
EMReadScreen PEND_EMER_check,	1, MAXIS_row, 68
EMReadScreen PEND_GRH_check, 1, MAXIS_row, 72

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'EMReadScreen application_date, 8, 6, 33

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

'trims excess spaces of pending_progs
pending_progs = trim(pending_progs)
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
Elseif (DateDiff("d", application_date, date) > 1 AND DateDiff("d", application_date, date) < 9) then
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
'----------------------------------------------------------------------------------------------------dialogs
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 386, 185, "Application Check: "  & application_check
  DropListBox 75, 15, 80, 15, "Select One:"+chr(9)+"ApplyMN"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Pop"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer", app_type
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
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If other_app_notes <> "" Then application_status_droplist = application_status_droplist & ", " & other_app_notes

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
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Recieved", verifs_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("Pending Verifications", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

If Outlook_reminder_checkbox = CHECKED THEN
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	CALL create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "Application check: " & reminder_text & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
End if

If application_status_droplist <> "Case is ready to approve or deny" then
   Call navigate_to_MAXIS_screen("DAIL", "WRIT")
   CALL create_MAXIS_friendly_date(reminder_date, 0, 5, 18)   'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
   CALL write_variable_in_TIKL("Application check: " & reminder_text & ". Review case and ECF including verification requests, and take appropriate action.")
   PF3		'Exits and saves TIKL
End if

script_end_procedure_with_error_report("Application check completed, a case note made, and a TIKL has been set.")
