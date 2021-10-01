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
BeginDialog Dialog1, 0, 0, 276, 125, "Notes On Demand"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  CheckBox 190, 10, 80, 10, "Update STAT/PROG", prog_updated_CHECKBOX
  DropListBox 55, 25, 215, 15, "Select One:"+chr(9)+"Case was not pended timely"+chr(9)+"Client completed application interview"+chr(9)+"Client has not completed application interview"+chr(9)+"Denied programs for no interview"+chr(9)+"Interview not needed for MFIP to SNAP transition"+chr(9)+"Other(please describe)", case_status_dropdown
  EditBox 55, 45, 215, 15, other_notes
  EditBox 220, 65, 50, 15, interview_date
  EditBox 220, 85, 50, 15, case_note_date
  ButtonGroup ButtonPressed
    OkButton 165, 105, 50, 15
    CancelButton 220, 105, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 45, 10, "Case status:"
  Text 5, 70, 165, 10, "Date NOMI sent or interview completed case note:"
  Text 5, 50, 45, 10, "Other notes:"
  Text 5, 90, 90, 10, "Date of appt or case note:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		IF case_status_dropdown = "Case was not pended timely" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid NOMI date."
		IF case_status_dropdown = "Denied programs for no interview" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid NOMI date."
		IF case_status_dropdown = "Client completed application interview" and prog_updated_CHECKBOX <> CHECKED THEN err_msg = err_msg & vbNewLine & "* Please update PROG with interview date."
		IF case_status_dropdown = "Client completed application interview" and case_note_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case note date."
		IF case_status_dropdown = "Client completed application interview" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid interview date."
		IF case_status_dropdown = "Client has not completed application interview" THEN
		 	IF interview_date = "" or isDate(interview_date) = False THEN err_msg = err_msg & vbcr & "* Pleasse enter a valid interview date."
		END IF
		IF case_status_dropdown = "Other(please describe)" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a description of what occurred."
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Checking for PRIV cases.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
MAXIS_background_check      'Making sure we are out of background.

'Grabbing case and program status information from MAXIS.
'For tis script to work correctly, these must be correct BEFORE running the script.
CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
EMReadScreen case_status, 15, 8, 9                  'Now we are reading the CASE STATUS string from the panel - we want to make sure this does NOT read CAF1 PENDING
EMReadScreen pnd2_appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
ive_status = "INACTIVE"                             'There are some programs that are NOT read from the function and are pretty specific to this script/functionality
cca_status = "INACTIVE"                             'defaulting these statuses to 'INACTIVE' until they are read from the panel
ega_status = "INACTIVE"
ea_status = "INACTIVE"
'This functionality is how the above function reads for program information - just pulled out for these specific programs
row = 1                                             'Looking for IV-E information
col = 1
EMSearch "IV-E:", row, col
If row <> 0 Then
    EMReadScreen ive_status, 9, row, col + 6
    ive_status = trim(ive_status)
    If ive_status = "ACTIVE" or ive_status = "APP CLOSE" or ive_status = "APP OPEN" Then ive_status = "ACTIVE"
    If ive_status = "PENDING" Then case_pending = TRUE      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for CCAP information
col = 1
EMSearch "CCAP", row, col
If row <> 0 Then
    EMReadScreen cca_status, 9, row, col + 6
    cca_status = trim(cca_status)
    If cca_status = "ACTIVE" or cca_status = "APP CLOSE" or cca_status = "APP OPEN" Then cca_status = "ACTIVE"
    If cca_status = "PENDING" Then case_pending = TRUE      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for EGA information
col = 1
EMSearch "EGA", row, col
If row <> 0 THEN
    EMReadScreen ega_status, 9, row, col + 6
    ega_status = trim(ega_status)
    If ega_status = "ACTIVE" or ega_status = "APP CLOSE" or ega_status = "APP OPEN" Then ega_status = "ACTIVE"
    If ega_status = "PENDING" Then case_pending = TRUE      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for EA information
col = 1
EMSearch "EA: ", row, col
If row <> 0 THEN
    EMReadScreen ea_status, 9, row, col + 5
    ea_status = trim(ea_status)
    If ea_status = "ACTIVE" or ea_status = "APP CLOSE" or ea_status = "APP OPEN" THEN ea_status = "ACTIVE"
    If ea_status = "PENDING" THEN case_pending = TRUE      'Updating the case_pending variable from the function
End If

case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "CASE IS PENDING - " & case_pending        'Adding details about CASE/CURR information to a script report out to BZST
If case_status = "CAF1 PENDING" OR case_pending = False THEN                    'The case MUST be pending and NOT in PND1 to continue.
    call script_end_procedure_with_error_report("This case is not in PND2 status. Current case status in MAXIS is " & case_status & ". Update MAXIS to put this case in PND2 status. Script will not navigate.")
End If

call back_to_SELF           'resetting

multiple_app_dates = False                          'defaulting the Boolean about multiple application dates to FALSE
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
script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "MULTIPLE APP DATES" & additional_application_check        'Adding details about CASE/CURR information to a script report out to BZST

IF IsDate(application_date) = False THEN                   'If we could NOT find the application date - then it will use the PND2 application date.
    application_date = pnd2_appl_date
END IF

active_programs = ""        'Creates a variable that lists all the active programs on the case.
If ga_status = "ACTIVE" Then active_programs = active_programs & "GA, "
If msa_status = "ACTIVE" Then active_programs = active_programs & "MSA, "
If mfip_status = "ACTIVE" Then active_programs = active_programs & "MFIP, "
If dwp_status = "ACTIVE" Then active_programs = active_programs & "DWP, "
If ive_status = "ACTIVE" Then active_programs = active_programs & "IV-E, "
If grh_status = "ACTIVE" Then active_programs = active_programs & "GRH, "
If snap_status = "ACTIVE" Then active_programs = active_programs & "SNAP, "
If ega_status = "ACTIVE" Then active_programs = active_programs & "EGA, "
If ea_status = "ACTIVE" Then active_programs = active_programs & "EA, "
If cca_status = "ACTIVE" Then active_programs = active_programs & "CCA, "
If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then active_programs = active_programs & "HC, "

active_programs = trim(active_programs)  'trims excess spaces of active_programs
If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

programs_applied_for = ""        'Creates a variable that lists all the pending programs on the case.
IF unknown_cash_pending = TRUE THEN programs_applied_for = programs_applied_for & "Cash, "
IF ga_status = "PENDING" THEN programs_applied_for = programs_applied_for & "GA, "
IF msa_status = "PENDING" THEN programs_applied_for = programs_applied_for & "MSA, "
IF mfip_status = "PENDING" THEN programs_applied_for = programs_applied_for & "MFIP, "
IF dwp_status = "PENDING" THEN programs_applied_for = programs_applied_for & "DWP, "
IF ive_status = "PENDING" THEN programs_applied_for = programs_applied_for & "IV-E, "
IF grh_status = "PENDING" THEN programs_applied_for = programs_applied_for & "GRH, "
IF snap_status = "PENDING" THEN programs_applied_for = programs_applied_for & "SNAP, "
IF ega_status = "PENDING" THEN programs_applied_for = programs_applied_for & "EGA, "
IF ea_status = "PENDING" THEN programs_applied_for = programs_applied_for & "EA, "
IF cca_status = "PENDING" THEN programs_applied_for = programs_applied_for & "CCA, "
IF ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = TRUE THEN programs_applied_for = programs_applied_for & "HC, "

programs_applied_for = trim(programs_applied_for)  'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

CALL navigate_to_MAXIS_screen("STAT", "PROG")           'going here because this is a good background for the dialog to display against.

IF prog_updated_CHECKBOX = CHECKED THEN
    Do
    	EMReadscreen panel_check, 4, 2, 50
    	IF panel_check <> "PROG" THEN CALL navigate_to_MAXIS_screen("STAT", "PROG")
    Loop until HCRE_panel_check = "PROG"		'repeats until case is not in the HCRE panel
	PF9
	'interview_date = date & ""		'Defaults the date of the interview to today's date.
	intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
	intv_day = DatePart("d", interview_date)
	intv_yr = DatePart("yyyy", interview_date)

	intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
	intv_day = right("00"&intv_day, 2)
	intv_yr = right(intv_yr, 2)

	IF snap_status = "PENDING" THEN      'If SNAP interview to be updated
		EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
		EMWriteScreen intv_day, 10, 58
		EMWriteScreen intv_yr, 10, 61
	End If

	IF programs_applied_for = "Cash" THEN
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

	TRANSMIT                                    'Saving the panel
	closing_message = closing_message & " PROG has been updated, please review for accuracy."

END IF

IF IsDate(application_date) = False THEN
    stop_early_msg = "This script cannot continue as the application date could not be found from MAXIS."
    stop_early_msg = stop_early_msg & vbCr & vbCr & "CASE: " & MAXIS_case_number
    stop_early_msg = stop_early_msg & vbCr & "Application Date: " & application_date
    stop_early_msg = stop_early_msg & vbCr & "Programs applied for: " & programs_applied_for
    stop_early_msg = stop_early_msg & vbCr & vbCr & "If you are unsure why this happened, screenshot this and send it to HSPH.EWS.BlueZoneScripts@hennepin.us"
    CALL script_end_procedure_with_error_report(stop_early_msg)
END IF

'NOW WE START CASE NOTING - there are a few
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("~ Review of Application Received on " & application_date & " ~")
CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Pending Programs", additional_programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
IF case_status_dropdown = "Client completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " on "  & interview_date & " PROG updated ~")
	CALL write_variable_in_CASE_NOTE("* Completed by previous worker per case note dated: " & case_note_date)
ELSEIF case_status_dropdown = "Client has not completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " ~")
	CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
	CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & interview_date)
	CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview.")
	CALL write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice.")
'ELSEIF case_status_dropdown = "Client has not completed CASH application interview" THEN
	'CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " NOMI sent ~")
	'CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an     interview.")
	'CALL write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file     an application will receive a denial notice")
	'CALL write_variable_in_CASE_NOTE("* Face to face interview required because client has not been open on CASH in the     last 12 months")
 	'CALL write_variable_in_CASE_NOTE("* SNAP interview completed by previous worker per case note dated: " & case_note_date)
ELSEIF case_status_dropdown = "Case was not pended timely" THEN
    CALL write_variable_in_CASE_NOTE("~ Client has not completed application interview ~")
    CALL write_variable_in_CASE_NOTE("* Application date:" & application_date)
    CALL write_variable_in_CASE_NOTE("* NOMI sent to client on:" & interview_date)
    CALL write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it. Because the case was not pended timely a NOMI still needs to be sent and adequate time provided to the client to comply.")
ELSEIF case_status_dropdown = "Denied programs for no interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " - " & programs_applied_for & " for no interview" & " ~")
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
