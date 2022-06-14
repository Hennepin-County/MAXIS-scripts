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
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
correction_process = "On Demand Applications" 'for the next script'
'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
back_to_SELF' added to ensure we have the time to update and send the case in the background


'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 226, 120, "Notes On Demand"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  DropListBox 55, 25, 165, 15, "Select One:"+chr(9)+"Additional application pended prior to interview being completed"+chr(9)+"Case was not pended timely"+chr(9)+"Clear case note for other pending program pending not mentioned"+chr(9)+"Client completed application interview"+chr(9)+"Client has not completed application interview"+chr(9)+"Client has not completed F2F CASH application interview"+chr(9)+"Client contact was made no interview offered"+chr(9)+"Denied programs for no interview"+chr(9)+"Interview completed on NOMI day and NOMI was not cancelled"+chr(9)+"Interview not needed for MFIP to SNAP transition"+chr(9)+"Script-Application Received-not used when pending case"+chr(9)+"Script-CAF-used but no interview completed"+chr(9)+"Script-CAF-not used but approval made"+chr(9)+"Worker completed denial for no interview in ELIG"+chr(9)+"Other(please describe)", case_status_dropdown
  EditBox 55, 40, 165, 15, other_notes
  EditBox 170, 60, 50, 15, interview_date
  EditBox 170, 80, 50, 15, case_note_date
  CheckBox 5, 105, 85, 10, "Send Email Correction", correction_email_CHECKBOX
  ButtonGroup ButtonPressed
    OkButton 115, 100, 50, 15
    CancelButton 170, 100, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 45, 10, "Case status:"
  Text 5, 65, 165, 10, "Date NOMI sent or interview completed case note:"
  Text 5, 45, 45, 10, "Other notes:"
  Text 5, 85, 90, 10, "Date of appt or case note:"
  CheckBox 140, 10, 80, 10, "Update STAT/PROG", prog_updated_CHECKBOX
EndDialog

'Runs the first dialog
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
		IF case_status_dropdown = "Client has not completed application interview" and interview_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date that the NOMI was sent."
		IF case_status_dropdown = "Other(please describe)" and issue_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a description of what occured."
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'----------------------------------------------------------------------------------------------------'pending & active programs information
'information gathering to auto-populate the application date
back_to_self
EMWriteScreen MAXIS_case_number, 18, 43
Call navigate_to_MAXIS_screen("REPT", "PND2")

limit_reached = FALSE
row = 1
col = 1
EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
If row <> 0 Then
    transmit
    limit_reached = TRUE
End If

'Ensuring that the user is in REPT/PND2
Do
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check <> "PND2" then
		back_to_SELF
		Call navigate_to_MAXIS_screen("REPT", "PND2")
	End if
LOOP until PND2_check = "PND2"

'checking the case to make sure there is a pending case.  If not script will end & inform the user no pending case exists in PND2 this does not tell you if it is active or pnd1
EMReadScreen not_pending_check, 5, 24, 2
If not_pending_check = "CASE " THEN script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

If limit_reached = TRUE Then
    MAXIS_row = 7
    Do
        EMReadScreen PND2_case_number, 8, MAXIS_row, 5
        if trim(PND2_case_number) = MAXIS_case_number Then Exit Do
        MAXIS_row = MAXIS_row + 1
    Loop until MAXIS_row = 19
Else
    EMGetCursor MAXIS_row, MAXIS_col
End If
If MAXIS_row > 18 Then script_end_procedure("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

'grabs row and col number that the cursor is at this is where the application date comes in
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

MAXIS_footer_month = right("00" & DatePart("m", application_date), 2)
MAXIS_footer_year = right(DatePart("yyyy", application_date), 2)

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
programs_applied_for = ""   'Creates a variable that lists all pending cases.
additional_programs_applied_for = ""
'cash I
IF cash1_status_check = "PEND" then
    If cash1_app_date = application_date THEN
        cash_pends = TRUE
        programs_applied_for = programs_applied_for & "CASH, "
    Else
        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
    End if
End if
'cash II
IF cash2_status_check = "PEND" then
    if cash2_app_date = application_date THEN
        cash2_pends = TRUE
        programs_applied_for = programs_applied_for & "CASH, "
    Else
        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
    End if
End if
'SNAP
IF snap_status_check  = "PEND" then
    If snap_app_date  = application_date THEN
        SNAP_pends = TRUE
        programs_applied_for = programs_applied_for & "SNAP, "
    else
        additional_programs_applied_for = additional_programs_applied_for & "SNAP, "
    end if
End if
'GRH
IF grh_status_check = "PEND" then
    If grh_app_date = application_date THEN
        grh_pends = TRUE
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
        IF emer_prog_check = "EG" THEN programs_applied_for = programs_applied_for & "EGA, "
        IF emer_prog_check = "EA" THEN programs_applied_for = programs_applied_for & "EA, "
    else
        IF emer_prog_check = "EG" THEN additional_programs_applied_for = additional_programs_applied_for & "EGA, "
        IF emer_prog_check = "EA" THEN additional_programs_applied_for = additional_programs_applied_for & "EA, "
    End if
End if

programs_applied_for = trim(programs_applied_for)       'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

additional_programs_applied_for = trim(additional_programs_applied_for)       'trims excess spaces of programs_applied_for
If right(additional_programs_applied_for, 1) = "," THEN additional_programs_applied_for = left(additional_programs_applied_for, len(additional_programs_applied_for) - 1)

IF programs_applied_for = "" THEN
    DO
    	prog_confirmation = MsgBox("Press YES to confirm this application is PND1 and has no programs selected. If this is not the case select NO and run the script again.", vbYesNo, "Program confirmation")
    	IF prog_confirmation = vbNo THEN script_end_procedure_with_error_report("The script has ended. The application has not been acted on.")
    	IF prog_confirmation = vbYes THEN
    		EXIT DO
    	END IF
    Loop
END IF
'pended_date = date need for future reference

'intv_day = right("00" & DatePart("d", date), 2)
'Intv_mo  = right("00" & DatePart("m", date), 2)
'intv_yr  = right(DatePart("yyyy", date), 2)
If cash_pends = TRUE Then
    EmReadscreen PROG_interview_date, 8, 6, 55
    'If interview_date = "__ __ __" Then
        'EmWriteScreen intv_mo, 6, 55
        'EmWriteScreen intv_day, 6, 58
        'EmWriteScreen intv_yr, 6, 61
    'End If
End If
If cash2_pends = TRUE Then
    EmReadscreen PROG_interview_date, 8, 7, 55
    'If interview_date = "__ __ __" Then
        'EmWriteScreen intv_mo, 7, 55
        'EmWriteScreen intv_day, 7, 58
        'EmWriteScreen intv_yr, 7, 61
    'End If
End If
If SNAP_pends = TRUE Then
    EmReadscreen PROG_interview_date, 8, 10, 55
    'If interview_date = "__ __ __" Then
    '    EmWriteScreen intv_mo, 10, 55
    '    EmWriteScreen intv_day, 10, 58
    '    EmWriteScreen intv_yr, 10, 61
    'End If
End If
case_note = TRUE
DropListBox 55, 25, 215, 15, "Select One:"+chr(9)+"Additional application pended prior to interview being completed"+chr(9)+"Case was not pended timely"+chr(9)+"Clear case note for other pending program pending not mentioned"+chr(9)+"Client completed application interview"+chr(9)+"Client has not completed application interview"+chr(9)+"Client has not completed F2F CASH application interview"+chr(9)+"Client contact was made no interview offered"+chr(9)+"Denied programs for no interview"+chr(9)+"Interview completed on NOMI day and NOMI was not cancelled"+chr(9)+"Interview not needed for MFIP to SNAP transition"+chr(9)+"Script-Application Received-not used when pending case"+chr(9)+"Script-CAF-used but no interview completed"+chr(9)+"Script-CAF-not used but approval made"+chr(9)+"Worker completed denial for no interview in ELIG"+chr(9)+"Other(please describe)", case_status_dropdown


IF case_status_dropdown = "Additional application pended prior to interview being completed" OR case_status_dropdown = 	"Clear case note for other pending program pending not mentioned" OR case_status_dropdown = "Client contact was made no interview offered" OR case_status_dropdown = "Interview completed on NOMI day and NOMI was not cancelled" OR case_status_dropdown = "Interview not needed for MFIP to SNAP transition" OR case_status_dropdown = "Script-Application Received-not used when pending case" OR case_status_dropdown = "Script-CAF-used but no interview completed" OR case_status_dropdown = "Script-CAF-not used but approval made" OR case_status_dropdown = "Worker completed denial for no interview in ELIG" THEN
	case_note = FALSE

	IF case_status_dropdown = "Script-Application Received-not used when pending case" THEN  app_recvd_script_not_used_checkbox = CHECKED

	IF case_status_dropdown = "Client completed application interview" THEN prog_not_updated_interview_completed_checkbox = CHECKED

	IF case_status_dropdown = "Additional application pended prior to interview being completed"  THEN email_notes = case_status_dropdown

	IF case_status_dropdown = 	"Clear case note for other pending program pending not mentioned" THEN  cash_detail_not_covered_checkbox = CHECKED

	IF case_status_dropdown = "Cash and SNAP were applied for, Cash needs a face to face, SNAP should have had a phone interview offered." THEN snap_phone_interview_should_have_been_offered_checkbox = CHECKED

	IF case_status_dropdown = "Worker completed denial for no interview in ELIG" THEN denied_for_no_interview_NOT_on_PND2_checkbox = CHECKED

	IF case_status_dropdown = "Script-CAF-not used but approval made" THEN  caf_script_used_no_interview_checkbox = CHECKED

	IF case_status_dropdown = "Client contact was made no interview offered" THEN  client_contacted_should_have_been_offered_interview_checkbox = CHECKED

	IF case_status_dropdown = "Interview completed on NOMI day and NOMI was not cancelled" THEN  interview_completed_same_day_NOMI_sent_checkbox = CHECKED
	CALL run_another_script("C:\MAXIS-scripts\admin\send-email-correction.vbs")
END IF

IF case_note = TRUE THEN
	start_a_blank_CASE_NOTE
    IF case_status_dropdown = "Client completed application interview" THEN
    	Call write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " on "  & interview_date & " PROG updated ~")
    	Call write_variable_in_CASE_NOTE("* Completed by previous worker per case note dated: " & case_note_date)
    ELSEIF case_status_dropdown = "Client has not completed application interview" THEN
    	Call write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " ~")
    	Call write_variable_in_CASE_NOTE("* Application date: " & application_date)
    	Call write_variable_in_CASE_NOTE("* NOMI sent to client on: " & interview_date  )
    	Call write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview.")
    	Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice.")
    'ELSEIF case_status_dropdown = "Client has not completed CASH application interview" THEN
    	'Call write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " NOMI sent ~")
    	'Call write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an     interview.")
    	'Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file     an application will receive a denial notice")
    	'Call write_variable_in_CASE_NOTE("* Face to face interview required because client has not been open on CASH in the     last 12 months")
     	'Call write_variable_in_CASE_NOTE("* SNAP interview completed by previous worker per case note dated: " &     case_note_date)
    ELSEIF case_status_dropdown = "Case was not pended timely" THEN
        Call write_variable_in_CASE_NOTE("~ Client has not completed application interview ~")
        Call write_variable_in_CASE_NOTE("* Application date:" & application_date)
        Call write_variable_in_CASE_NOTE("* NOMI sent to client on:" & interview_date  )
        Call write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it. Because the case was not pended timely a NOMI still needs to be sent and adequate time provided to the client to comply.")
    ELSEIF case_status_dropdown = "Denied programs for no interview" THEN
    	Call write_variable_in_CASE_NOTE("~ " & case_status_dropdown & " - " & programs_applied_for & " for no interview" & " ~")
        Call write_variable_in_CASE_NOTE("* Application date:" & application_date)
        Call write_variable_in_CASE_NOTE("* Reason for denial: interview was not completed timely")
        Call write_variable_in_CASE_NOTE("* NOMI sent to client on:" & interview_date  )
       	Call write_variable_in_CASE_NOTE("* Confirmed client was provided sufficient 10 day notice.")
    ELSEIF case_status_dropdown = "Interview not needed for MFIP to SNAP transition" THEN
    	Call write_variable_in_CASE_NOTE("~ " & case_status_dropdown  & " ~")
    	Call write_variable_in_CASE_NOTE("* MFIP to SNAP transition no interview required updated PROG to reflect this")
    ELSEIF case_status_dropdown = "Other(please describe)" THEN
    	Call write_variable_in_CASE_NOTE("~ " & issue_notes  & " ~")
    END IF
    CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
    PF3
END IF
script_end_procedure_with_error_report ("Case note has been entered please review.")
