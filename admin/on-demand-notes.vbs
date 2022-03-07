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
CALL changelog_update("02/10/2022", "Removed confirmation when hitting cancel. Added handing for subsequent applications. ", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/01/2021", "GitHub #189 Updated script to remove correction email process.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		IF HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function
'---------------------------------------------------------------------------------------The script

EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
CALL Check_for_MAXIS(false)                         'Ensuring we are not passworded out

closing_message = "On Demand Application Waiver review has been case noted." 'setting up closing_message variable for possible additions later based on conditions

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 286, 145, "Notes On Demand"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 175, 5, 50, 15, application_date
  ButtonGroup ButtonPressed
    PushButton 230, 5, 55, 15, "STAT/PROG", PROG_button
    PushButton 230, 25, 55, 15, "CASE/NOTE", NOTE_button
  CheckBox 55, 30, 105, 10, "Case was not pended timely", pended_checkbox
  CheckBox 55, 40, 140, 10, "Client completed application interview", completed_interview_checkbox
  CheckBox 55, 50, 175, 10, "Client has not completed application interview", incomplete_interview_checkbox
  CheckBox 55, 60, 140, 10, "Denied programs for no interview", denial_checkbox
  CheckBox 55, 70, 120, 10, "Subsequent application received", subsequent_app_checkbox
  CheckBox 55, 80, 170, 10, "Interview not needed for MFIP to SNAP transition", MTAF_checkbox
  CheckBox 55, 90, 90, 10, "Other(please describe)", other_notes_checkbox
  EditBox 55, 105, 170, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 180, 125, 50, 15
    CancelButton 235, 125, 50, 15
  Text 5, 30, 45, 10, "Case status:"
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 110, 45, 10, "Other notes:"
  Text 110, 10, 65, 10, "Date of application:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		CALL MAXIS_dialog_navigation()
		IF IsNumeric(maxis_case_number) = FALSE or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF IsDate(application_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter a valid application date."
		IF other_notes_checkbox = CHECKED THEN
			case_status = "Other"
			IF case_status = "Other(please describe)" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a description of what occurred."
		END IF
		IF ButtonPressed = NOTE_button or ButtonPressed = PROG_button THEN 'need the error message to not be blank so that it wont message box but it will not leave '
			err_msg = "Loop"
		ELSE
			IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine 'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		END IF
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month
CALL convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)

'Checking for PRIV cases.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
IF is_this_priv = True THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
MAXIS_background_check      'Making sure we are out of background.

'Grabbing case and program status information from MAXIS.
'For tis script to work correctly, these must be correct BEFORE running the script.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
EMReadScreen program_status, 15, 8, 9                  'Now we are reading the CASE STATUS string from the panel - we want to make sure this does NOT read CAF1 PENDING
EMReadScreen pnd2_appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
ive_status = "INACTIVE"                             'There are some programs that are NOT read from the function and are pretty specific to this script/functionality
cca_status = "INACTIVE"                             'defaulting these statuses to 'INACTIVE' until they are read from the panel
ega_status = "INACTIVE"
ea_status = "INACTIVE"
'\This functionality is how the above function reads for program information - just pulled out for these specific programs
row = 1                                             'Looking for IV-E information
col = 1
EMSearch "IV-E:", row, col
If row <> 0 Then
    EMReadScreen ive_status, 9, row, col + 6
    ive_status = trim(ive_status)
    If ive_status = "ACTIVE" or ive_status = "APP CLOSE" or ive_status = "APP OPEN" Then ive_status = "ACTIVE"
    If ive_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for CCAP information
col = 1
EMSearch "CCAP", row, col
If row <> 0 Then
    EMReadScreen cca_status, 9, row, col + 6
    cca_status = trim(cca_status)
    If cca_status = "ACTIVE" or cca_status = "APP CLOSE" or cca_status = "APP OPEN" Then cca_status = "ACTIVE"
    If cca_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for EGA information
col = 1
EMSearch "EGA", row, col
If row <> 0 Then
    EMReadScreen ega_status, 9, row, col + 6
    ega_status = trim(ega_status)
    If ega_status = "ACTIVE" or ega_status = "APP CLOSE" or ega_status = "APP OPEN" Then ega_status = "ACTIVE"
    If ega_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If
row = 1                                             'Looking for EA information
col = 1
EMSearch "EA: ", row, col
If row <> 0 Then
    EMReadScreen ea_status, 9, row, col + 5
    ea_status = trim(ea_status)
    If ea_status = "ACTIVE" or ea_status = "APP CLOSE" or ea_status = "APP OPEN" Then ea_status = "ACTIVE"
    If ea_status = "PENDING" Then case_pending = True      'Updating the case_pending variable from the function
End If

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

pending_programs = ""        'Creates a variable that lists all the pending programs on the case.
If unknown_cash_pending = True Then pending_programs = pending_programs & "Cash, "
If ga_status = "PENDING" Then pending_programs = pending_programs & "GA, "
If msa_status = "PENDING" Then pending_programs = pending_programs & "MSA, "
If mfip_status = "PENDING" Then pending_programs = pending_programs & "MFIP, "
If dwp_status = "PENDING" Then pending_programs = pending_programs & "DWP, "
If ive_status = "PENDING" Then pending_programs = pending_programs & "IV-E, "
If grh_status = "PENDING" Then pending_programs = pending_programs & "GRH, "
If snap_status = "PENDING" Then pending_programs = pending_programs & "SNAP, "
If ega_status = "PENDING" Then pending_programs = pending_programs & "EGA, "
If ea_status = "PENDING" Then pending_programs = pending_programs & "EA, "
If cca_status = "PENDING" Then pending_programs = pending_programs & "CCA, "
If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then pending_programs = pending_programs & "HC, "

pending_programs = trim(pending_programs)  'trims excess spaces of pending_programs
If right(pending_programs, 1) = "," THEN pending_programs = left(pending_programs, len(pending_programs) - 1)

IF pended_checkbox = CHECKED THEN
	case_status = "Case was not pended timely"
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 181, 85, "Case was not pended timely"
    EditBox 120, 5, 50, 15, appointment_letter_date
    EditBox 120, 25, 50, 15, NOMI_date
    EditBox 120, 45, 50, 15, denial_date
	ButtonGroup ButtonPressed
	  OkButton 65, 65, 50, 15
	  CancelButton 120, 65, 50, 15
    Text 5, 10, 80, 10, "Appointment letter date:"
    Text 5, 30, 115, 10, "Notice of missed interview date:"
    Text 5, 50, 45, 10, "Denial date:"
    EndDialog
    DO
    	DO
    		err_msg = ""
    		Dialog Dialog1
    		IF IsDate(appointment_letter_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter the date the appointment letter was sent."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF completed_interview_checkbox = CHECKED THEN
	case_status = "Client completed application interview"
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 141, 85, "Client completed application interview"
	  EditBox 65, 5, 50, 15, case_note_date
	  EditBox 65, 25, 50, 15, interview_date
	  DropListBox 65, 45, 70, 15, "NO"+chr(9)+"YES", confirm_update_prog
	  ButtonGroup ButtonPressed
	    OkButton 30, 65, 50, 15
	    CancelButton 85, 65, 50, 15
	  Text 5, 10, 55, 10, "Case note date:"
	  Text 5, 50, 50, 10, "Update PROG:"
	  Text 5, 30, 55, 10, "Interview date:"
	EndDialog
    DO
        DO
        	err_msg = ""
        	Dialog Dialog1
    		IF IsDate(case_note_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case note date."
    		IF IsDate(interview_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter a valid interview date."
        	IF confirm_update_prog = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please advise if STAT/PROG needs interview date entered."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF incomplete_interview_checkbox = CHECKED THEN
	case_status = "Client has not completed application interview"
	Dialog1 = "" 'Blanking out previous dialog detail
 	BeginDialog Dialog1, 0, 0, 181, 85, "Client has not completed application interview"
   	EditBox 120, 5, 50, 15, appointment_letter_date
   	EditBox 120, 25, 50, 15, NOMI_date
   	EditBox 120, 45, 50, 15, denial_date
   	ButtonGroup ButtonPressed
     	OkButton 65, 65, 50, 15
     	CancelButton 120, 65, 50, 15
   	Text 5, 10, 80, 10, "Appointment letter date:"
   	Text 5, 30, 115, 10, "Notice of missed interview date:"
   	Text 5, 50, 45, 10, "Denial date:"
 	EndDialog
    DO
    	DO
        	err_msg = ""
        	Dialog Dialog1
        	IF IsDate(appointment_letter_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter the date the appointment letter was sent."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF denial_checkbox = CHECKED THEN
    case_status = "Denied programs for no interview"
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 181, 65, "Denied programs for no interview"
      EditBox 120, 5, 50, 15, appointment_letter_date
      EditBox 120, 25, 50, 15, NOMI_date
      ButtonGroup ButtonPressed
        OkButton 65, 45, 50, 15
        CancelButton 120, 45, 50, 15
      Text 5, 30, 115, 10, "Notice of missed interview date:"
      Text 5, 10, 80, 10, "Appointment letter date:"
    EndDialog
    DO
        DO
        	err_msg = ""
        	Dialog Dialog1
            IF case_status = "Denied programs for no interview" THEN
            	IF IsDate(application_date) = TRUE THEN
            		IF datediff("d", application_date, date) < 30 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid application date, the resident must be provided 30 days from the date of application."    'confirming that these cases meet all the criteria for denial
            	END IF
        		IF IsDate(appointment_letter_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter the date the appointment letter was sent."
            	IF IsDate(NOMI_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter a valid NOMI date."
            END IF
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF subsequent_app_checkbox = CHECKED THEN
	case_status = "Subsequent application received"
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 191, 225, "Application Received"
	  Text 120, 20, 60, 10, application_date
	  EditBox 120, 35, 60, 15, subsequent_app_date
	  DropListBox 85, 55, 95, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Mystery Doc Queue"+chr(9)+"Online"+chr(9)+"Phone-Verbal Request"+chr(9)+"Request to APPL Form"+chr(9)+"Virtual Drop Box", how_application_rcvd
	  DropListBox 85, 70, 95, 15, "Select One:"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Populations"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer"+chr(9)+"MN Benefits"+chr(9)+"N/A"+chr(9)+"Verbal Request", application_type
	  EditBox 85, 85, 95, 15, confirmation_number
	  GroupBox 5, 5, 180, 115, "Application Information"
	  Text 10, 20, 65, 10, "Date of Application:"
	  Text 10, 40, 105, 10, "Subsequent application date:"
	  Text 10, 60, 70, 10, "Application Received:"
	  Text 10, 75, 65, 10, "Type of Application:"
	  Text 10, 90, 50, 10, "Confirmation #:"
	  Text 10, 105, 65, 10, "Pending Programs: "
	    Text 85, 105, 60, 10, pending_programs
	  ButtonGroup ButtonPressed
	    PushButton 50, 175, 95, 15, "Open CM 05.09.06", cm_05_09_06_btn
	    OkButton 75, 125, 50, 15
	    CancelButton 130, 125, 50, 15
	  Text 10, 145, 170, 25, "Per CM 0005.09.06 - if a case is pending and a new app is received you should use the original application date."
	  Text 10, 195, 170, 30, "Please contact Knowledge Now or your Supervisor if you have questions about dates to enter in MAXIS for applications."
	EndDialog

	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
			If ButtonPressed = cm_05_09_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00050906"
		Loop until ButtonPressed = -1
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF MTAF_checkbox = CHECKED THEN
	case_status = "Interview not needed for MFIP to SNAP transition"
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 151, 125, " Minnesota Transition Application Form"
	  ButtonGroup ButtonPressed
	    PushButton 30, 45, 95, 15, "Open CM 0005.10 ", cm_05_10_btn
	    OkButton 40, 105, 50, 15
	    CancelButton 95, 105, 50, 15
	  Text 5, 5, 140, 35, "Per CM 0005.10  - Review the MTAF for completeness. A complete MTAF is signed and dated with all questions answered. No interview is needed."
	  Text 5, 70, 145, 30, "Please contact Knowledge Now or your Supervisor if you have questions about dates to enter in MAXIS for applications."
	EndDialog
	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
			If ButtonPressed = cm_05_10_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000510"
		Loop until ButtonPressed = -1
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF other_notes_checkbox = CHECKED THEN case_status = "Other"

'Checking for PRIV cases.
IF case_status = "Client completed application interview" THEN          'Interviews are only required for Cash and SNAP
	IF confirm_update_prog = "YES" THEN
	    Call navigate_to_MAXIS_screen("STAT", "PROG")
	    PF9                                         'Edit
        intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
        intv_day = DatePart("d", interview_date)
        intv_yr = DatePart("yyyy", interview_date)
        intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
        intv_day = right("00"&intv_day, 2)
        intv_yr = right(intv_yr, 2)
        intv_date_to_check = intv_mo & " " & intv_day & " " & intv_yr
        EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
        EMReadScreen cash_two_app, 8, 7, 33
        EMReadScreen emergency_app, 8, 8, 33
		EMReadScreen emer_prog_check, 2, 8, 67
		EMReadScreen grh_cash_app, 8, 9, 33
		EMReadScreen snap_app, 8, 10, 33
        cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
        cash_two_app = replace(cash_two_app, " ", "/")
		emergency_app = replace(emergency_app, " ", "/")
        grh_cash_app = replace(grh_cash_app, " ", "/")
		snap_app = replace(snap_app, " ", "/")
        If cash_one_app <> "__/__/__" THEN              'Comparing them to the date of application to determine which row to use
            If IsDate(cash_one_app) = TRUE THEN
                if DateDiff("d", cash_one_app, application_date) = 0 then prog_row = 6
            End If
    		EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
    		EMWriteScreen intv_day, prog_row, 58
    		EMWriteScreen intv_yr, prog_row, 61
        End If
        If cash_two_app <> "__/__/__" THEN
            If IsDate(cash_two_app) = TRUE THEN
                if DateDiff("d", cash_two_app, application_date) = 0 then prog_row = 7
            End If
    		EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
    		EMWriteScreen intv_day, prog_row, 58
    		EMWriteScreen intv_yr, prog_row, 61
        End If
		If emergency_app <> "__/__/__" and emer_prog_check <> "EA" THEN              'Comparing them to the date of application to determine which row to use
			If IsDate(emergency_app) = TRUE THEN
				if DateDiff("d", emergency_app, application_date) = 0 then prog_row = 8
			End If
			EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
			EMWriteScreen intv_day, prog_row, 58
			EMWriteScreen intv_yr, prog_row, 61
		End If
		If grh_cash_app <> "__/__/__" THEN
			If IsDate(grh_cash_app) = TRUE THEN
				if DateDiff("d", grh_cash_app, application_date) = 0 then prog_row = 9
			End If
			EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
			EMWriteScreen intv_day, prog_row, 58
			EMWriteScreen intv_yr, prog_row, 61
		End If
		If snap_app <> "__/__/__" THEN              'Comparing them to the date of application to determine which row to use
			If IsDate(snap_app) = TRUE THEN
				if DateDiff("d", snap_app, application_date) = 0 then prog_row = 10
			End If
            EMWriteScreen intv_mo, prog_row, 55               'SNAP is easy because there is only one area for interview - the variables go there
            EMWriteScreen intv_day, prog_row, 58
            EMWriteScreen intv_yr, prog_row, 61
		End If

        TRANSMIT                                  'Saving the panel
		EMReadScreen error_check,  7, 24, 02
		EMReadScreen MISC_error_check,  74, 24, 02
		'WARNING: EMER INTERVIEW DATE IS MISSING OR INVALID
		IF error_check <> "WARNING" THEN
	    	IF trim(MISC_error_check) = "" THEN
	    		case_note_only = FALSE
	    	else
	    		maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion, "Message handling")
	    		IF maxis_error_check = vbYes THEN
	    			case_note_only = TRUE 'this will case note only'
	    		END IF
	    		IF maxis_error_check= vbNo THEN
	    			case_note_only = FALSE 'this will update the panels and case note'
	    		END IF
	    	END IF
            Call HCRE_panel_bypass()
            Call MAXIS_background_check
	    	Call navigate_to_MAXIS_screen("STAT", "PROG")
            If intv_date_needed = TRUE THEN         'If previous code has determined that PROG needs to be updated
                snap_intv_date_updated = FALSE
                cash_intv_date_updated = FALSE
                show_prog_update_failure = FALSE
                Call back_to_SELF
                CALL navigate_to_MAXIS_screen("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
            	IF confirm_update_prog = "Snap" or confirm_update_prog = "Cash & Snap" THEN
                    EMReadScreen new_snap_intv_date, 8, 10, 55
                    If new_snap_intv_date = intv_date_to_check Then snap_intv_date_updated = TRUE
                    If snap_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
            	END IF
                If confirm_update_prog = "Cash" or confirm_update_prog = "Cash & Snap" THEN
                    EMReadScreen new_cash_intv_date, 8, prog_row, 55
                    If new_cash_intv_date = intv_date_to_check Then cash_intv_date_updated = TRUE
                    If cash_intv_date_updated = FALSE Then show_prog_update_failure = TRUE
                End If
                If show_prog_update_failure = TRUE THEN
                    fail_msg = "You have requested the script update PROG for "
                    If confirm_update_prog = "Cash & Snap" THEN
                        fail_msg = fail_msg & "Cash and SNAP "
                    ElseIf confirm_update_prog = "Snap"  THEN
                        fail_msg = fail_msg & "SNAP "
                    ElseIf confirm_update_prog = "Cash"  THEN
                        fail_msg = fail_msg & "Cash "
                    End If
                    fail_msg = fail_msg & "to enter the interview date on PROG." & vbCr & vbCr & "The script was unable to update PROG completely." & vbCr
                    If confirm_update_prog = "Snap" THEN
                        fail_msg = fail_msg & " - The SNAP Interview Date was not entered." & vbCr
                    ElseIf confirm_update_prog = "Cash" THEN
                        fail_msg = fail_msg & " - The Cash Interview Date was not entered." & vbCr
                    End If
                    fail_msg = fail_msg & closing_message & "The PROG panel will need to be updated manually with the interview information."
                END IF
	        END IF
        END IF
	END IF
END IF
'denial_date = dateadd("d", 0, denial_date) ' if needed this will help this the script recognize that the date is a date'
'this to remind workers that we must give clients 10 days when we are outside of that 30 day window for applications'
IF NOMI_date <> "" THEN denial_date = dateadd("d", 10, NOMI_date)
IF denial_date < date then denial_date = dateadd("d", 10, date)

'NOW WE START CASE NOTING - there are a few
start_a_blank_case_note
IF case_status = "Case was not pended timely" THEN
    CALL write_variable_in_CASE_NOTE("~ Client has not completed application interview ~")
    CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
    CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & NOMI_date)
    CALL write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it, because the case was not pended timely a NOMI still needs to be sent and adequate time provided to the client to comply. Denial can be done after " & denial_date)
ELSEIF case_status = "Client completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status & " on "  & interview_date & " PROG updated ~")
	CALL write_variable_in_CASE_NOTE("* Completed by previous worker per case note dated: " & case_note_date)
ELSEIF case_status = "Client has not completed application interview" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status  & " ~")
	CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
	CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & NOMI_date)
	CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview.")
	CALL write_variable_in_CASE_NOTE("* Interview is still needed, client has 30 days from date of application to complete it.")
ELSEIF case_status = "Denied programs for no interview" THEN
	CALL write_variable_in_CASE_NOTE("~ Denied programs for no interview ~")
    CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
    CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & NOMI_date)
	CALL write_variable_in_CASE_NOTE("* Reason for denial: interview was not completed timely.")
   	CALL write_variable_in_CASE_NOTE("* Confirmed client was provided sufficient 10 day notice.")
ELSEIF case_status = "Interview not needed for MFIP to SNAP transition" THEN
	CALL write_variable_in_CASE_NOTE("~ " & case_status & " ~")
	CALL write_variable_in_CASE_NOTE("* MFIP to SNAP transition no interview required updated PROG to reflect this Per CM 0005.10.")
ELSEIF case_status = "Other" THEN
	CALL write_variable_in_CASE_NOTE("~ Application review (on demand) ~")
	CALL write_variable_in_CASE_NOTE("* Application date: " & application_date)
	CALL write_variable_in_CASE_NOTE("* NOMI sent to client on: " & NOMI_date)
ELSEIF case_status = "Subsequent application received" THEN
	CALL write_variable_in_CASE_NOTE ("~ Subsequent Application Received (" &  application_type & ") via " & how_application_rcvd & " for " & subsequent_app_date & " ~")
    CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
	CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", pending_programs)
	CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Pending Programs:", pending_programs)
    CALL write_variable_in_CASE_NOTE("* Aligned dates on STAT/PROG to match current pending program(s) per CM 0005.09.12 - APPLICATION - PENDING CASES")
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
'--Comment Code-----------------------------------------------------------------11/01/2021
'--Update Changelog for release/update------------------------------------------11/01/2021
'--Remove testing message boxes-------------------------------------------------11/01/2021
'--Remove testing code/unnecessary code-----------------------------------------11/01/2021
'--Review/update SharePoint instructions----------------------------------------10/01/2021
'--Review Best Practices using BZS page ----------------------------------------10/01/2021
'--Review script information on SharePoint BZ Script List-----------------------11/01/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/01/2021
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/01/2021
'--Complete misc. documentation (if applicable)---------------------------------10/01/2021
'--Update project team/issue contact (if applicable)----------------------------10/01/2021
