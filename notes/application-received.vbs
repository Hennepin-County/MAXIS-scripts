'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
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
call changelog_update("01/29/2021", "Updated Request for APPL handling per case assignement request. Issue #322", "MiKayla Handley, Hennepin County")
call changelog_update("01/07/2021", "Updated worker signature as a mandatory field.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Temporarily removed in person option for how applications are received.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Update to make confirmation number mandatory for MN Benefit application.", "MiKayla Handley, Hennepin County")
call changelog_update("11/15/2020", "Updated droplist to add virtual drop box option to how the application was received.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/13/2020", "Enhanced date evaluation functionality when which determining HEST standards to use.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/24/2020", "Added MN Benefits application and removed SHIBA and apply MN options.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/01/2020", "Updated Standard Utility Allowances for 10/2020.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/24/2020", "Added SHIBA application and combined CA and NOTES scripts.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/10/2020", "Email functionality removed for Triagers.", "MiKayla Handley, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2020", "Update to the notice wording. Information and direction for in-person interview option removed. County offices are not currently open due to the COVID-19 Peacetime Emergency.", "Casey Love, Hennepin County")
call changelog_update("03/09/2020", "Per project request- Updated checkbox for the METS Retro Request to team 601.", "MiKayla Handley, Hennepin County")
call changelog_update("01/13/2020", "Updated requesting worker for the Request to APPL form process.", "MiKayla Handley, Hennepin County")
call changelog_update("11/04/2019", "New version pulled to support the request for APPL process.", "MiKayla Handley, Hennepin County")
call changelog_update("10/01/2019", "Updated the utility standards for SNAP.", "Casey Love, Hennepin County")
call changelog_update("08/27/2019", "Added handling to push the case into background to ensure pending programs are read.", "MiKayla Handley, Hennepin County")
call changelog_update("08/27/2019", "Added GRH to appointment letter handling for future enhancements.", "MiKayla Handley, Hennepin County")
call changelog_update("08/20/2019", "Bug on the script when a large PND2 list is accessed.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/26/2019", "Reverted the script to not email Team 603 for METS cases. CA workers will need to manually complete the email to: field.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/24/2019", "Removed Mail & Fax option and added MDQ per request.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/22/2019", "Updated the script to automatically email Team 603 for METS cases.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
CALL changelog_update("02/05/2019", "Updated case correction handling.", "Casey Love, Hennepin County")
CALL changelog_update("11/15/2018", "Enhanced functionality for SameDay interview cases.", "Casey Love, Hennepin County")
CALL changelog_update("11/06/2018", "Updated handling for HC only applications.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/25/2018", "Updated script to add handling for case correction.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/17/2018", "Updated appointment letter to address EGA programs.", "MiKayla Handley, Hennepin County")
CALL changelog_update("09/01/2018", "Updated Utility standards that go into effect for 10/01/2018.", "Ilse Ferris, Hennepin County")
CALL changelog_update("07/20/2018", "Changed wording of the Appointment Notice and changed default interview date to 10 days from application for non-expedidted cases.", "Casey Love, Hennepin County")
CALL changelog_update("07/16/2018", "Bug Fix that was preventing notices from being sent.", "Casey Love, Hennepin County")
CALL changelog_update("03/28/2018", "Updated appt letter case note for bulk script process.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/21/2018", "Added on demand waiver handling.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/16/2018", "Added case transfer confirmation coding.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/03/2017", "Email functionality - only expedited emails will be sent to Triagers.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/25/2017", "Email functionality - will create email, and send for all CASH and FS applications.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/12/2017", "Email functionality will create email, but not send it. Staff will need to send email after reviewing email.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
back_to_self' added to ensure we have the time to update and send the case in the background

'Checking for PRIV cases.
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
IF priv_check = "PRIV" THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 116, 45, "Application Received"
  EditBox 65, 5, 45, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 5, 25, 50, 15
    CancelButton 60, 25, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'---------------------------------------------------------------------------------------------'pending & active programs information
'information gathering to auto-populate the application date

EMWriteScreen MAXIS_case_number, 18, 43
limit_reached = FALSE
Call navigate_to_MAXIS_screen("REPT", "PND2")
EmReadScreen limit_reached, 8, 6, 21
IF limit_reached = "The REPT" THEN
    limit_reached_confirmation = MsgBox("Press YES to confirm this case has been pended."  & vbNewLine & " ", vbYesNoCancel, "The REPT:PND2 Display Limit Has Been Reached.")
    IF limit_reached_confirmation = vbNo THEN
    	limit_reached = TRUE
        script_end_procedure_with_error_report("Limit reached in basket. Please pend the case and run the script again.")
    END IF
    IF limit_reached_confirmation = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The application has not been acted on.")
    IF limit_reached_confirmation = vbYes THEN
       	limit_reached = FALSE
	    transmit
	    PF3
	    EMWriteScreen MAXIS_case_number, 18, 43
    END IF
ELSEIF limit_reached <> "The REPT" THEN
    'Ensuring that the user is in REPT/PND2
    Do
    	EMReadScreen PND2_check, 4, 2, 52
    	If PND2_check <> "PND2" then
    		back_to_SELF
    		Call navigate_to_MAXIS_screen("REPT", "PND2")
    	End if
    LOOP until PND2_check = "PND2"

    'checking the case to make sure there is a pending case.  If not script will end & inform the user no pending case exists in PND2
    EMReadScreen not_pending_check, 4, 24, 2
    If not_pending_check = "CASE" THEN script_end_procedure_with_error_report("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")
    'grabs row and col number that the cursor is at
    EMGetCursor MAXIS_row, MAXIS_col
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
END IF 'this will ensire if the basket limit is reached a worker can still use the script'

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
IF application_date = "" THEN 'if the rept/pnd2 is full this will allow the use to get to PROG '
	Row = 6
	DO
		EMReadScreen application_date, 8, row, 33
		IF application_date = "" THEN script_end_procedure_with_error_report("A application for the selected period could not be found. The script will now end.")
		application_date_confirmation = MsgBox("Press YES to confirm this is the application you wish to act on." & vbNewLine & "For the next application, press NO." & vbNewLine & vbNewLine & _
		" " & "Application date: " & application_date, vbYesNoCancel, "Please confirm this application")
		IF application_date_confirmation = vbNo THEN
			row = row + 1
			IF row = 17 THEN
				PF8
				row = 7
				EMReadScreen application_date, 11, row, 47
			END IF
		END IF
		IF application_date_confirmation = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The application has not been acted on.")
		IF application_date_confirmation = vbYes THEN 	EXIT DO
	LOOP UNTIL application_date_confirmation = vbYes
	application_date = replace(application_date, " ", "/")
	EMReadScreen err_msg, 7, 24, 02
    IF err_msg = "BENEFIT" THEN	script_end_procedure_with_error_report ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")
END IF
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

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 306, 125,  "Application Received for: "  & programs_applied_for &   " on "   & application_date
  DropListBox 85, 10, 95, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mystery Doc Queue"+chr(9)+"Online"+chr(9)+"Phone-Verbal Request"+chr(9)+"Request to APPL Form"+chr(9)+"Virtual Drop Box", how_application_rcvd
  DropListBox 85, 30, 95, 15, "Select One:"+chr(9)+"ApplyMN"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Populations"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer"+chr(9)+"MN Benefits"+chr(9)+"N/A"+chr(9)+"Verbal Request", application_type
  EditBox 250, 30, 45, 15, confirmation_number
  EditBox 50, 60, 20, 15, transfer_to_worker
  CheckBox 140, 65, 155, 10, "Check if the case does not require a transfer ", no_transfer_checkbox
  EditBox 55, 85, 245, 15, other_notes
  EditBox 70, 105, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 105, 50, 15
    CancelButton 250, 105, 50, 15
  Text 185, 35, 50, 10, "Confirmation #:"
  Text 5, 110, 60, 10, "Worker Signature:"
  GroupBox 5, 50, 295, 30, "Transfer Information"
  Text 5, 90, 45, 10, "Other Notes:"
  Text 10, 65, 40, 10, "Transfer to:"
  Text 75, 65, 60, 10, "(last 3 digit of X#)"
  Text 10, 15, 70, 10, "Application Received:"
  Text 10, 35, 65, 10, "Type of Application:"
  GroupBox 5, 0, 295, 50, "Application Information"
EndDialog
'------------------------------------------------------------------------------------DIALOG APPL
Do
	Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
	    IF how_application_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
        'IF application_type = "N/A" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter in other notes what type of application was received to the agency."
	    IF application_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
        IF application_type = "ApplyMN" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If an ApplyMN app was received, you must enter the confirmation number and time received."
        IF application_type = "MN Benefits" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If a MN Benefits app was received, you must enter the confirmation number and time received."
	    IF no_transfer_checkbox = UNCHECKED AND transfer_to_worker = "" then err_msg = err_msg & vbNewLine & "* You must enter the basket number the case to be transferred by the script or check that no transfer is needed."
	    IF no_transfer_checkbox = CHECKED and transfer_to_worker <> "" then err_msg = err_msg & vbNewLine & "* You have checked that no transfer is needed, please remove basket number from transfer field."
	    IF no_transfer_checkbox = UNCHECKED AND len(transfer_to_worker) > 3 AND isnumeric(transfer_to_worker) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the last 3 digits of the worker number for transfer."
	    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

IF how_application_rcvd = "Request to APPL Form" THEN
    DO
        CALL HH_member_custom_dialog(HH_member_array)
        IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
    LOOP UNTIL uBound(HH_member_array) <> -1

    'CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    'FOR EACH person IN HH_member_array
    '    IF person <> "" THEN
    '        CALL write_value_and_transmit(person, 20, 76)
    '        EMReadScreen MEMB_number, 2, 4, 33

    '    END IF
    '    MEMB_number = MEMB_number & ", "
    'NEXT
    household_persons = ""
    pers_count = 0
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		IF pers_count = uBound(HH_member_array) THEN
    			IF pers_count = 0 THEN
    				household_persons = household_persons & person & ", "
    			ELSE
    				household_persons = household_persons & " " & person
    			END IF
    		ELSE
    			household_persons = household_persons & person & ", "
    			pers_count = pers_count + 1
    		END IF
    	END IF
    NEXT

    '-------------------------------------------------------------------------------------------------DIALOG
    BeginDialog Dialog1, 0, 0, 186, 135, "Request to Appl"
      EditBox 85, 15, 45, 15, request_date
      EditBox 85, 35, 45, 15, request_worker_number
      EditBox 85, 55, 45, 15, METS_case_number
      CheckBox 15, 80, 55, 10, "MA Transition", MA_transition_request_checkbox
      CheckBox 15, 95, 60, 10, "Auto Newborn", Auto_Newborn_checkbox
      CheckBox 85, 80, 85, 10, "METS Retro Coverage", METS_retro_checkbox
      CheckBox 85, 95, 85, 10, "Team 601 will process", team_601_email_checkbox
      GroupBox 5, 5, 175, 105, "Request to Appl Information"
      Text 15, 20, 60, 10, "Submission Date:"
      Text 15, 40, 60, 10, "Requested By X#:"
      Text 15, 60, 55, 10, "METS Case #:"
      ButtonGroup ButtonPressed
        OkButton 75, 115, 50, 15
        CancelButton 130, 115, 50, 15
    EndDialog


    '------------------------------------------------------------------------------------DIALOG APPL
    Do
    	Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
    	    IF request_date = "" THEN err_msg = err_msg & vbNewLine & "* If a request to APPL was received, you must enter the date the form was submitted."
    	    IF METS_retro_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg & vbNewLine & "* You have checked that this is a METS Retro Request, please enter a METS IC #."
    	    IF MA_transition_request_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg &  vbNewLine & "* You have checked that this is a METS Transition Request, please enter a METS IC #."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has     not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = FALSE
END IF

transfer_to_worker = trim(transfer_to_worker)
transfer_to_worker = Ucase(transfer_to_worker)
request_worker_number = trim(request_worker_number)
request_worker_number = Ucase(request_worker_number)
'msgbox transfer_to_worker
pended_date = date

HC_applied_for = FALSE
IF application_type = "6696" or application_type = "HCAPP" or application_type = "HC-Certain Pop" or application_type = "LTC" or application_type = "MHCP B/C Cancer" THEN HC_applied_for = TRUE

If how_application_rcvd = "Phone-Verbal Request" THEN how_application_rcvd = replace(how_application_rcvd, "Phone-Verbal Request", "Phone")
IF how_application_rcvd = "Request to APPL Form" THEN
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED AND MA_transition_request_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED AND METS_retro_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND METS_retro_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "METS Retro")
    IF application_type = "N/A" AND MA_transition_request_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "MA Transition")
END IF
'--------------------------------------------------------------------------------initial case note
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("~ Application Received (" &  application_type & ")via " & how_application_rcvd & " for " & application_date & " ~")
CALL write_bullet_and_variable_in_CASE_NOTE("Requesting HC for MEMBER(S) ", household_persons)
CALL write_bullet_and_variable_in_CASE_NOTE("Request to APPL Form received on ", request_date)
IF how_application_rcvd = "Request to APPL Form" THEN
	IF team_601_email_checkbox = UNCHECKED  THEN CALL write_variable_in_CASE_NOTE("* Emailed " & request_worker_number & " to let them know the request was processed.")
    IF team_601_email_checkbox = CHECKED  THEN CALL write_variable_in_CASE_NOTE("* Emailed team 601 to let them know the retro request was processed.")
END IF
CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
IF application_type = "6696" THEN write_variable_in_CASE_NOTE ("* Form Received: METS Application for Health Coverage and Help Paying Costs (DHS-6696) ")
IF application_type = "HCAPP" THEN write_variable_in_CASE_NOTE ("* Form Received: Health Care Application (HCAPP) (DHS-3417) ")
IF application_type = "HC-Certain Pop" THEN write_variable_in_CASE_NOTE ("* Form Received: MHC Programs Application for Certain Populations (DHS-3876) ")
IF application_type = "LTC" THEN write_variable_in_CASE_NOTE ("* Form Received: Application for Medical Assistance for Long Term Care Services (DHS-3531) ")
IF application_type = "MHCP B/C Cancer" THEN write_variable_in_CASE_NOTE ("* Form Received: Minnesota Health Care Programs Application and Renewal Form Medical Assistance for Women with Breast or Cervical Cancer (DHS-3525) ")
IF application_type = "Verbal Request" THEN write_variable_in_CASE_NOTE ("* Verbal Request was made for programs. CAF will be completed with resident over the phone.")
CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Pending Programs", additional_programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
If transfer_to_worker <> "" THEN CALL write_variable_in_CASE_NOTE ("* Application assigned to X127" & transfer_to_worker)
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)

CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)
PF3 ' to save Case note

'----------------------------------------------------------------------------------------------------EXPEDITED SCREENING!
IF snap_pends = TRUE THEN
    'DATE BASED LOGIC FOR UTILITY AMOUNTS: variables need to change every October per CM.18.15.09------------------------------------------------------------------------------------------
    If DateDiff("d",application_date,"10/01/2020") <= 0 then
        'October 2020 amounts
        heat_AC_amt = 496
        electric_amt = 154
        phone_amt = 56
    Else
        'October 2019 amounts
        heat_AC_amt = 490
        electric_amt = 143
        phone_amt = 49
    End if

    'Ensuring case number carries thru
    CALL MAXIS_case_number_finder(MAXIS_case_number)
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 181, 165, "Expedited Screening"
     	EditBox 100, 5, 50, 15, MAXIS_case_number
     	EditBox 100, 25, 50, 15, income
     	EditBox 100, 45, 50, 15, assets
     	EditBox 100, 65, 50, 15, rent
     	CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
     	CheckBox 75, 95, 45, 10, "Electricity", electric_check
     	CheckBox 130, 95, 35, 10, "Phone", phone_check
     	ButtonGroup ButtonPressed
    	OkButton 70, 115, 50, 15
    	CancelButton 125, 115, 50, 15
     	Text 10, 140, 160, 15, "The income, assets and shelter costs fields will default to $0 if left blank. "
     	Text 5, 30, 95, 10, "Income received this month:"
     	Text 5, 50, 95, 10, "Cash, checking, or savings: "
     	Text 5, 70, 90, 10, "AMT paid for rent/mortgage:"
     	GroupBox 5, 85, 170, 25, "Utilities claimed (check below):"
     	Text 50, 10, 50, 10, "Case number: "
     	GroupBox 0, 130, 175, 30, "**IMPORTANT**"
    EndDialog

    '----------------------------------------------------------------------------------------------------THE SCRIPT
    CALL MAXIS_case_number_finder(MAXIS_case_number)
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_confirmation
    		If isnumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbnewline & "* You must enter a valid case number."
    		If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) THEN err_msg = err_msg & vbnewline & "* The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
    		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
    '----------------------------------------------------------------------------------------------------LOGIC AND CALCULATIONS
    'Logic for figuring out utils. The highest priority for the if...THEN is heat/AC, followed by electric and phone, followed by phone and electric separately.
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

    'Calculates expedited status based on above numbers
    IF (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) THEN expedited_status = "Client Appears Expedited"
    IF (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) THEN expedited_status = "Client Does Not Appear Expedited"
    '----------------------------------------------------------------------------------------------------checking DISQ
    CALL navigate_to_MAXIS_screen("STAT", "DISQ")
    'grabbing footer month and year
    CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
    'Reads the DISQ info for the case note.
    EMReadScreen DISQ_member_check, 34, 24, 2
    IF DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" THEN
      	has_DISQ = False
    ELSE
      	has_DISQ = True
    END IF

	'expedited_status = "Client Appears Expedited"
	'expedited_status = "Client Does Not Appear Expedited"
    'Reads MONY/DISB to see if EBT account is open
    'IF expedited_status = "Client Appears Expedited" THEN
  	'	CALL navigate_to_MAXIS_screen("MONY", "DISB")
  	'	EMReadScreen EBT_account_status, 1, 14, 27
	'    same_day_offered = FALSE
    '    If  = TRUE Then same_day_offered = TRUE
    '    If  = FALSE Then
    '        offer_same_date_interview = MsgBox("This client appears EXPEDITED. A same-day needs to be offered." & vbNewLine & vbNewLine & "Has the 'client been offered a Same Day Interview?", vbYesNo + vbQuestion, "SameDay Offered?")
    '        if offer_same_date_interview = vbYes Then same_day_offered = TRUE
    '    End If
  	'	'MsgBox "This Client Appears EXPEDITED. A same-day interview needs to be offered."
	'	'same_day_interview = TRUE
	'	'Send_email = TRUE
    'END IF
	'IF expedited_status = "Client Does Not Appear Expedited" THEN MsgBox "This client does NOT appear expedited. A same-day interview does not need to be offered."
    '-----------------------------------------------------------------------------------------------EXPCASENOTE
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
    CALL write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
    CALL write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
    CALL write_variable_in_CASE_NOTE("        Utilities (AMT/HEST claimed): $" & utilities)
    CALL write_variable_in_CASE_NOTE("---")
    IF has_DISQ = TRUE THEN CALL write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
    IF has_DISQ = FALSE THEN CALL write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
    IF expedited_status = "Client Appears Expedited" AND EBT_account_status = "Y" THEN CALL write_variable_in_CASE_NOTE("* EBT Account IS open.  Recipient will NOT be able to get a replacement card in the agency.  Rapid Electronic Issuance (REI) with caution.")
    IF expedited_status = "Client Appears Expedited" AND EBT_account_status = "N" THEN CALL write_variable_in_CASE_NOTE("* EBT Account is NOT open.  Recipient is able to get initial card in the agency.  Rapid Electronic Issuance (REI) can be used, but only to avoid an emergency issuance or to meet EXP criteria.")
    IF expedited_status = "Client Does Not Appear Expedited" THEN CALL write_variable_in_CASE_NOTE("Client does not appear expedited. Application sent to ECF.")
	IF expedited_status = "Client Appears Expedited" THEN CALL write_variable_in_CASE_NOTE("Client appears expedited. Application sent to ECF.")
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
IF transfer_to_worker = "" and no_transfer_checkbox = CHECKED THEN
	transfer_case = False
    action_completed = TRUE     'This is to decide if the case was successfully transferred or not
ELSE
	transfer_case = True
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMreadscreen servicing_worker, 3, 18, 65
	servicing_worker = trim(servicing_worker)
	'msgbox servicing_worker
	IF servicing_worker = transfer_to_worker THEN
		MsgBox "This case is already in the requested worker's number."
		action_completed = False
		PF10 'backout
		PF3 'SPEC menu
		PF3 'SELF Menu'
	ELSE
	    EMWriteScreen "X127" & transfer_to_worker, 18, 61
	    transmit
		'msgbox "stop"
	    EMReadScreen worker_check, 9, 24, 2
	    IF worker_check = "SERVICING" THEN
            action_completed = False
	    	PF10 'backout
			PF3 'SPEC menu
			PF3 'SELF Menu'
	    END IF
		'msgbox "stop"
        EMReadScreen transfer_confirmation, 16, 24, 2
        IF transfer_confirmation = "CASE XFER'D FROM" then
        	action_completed = True
        Else
            action_completed = False
			'msgbox action_completed
        End if
	END IF
END IF



'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)

'IF send_email = True THEN CALL create_outlook_email("HSPH.EWS.Triagers@hennepin.us", "", "Case #" & maxis_case_number & " Expedited case to be assigned, transferred to team. " & worker_number & "  EOM.", "", "", TRUE)
IF how_application_rcvd = "Request to APPL Form" and METS_retro_checkbox = UNCHECKED and team_601_email_checkbox =  UNCHECKED and MA_transition_request_checkbox = UNCHECKED and Auto_Newborn_checkbox = UNCHECKED THEN
    CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & " Request to APPL form received-APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)
    ELSEIF Auto_Newborn_checkbox = CHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & " Request to APPL form received-APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)
END IF

IF METS_retro_checkbox = CHECKED and team_601_email_checkbox = UNCHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " Retro Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

IF METS_retro_checkbox = CHECKED and team_601_email_checkbox = CHECKED THEN CALL create_outlook_email("HSPH.EWS.TEAM.601@hennepin.us", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " Retro Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

IF MA_transition_request_checkbox = CHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " MA Transition Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

'----------------------------------------------------------------------------------------------------NOTICE APPT LETTER Dialog
send_appt_ltr = FALSE
IF cash_pends = TRUE or cash2_pends = TRUE or SNAP_pends = TRUE or grh_pends or instr(programs_applied_for, "EGA") THEN send_appt_ltr = TRUE

IF send_appt_ltr = TRUE THEN
	IF expedited_status = "Client Appears Expedited" THEN
        'creates interview date for 7 calendar days from the CAF date
    	interview_date = dateadd("d", 7, application_date)
    	If interview_date <= date then interview_date = dateadd("d", 7, date)
    ELSE
        'creates interview date for 7 calendar days from the CAF date
    	interview_date = dateadd("d", 10, application_date)
    	If interview_date <= date then interview_date = dateadd("d", 10, date)
    END IF

    Call change_date_to_soonest_working_day(interview_date)

    application_date = application_date & ""
    interview_date = interview_date & ""		'turns interview date into string for variable
 	'need to handle for if we dont need an appt letter, which would be...'
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 266, 80, "APPOINTMENT LETTER"
	EditBox 185, 20, 55, 15, interview_date
	ButtonGroup ButtonPressed
		OkButton 155, 60, 50, 15
		CancelButton 210, 60, 50, 15
	EditBox 50, 20, 55, 15, application_date
	Text 120, 25, 60, 10, "Appointment date:"
	GroupBox 5, 5, 255, 35, "Enter a new appointment date only if it's a date county offices are not open."
	Text 15, 25, 35, 10, "CAF date:"
	Text 25, 45, 205, 10, "If interview is being completed please use today's date"
    EndDialog

	Do
		Do
    		err_msg = ""
    		dialog Dialog1
    		cancel_confirmation
			If isdate(application_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid application date."
    		If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid interview date."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    'This checks to make sure the case is not in background and is in the correct footer month for PND1 cases.
    Do
    	call navigate_to_MAXIS_screen("STAT", "SUMM")
    	EMReadScreen month_check, 11, 24, 56 'checking for the error message when PND1 cases are not in APPL month
    	IF left(month_check, 5) = "CASES" THEN 'this means the case can't get into stat in current month
    		EMWriteScreen mid(month_check, 7, 2), 20, 43 'writing the correct footer month (taken from the error message)
    		EMWriteScreen mid(month_check, 10, 2), 20, 46 'writing footer year
    		EMWriteScreen "STAT", 16, 43
    		EMWriteScreen "SUMM", 21, 70
    		transmit 'This transmit should take us to STAT / SUMM now
    	END IF
    	'This section makes sure the case isn't locked by background, if it is it will loop and try again
    	EMReadScreen SELF_check, 4, 2, 50
    	If SELF_check = "SELF" then
    		PF3
    		Pause 2
    	End if
    Loop until SELF_check <> "SELF"

	last_contact_day = DateAdd("d", 30, application_date)
	If DateDiff("d", interview_date, last_contact_day) < 0 Then last_contact_day = interview_date

	'Navigating to SPEC/MEMO
	Call start_a_new_spec_memo		'Writes the appt letter into the MEMO.
    Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
    Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & interview_date & ". **")
    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday thru Friday.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    ' Call write_variable_in_SPEC_MEMO(" ")    'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
    ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **")
    CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
    ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
	PF4

    start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
    Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
    Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
    Call write_variable_in_CASE_NOTE("* A link to the Domestic Violence Brochure sent to client in SPEC/MEMO as part of notice.")
    Call write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
	PF3
END IF

'IF same_day_offered = TRUE and how_application_rcvd = "Office" THEN
'   	start_a_blank_CASE_NOTE
'  	Call write_variable_in_CASE_NOTE("~ Same-day interview offered ~")
'  	Call write_variable_in_CASE_NOTE("* Agency informed the client of needed interview.")
'  	Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive 'a denial notice")
'  	Call write_variable_in_CASE_NOTE("* A Domestic Violence Brochure has been offered to client as part of application packet.")
'  	Call write_variable_in_CASE_NOTE("---")
'  	CALL write_variable_in_CASE_NOTE (worker_signature)
'	PF3
'END IF

IF action_completed = False and servicing_worker <> transfer_to_worker THEN
    script_end_procedure_with_error_report ("Warning! Case did not transfer. Transfer the case manually. Script was able to complete all other steps.")
ELSEIF action_completed = False and servicing_worker = transfer_to_worker THEN
	script_end_procedure_with_error_report ("Warning! Case was already in requested worker's number. Script was able to complete all other steps.")
Else
    script_end_procedure_with_error_report ("Success! CASE/NOTE has been updated please review to ensure case was processed correctly.")
END IF
