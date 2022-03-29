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
call changelog_update("03/29/2022", "Removed APPLYMN as application option.", "Ilse Ferris, Hennepin County")
call changelog_update("03/11/2022", "Added randomizer functionality for Adults appplications that appear expedited. Caseloads suggested will be either EX1 or EX2", "Ilse Ferris, Hennepin County")
call changelog_update("03/07/2022", "Updated METS retro contact from Team 601 to Team 603.", "Ilse Ferris, Hennepin County")
call changelog_update("1/6/2022", "The script no longer allows you to change the Appointment Notice date if one is required based on the pending programs. This change is to ensure compliance with notification requirements of the On Demand Waiver.", "Casey Love, Hennepin County")
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
call changelog_update("09/29/2021", "Added functionality to determine HEST utility allowances based on application date. ", "Ilse Ferris, Hennepin County")
call changelog_update("09/17/2021", "Removed the field for 'Requested by X#' in the 'Request to APPL' option as this information will no longer be in the CASE/NOTE as this information is not pertinent to case actions/decisions.##~##", "Casey Love, Hennepin County")
call changelog_update("09/10/2021", "This is a very large update for the script.##~## ##~##We have reordered the functionality and consolidated the dialogs to have fewer interruptions in the process and to support the natural order of completing a pending update.##~##", "Casey Love, Hennepin County")
call changelog_update("08/03/2021", "GitHub Issue #547, added Mail as an option for how an application can be received.", "MiKayla Handley, Hennepin County")
call changelog_update("08/01/2021", "Changed the notices sent in 2 ways:##~## ##~## - Updated verbiage on how to submit documents to Hennepin.##~## ##~## - Appointment Notices will now be sent with a date of 5 days from the date of application.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("01/29/2021", "Updated Request for APPL handling per case assignement request. Issue #322", "MiKayla Handley, Hennepin County")
call changelog_update("01/07/2021", "Updated worker signature as a mandatory field.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Temporarily removed in person option for how applications are received.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Update to make confirmation number mandatory for MN Benefit application.", "MiKayla Handley, Hennepin County")
call changelog_update("11/15/2020", "Updated droplist to add virtual drop box option to how the application was received.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/13/2020", "Enhanced date evaluation functionality when which determining HEST standards to use.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/24/2020", "Added Mnbenefits application and removed SHIBA and apply MN options.", "MiKayla Handley, Hennepin County")
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
EMConnect ""                                        'Connecting to BlueZone
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

'Grabbing case and program status information from MAXIS.
'For tis script to work correctly, these must be correct BEFORE running the script.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
EMReadScreen case_status, 15, 8, 9                  'Now we are reading the CASE STATUS string from the panel - we want to make sure this does NOT read CAF1 PENDING
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

case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "CASE IS PENDING - " & case_pending        'Adding details about CASE/CURR information to a script report out to BZST
If case_status = "CAF1 PENDING" OR case_pending = False Then                    'The case MUST be pending and NOT in PND1 to continue.
    call script_end_procedure_with_error_report("This case is not in PND2 status. Current case status in MAXIS is " & case_status & ". Update MAXIS to put this case in PND2 status and then run the script again.")
End If

call back_to_SELF           'resetting

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
If unknown_cash_pending = True Then programs_applied_for = programs_applied_for & "Cash, "
If ga_status = "PENDING" Then programs_applied_for = programs_applied_for & "GA, "
If msa_status = "PENDING" Then programs_applied_for = programs_applied_for & "MSA, "
If mfip_status = "PENDING" Then programs_applied_for = programs_applied_for & "MFIP, "
If dwp_status = "PENDING" Then programs_applied_for = programs_applied_for & "DWP, "
If ive_status = "PENDING" Then programs_applied_for = programs_applied_for & "IV-E, "
If grh_status = "PENDING" Then programs_applied_for = programs_applied_for & "GRH, "
If snap_status = "PENDING" Then programs_applied_for = programs_applied_for & "SNAP, "
If ega_status = "PENDING" Then programs_applied_for = programs_applied_for & "EGA, "
If ea_status = "PENDING" Then programs_applied_for = programs_applied_for & "EA, "
If cca_status = "PENDING" Then programs_applied_for = programs_applied_for & "CCA, "
If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then programs_applied_for = programs_applied_for & "HC, "

programs_applied_for = trim(programs_applied_for)  'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "PROG")           'going here because this is a good background for the dialog to display against.

IF IsDate(application_date) = False THEN
    stop_early_msg = "This script cannot continue as the application date could not be found from MAXIS."
    stop_early_msg = stop_early_msg & vbCr & vbCr & "CASE: " & MAXIS_case_number
    stop_early_msg = stop_early_msg & vbCr & "Application Date: " & application_date
    stop_early_msg = stop_early_msg & vbCr & "Programs applied for: " & programs_applied_for
    stop_early_msg = stop_early_msg & vbCr & vbCr & "If you are unsure why this happened, screenshot this and send it to HSPH.EWS.BlueZoneScripts@hennepin.us"
    Call script_end_procedure_with_error_report(stop_early_msg)
End If

'THIS COMMENTED OUT DIALOG IS THE DLG EDITOR FRIENDLY VERSION SINCE THERE IS LOGIC IN THE DIALOG
'-------------------------------------------------------------------------------------------------DIALOG
' BeginDialog Dialog1, 0, 0, 266, 335, "Application Received for: & programs_applied_for & on & date"
'   GroupBox 5, 5, 255, 120, "Application Information"
'   DropListBox 85, 40, 95, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Mystery Doc Queue"+chr(9)+"Online"+chr(9)+"Phone-Verbal Request"+chr(9)+"Request to APPL Form"+chr(9)+"Virtual Drop Box", how_application_rcvd
'   DropListBox 85, 60, 95, 15, "Select One:"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Populations"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer"+chr(9)+"Mnbenefits"+chr(9)+"N/A"+chr(9)+"Verbal Request", application_type
'   EditBox 85, 85, 95, 15, confirmation_number
'   DropListBox 85, 105, 170, 45, "", List2
'   Text 15, 20, 65, 10, "Date of Application:"
'   Text 85, 20, 60, 10, "date_of_application"
'   Text 185, 20, 65, 10, "Pending Programs:"
'   Text 195, 35, 50, 10, "Text14"
'   Text 195, 45, 50, 10, "Text14"
'   Text 195, 55, 50, 10, "Text14"
'   Text 195, 65, 50, 10, "Text14"
'   Text 195, 75, 50, 10, "Text14"
'   Text 10, 45, 70, 10, "Application Received:"
'   Text 10, 65, 65, 10, "Type of Application:"
'   Text 85, 75, 50, 10, "Confirmation #:"
'   Text 10, 110, 70, 10, "Population/Specialty"
'   GroupBox 5, 130, 255, 105, "Expedited Screening"
'   EditBox 130, 145, 50, 15, income
'   EditBox 130, 165, 50, 15, assets
'   EditBox 130, 185, 50, 15, rent
'   CheckBox 15, 215, 55, 10, "Heat (or AC)", heat_AC_check
'   CheckBox 85, 215, 45, 10, "Electricity", electric_check
'   CheckBox 140, 215, 35, 10, "Phone", phone_check
'   Text 25, 150, 95, 10, "Income received this month:"
'   Text 30, 170, 95, 10, "Cash, checking, or savings: "
'   Text 30, 190, 90, 10, "AMT paid for rent/mortgage:"
'   Text 195, 155, 60, 45, "The income, assets and shelter costs fields will default to $0 if left blank. "
'   Text 5, 300, 60, 10, "Worker Signature:"
'   Text 5, 280, 45, 10, "Other Notes:"
'   ButtonGroup ButtonPressed
'     OkButton 155, 315, 50, 15
'     CancelButton 210, 315, 50, 15
'   EditBox 70, 295, 190, 15, worker_signature
'   GroupBox 10, 205, 170, 25, "Utilities claimed (check below):"
'   GroupBox 190, 140, 65, 60, "**IMPORTANT**"
'   CheckBox 15, 240, 220, 10, "Check here if a HH Member is active on another MAXIS Case.", Check4
'   CheckBox 15, 255, 220, 10, "Check here if only CAF1 is completed on the application.", Check5
'   EditBox 55, 275, 205, 15, other_notes
' EndDialog

'since this dialog has different displays for SNAP cases vs non-snap cases - there are differences in the dialog size
dlg_len = 225
If snap_status = "PENDING" Then dlg_len = 335

'This is the dialog with the application information.
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 266, dlg_len, "Application Received for: " & programs_applied_for & " on " & application_date
  GroupBox 5, 5, 255, 120, "Application Information"
  DropListBox 85, 40, 95, 15, "Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Mystery Doc Queue"+chr(9)+"Online"+chr(9)+"Phone-Verbal Request"+chr(9)+"Request to APPL Form"+chr(9)+"Virtual Drop Box", how_application_rcvd
  DropListBox 85, 60, 95, 15, "Select One:"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Populations"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer"+chr(9)+"Mnbenefits"+chr(9)+"N/A"+chr(9)+"Verbal Request", application_type
  EditBox 85, 85, 95, 15, confirmation_number
  DropListBox 85, 105, 170, 45, "Select One:"+chr(9)+"Adults"+chr(9)+"Families"+chr(9)+"Specialty", population_of_case
  Text 15, 25, 65, 10, "Date of Application:"
  Text 85, 25, 60, 10, application_date
  Text 185, 20, 65, 10, "Pending Programs:"
  y_pos = 30
  If unknown_cash_pending = True Then
    Text 195, y_pos, 50, 10, "Cash"
    y_pos = y_pos + 10
  End If
  If ga_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "GA"
    y_pos = y_pos + 10
  End If
  If msa_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "MSA"
    y_pos = y_pos + 10
  End If
  If mfip_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "MFIP"
    y_pos = y_pos + 10
  End If
  If dwp_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "DWP"
    y_pos = y_pos + 10
  End If
  If ive_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "IV-E"
    y_pos = y_pos + 10
  End If
  If grh_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "GRH"
    y_pos = y_pos + 10
  End If
  If snap_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "SNAP"
    y_pos = y_pos + 10
  End If
  If ea_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "EA"
    y_pos = y_pos + 10
  End If
  If ega_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "EGA"
    y_pos = y_pos + 10
  End If
  If cca_status = "PENDING" Then
    Text 195, y_pos, 50, 10, "CCA"
    y_pos = y_pos + 10
  End If
  If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then
    Text 195, y_pos, 50, 10, "HC"
    y_pos = y_pos + 10
  End If
  Text 10, 45, 70, 10, "Application Received:"
  Text 10, 65, 65, 10, "Type of Application:"
  Text 85, 75, 50, 10, "Confirmation #:"
  Text 10, 110, 70, 10, "Population/Specialty"
  y_pos = 135
  If snap_status = "PENDING" Then
      GroupBox 5, 130, 255, 105, "Expedited Screening"
      EditBox 130, 145, 50, 15, income
      EditBox 130, 165, 50, 15, assets
      EditBox 130, 185, 50, 15, rent
      CheckBox 15, 215, 55, 10, "Heat (or AC)", heat_AC_check
      CheckBox 85, 215, 45, 10, "Electricity", electric_check
      CheckBox 140, 215, 35, 10, "Phone", phone_check
      Text 25, 150, 95, 10, "Income received this month:"
      Text 30, 170, 95, 10, "Cash, checking, or savings: "
      Text 30, 190, 90, 10, "AMT paid for rent/mortgage:"
      GroupBox 10, 205, 170, 25, "Utilities claimed (check below):"
      GroupBox 185, 140, 70, 65, "**IMPORTANT**"
      Text 190, 155, 60, 45, "The income, assets and shelter costs fields will default to $0 if left blank. "
      y_pos = 245
  End If
  CheckBox 15, y_pos, 220, 10, "Check here if a HH Member is active on another MAXIS Case.", hh_memb_on_active_case_checkbox
  y_pos = y_pos + 15
  CheckBox 15, y_pos, 220, 10, "Check here if only CAF1 is completed on the application.", only_caf1_recvd_checkbox
  y_pos = y_pos + 15
  EditBox 55, y_pos, 205, 15, other_notes
  Text 5, y_pos + 5, 45, 10, "Other Notes:"
  y_pos = y_pos + 20
  EditBox 70, y_pos, 190, 15, worker_signature
  Text 5, y_pos + 5, 60, 10, "Worker Signature:"
  y_pos = y_pos + 20
  ButtonGroup ButtonPressed
    OkButton 155, y_pos, 50, 15
    CancelButton 210, y_pos, 50, 15
EndDialog

'Displaying the dialog
Do
	Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
	    IF how_application_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
        'IF application_type = "N/A" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter in other notes what type of application was received to the agency."
	    IF application_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
        IF application_type = "Mnbenefits" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If a Mnbenefits app was received, you must enter the confirmation number and time received."
        If population_of_case = "Select One:" then err_msg = err_msg & vbNewLine & "* Please indicate the population or specialty of the case."
        If snap_status = "PENDING" Then
            If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) THEN err_msg = err_msg & vbnewline & "* The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
        End If
	    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        If snap_status = "PENDING" Then
        End If
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

app_date_with_banks = replace(application_date, "/", " ")                       'creating a variable formatted with spaces instead of '/' for reading on HCRE if needed later in the script

Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)      'We want to be acting in the application month generally

Call hest_standards(heat_AC_amt, electric_amt, phone_amt, application_date) 'function to determine the hest standards depending on the application date.

send_appt_ltr = FALSE                                           'Now we need to determine if this case needs an appointment letter based on the program(s) pending
If unknown_cash_pending = True Then send_appt_ltr = TRUE
If ga_status = "PENDING" Then send_appt_ltr = TRUE
If msa_status = "PENDING" Then send_appt_ltr = TRUE
If mfip_status = "PENDING" Then send_appt_ltr = TRUE
If dwp_status = "PENDING" Then send_appt_ltr = TRUE
If grh_status = "PENDING" Then send_appt_ltr = TRUE
If snap_status = "PENDING" Then send_appt_ltr = TRUE
If ega_status = "PENDING" Then send_appt_ltr = TRUE
' If ea_status = "PENDING" Then send_appt_ltr = TRUE

If ega_status = "PENDING" Then transfer_to_worker = "EP8"           'defaulting the transfer working for EGA cases as these are to be sent to this basket'

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

'Calculates expedited status based on above numbers - only for snap pending cases
If snap_status = "PENDING" Then
    IF (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) THEN
        If population_of_case = "Families" Then transfer_to_worker = "EZ1"      'cases that screen as expedited are defaulted to expedited specific baskets based on population
        If population_of_case = "Adults" Then
            'making sure that Adults EXP baskets are not at limit
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
    End If
    IF (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) THEN expedited_status = "Client Does Not Appear Expedited"
End If

'if the case is determined to need an appointment letter the script will default the interview date
IF send_appt_ltr = TRUE THEN
    interview_date = dateadd("d", 5, application_date)
    If interview_date <= date then interview_date = dateadd("d", 5, date)
    Call change_date_to_soonest_working_day(interview_date, "FORWARD")

    application_date = application_date & ""
    interview_date = interview_date & ""                                        'turns interview date into string for variable
End If

If population_of_case = "Families" Then                                         'families cases that have cash pending need to to to these specific baskets
    If unknown_cash_pending = True Then transfer_to_worker = "EY9"
    If mfip_status = "PENDING" Then transfer_to_worker = "EY9"
    If dwp_status = "PENDING" Then transfer_to_worker = "EY9"
End if

'The familiy cash basket has a backup if it has hit the display limit.
If transfer_to_worker = "EY9" Then
    Call navigate_to_MAXIS_screen("REPT", "PND2")
    EMWriteScreen "EY9", 21, 17
    transmit
    EMReadScreen pnd2_disp_limit, 13, 6, 35
    If pnd2_disp_limit = "Display Limit" Then transfer_to_worker = "EY8"
End If

'TODO - add more defaults to the transfer_to_worker as we confirm procedure

dlg_len = 75                'this is another dynamic dialog that needs different sizes based on what it has to display.
IF send_appt_ltr = TRUE THEN dlg_len = dlg_len + 95
IF how_application_rcvd = "Request to APPL Form" THEN dlg_len = dlg_len + 80

back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

'defining the actions dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, dlg_len, "Actions in MAXIS"
  EditBox 95, 15, 30, 15, transfer_to_worker
  CheckBox 20, 35, 185, 10, "Check here if this case does not require a transfer.", no_transfer_checkbox
  If expedited_status = "Client Appears Expedited" Then Text 130, 20, 130, 10, "This case screened as EXPEDITED."
  If expedited_status = "Client Does Not Appear Expedited" Then Text 130, 20, 130, 10, "Case screened as NOT EXPEDITED."
  GroupBox 5, 5, 255, 45, "Transfer Information"
  Text 10, 20, 85, 10, "Transfer the case to x127"
  y_pos = 55
  IF send_appt_ltr = TRUE THEN
      GroupBox 5, 55, 255, 90, "Appointment Notice"
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
  ButtonGroup ButtonPressed
    OkButton 155, y_pos, 50, 15
    CancelButton 210, y_pos, 50, 15
    IF send_appt_ltr = TRUE THEN PushButton 50, odw_btn_y_pos, 125, 13, "HSR Manual - On Demand Waiver", on_demand_waiver_button
EndDialog
'THIS COMMENTED OUT DIALOG IS THE DLG EDITOR FRIENDLY VERSION SINCE THERE IS LOGIC IN THE DIALOG
'-------------------------------------------------------------------------------------------------DIALOG
' BeginDialog Dialog1, 0, 0, 266, 220, "Request to Appl"
'   EditBox 95, 15, 30, 15, transfer_to_worker
'   CheckBox 20, 35, 185, 10, "Check here if this case does not require a transfer.", no_transfer_checkbox
'   GroupBox 5, 5, 255, 45, "Transfer Information"
'   Text 10, 20, 85, 10, "Transfer the case to x127"
'   GroupBox 5, 55, 255, 60, "Appointment Notice"
'   Text 50, 70, 55, 15, "application_date"
'   EditBox 185, 65, 55, 15, interview_date
'   Text 15, 70, 35, 10, "CAF date:"
'   Text 120, 70, 60, 10, "Appointment date:"
'   Text 50, 85, 185, 10, "If interview is being completed please use today's date."
'   Text 50, 95, 190, 20, "Enter a new appointment date only if it's a date county offices are not open."
'   GroupBox 5, 120, 255, 75, "Request to Appl Information"
'   EditBox 85, 130, 45, 15, request_date
'   EditBox 85, 150, 45, 15, request_worker_number
'   EditBox 85, 170, 45, 15, METS_case_number
'   CheckBox 150, 130, 55, 10, "MA Transition", MA_transition_request_checkbox
'   CheckBox 150, 145, 60, 10, "Auto Newborn", Auto_Newborn_checkbox
'   CheckBox 150, 160, 85, 10, "METS Retro Coverage", METS_retro_checkbox
'   CheckBox 150, 175, 85, 10, "Team 603 will process", team_603_email_checkbox
'   Text 15, 135, 60, 10, "Submission Date:"
'   Text 15, 155, 60, 10, "Requested By X#:"
'   Text 15, 175, 55, 10, "METS Case #:"
'   ButtonGroup ButtonPressed
'     OkButton 155, 200, 50, 15
'     CancelButton 210, 200, 50, 15
' EndDialog

'displaying the dialog
Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        IF no_transfer_checkbox = UNCHECKED AND transfer_to_worker = "" then err_msg = err_msg & vbNewLine & "* You must enter the basket number the case to be transferred by the script or check that no transfer is needed."
        IF no_transfer_checkbox = CHECKED and transfer_to_worker <> "" then err_msg = err_msg & vbNewLine & "* You have checked that no transfer is needed, please remove basket number from transfer field."
        IF no_transfer_checkbox = UNCHECKED AND len(transfer_to_worker) > 3 AND isnumeric(transfer_to_worker) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the last 3 digits of the worker number for transfer."
        IF send_appt_ltr = TRUE THEN
            If IsDate(interview_date) = False Then err_msg = err_msg & vbNewLine & "* The Interview Date needs to be entered as a valid date."
        End If
        IF how_application_rcvd = "Request to APPL Form" THEN
            IF request_date = "" THEN err_msg = err_msg & vbNewLine & "* If a request to APPL was received, you must enter the date the form was submitted."
            IF METS_retro_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg & vbNewLine & "* You have checked that this is a METS Retro Request, please enter a METS IC #."
            IF MA_transition_request_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg &  vbNewLine & "* You have checked that this is a METS Transition Request, please enter a METS IC #."
        End If
        If ButtonPressed = on_demand_waiver_button Then
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/On_Demand_Waiver.aspx"
            err_msg = "LOOP"
        Else
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        End If
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has     not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE

transfer_to_worker = trim(transfer_to_worker)               'formatting the information entered in the dialog
transfer_to_worker = Ucase(transfer_to_worker)
request_worker_number = trim(request_worker_number)
request_worker_number = Ucase(request_worker_number)
f = date

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

'specific formatting for certain selections
If how_application_rcvd = "Phone-Verbal Request" THEN how_application_rcvd = replace(how_application_rcvd, "Phone-Verbal Request", "Phone")
IF how_application_rcvd = "Request to APPL Form" THEN
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED AND MA_transition_request_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND Auto_Newborn_checkbox = CHECKED AND METS_retro_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "Auto Newborn")
    IF application_type = "N/A" AND METS_retro_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "METS Retro")
    IF application_type = "N/A" AND MA_transition_request_checkbox = CHECKED THEN application_type = replace(application_type, "N/A", "MA Transition")
END IF

'NOW WE START CASE NOTING - there are a few
'Initial application CNOTE - all cases get these ones
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("~ Application Received (" &  application_type & ") via " & how_application_rcvd & " for " & application_date & " ~")
CALL write_bullet_and_variable_in_CASE_NOTE("Requesting HC for MEMBER(S) ", household_persons)
CALL write_bullet_and_variable_in_CASE_NOTE("Request to APPL Form received on ", request_date)
' IF how_application_rcvd = "Request to APPL Form" THEN
' 	IF team_603_email_checkbox = UNCHECKED  THEN CALL write_variable_in_CASE_NOTE("* Emailed " & request_worker_number & " to let them know the request was processed.")
'     IF team_603_email_checkbox = CHECKED  THEN CALL write_variable_in_CASE_NOTE("* Emailed team 603 to let them know the retro request was processed.")
' END IF
CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
Call write_bullet_and_variable_in_CASE_NOTE ("Case Population", population_of_case)
IF application_type = "6696" THEN write_variable_in_CASE_NOTE ("* Form Received: METS Application for Health Coverage and Help Paying Costs (DHS-6696) ")
IF application_type = "HCAPP" THEN write_variable_in_CASE_NOTE ("* Form Received: Health Care Application (HCAPP) (DHS-3417) ")
IF application_type = "HC-Certain Pop" THEN write_variable_in_CASE_NOTE ("* Form Received: MHC Programs Application for Certain Populations (DHS-3876) ")
IF application_type = "LTC" THEN write_variable_in_CASE_NOTE ("* Form Received: Application for Medical Assistance for Long Term Care Services (DHS-3531) ")
IF application_type = "MHCP B/C Cancer" THEN write_variable_in_CASE_NOTE ("* Form Received: Minnesota Health Care Programs Application and Renewal Form Medical Assistance for Women with Breast or Cervical Cancer (DHS-3525) ")
IF application_type = "Verbal Request" THEN write_variable_in_CASE_NOTE ("* Verbal Request was made for programs. CAF will be completed with resident over the phone.")
CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
' CALL write_bullet_and_variable_in_CASE_NOTE ("Other Pending Programs", additional_programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
If transfer_to_worker <> "" THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to X127" & transfer_to_worker)
If hh_memb_on_active_case_checkbox = checked Then Call write_variable_in_CASE_NOTE("* A Member on this case is active on another MAXIS Case.")
If only_caf1_recvd_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Case Pended with only information on CAF1 of the Application.")
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)

CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)
PF3 ' to save Case note

'Functionality to send emails in certain situations
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)

'IF send_email = True THEN CALL create_outlook_email("HSPH.EWS.Triagers@hennepin.us", "", "Case #" & maxis_case_number & " Expedited case to be assigned, transferred to team. " & worker_number & "  EOM.", "", "", TRUE)
IF how_application_rcvd = "Request to APPL Form" and METS_retro_checkbox = UNCHECKED and team_603_email_checkbox =  UNCHECKED and MA_transition_request_checkbox = UNCHECKED and Auto_Newborn_checkbox = UNCHECKED THEN
    CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & " Request to APPL form received-APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)
    ELSEIF Auto_Newborn_checkbox = CHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & " Request to APPL form received-APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)
END IF

IF METS_retro_checkbox = CHECKED and team_603_email_checkbox = UNCHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " Retro Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

IF METS_retro_checkbox = CHECKED and team_603_email_checkbox = CHECKED THEN CALL create_outlook_email("HSPH.EWS.TEAM.603@hennepin.us", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " Retro Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

IF MA_transition_request_checkbox = CHECKED THEN CALL create_outlook_email("", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " MA Transition Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

IF MA_transition_request_checkbox = CHECKED and team_603_email_checkbox = CHECKED THEN CALL create_outlook_email("HSPH.EWS.TEAM.603@hennepin.us", "", "MAXIS case #" & maxis_case_number & "/METS IC #" & METS_case_number & " MA Transition Request APPL'd in MAXIS-ACTION REQUIRED.", "", "", FALSE)

'Expedited Screening CNOTE for cases where SNAP is pending
If snap_status = "PENDING" Then
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
    CALL write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
    CALL write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
    CALL write_variable_in_CASE_NOTE("        Utilities (AMT/HEST claimed): $" & utilities)
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
IF send_appt_ltr = TRUE THEN        'If we are supposed to be sending an applientment letter - it will do it here - this matches the information in ON DEMAND functionality
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
    Call write_variable_in_SPEC_MEMO(" ")
    CALL write_variable_in_SPEC_MEMO("** You can submit documents Online at www.MNBenefits.mn.gov **")
    CALL write_variable_in_SPEC_MEMO("Other options for submitting documents to Hennepin County:")
    CALL write_variable_in_SPEC_MEMO(" - Mail, Fax, or Drop Boxes at service centers")
    CALL write_variable_in_SPEC_MEMO(" - Email with document attachment.EMAIL: hhsews@hennepin.us")
    CALL write_variable_in_SPEC_MEMO("   (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types)")
    ' CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
    ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
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

'THIS IS FUNCTIONALITY WE WILL NEED TO ADD BACK IN WHEN WE RETURN TO IN PERSON.
'removal of in person functionality during the COVID-19 PEACETIME STATE OF EMERGENCY'
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

'Now we create some messaging to explain what happened in the script run.
end_msg = "Application Received has been noted."
end_msg = end_msg & vbCr & "Programs requested: " & programs_applied_for & " on " & application_date
If snap_status = "PENDING" Then end_msg = end_msg & vbCr & vbCr & "Since SNAP is pending, an Expedtied SNAP screening has been completed and noted based on resident reported information from CAF1."

IF send_appt_ltr = TRUE AND memo_found = True THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice has been sent to the resident to alert them to the need for an interview for their requested programs."
IF send_appt_ltr = TRUE AND memo_found = False THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice about the Interview appears to have failed. Contact QI Knowledge Now to have one sent manually."

If transfer_message = "" Then
    If transfer_case = True Then end_msg = end_msg & vbCr & vbCr & "Case transfer has been completed to x127" & transfer_to_worker
Else
    end_msg = end_msg & vbCr & vbCr & "FAILED CASE TRANSFER:" & vbCr & transfer_message
End If
If transfer_case = False Then end_msg = end_msg & vbCr & vbCr & "NO TRANSFER HAS BEEN REQUESTED."
IF how_application_rcvd = "Request to APPL Form" Then end_msg = end_msg & vbCr & vbCr & "CASE PENDED from a REQUEST TO APPL FORM"
script_run_lowdown = script_run_lowdown & vbCr & "END Message: " & vbCr & end_msg
Call script_end_procedure_with_error_report(end_msg)


'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/10/2021
'--Tab orders reviewed & confirmed----------------------------------------------09/10/2021
'--Mandatory fields all present & Reviewed--------------------------------------09/10/2021
'--All variables in dialog match mandatory fields-------------------------------09/10/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/10/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------09/10/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/10/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------09/10/2021
'--PRIV Case handling reviewed -------------------------------------------------09/10/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/10/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/10/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------09/10/2021
'--Script name reviewed---------------------------------------------------------09/10/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/10/2021
'--comment Code-----------------------------------------------------------------09/13/2021
'--Update Changelog for release/update------------------------------------------09/10/2021
'--Remove testing message boxes-------------------------------------------------09/10/2021
'--Remove testing code/unnecessary code-----------------------------------------09/10/2021
'--Review/update SharePoint instructions----------------------------------------09/13/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/10/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/10/2021
