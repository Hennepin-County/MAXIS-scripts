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
call changelog_update("03/22/2023", "Updated form names and simplified selections for how an application is received by the Case Assignment team. Updated email verbiage on a response to the 'Request for APPL' form. These updates are meant to align the script to official language and information.", "Casey Love, Hennepin County")
call changelog_update("03/21/2023", "Removed the functionality to e-mail the CCAP team if CCAP was requested with other programs on MNbenefits. This process is now supported in ECF Next and the manual e-mail process is no longer required.", "Casey Love, Hennepin County")
call changelog_update("02/23/2023", "BUG FIX for cases with a second application to better determine which application is for HC and which is for CAF Based Programs.", "Casey Love, Hennepin County")
call changelog_update("02/21/2023", "BUG FIX for cases with a second application that is for HC. These cases are not subsequent applications, and need to be handled within this script and not duplicate MEMOs and Screenings.", "Casey Love, Hennepin County")
call changelog_update("01/30/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("01/30/2023", "Script will now redirect to new script (NOTES - Subsequent Application) for cases that are already pending with a new application form received. This new script supports case actions needed for subsequent applications received.", "Casey Love, Hennepin County")
CALL changelog_update("09/12/2022", "Updated EBT card availibilty in the office direction for expedited cases.", "Ilse Ferris, Hennepin County")
call changelog_update("05/24/2022", "CASE/NOTE format updated to exclude the 'How App Received' detail. This information is important for the script operation, but is not necessary to be included in the CASE/NOTE", "Casey Love, Hennepin County")   '#799
call changelog_update("05/01/2022", "Updated the Appointment Notice to have information for residents about in person support.", "Casey Love, Hennepin County")
call changelog_update("04/27/2022", "The Application Received script is updated to check cases to find if the ADDR panel is missing or has an error. The script will stop if it discovers a possible issue with an ADDR panel as that is a mandatory panel for all cases.", "Casey Love, Hennepin County")
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
      	IF IsNumeric(MAXIS_case_number) = false or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
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

Call navigate_to_MAXIS_screen("CASE", "PERS")               'Getting client eligibility of HC from CASE PERS
pers_row = 10                                               'This is where client information starts on CASE PERS
clt_hc_is_pending = False                                   'defining this at the beginning of each row of CASE PERS
HH_members_pending = ""
Do
	EMReadScreen clt_hc_ref_numb, 2, pers_row, 3     'this reads for the end of the list
	EMReadScreen clt_hc_status, 1, pers_row, 61             'reading the HC status of each client
	'MsgBox clt_hc_status
	If clt_hc_status = "P" Then
		clt_hc_is_pending = True                             'if HC is active then we will add this client to the array to find additional information
		HH_members_pending = HH_members_pending & ", MEMB " & clt_hc_ref_numb
	End If

	pers_row = pers_row + 3         'next client information is 3 rows down
	If pers_row = 19 Then           'this is the end of the list of client on each list
		PF8                         'going to the next page of client information
		on_page = on_page + 1       'saving that we have gone to a new page
		pers_row = 10               'resetting the row to read at the top of the next page
		EMReadScreen end_of_list, 9, 24, 14
		If end_of_list = "LAST PAGE" Then Exit Do
	End If
	EMReadScreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list
	' MsgBox "next_pers_ref_numb - " & next_pers_ref_numb & vbCr & "clt_hc_status - " & clt_hc_status
Loop until next_pers_ref_numb = "  "
If left(HH_members_pending, 1) = "," Then HH_members_pending = right(HH_members_pending, len(HH_members_pending)-1)
HH_members_pending = trim(HH_members_pending)
PF3
If clt_hc_is_pending = True and InStr(programs_applied_for, "HC") = 0 Then
	If programs_applied_for <> "" Then programs_applied_for = programs_applied_for & ", HC"
	If programs_applied_for = "" Then programs_applied_for = "HC"
End If

case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "CASE IS PENDING - " & case_pending        'Adding details about CASE/CURR information to a script report out to BZST
If case_status = "CAF1 PENDING" OR (case_pending = False and clt_hc_is_pending = False) Then                    'The case MUST be pending and NOT in PND1 to continue.
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
oldest_app_date = application_date
EMReadScreen CA_1_code, 1, pnd2_row, 54                                         'reading the pending codes by program for the application date line.
EMReadScreen FS_1_code, 1, pnd2_row, 62
EMReadScreen HC_1_code, 1, pnd2_row, 65
EMReadScreen EA_1_code, 1, pnd2_row, 68
EMReadScreen GR_1_code, 1, pnd2_row, 72

EMReadScreen additional_application_check, 14, pnd2_row + 1, 17                 'looking to see if this case has a secondary application date entered
If additional_application_check = "ADDITIONAL APP" THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.
    multiple_app_dates = True           'identifying that this case has multiple application dates - this is not used specifically yet but is in place so we can output information for managment of case handling in the future.

    EMReadScreen additional_application_date, 8, pnd2_row + 1, 38               'reading the app date from the other application line
    additional_application_date = replace(additional_application_date, " ", "/")
    newest_app_date = additional_application_date
    EMReadScreen CA_2_code, 1, pnd2_row, 54                                     'reading the pending codes by program for the second application date line.
    EMReadScreen FS_2_code, 1, pnd2_row, 62
    EMReadScreen HC_2_code, 1, pnd2_row, 65
    EMReadScreen EA_2_code, 1, pnd2_row, 68
    EMReadScreen GR_2_code, 1, pnd2_row, 72


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

app_recvd_note_found = False
Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
too_old_date = DateAdd("D", -1, oldest_app_date)              'We don't need to read notes from before the CAF date

note_row = 5
previously_pended_progs = ""
MEMO_NOTE_found = False
screening_found = False
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
    If left(note_title, 33) = "~ Appointment letter sent in MEMO" Then MEMO_NOTE_found = True       'MEMO case note
    If left(note_title, 31) = "~ Received Application for SNAP" Then screening_found = True         'Exp screening case note

    if note_date = "        " then Exit Do

    note_row = note_row + 1                         'Going to the next row of the CASE/NOTEs to read the next NOTE
    if note_row = 19 then
        note_row = 5
        PF8
        EMReadScreen check_for_last_page, 9, 24, 14
        If check_for_last_page = "LAST PAGE" Then Exit Do
    End If
    EMReadScreen next_note_date, 8, note_row, 6
    If next_note_date = "        " then Exit Do                         'if we are out of notes to read - leave the loop
Loop until DateDiff("d", too_old_date, next_note_date) <= 0             'once we are past the first application date, we stop reading notes

'If we have found the application received CASE/NOTE, we want to evaluate for if we need a subsequent application or app received run
If app_recvd_note_found = True Then
    skip_start_of_subsequent_apps = True                                'defaults
    hc_case = False
    If unknown_hc_pending = True Then hc_case = True                    'finding if the case has HC pending
    If ma_status = "PENDING" Then hc_case = True
    If msp_status = "PENDING" Then hc_case = True
	If clt_hc_is_pending = True Then hc_case = True

    'if HC is pending, we need to confirm that there are 2 different applications to process.
    If hc_case = True Then hc_request_on_second_app = MsgBox("It appears this case has already had the 'Application Received' script on this case. For CAF based programs, we should only run Application Received once since the application dates need to be aligned." & vbCr & vbCR &_
															 "Case currently has the following programs pending: " & programs_applied_for & vbCr & vbCR &_
															 "The following household members have Health Care pending: " & HH_members_pending & vbCr & vbCR &_
                                                             "Are there 2 seperate applications? One for Health Care and another for CAF based program(s)?", vbquestion + vbYesNo, "Type of Application Process")
    'If no HC or if answered 'No' we need to run Subsequent Application instead
    If hc_case = False or hc_request_on_second_app = vbNo Then  call run_from_GitHub(script_repository & "notes/subsequent-application.vbs")
    If hc_case = True and hc_request_on_second_app = vbYes Then     'if this is a HC application and CAF application, we need to determine which is which.
        If application_date = oldest_app_date Then                  'defaulting the program selection based on the application dates and programs
            not_processed_app_date = newest_app_date
            If HC_1_code = "P" Then processing_application_program = "Health Care Programs"
            If HC_1_code = "P" Then other_application_program = "CAF Based Programs"
            If HC_2_code = "P" Then other_application_program = "Health Care Programs"
            If HC_2_code = "P" Then processing_application_program = "CAF Based Programs"
        End If
        If application_date = newest_app_date Then
            not_processed_app_date = oldest_app_date
            If HC_2_code = "P" Then processing_application_program = "Health Care Programs"
            If HC_2_code = "P" Then other_application_program = "CAF Based Programs"
            If HC_1_code = "P" Then other_application_program = "Health Care Programs"
            If HC_1_code = "P" Then processing_application_program = "CAF Based Programs"
        End If

        'this dialog will allow the worker to assign the type of application to the correct application date so the rest of the secipt
        Do
            Do
                err_msg = ""

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 321, 135, "On Demand Applications Dashboard"
                  DropListBox 215, 60, 95, 45, "Select One..."+chr(9)+"Health Care Programs"+chr(9)+"CAF Based Programs", processing_application_program
                  DropListBox 215, 85, 95, 45, "Select One..."+chr(9)+"Health Care Programs"+chr(9)+"CAF Based Programs", other_application_program
                  ButtonGroup ButtonPressed
                    OkButton 210, 115, 50, 15
                    CancelButton 265, 115, 50, 15
                  Text 130, 10, 90, 10, "Multiple Application Dates"
                  Text 10, 30, 300, 20, "This case has Health Care pending and multiple application dates. We need to determine if this run of the script is for a seperate Health Care application or a CAF application."
                  GroupBox 5, 50, 310, 30, "THIS APPLICATION"
                  Text 10, 65, 135, 10, "Application we are currently processing:"
                  Text 165, 65, 40, 10, application_date
                  Text 70, 90, 75, 10, " Previous Application:"
                  Text 165, 90, 40, 10, not_processed_app_date
                  Text 10, 110, 150, 20, "*** CAF Based Programs mean Cash, SNAP,         Emergency, or Housing Support. "
                EndDialog


                dialog Dialog1
                cancel_confirmation

                If processing_application_program = "Select One..." Then err_msg = err_msg & vbCr & "* Please indicate what types of programs are requested on the application you are currently processing."
                If other_application_program = "Select One..." Then err_msg = err_msg & vbCr & "* Please indicate what types of programs are requested on the application that wass previously worked on."
                If processing_application_program = other_application_program Then err_msg = err_msg & vbCr & "* If both applications are for the same types of programs, there should not be seperate application dates. Review the answers and update if incorrect. If correct, cancel the script and call the TSS Help Desk to remove the second application date."

                If err_msg <> "" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine

            Loop until err_msg = ""
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

        script_run_lowdown = script_run_lowdown & vbCr & "processing_application_program - " & processing_application_program

        'reset the programs based on what was answered for the current application and force the processing in the right way.
        If processing_application_program = "Health Care Programs" Then
            unknown_cash_pending = False
            ga_status = ""
            msa_status = ""
            mfip_status = ""
            dwp_status = ""
            grh_status = ""
            snap_status = ""
            emer_status = ""
            emer_type = ""
            programs_applied_for = "HC"
        End If

        If processing_application_program = "CAF Based Programs" Then
            unknown_hc_pending = False
            ma_status = ""
            msp_status = ""
            msp_type = ""

            programs_applied_for = replace(programs_applied_for, "HC", "")
            programs_applied_for = replace(programs_applied_for, ", ,", "")
            programs_applied_for = trim(programs_applied_for)

            If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)
        End If
    End If
End If

Call navigate_to_MAXIS_screen("SPEC", "MEMO")                   'checking to make sure the ADDR panel exists. the MEMO functionality doesnt list 'client' if the ADDR is missing
PF5
EMReadScreen recipient_type, 6, 5, 15
PF3
If recipient_type <> "CLIENT" Then
    script_run_lowdown = script_run_lowdown & vbCr & "First MEMO Recipient was: " & recipient_type
    Call script_end_procedure_with_error_report("This case appears to have an issue with the ADDR panel. Proper case actions cannot occur if the ADDR panel is missing, blank, or has another error.")
End If

Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadscreen last_name, 25, 6, 30
EMReadscreen first_name, 12, 6, 63
last_name = replace(last_name, "_", "")
first_name = replace(first_name, "_", "")
case_name = last_name & ", " & first_name
Call navigate_to_MAXIS_screen("STAT", "PROG")           'going here because this is a good background for the dialog to display against.

IF IsDate(application_date) = False THEN
    stop_early_msg = "This script cannot continue as the application date could not be found from MAXIS."
    stop_early_msg = stop_early_msg & vbCr & vbCr & "CASE: " & MAXIS_case_number
    stop_early_msg = stop_early_msg & vbCr & "Application Date: " & application_date
    stop_early_msg = stop_early_msg & vbCr & "Programs applied for: " & programs_applied_for
    stop_early_msg = stop_early_msg & vbCr & vbCr & "If you are unsure why this happened, screenshot this and send it to HSPH.EWS.BlueZoneScripts@hennepin.us"
    Call script_end_procedure_with_error_report(stop_early_msg)
End If

'since this dialog has different displays for SNAP cases vs non-snap cases - there are differences in the dialog size
dlg_len = 220
If snap_status = "PENDING" Then dlg_len = 330

'This is the dialog with the application information.
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 266, dlg_len, "Application Received for: " & programs_applied_for & " on " & application_date
  GroupBox 5, 5, 255, 120, "Application Information"
  DropListBox 85, 40, 95, 45, "Select One:"+chr(9)+"ECF"+chr(9)+"Online"+chr(9)+"Request to APPL Form"+chr(9)+"In Person", how_application_rcvd
  DropListBox 85, 60, 95, 45, "Select One:"+chr(9)+"Adults"+chr(9)+"Families"+chr(9)+"Specialty", population_of_case
  DropListBox 85, 80, 170, 45, "Select One:"+chr(9)+"CAF - 5223"+chr(9)+"MNbenefits CAF - 5223"+chr(9)+"SNAP App for Seniors - 5223F"+chr(9)+"MNsure App for HC - 6696"+chr(9)+"MHCP App for Certain Populations - 3876"+chr(9)+"App for MA for LTC - 3531"+chr(9)+"MHCP App for B/C Cancer - 3523"+chr(9)+"EA/EGA Application"+chr(9)+"No Application Required", application_type
  EditBox 85, 105, 105, 15, confirmation_number
  Text 15, 25, 65, 10, "Date of Application:"
  Text 85, 25, 60, 10, application_date
  Text 185, 15, 65, 10, "Pending Programs:"
  y_pos = 25
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
  If emer_status = "PENDING" Then
    Text 195, y_pos, 50, 10, emer_type
    y_pos = y_pos + 10
  End If
  If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True OR clt_hc_is_pending = True Then
    Text 195, y_pos, 50, 10, "HC"
    y_pos = y_pos + 10
  End If
  Text 10, 45, 70, 10, "Application Received:"
  Text 10, 65, 70, 10, "Population/Specialty"
  Text 15, 85, 65, 10, "Type of Application:"
  Text 85, 95, 50, 10, "Confirmation #:"
  y_pos = 130
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
      y_pos = 240
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
        If application_type = "MNbenefits CAF - 5223" AND how_application_rcvd = "Select One:" Then how_application_rcvd = "Online"
		If application_type = "No Application Required" AND how_application_rcvd <> "Request to APPL Form" Then err_msg = err_msg & vbNewLine & "* No Application cases can only be processed with a 'Request to APPL' form."
	    IF how_application_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
        'IF application_type = "N/A" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter in other notes what type of application was received to the agency."
	    IF application_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
        IF application_type = "MNbenefits CAF - 5223" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If a MNbenefits app was received, you must enter the confirmation number and time received."
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

app_date_with_blanks = replace(application_date, "/", " ")                       'creating a variable formatted with spaces instead of '/' for reading on HCRE if needed later in the script

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
If emer_status = "PENDING" and emer_type = "EGA" Then send_appt_ltr = TRUE
' If emer_status = "PENDING" and emer_type = "EA" Then send_appt_ltr = TRUE

If emer_status = "PENDING" and emer_type = "EGA" Then transfer_to_worker = "EP8"           'defaulting the transfer working for EGA cases as these are to be sent to this basket'

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

    'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
    Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)
    Call navigate_to_MAXIS_screen("STAT", "DISQ")
    EMReadScreen DISQ_member_check, 34, 24, 2   'Reads the DISQ info for the case note.
    If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then
    	has_DISQ = False
    Else
    	has_DISQ = True
    End if

    'Reads MONY/DISB 'Head of Household" coding to see if a card has been issued. B, H, p and R codes mean that a resident has already received a card and cannot get another in office.
    'DHS webinar meeting 07/20/2022
    in_office_card = True   'Defaulting to true
    IF expedited_status = "client appears expedited" THEN
        Call navigate_to_MAXIS_screen("MONY", "DISB")
        EmReadscreen HoH_card_status, 1, 15, 27
        If HoH_card_status = "B" or _
           HoH_card_status = "H" or _
           HoH_card_status = "P" or _
           HoH_card_status = "R" then
           in_office_card = False
        End if
    End if
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
            If hcre_app_date = app_date_with_blanks Then household_persons = household_persons & hcre_ref_nbr & ", "

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

'creating a variable for a shortened form of the application form for the CASE/NOTE header
If application_type = "CAF - 5223" Then short_form_info = "CAF"
If application_type = "MNbenefits CAF - 5223" Then short_form_info = "CAF from MNbenefits"
If application_type = "SNAP App for Seniors - 5223F" Then short_form_info = "Sr SNAP App"
If application_type = "MNsure App for HC - 6696" Then short_form_info = "MNsure HCAPP"
If application_type = "MHCP App for Certain Populations - 3876" Then short_form_info = "HC - Certain Populations"
If application_type = "App for MA for LTC - 3531" Then short_form_info = "LTC HCAPP"
If application_type = "MHCP App for B/C Cancer - 3523" Then short_form_info = "HCAPP for B/C Cancer"
If application_type = "EA/EGA Application" Then short_form_info = "EA/EGA Application"

'NOW WE START CASE NOTING - there are a few
'Initial application CNOTE - all cases get these ones
start_a_blank_case_note
If application_type = "No Application Required" Then
	'this header is for pending a case when no form is received or needed.
	MX_pend_reason = ""
	If Auto_Newborn_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "Auto Newborn & "
	If METS_retro_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "METS Retro Request & "
	If MA_transition_request_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "MA Transition & "
	MX_pend_reason = trim(MX_pend_reason)
	If right(MX_pend_reason, 1) = "&" Then MX_pend_reason = left(MX_pend_reason, len(MX_pend_reason)-1)
	MX_pend_reason = trim(MX_pend_reason)
	CALL write_variable_in_CASE_NOTE ("~ HC Pended from a METS case for " & MX_pend_reason & " effective " & application_date & " ~")
Else
	If application_type <> "EA/EGA Application" Then application_type = replace(application_type, " - ", " (DHS-") & ")"
	CALL write_variable_in_CASE_NOTE ("~ Application Received (" &  short_form_info & ") pended for " & application_date & " ~")
	CALL write_bullet_and_variable_in_CASE_NOTE("Application Form Received", application_type)
End If
CALL write_bullet_and_variable_in_CASE_NOTE("Requesting HC for MEMBER(S) ", household_persons)
CALL write_bullet_and_variable_in_CASE_NOTE("Request to APPL Form received on ", request_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
Call write_bullet_and_variable_in_CASE_NOTE ("Case Population", population_of_case)
CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
If transfer_to_worker <> "" THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to X127" & transfer_to_worker)
If hh_memb_on_active_case_checkbox = checked Then Call write_variable_in_CASE_NOTE("* A Member on this case is active on another MAXIS Case.")
If only_caf1_recvd_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Case Pended with only information on CAF1 of the Application.")
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)

CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

PF3 ' to save Case note


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

	CALL create_outlook_email(send_email_to, cc_email_to, email_subject, email_body, "", FALSE)
	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
End If

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
    If has_DISQ = True then CALL write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
    If has_DISQ = False then CALL write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
    If in_office_card = False then CALL write_variable_in_CASE_NOTE("Recipient will NOT be able to get an EBT card in an agency office. An EBT card has previously been provided to the household.")
    CALL write_variable_in_CASE_NOTE("---")
    IF expedited_status = "Client Does Not Appear Expedited" THEN CALL write_variable_in_CASE_NOTE("Client does not appear expedited. Application sent to case file.")
    IF expedited_status = "Client Appears Expedited" THEN CALL write_variable_in_CASE_NOTE("Client appears expedited. Application sent to case file.")
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
IF send_appt_ltr = TRUE THEN        'If we are supposed to be sending an appointment letter - it will do it here - this matches the information in ON DEMAND functionality
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

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the BZST connection to SQL Database'
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

    'delete a record if the case number matches
    ' objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection
    'Add a new record with this case information'
    objRecordSet.Open "INSERT INTO ES.ES_CasesPending (WorkerID, CaseNumber, CaseName, ApplDate, FSStatusCode, CashStatusCode, HCStatusCode, GAStatusCode, GRStatusCode, EAStatusCode, MFStatusCode, IsExpSnap, UpdateDate)" &  _
                      "VALUES ('" & worker_id_for_data_table & "', '" & eight_digit_case_number & "', '" & case_name_for_data_table & "', '" & application_date & "', '" & snap_stat_code & "', '" & cash_stat_code & "', '" & hc_stat_code & "', '" & ga_stat_code & "', '" & grh_stat_code & "', '" & emer_stat_code & "', '" & mfip_stat_code & "', '" & 1 & "', '" & date & "')", objConnection, adOpenStatic, adLockOptimistic

    'close the connection and recordset objects to free up resources
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
End If

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
'--All variables are CASE:NOTEing (if required)---------------------------------05/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------05/24/2022
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
'--Remove testing code/unnecessary code-----------------------------------------05/01/2022                  We were holding old NOTICE details for in person return. Removed as this detail is drastically different.
'--Review/update SharePoint instructions----------------------------------------09/13/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/10/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/10/2021
