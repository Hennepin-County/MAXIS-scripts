'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TRANSFER CASE.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 229                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/21/2021", "Updated browser to default when opening SIR from Internet Explorer to Edge.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/09/2020", "No issues with SPEC/MEMO for out-of-county cases. SIR Announcement from 11/05/20 stated an issue was identified. Hennepin County's script project is seperate from DHS's script project. We are not experiencing the reported issue. Thank you!", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/20/2020", "Updated link for out-of-county use form.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/16/2020", "Added closing message box with information if a case is identified as being transferred to the MNPrairie Bank of counties.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/20/2020", "Rewrite.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
EMConnect ""
check_for_MAXIS(True)
call MAXIS_case_number_finder(MAXIS_case_number)
'VARIABLES TO DECLARE----------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------DIALOG
DO
	BeginDialog Dialog1, 0, 0, 231, 145, "Case Transfer"
    EditBox 55, 5, 50, 15, MAXIS_case_number
    EditBox 175, 5, 50, 15, servicing_worker
    CheckBox 40, 30, 55, 10, "out of county", out_of_county_checkbox
    CheckBox 95, 30, 55, 10, "within county", in_county_checkbox
    EditBox 105, 45, 120, 15, transfer_reason
    EditBox 105, 65, 120, 15, action_taken
    EditBox 105, 85, 120, 15, requested_verifs
    EditBox 105, 105, 120, 15, other_notes
    'CheckBox 5, 130, 75, 10, "Send SPEC/MEMO", SPEC_MEMO_checkbox
    ButtonGroup ButtonPressed
    OkButton 120, 125, 50, 15
    CancelButton 175, 125, 50, 15
    Text 115, 10, 55, 10, "Transferring to:"
    Text 5, 30, 35, 10, "Transfer:"
    Text 5, 50, 70, 10, "Reason for transfer:"
    Text 5, 70, 70, 10, "Actions taken:"
    Text 5, 90, 95, 10, "Actions/verifications needed:"
    Text 5, 110, 45, 10, "Other notes:"
    Text 5, 10, 50, 10, "Case number:"
    Text 175, 20, 50, 10, " Ex. (X127ICT)"
    EndDialog

	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
	      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "Please enter a valid case number."
			IF servicing_worker = "" or len(servicing_worker) > 7 THEN err_msg = err_msg & vbNewLine & "Please enter a valid  worker number (EX. X120ICT)."
			IF out_of_county_checkbox <> CHECKED and in_county_checkbox <> CHECKED THEN err_msg = err_msg & vbNewLine & "Please select the type of transfer."
			IF transfer_reason = "" THEN err_msg = err_msg & vbNewLine & "Please enter a reason for transfer."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
		'----------Checks that the worker or agency is valid---------- 'must find user information before transferring to account for privileged cases.
	servicing_worker = ucase(servicing_worker)
	call navigate_to_MAXIS_screen("REPT", "USER")
	EMWriteScreen servicing_worker, 21, 12
	TRANSMIT
	EMReadScreen worker_found, 15, 24, 2
	EMReadScreen inactive_worker, 8, 7, 38
	IF inactive_worker = "INACTIVE" THEN MsgBox "The worker or agency selected is not active. Please try again."
	IF worker_found = "NO WORKER FOUND" THEN MsgBox "The worker or agency selected does not exist. Please try again."
LOOP UNTIL inactive_worker <> "INACTIVE" AND worker_found <> "NO WORKER FOUND"

EMWriteScreen "X", 7, 3 ' navigating to read the worker information'
TRANSMIT
EMReadScreen worker_agency_name, 43, 8, 27
worker_agency_name = trim(worker_agency_name)
IF worker_agency_name = "" THEN 						'If we are unable to find the alias for the worker we will just use the worker name as it is what is used on notices anyway
	EMReadScreen worker_agency_name, 43, 7, 27
	worker_agency_name = trim(worker_agency_name)
	name_length = len(worker_agency_name)
	comma_location = InStr(worker_agency_name, ",")
	worker_agency_name = right(worker_agency_name, (name_length - comma_location)) & " " & left(worker_agency_name, (comma_location - 1)) 'this section will reorder the name of the worker since it is stored here as last, first. the comma_location - 1 removes the comma from the "last,"
END IF
EMReadScreen mail_addr_line_one, 43, 9, 27
	mail_addr_line_one = trim(mail_addr_line_one)
EMReadScreen mail_addr_line_two, 43, 10, 27
	mail_addr_line_two = trim(mail_addr_line_two)
EMReadScreen mail_addr_line_three, 43, 11, 27
	mail_addr_line_three = trim(mail_addr_line_three)
EMReadScreen mail_addr_line_four, 43, 12, 27
	mail_addr_line_four = trim(mail_addr_line_four)
EMReadScreen worker_agency_phone, 14, 13, 27
EMReadScreen worker_county_code, 2, 15, 32

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
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
'msgbox active_programs
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

EMReadScreen previous_servicing_worker, 7, 21, 21
IF previous_servicing_worker = servicing_worker THEN script_end_procedure_with_error_report("This case is already in this worker's number, the case will not transfer." & vbcr & "Current worker: " & vbcr & previous_servicing_worker & vbcr & "Requested worker: " & vbcr & servicing_worker )

'goes to MEMO and create a new memo to be send out
If in_county_checkbox = CHECKED THEN
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("~ Case transferred within the county ~")
    Call write_bullet_and_variable_in_CASE_NOTE("Active programs", active_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", pend_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason case was transferred", transfer_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("Transferred to",  servicing_worker)
	Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verifs)
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
    PF3
    '-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
    'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
    EMWriteScreen "x", 7, 16
    TRANSMIT
    PF9
    EMWriteScreen servicing_worker, 18, 61
    TRANSMIT
    EMReadScreen worker_check, 9, 24, 2
    IF worker_check = "SERVICING" THEN
    	action_completed = False
    	PF10
    END IF
    EMReadScreen transfer_confirmation, 16, 24, 2
    IF transfer_confirmation = "CASE XFER'D FROM" then
    	action_completed = True
    Else
    	action_completed = False
    End if
    PF3
    IF action_completed = TRUE THEN script_end_procedure("Case has been transferred and a case note created.")
    IF action_completed = FALSE THEN script_end_procedure_with_error_report("Case did not transfer, please review for where the error occured.")
END If	'END OF IN COUNTY TRANSFER-------------------------------------------------------------------------------------------------------------------------------------

IF out_of_county_checkbox = CHECKED THEN 'BEGINNING OF OUT OF COUNTY TRANSFER----------------------------------------------------------------------------------------------------
county_financial_responsibilty = "27"

DO
	DO
		DO
		   	'-------------------------------------------------------------------------------------------------DIALOG
		    Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 206, 265, "Out of county transfer"
              EditBox 150, 15, 45, 15, move_date
              DropListBox 65, 35, 40, 15, "No"+chr(9)+"Yes", excluded_time
              EditBox 150, 35, 45, 15, excluded_date
              CheckBox 10, 55, 95, 10, "HC CFR is not changing", hc_cfr_no_change_checkbox
              CheckBox 10, 70, 100, 10, "CASH CFR is not changing.", cash_cfr_no_change_checkbox
              CheckBox 10, 85, 170, 10, "Manually set the CFR and change date for CASH.", manual_cfr_cash_checkbox
              CheckBox 10, 100, 170, 10, "Manually set the CFR and change date for HC.", manual_cfr_hc_checkbox
              EditBox 55, 115, 20, 15, county_financial_responsibilty
              EditBox 165, 115, 15, 15, cfr_month
              EditBox 180, 115, 15, 15, cfr_year
              CheckBox 10, 150, 80, 10, "Active METS case", mnsure_active_checkbox
              CheckBox 10, 160, 95, 10, "Pending METS application", mnsure_pend_checkbox
              CheckBox 10, 170, 100, 10, "Active on Minnesota Care", mcre_active_checkbox
              EditBox 155, 150, 40, 15, METS_case_number
              CheckBox 130, 170, 60, 10, "Updated MMIS", mmis_update_checkbox
              CheckBox 10, 195, 95, 10, "Change Report Form sent", crf_sent_checkbox
              CheckBox 10, 205, 105, 10, "DHS 3195 transfer form sent", transfer_form_checkbox
              CheckBox 10, 215, 155, 10, "Add a programs closure date to the MEMO", closure_date_checkbox
              CheckBox 10, 225, 185, 10, "Send a SPEC/MEMO to the client with new worker info", SPEC_MEMO_checkbox
              EditBox 40, 245, 70, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 115, 245, 40, 15
                CancelButton 160, 245, 40, 15
              Text 110, 20, 40, 10, "Move Date:"
              Text 10, 40, 55, 10, "Excluded time?"
              Text 110, 40, 40, 10, "Begin Date"
              GroupBox 5, 5, 195, 135, "Current Financial Responsibilty"
              Text 10, 120, 45, 10, "Current CFR:"
              Text 90, 120, 75, 10, "Change Date MM/YY:"
              GroupBox 5, 140, 195, 45, "Health Care Information"
              GroupBox 5, 185, 195, 55, "Client Contact Information"
              Text 125, 155, 25, 10, "Case #:"
              Text 5, 250, 35, 10, "Signature:"
              ButtonGroup ButtonPressed
                PushButton 10, 15, 85, 15, "NAV to SPEC/XFER", nav_to_xfer_button
              Text 135, 55, 60, 10, "*See CM 0006.06"
            EndDialog

	    	IF ButtonPressed = nav_to_xfer_button THEN
				CALL navigate_to_MAXIS_screen("SPEC", "XFER")
				EMWriteScreen "X", 9, 16
				TRANSMIT
			END IF

			err_msg = ""
			dialog Dialog1
			cancel_confirmation
			IF excluded_time = "Yes" AND isdate(excluded_date) = FALSE THEN MsgBox "Please enter a valid date for the start of excluded time or double check that the client's absense is due to excluded time."
			IF isdate(move_date) = FALSE THEN MsgBox "Please enter a valid date for client move."
			IF ucase(left(servicing_worker, 4)) = ucase(transferring_worker_county_code) THEN MsgBox "You must use the ''Within the Agency'' script to transfer the case within the agency. The Worker/Agency you have selected indicates you are trying to transfer within your agency."
			IF (hc_status_check = "ACTV" AND excluded_time = "") THEN MsgBox "Please select whether the client is on excluded time."
			IF manual_cfr_cash_checkbox = CHECKED AND cash_cfr_no_change_checkbox = CHECKED THEN MsgBox ("Please select whether the CFR for CASH is changing or not. Review input.")
			IF manual_cfr_hc_checkbox = CHECKED AND hc_cfr_no_change_checkbox = CHECKED THEN MsgBOx ("Please select whether the CFR for HC is changing or not. Review input.")
			If mets_active_checkbox = CHECKED and METS_case_number = "" then err_msg = err_msg & vbNewLine & "* Please enter a METS case number."
			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false
						'loops until user passwords back in
		IF closure_date_checkbox = CHECKED THEN
			closure_date = dateadd("m", 1, date)
			closure_date = datepart("M", closure_date) & "/01/" & datepart("YYYY", closure_date)
			IF datediff("d", date, closure_date) < 10 THEN
				closure_date = dateadd("m", 1, closure_date)
			END IF
			closure_date = dateadd("d", -1, closure_date)
		END IF
	call navigate_to_MAXIS_screen("CASE", "NOTE")
	PF9
	EMReadScreen write_access, 9, 24, 12
	IF write_access = "READ ONLY" THEN MsgBox("You do not have access to modify this case. Please double check your case number and try again." & chr(13) & chr(13) & "Alternatively, you may be in INQUIRY MODE.")
LOOP UNTIL write_access <> "READ ONLY"

PF3
	'----------Sending the Client a SPEC/MEMO notifying them of the details of the transfer----------
	If SPEC_MEMO_checkbox = CHECKED then
	    Call start_a_new_spec_memo		'Writes the appt letter into the MEMO.
	    Call write_variable_in_SPEC_MEMO("Your case has been transferred. Your new agency/worker is: " & worker_agency_name & "")
	    Call write_variable_in_SPEC_MEMO("If you have any questions, or to send in requested proofs,")
	    Call write_variable_in_SPEC_MEMO("please direct all communications to the agency listed.")
	    Call write_variable_in_SPEC_MEMO(worker_agency_name)
	    Call write_variable_in_SPEC_MEMO(mail_addr_line_one)
	    Call write_variable_in_SPEC_MEMO(mail_addr_line_two)
	    Call write_variable_in_SPEC_MEMO(mail_addr_line_three)
	    Call write_variable_in_SPEC_MEMO(mail_addr_line_four)
	    Call write_variable_in_SPEC_MEMO(worker_agency_phone)
		IF closure_date_checkbox = CHECKED THEN
			Call write_variable_in_SPEC_MEMO("If you fail to provide required proofs to your new worker")
			Call write_variable_in_SPEC_MEMO("or agency then your benefits will close on " & closure_date & ".")
		END IF
	    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
	    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")

		'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
    	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
    	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
    	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
    	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
    	    PF5                                                     'PF5s again to initiate the new memo process
		END IF
		'Checking for SWKR
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
    	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
    	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
    	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
    	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
    	    PF5                                           'PF5s again to initiate the new memo process
		END IF
		EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		TRANSMIT                                       'TRANSMITs to start the memo writing process
		PF4'save and exit
	End if

    'MNPrairie Bank Support - MNPrairie Bank cases all go to Steele (county code 74)'s ICT transfer. 
    'Agencies in the MNPrairie Bank are Dodge (county code 20), Steele (county code 74), and Waseca (county code 81)
	IF servicing_worker = "X120ICT" OR servicing_worker = "X181ICT" THEN servicing_worker = "X174ICT"
    
    'Using move date to determine CRF change date.
	cfr_date = dateadd("M", 1, move_date)
	cfr_date = datepart("M", cfr_date) & "/01/" & datepart("YYYY", cfr_date)
	cfr_date = dateadd("M", 2, cfr_date)
	cfr_month = datepart("M", cfr_date)
	IF len(cfr_month) <> 2 THEN cfr_month = "0" & cfr_month
	cfr_year = datepart("YYYY", cfr_date)
	cfr_year = right(cfr_year, 2)

	'----------The case note of the reason for the XFER----------
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("~ Case transferred to " & servicing_worker & " ~")
    Call write_bullet_and_variable_in_CASE_NOTE("Active programs", active_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", pend_programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason case was transferred", transfer_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verifs)
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	IF mmis_update_checkbox = CHECKED THEN Call write_variable_in_CASE_NOTE("* Updated MMIS")
	IF mets_active_checkbox = CHECKED THEN call write_variable_in_case_note("* Client is active on HC through METS case number:", METS_case_number)
	IF mcre_active_checkbox = CHECKED THEN call write_variable_in_case_note("* Client is active on HC through Minnesota Care")
	IF pend_programs <> "" THEN call write_bullet_and_variable_in_case_note("Pending Programs", (left(pend_programs, (len(pend_programs) - 1))))
	IF mnsure_pend_checkbox = CHECKED THEN call write_variable_in_case_note("* Client has pending HC application through METS")
	call write_bullet_and_variable_in_case_note("Client Move Date", Client_move_date)
	IF crf_sent_checkbox = CHECKED THEN
		crf_sent_date = date
		call write_bullet_and_variable_in_case_note("Change Report Sent", crf_sent_date)
		call write_bullet_and_variable_in_case_note("Case File Sent:", crf_sent_date)
	END IF
	IF excluded_time = "Yes" THEN
		excluded_time = excluded_time & ", Begins " & excluded_date
		call write_bullet_and_variable_in_case_note("Excluded Time" , excluded_time)
		excluded_time = "Yes"
	ELSEIF excluded_time = "No" THEN
		call write_bullet_and_variable_in_case_note("Excluded Time", excluded_time)
	ELSEIF excluded_time = "" THEN
		call write_variable_in_case_note("* Excluded Time: N/A")
	END IF
	IF hc_status_check = "ACTV" THEN
		CALL write_bullet_and_variable_in_case_note("HC County of Financial Responsibility", hc_cfr)
		IF hc_cfr_no_change_checkbox = 0 THEN
			CALL write_bullet_and_variable_in_case_note("HC CFR Change Date", (cfr_month & "/" & cfr_year))
		ELSE
			CALL write_bullet_and_variable_in_case_note("HC CFR", "Not changing")
		END IF
	END IF
	IF cash1_status_check = "ACTV" OR cash2_status_check = "ACTV" THEN
		CALL write_bullet_and_variable_in_case_note("CASH County of Financial Responsibility", county_financial_responsibilty)
		IF cash_cfr_no_change_checkbox = 0 THEN
			CALL write_bullet_and_variable_in_case_note("CASH CFR Change Date", (cfr_month & "/" & cfr_year))
		ELSE
			CALL write_bullet_and_variable_in_case_note("CASH CFR", "Not changing")
		END IF
	END IF
	IF closure_date_checkbox = CHECKED THEN call write_variable_in_case_note("* Client has until " & closure_date & " to provide required proofs or the case will close.")
	If transfer_form_checkbox = CHECKED THEN call write_variable_in_case_note("* DHS 3195 Inter Agency Case Transfer Form completed and sent.")
	IF SPEC_MEMO_checkbox = CHECKED THEN call write_variable_in_case_note("* SPEC/MEMO sent to client with new worker information.")
	IF forms_to_arep = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to AREP.")
	IF forms_to_swkr = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to social worker.")
    Call write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
    PF3
	'----------The business end of the script (DON'T POINT THIS SCRIPT AT ANYTHING YOU DON'T WANT TRANSFERRED!!)----------
	call navigate_to_MAXIS_screen("SPEC", "XFER")
	EMWriteScreen "X", 9, 16
	TRANSMIT
	PF9

	'Writing client move date
	call create_MAXIS_friendly_date(move_date, 0, 4, 28)

	'Writing CRF date, if CRF check is CHECKED
	IF crf_sent_checkbox = CHECKED THEN call create_MAXIS_friendly_date(crf_sent_date, 0, 4, 61)
	IF crf_sent_checkbox = CHECKED THEN call create_MAXIS_friendly_date(crf_sent_date, 0, 5, 61)
	'Writes the excluded time info. Only need the left character (it's a dropdown)
	EMWriteScreen left(excluded_time, 1), 5, 28

	'If there's excluded time, need to write the info
	IF excluded_time = "Yes" THEN
		call create_MAXIS_friendly_date(excluded_date, 0, 6, 28)
		EMWriteScreen county_financial_responsibilty, 15, 39
	END IF

	IF excluded_date = "" AND excluded_time = "No" THEN
		EMWriteScreen "__", 6, 28
		EMWriteScreen "__", 6, 31
		EMWriteScreen "__", 6, 34
	END IF

	IF hc_status_check = "ACTV" THEN
		EMWriteScreen county_financial_responsibilty, 14, 39
		EMWriteScreen cfr_month, 14, 53
		EMWriteScreen cfr_year, 14, 59
	END IF

	IF cash1_status_check = "ACTV" THEN
		EMWriteScreen county_financial_responsibilty, 11, 39
		EMWriteScreen cfr_month, 11, 53
		EMWriteScreen cfr_year, 11, 59
	END IF


	IF cash2_status_check = "ACTV" AND cash_cfr_no_change_check = 0 THEN
		EMWriteScreen county_financial_responsibilty, 12, 39
		EMWriteScreen cfr_month, 12, 53
		EMWriteScreen cfr_year, 12, 59
	END IF

	EMReadScreen primary_worker, 7, 21, 16
	EMWriteScreen primary_worker, 18, 28
	EMWriteScreen servicing_worker, 18, 61
	TRANSMIT
	IF MAXIS_error_message <> "" THEN
		EMReadScreen MAXIS_error_message, 75, 24, 02
		MAXIS_error_message = trim(MAXIS_error_message)

		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 231, 95, "Maxis Message, please screen shot"
			ButtonGroup ButtonPressed
			OkButton 135, 75, 45, 15
			CancelButton 180, 75, 45, 15
			GroupBox 5, 0, 220, 50, "You can update maxis if there is an error, THEN hit ok to continue."
			Text 15, 10, 190, 35, MAXIS_error_message
			EditBox 50, 55, 175, 15, email_BZST
			Text 5, 60, 45, 10, "Email BZST:"
		EndDialog

		'Showing case number dialog
		Do
		  Dialog Dialog1
		  cancel_without_confirmation
		  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in
		IF email_BZST <> "" THEN CALL create_outlook_email("Mikayla.Handley@hennepin.us", "", "Case #" & maxis_case_number & " Error message: " & MAXIS_error_message & "  EOM.", "", "", TRUE)
	END IF

	PF3

	navigate_decision = Msgbox("Do you want to navigate to open a the case transfer useform?" & vbcr & "If you are using Chrome you will be asked to open in Adobe.", vbQuestion + vbYesNo, "Navigate to Useform?")
	If navigate_decision = vbYes then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://aem.hennepin.us/rest/services/HennepinCounty/Processes/ServletRenderForm:1.0?formName=HSPH5069_1-0.xdp&interactive=1"
	If navigate_decision = vbNo then navigate_to_form = False
	If servicing_worker = "X120ICT" or servicing_worker = "X181ICT" then 
        script_end_procedure_with_error_report("You've transferred a case to the MNPrairie county bank. The case will be transferred to X174ICT per their process." & vbcr & vbcr & "Please review to ensure case has been transferred, a memo sent, and a case note created.")
    Else 
        script_end_procedure_with_error_report("Please review to ensure case has been transferred, a memo sent, and a case note created.")
    End If 
END IF
