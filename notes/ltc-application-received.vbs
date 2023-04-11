'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'The HC script is testing a new process. - GH 1218
Call select_testing_file("ALL", "", "notes/health-care-evaluation.vbs", "1218-hc-apps-rewrite", True, False)

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/11/2023", "This script will be replaced with a new script 'NOTES - Health Care Evaluation' in the coming weeks. Look for additional information next week.", "Casey Love, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("01/06/2020", "Updated support of mandatory fields and password handling in dialogs.", "Casey Love, Hennepin County")
call changelog_update("03/12/2018", "Fixed bug that caused the script to fail if a TIKL was set on old cases.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
HH_memb_row = 05
'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Searching for case number.
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 156, 70, "Case number dialog"
  EditBox 60, 5, 90, 15, MAXIS_case_number
  EditBox 60, 25, 30, 15, MAXIS_footer_month
  EditBox 120, 25, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 100, 50, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 95, 30, 20, 10, "Year:"
EndDialog

DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid case number."		'mandatory fields
        If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Now it checks to make sure MAXIS is running on this screen.
Call check_for_MAXIS(False)

'Navigating to STAT/HCRE so we can grab the app date
call navigate_to_MAXIS_screen("stat", "hcre")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabs autofill info from STAT
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", appl_date)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", MEDI)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "SWKR", SWKR)

'Now, because INSA and MEDI will go on the same variable, we're going to add INSA to MEDI. To separate them in the case note, we have to add a semicolon (assuming both have data).
If INSA <> "" and MEDI <> "" then
  INSA = INSA & "; " & MEDI
Else
  INSA = INSA & MEDI
End if
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 361, 410, "LTC application received dialog"
  EditBox 75, 5, 65, 15, appl_date
  EditBox 75, 25, 65, 15, appl_type
  CheckBox 150, 15, 105, 10, "A transfer has been reported", transfer_reported_check
  CheckBox 150, 30, 140, 10, "Spousal allocation has been requested", spousal_allocation_check
  EditBox 160, 45, 120, 15, forms_needed
  EditBox 30, 65, 30, 15, CFR
  EditBox 110, 65, 170, 15, HH_comp
  EditBox 65, 85, 215, 15, pre_FACI_ADDR
  DropListBox 65, 105, 215, 15, "Select one..."+chr(9)+"Age 65 or older"+chr(9)+"Adult without children"+chr(9)+"Blind/disabled"+chr(9)+"Child under 21"+chr(9)+"Parent/Caretaker"+chr(9)+"Pregnant", basis_of_elig_droplist
  EditBox 35, 125, 245, 15, FACI
  EditBox 60, 145, 220, 15, retro_request
  EditBox 35, 165, 245, 15, AREP
  CheckBox 285, 170, 60, 10, "Forms to AREP", forms_to_arep_checkbox
  EditBox 60, 185, 220, 15, SWKR
  CheckBox 285, 190, 65, 10, "Forms to SWKR", forms_to_skwr_checkbox
  EditBox 60, 205, 220, 15, INSA
  CheckBox 75, 230, 210, 10, "Check here to update additional PNLR/PNLI/PNLE information.", add_detail_from_app_checkbox
  EditBox 60, 250, 220, 15, adult_signatures
  EditBox 50, 270, 230, 15, veteran_info
  EditBox 50, 290, 230, 15, LTCC
  EditBox 55, 310, 225, 15, actions_taken
  CheckBox 5, 340, 220, 10, "Check here to have the script update PND2 to show client delay.", update_PND2_check
  CheckBox 5, 355, 280, 10, "Check here to have the script create a TIKL to deny at the 45 day mark (NON-DISA).", TIKL_45_day_check
  CheckBox 5, 370, 265, 10, "Check here to have the script create a TIKL to deny at the 60 day mark (DISA).", TIKL_60_day_check
  EditBox 65, 390, 170, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 390, 50, 15
    CancelButton 305, 390, 50, 15
  GroupBox 295, 5, 55, 120, "STAT navigation"
  ButtonGroup ButtonPressed
    PushButton 300, 20, 45, 10, "prev. panel", prev_panel_button
    PushButton 300, 30, 45, 10, "next panel", next_panel_button
    PushButton 300, 45, 45, 10, "prev. memb", prev_memb_button
    PushButton 300, 55, 45, 10, "next memb", next_memb_button
    PushButton 300, 70, 25, 10, "TYPE", TYPE_button
    PushButton 300, 80, 25, 10, "PROG", PROG_button
    PushButton 300, 90, 25, 10, "HCRE", HCRE_button
    PushButton 300, 100, 25, 10, "REVW", REVW_button
    PushButton 300, 110, 25, 10, "MEMB", MEMB_button
    PushButton 5, 130, 25, 10, "FACI:", FACI_button
    PushButton 5, 170, 25, 10, "AREP:", AREP_button
    PushButton 25, 190, 30, 10, "SWKR:", SWKR_button
    PushButton 5, 210, 25, 10, "INSA/", INSA_button
    PushButton 30, 210, 25, 10, "MEDI:", MEDI_button
  Text 5, 10, 55, 10, "Application date:"
  Text 5, 30, 65, 10, "Appl type received:"
  Text 5, 50, 150, 10, "Forms needed? 1503, 3543, 3050, 5181, AA:"
  Text 5, 70, 20, 10, "CFR:"
  Text 70, 70, 40, 10, "HH Comp:"
  Text 5, 90, 60, 10, "Pre FACI address:"
  Text 5, 110, 60, 10, "Basis of eligibilty:"
  Text 5, 150, 55, 10, "Retro requested:"
  Text 5, 190, 20, 10, "PHN/"
  Text 5, 255, 55, 10, "Adult signatures:"
  Text 5, 275, 45, 10, "Veteran info:"
  Text 10, 295, 35, 10, "LTCC info:"
  Text 5, 315, 50, 10, "Actions taken:"
  GroupBox 0, 330, 285, 55, "Actions"
  Text 5, 395, 60, 10, "Worker signature:"
EndDialog

'The main dialog
Do
    Do
        err_msg = ""
        Do
    	    Dialog Dialog1
    	    cancel_confirmation
            MAXIS_dialog_navigation
        Loop until ButtonPressed = -1
    	IF trim(appl_date) = "" or isdate(appl_date) = False THEN err_msg = err_msg & vbCr & "* Please enter an application date."
    	IF basis_of_elig_droplist = "Select one..." THEN err_msg = err_msg & vbCr & "* Please select an MA basis of eligibility."
        If trim(adult_signatures) = "" THEN err_msg = err_msg & vbCr & "* Enter who signed the LTC application."
    	IF trim(actions_taken) = "" THEN err_msg = err_msg & vbCr & "* Please discuss the actions taken."
    	IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
    	IF (TIKL_45_day_check = 1 AND TIKL_60_day_check = 1) AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* You cannot TIKL for both 45 and 60 days. Please select one or neither."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    Loop UNTIL err_msg = "" AND ButtonPressed = -1
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'checking for an active MAXIS session
Call check_for_MAXIS(False)

If add_detail_from_app_checkbox = checked Then
    call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", accounts)
    call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
    call autofill_editbox_from_MAXIS(HH_member_array, "CASH", accounts)
    call autofill_editbox_from_MAXIS(HH_member_array, "CARS", vehicles)
    call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
    call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
    call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
    call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
    call autofill_editbox_from_MAXIS(HH_member_array, "SECU", securities)
    call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_asset)
    call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
    call autofill_editbox_from_MAXIS(HH_member_array, "REST", real_estate)
    call autofill_editbox_from_MAXIS(HH_member_array, "STWK", STWK)
    call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 446, 300, "LTC Application Detail"
	  EditBox 45, 25, 225, 15, cit_id
	  EditBox 45, 45, 395, 15, accounts
	  EditBox 45, 65, 395, 15, securities
	  EditBox 45, 85, 395, 15, other_asset
	  EditBox 45, 105, 395, 15, vehicles
	  EditBox 50, 125, 390, 15, real_estate
	  EditBox 60, 145, 380, 15, earned_income
	  EditBox 70, 165, 370, 15, unearned_income
	  EditBox 35, 185, 245, 15, STWK
	  EditBox 350, 185, 90, 15, COEX_DCEX
	  EditBox 70, 205, 370, 15, notes_on_income
	  EditBox 155, 225, 285, 15, is_any_work_temporary
	  EditBox 55, 255, 385, 15, verifs_needed
	  ButtonGroup ButtonPressed
	    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
	    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
	    PushButton 335, 25, 45, 10, "next panel", next_panel_button
	    PushButton 395, 25, 45, 10, "next memb", next_memb_button
	    PushButton 5, 190, 25, 10, "STWK:", STWK_button
	    PushButton 295, 190, 25, 10, "COEX/", COEX_button
	    PushButton 320, 190, 25, 10, "DCEX:", DCEX_button
	    PushButton 5, 210, 60, 10, "Notes on Income", notes_on_income_button
	    PushButton 10, 285, 25, 10, "BUSI", BUSI_button
	    PushButton 35, 285, 25, 10, "JOBS", JOBS_button
	    PushButton 60, 285, 25, 10, "RBIC", RBIC_button
	    PushButton 85, 285, 25, 10, "UNEA", UNEA_button
	    PushButton 125, 285, 25, 10, "ACCT", ACCT_button
	    PushButton 150, 285, 25, 10, "SECU", SECU_button
	    PushButton 175, 285, 25, 10, "CARS", CARS_button
	    PushButton 200, 285, 25, 10, "REST", REST_button
	    PushButton 250, 285, 35, 10, "ELIG/HC", ELIG_HC_button
	  Text 45, 10, 260, 10, "Information as entered in STAT - this may change as verifications are received."
	  GroupBox 330, 5, 115, 35, "STAT-based navigation"
	  Text 5, 30, 40, 10, "Cit/ID/imig:"
	  Text 5, 50, 35, 10, "Accounts:"
	  Text 5, 70, 35, 10, "Securities:"
	  Text 5, 90, 25, 10, "Other:"
	  Text 5, 110, 35, 10, "Vehicles:"
	  Text 5, 130, 40, 10, "Real Estate:"
	  Text 5, 150, 55, 10, "Earned income:"
	  Text 5, 170, 65, 10, "Unearned income:"
	  Text 5, 230, 150, 10, "Is any work temporary? If so, explain details:"
	  Text 5, 260, 50, 10, "Verifs needed:"
	  GroupBox 5, 275, 110, 25, "Income panels"
	  GroupBox 120, 275, 115, 25, "Asset panels"
	  ButtonGroup ButtonPressed
	    OkButton 335, 280, 50, 15
	    CancelButton 390, 280, 50, 15
	EndDialog
    Do
        Do
        	err_msg = ""
        	Dialog Dialog1
        	cancel_confirmation
            MAXIS_dialog_navigation
        Loop UNTIL ButtonPressed = -1
        Call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
End If

'UPDATING PND2----------------------------------------------------------------------------------------------------
If update_PND2_check = 1 then
	call navigate_to_MAXIS_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
	If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
	EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
	If PND2_HC_status_check = "P" then
		EMWriteScreen "x", PND2_row, 3
		transmit
		person_delay_row = 7
		Do
			EMReadScreen person_delay_check, 1, person_delay_row, 39
			If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
			person_delay_row = person_delay_row + 2
		Loop until person_delay_check = " " or person_delay_row > 20
		PF3
	End if
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_check = 0
	End if
End if

'THE TIKL's----------------------------------------------------------------------------------------------------
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If TIKL_45_day_check = 1 then Call create_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.", 45, appl_date, False, TIKL_note_text)
If TIKL_60_day_check = 1 then Call create_TIKL("HC pending 60 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 10 days and reTIKL out.", 60, appl_date, False, TIKL_note_text)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***LTC intake***")
call write_bullet_and_variable_in_CASE_NOTE("Application date", appl_date)
call write_bullet_and_variable_in_CASE_NOTE("Application type received", appl_type)
call write_bullet_and_variable_in_CASE_NOTE("Forms Needed", forms_needed)
If forms_to_arep_checkbox = checked then CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
If forms_to_skwr_checkbox = checked then CALL write_variable_in_CASE_NOTE("* Sent form(s) to Social Worker.")
call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
call write_bullet_and_variable_in_CASE_NOTE("CFR", CFR)
call write_bullet_and_variable_in_CASE_NOTE("Pre FACI address", pre_FACI_ADDR)
call write_bullet_and_variable_in_CASE_NOTE("Basis of eligibility", basis_of_elig_droplist)
call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
call write_bullet_and_variable_in_CASE_NOTE("Retro request", retro_request)
call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
call write_bullet_and_variable_in_CASE_NOTE("PHN/SWKR", SWKR)
call write_bullet_and_variable_in_CASE_NOTE("INSA/MEDI", INSA)
call write_bullet_and_variable_in_CASE_NOTE("Adult signatures", adult_signatures)
call write_bullet_and_variable_in_CASE_NOTE("LTCC info", LTCC)
call write_bullet_and_variable_in_CASE_NOTE("Veteran information", veteran_info)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If transfer_reported_check = 1 THEN call write_variable_in_CASE_NOTE("* A transfer has been reported.")
IF spousal_allocation_check = 1 THEN Call write_variable_in_CASE_NOTE("* Spousal allocation has been requested.")
If update_PND2_check = 1 THEN Call write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
IF TIKL_45_day_check = 1 Then call write_variable_in_CASE_NOTE(TIKL_note_text)
IF TIKL_60_day_check = 1 Then call write_variable_in_CASE_NOTE(TIKL_note_text)
If add_detail_from_app_checkbox = checked Then
    IF move_verifs_needed = True THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)
    CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
    call write_variable_in_case_note("-----FOLLOWING INFORMATION WAS REPORTED ON APPLICATION-----")
    call write_variable_in_Case_note("* The actual amounts may change onceverifs/forms have been rec'vd. *")
    If earned_income <> "" or unearned_income <> "" or STWK <> "" Then Call write_variable_in_CASE_NOTE("- - - Income - - -")
    call write_bullet_and_variable_in_CASE_NOTE("Earned Income", earned_income)
    call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", unearned_income)
    call write_bullet_and_variable_in_CASE_NOTE("Notes on Income", notes_on_income)
    call write_bullet_and_variable_in_CASE_NOTE("Stop Work", STWK)
    CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary?", is_any_work_temporary)
    If accounts <> "" or securities <> "" or other_asset <> "" or vehicles <> "" or real_estate <> "" Then Call write_variable_in_CASE_NOTE("- - - Assets - - -")
    call write_bullet_and_variable_in_CASE_NOTE("Accounts", accounts)
    call write_bullet_and_variable_in_CASE_NOTE("Securities", securities)
    call write_bullet_and_variable_in_CASE_NOTE("Other", aother_asset)
    call write_bullet_and_variable_in_CASE_NOTE("Vehicles", vehicles)
    call write_bullet_and_variable_in_CASE_NOTE("Real Estate", real_estate)
    If COEX_DCEX <> "" Then Call write_variable_in_CASE_NOTE("- - -")
    CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
    IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)
End If
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure ("")
