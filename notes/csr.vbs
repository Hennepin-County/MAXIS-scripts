'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CSR.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("01/17/2017", "This script has been updated to clean up the case note. The script was case noting the ''Verifs Needed'' section twice. This has been resolved.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 171, 220, "Case number dialog"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 70, 25, 30, 15, MAXIS_footer_month
  EditBox 105, 25, 30, 15, MAXIS_footer_year
  CheckBox 30, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 80, 60, 30, 10, "GRH", GRH_checkbox
  CheckBox 130, 60, 25, 10, "HC", HC_checkbox
  CheckBox 10, 80, 90, 10, "Is this an exempt (*) IR?", paperless_checkbox
  EditBox 70, 95, 95, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 60, 115, 50, 15
    CancelButton 115, 115, 50, 15
  Text 5, 30, 65, 10, "Footer month/year:"
  GroupBox 10, 45, 155, 30, "Programs recertifying"
  Text 10, 100, 60, 10, "Worker Signature"
  GroupBox 10, 140, 155, 75, "Exempt IR checkbox warning:"
  Text 15, 155, 145, 25, "If you select ''Is this an exempt IR'', the case note will only read that the paperless IR was cleared (no case information listed)."
  Text 15, 190, 140, 20, " If you are processing a CSR with SNAP, you should NOT check that option."
  Text 20, 10, 45, 10, "Case number:"
EndDialog

BeginDialog CSR_dialog01, 0, 0, 451, 225, "CSR dialog"
  EditBox 65, 15, 50, 15, CSR_datestamp
  DropListBox 170, 15, 75, 15, "select one..."+chr(9)+"complete"+chr(9)+"incomplete", CSR_status
  EditBox 40, 35, 280, 15, HH_comp
  EditBox 65, 55, 380, 15, earned_income
  EditBox 70, 75, 375, 15, unearned_income
  ButtonGroup ButtonPressed
    PushButton 5, 100, 60, 10, "Notes on Income:", income_notes_button
  EditBox 70, 95, 375, 15, notes_on_income
  EditBox 65, 115, 380, 15, notes_on_abawd
  EditBox 40, 135, 405, 15, assets
  EditBox 60, 155, 95, 15, SHEL_HEST
  EditBox 225, 155, 95, 15, COEX_DCEX
  ButtonGroup ButtonPressed
    PushButton 340, 205, 50, 15, "Next", next_button
    CancelButton 395, 205, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 300, 15, 25, 10, "GRH", ELIG_GRH_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 160, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 160, 25, 10, "HEST:", HEST_button
    PushButton 160, 160, 30, 10, "COEX/", COEX_button
    PushButton 190, 160, 30, 10, "DCEX:", DCEX_button
    PushButton 10, 190, 25, 10, "BUSI", BUSI_button
    PushButton 35, 190, 25, 10, "JOBS", JOBS_button
    PushButton 75, 190, 25, 10, "ACCT", ACCT_button
    PushButton 100, 190, 25, 10, "CARS", CARS_button
    PushButton 125, 190, 25, 10, "CASH", CASH_button
    PushButton 150, 190, 25, 10, "OTHR", OTHR_button
    PushButton 190, 190, 25, 10, "MEMB", MEMB_button
    PushButton 215, 190, 25, 10, "MEMI", MEMI_button
    PushButton 240, 190, 25, 10, "REVW", REVW_button
    PushButton 35, 200, 25, 10, "UNEA", UNEA_button
    PushButton 75, 200, 25, 10, "REST", REST_button
    PushButton 100, 200, 25, 10, "SECU", SECU_button
    PushButton 125, 200, 25, 10, "TRAN", TRAN_button
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"
  Text 5, 20, 55, 10, "CSR datestamp:"
  Text 125, 20, 40, 10, "CSR status:"
  Text 5, 40, 35, 10, "HH comp:"
  Text 5, 60, 55, 10, "Earned income:"
  Text 5, 80, 60, 10, "Unearned income:"
  Text 5, 120, 60, 10, "Notes on WREG:"
  Text 5, 140, 30, 10, "Assets:"
  GroupBox 5, 180, 175, 35, "Income and asset panels"
  GroupBox 185, 180, 85, 25, "other STAT panels:"
  GroupBox 255, 5, 75, 25, "ELIG panels:"
EndDialog

BeginDialog CSR_dialog02, 0, 0, 451, 240, "CSR dialog"
  EditBox 100, 25, 150, 15, FIAT_reasons
  EditBox 50, 45, 395, 15, other_notes
  EditBox 45, 65, 400, 15, changes
  EditBox 60, 85, 385, 15, verifs_needed
  EditBox 60, 105, 385, 15, actions_taken
  CheckBox 190, 155, 110, 10, "Send forms to AREP?", sent_arep_checkbox
  CheckBox 190, 170, 175, 10, "Check here to case note grant info from ELIG/FS.", grab_FS_info_checkbox
  CheckBox 190, 185, 210, 10, "Check here if CSR and cash supplement were used as a HRF.", HRF_checkbox
  CheckBox 190, 200, 120, 10, "Check here if an eDRS was sent.", eDRS_sent_checkbox
  ButtonGroup ButtonPressed
    PushButton 275, 225, 60, 10, "Previous", previous_button
    OkButton 340, 220, 50, 15
    CancelButton 395, 220, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 300, 15, 25, 10, "GRH", ELIG_GRH_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 10, 140, 25, 10, "BUSI", BUSI_button
    PushButton 35, 140, 25, 10, "JOBS", JOBS_button
    PushButton 75, 140, 25, 10, "ACCT", ACCT_button
    PushButton 100, 140, 25, 10, "CARS", CARS_button
    PushButton 125, 140, 25, 10, "CASH", CASH_button
    PushButton 150, 140, 25, 10, "OTHR", OTHR_button
    PushButton 190, 140, 25, 10, "MEMB", MEMB_button
    PushButton 215, 140, 25, 10, "MEMI", MEMI_button
    PushButton 240, 140, 25, 10, "REVW", REVW_button
    PushButton 35, 150, 25, 10, "UNEA", UNEA_button
    PushButton 75, 150, 25, 10, "REST", REST_button
    PushButton 100, 150, 25, 10, "SECU", SECU_button
    PushButton 125, 150, 25, 10, "TRAN", TRAN_button
  EditBox 60, 180, 90, 15, MAEPD_premium
  ButtonGroup ButtonPressed
    PushButton 80, 200, 65, 10, "SIR mail", SIR_mail_button
  Text 5, 30, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 50, 40, 10, "Other notes:"
  Text 5, 70, 35, 10, "Changes?:"
  Text 5, 90, 50, 10, "Verifs needed:"
  Text 5, 110, 50, 10, "Actions taken:"
  GroupBox 5, 130, 175, 35, "Income and asset panels"
  GroupBox 185, 130, 85, 25, "other STAT panels:"
  GroupBox 5, 170, 150, 45, "If MA-EPD..."
  Text 10, 185, 50, 10, "New premium:"
  GroupBox 255, 5, 75, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"
  CheckBox 10, 200, 65, 10, "Emailed MADE?", MADE_checkbox
EndDialog

BeginDialog income_notes_dialog, 0, 0, 351, 185, "Explanation of Income"
  CheckBox 10, 30, 325, 10, "JOBS - Client has confirmed that JOBS income is expected to continue at this rate and hours.", jobs_anticipated_checkbox
  CheckBox 10, 45, 330, 10, "JOBS - This is a new job and actual check stubs are not available, advised client that if actual pay", new_jobs_checkbox
  CheckBox 10, 70, 325, 10, "BUSI - Client has confirmed that BUSI income is expected to continue at this rate and hours.", busi_anticipated_checkbox
  CheckBox 10, 85, 250, 10, "BUSI - Client has agreed to the self-employment budgeting method used.", busi_method_agree_checkbox
  CheckBox 10, 100, 325, 10, "RBIC - Client has confirmed that RBIC income is expected to continue at this rate and hours.", rbic_anticipated_checkbox
  CheckBox 10, 115, 325, 10, "UNEA - Client has confirmed that UNEA income is expected to continue at this rate and hours.", unea_anticipated_checkbox
  CheckBox 10, 130, 315, 10, "UNEA - Client has applied for unemployment benefits but no determination made at this time.", ui_pending_checkbox
  CheckBox 45, 140, 225, 10, "Check here to have the script set a TIKL to check UI in two weeks.", tikl_for_ui
  CheckBox 10, 155, 150, 10, "NONE - This case has no income reported.", no_income_checkbox
  ButtonGroup ButtonPressed
    PushButton 240, 165, 50, 15, "Insert", add_to_notes_button
    CancelButton 295, 165, 50, 15
  Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
  Text 45, 55, 315, 10, "varies significantly, client should provide proof of this difference to have benefits adjusted."
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Showing the case number dialog
Do
	DO
		err_msg = ""
		Dialog case_number_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "* You need to type a valid case number."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_checkbox = 1 then
	call start_a_blank_CASE_NOTE
	Call write_variable_in_case_note("***Cleared paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "***")
	Call write_variable_in_case_note("---")
	Call write_variable_in_case_note(worker_signature)
	call script_end_procedure("")
End if

'confirms that footer month/year from dialog matches footer month/year on MAXIS
Call MAXIS_footer_month_confirmation

'Navigating to STAT/REVW, checking for error prone cases
call navigate_to_MAXIS_screen("stat", "revw")
'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabbing SHEL/HEST first, and putting them in this special order that everyone seems to like
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
'If SHEL_HEST <> "" then SHEL_HEST = SHEL_HEST & "; "		'this is a temporary fix to resolve issues where a variable is "autofilled" by multiple functions in the same script
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)
'Autofilling HH comp
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
'Autofilling WREG status
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'Autofilling assets
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CSR_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'-----------------Creating text for case note
'Programs recertifying case noting info into variable
If GRH_checkbox = 1 then programs_recertifying = programs_recertifying & "GRH, "
If HC_checkbox = 1 then programs_recertifying = programs_recertifying & "HC, "
If SNAP_checkbox = 1 then programs_recertifying = programs_recertifying & "SNAP, "
programs_recertifying = trim(programs_recertifying)
if right(programs_recertifying, 1) = "," then programs_recertifying = left(programs_recertifying, len(programs_recertifying) - 1)

'Determining the CSR month for header
CSR_month = MAXIS_footer_month & "/" & MAXIS_footer_year

'Showing the case note dialog
DO
	DO
		Do
			err_msg = ""
			Do
				Do
					Dialog CSR_dialog01
					cancel_confirmation
					If ButtonPressed = SIR_mail_button then run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
					'If next_button = pressed THEN msgbox next_button
					If ButtonPressed = income_notes_button Then
						Dialog income_notes_dialog
						If ButtonPressed = add_to_notes_button Then
							If jobs_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from jobs to continue at this amount."
							If new_jobs_checkbox = checked Then notes_on_income = notes_on_income & "; This is a new job and actual check stubs have not been received, advised client to provide proof once pay is received if the income received differs significantly."
							If busi_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from self employment to continue at this amount."
							If busi_method_agree_checkbox = checked Then notes_on_income = notes_on_income & "; Explained to client the self employment budgeting methods and client agreed to the method used."
							If rbic_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects roomer/boarder income to continue at this amount."
							If unea_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects unearned income to continue at this amount."
							If ui_pending_checkbox = checked Then notes_on_income = notes_on_income & "; Client has applied for Unemployment Income recently but request is still pending, will need to be reviewed soon for changes."
							If tikl_for_ui = checked Then notes_on_income = notes_on_income & " TIKL set to request an update on Unemployment Income."
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; Client has reported they have no income and do not expect any changes to this at this time."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
				Loop until ButtonPressed <> no_cancel_button
				MAXIS_dialog_navigation
			LOOP until ButtonPressed = next_button
			IF CSR_datestamp = "" THEN 					err_msg = err_msg & vbCr & "* Please enter the date the CSR was received."
			IF CSR_status = "select one..." THEN 				err_msg = err_msg & vbCr & "* Please select the status of the CSR."
			IF HH_comp = "" THEN 						err_msg = err_msg & vbCr & "* Please enter household composition information."
			IF (earned_income <> "" AND notes_on_income = "") OR (unearned_income <> "" AND notes_on_income = "") THEN 					err_msg = err_msg & vbCr & "* You must provide some information about income. Please complete the 'Notes on Income' field."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		DO
			DO
				DO
					Dialog CSR_dialog02
					cancel_confirmation
					IF ButtonPressed = SIR_mail_button THEN run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
				LOOP UNTIL ButtonPressed <> no_cancel_button
				MAXIS_dialog_navigation
			LOOP UNTIL ButtonPressed = -1 OR ButtonPressed = previous_button
			err_msg = ""
			IF actions_taken = "" THEN 		err_msg = err_msg & vbCr & "* Please indicate the actions you have taken."
			IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = "" OR ButtonPressed = previous_button
	LOOP WHILE ButtonPressed = previous_button
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

IF grab_FS_info_checkbox = 1 THEN
	'grabbing information about elig/fs
	call navigate_to_MAXIS_screen("elig", "fs")
	EMReadScreen FSPR_check, 4, 3, 48
	If FSPR_check <> "FSPR" then
		MsgBox "The script couldn't find ELIG/FS. It will now jump to case note."
	Else
		EMWriteScreen "FSSM", 19, 70
		transmit
		EMReadScreen FSSM_line_01, 37, 13, 44
		EMReadScreen FSSM_line_02, 37, 8, 3
		EMReadScreen FSSM_line_03, 37, 10, 3
	End if
END IF

IF tikl_for_ui THEN
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	two_weeks_from_now = DateAdd("d", 14, date)
	call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
	call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
	PF3
END IF

'Writing the case note to MAXIS----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
call write_variable_in_case_note("***" & CSR_month & " CSR received " & CSR_datestamp & ": " & CSR_status & "***")
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
IF move_verifs_needed = TRUE THEN CALL write_variable_in_case_note("---")                               	                'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll add a line separator.
call write_bullet_and_variable_in_case_note("Programs recertifying", programs_recertifying)
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
call write_bullet_and_variable_in_case_note("Notes on Income", notes_on_income)
call write_bullet_and_variable_in_case_note("ABAWD Notes", notes_on_abawd)
call write_bullet_and_variable_in_case_note("Assets", assets)
call write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL_HEST)
call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Changes", changes)
If HRF_checkbox = checked then call write_variable_in_case_note("* CSR and cash supplement used as HRF.")
If eDRS_sent_checkbox = checked then call write_variable_in_case_note("* eDRS sent.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_bullet_and_variable_in_case_note("MA-EPD premium", MAEPD_premium)
If MADE_checkbox = checked then call write_variable_in_case_note("* Emailed MADE through DHS-SIR.")
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
call write_variable_in_case_note("---")
If grab_FS_info_checkbox = 1 AND FSPR_check = "FSPR" then
	call write_variable_in_case_note("   " & FSSM_line_01)
	call write_variable_in_case_note("   " & FSSM_line_02)
	call write_variable_in_case_note("   " & FSSM_line_03)
	call write_variable_in_case_note("---")
End if
call write_variable_in_case_note(worker_signature)

call script_end_procedure("")
