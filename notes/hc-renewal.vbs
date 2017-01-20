'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HC RENEWAL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 540          'manual run time in seconds
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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, MAXIS_footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog

BeginDialog HC_ER_dialog, 0, 0, 456, 300, "HC ER dialog"
  EditBox 75, 50, 50, 15, recert_datestamp
  DropListBox 185, 50, 75, 15, "(select one...)"+chr(9)+"complete"+chr(9)+"incomplete", recert_status
  EditBox 325, 50, 125, 15, HH_comp
  EditBox 60, 70, 390, 15, earned_income
  EditBox 70, 90, 380, 15, unearned_income
  EditBox 40, 110, 410, 15, assets
  EditBox 60, 130, 95, 15, COEX_DCEX
  EditBox 265, 130, 185, 15, CEI_availability
  EditBox 100, 150, 350, 15, FIAT_reasons
  EditBox 50, 170, 400, 15, other_notes
  EditBox 45, 190, 405, 15, changes
  EditBox 60, 210, 390, 15, verifs_needed
  EditBox 55, 230, 395, 15, actions_taken
  EditBox 60, 260, 90, 15, MAEPD_premium
  CheckBox 10, 280, 65, 10, "Emailed MADE?", MADE_check
  ButtonGroup ButtonPressed
    PushButton 85, 280, 65, 10, "SIR mail", SIR_mail_button
  CheckBox 175, 255, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
  CheckBox 175, 270, 85, 10, "MMIS updated?", MMIS_updated_checkbox
  EditBox 400, 250, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 345, 270, 50, 15
    CancelButton 400, 270, 50, 15
    PushButton 10, 20, 25, 10, "BUSI", BUSI_button
    PushButton 35, 20, 25, 10, "JOBS", JOBS_button
    PushButton 10, 30, 25, 10, "RBIC", RBIC_button
    PushButton 35, 30, 25, 10, "UNEA", UNEA_button
    PushButton 75, 20, 25, 10, "ACCT", ACCT_button
    PushButton 100, 20, 25, 10, "CARS", CARS_button
    PushButton 125, 20, 25, 10, "CASH", CASH_button
    PushButton 150, 20, 25, 10, "OTHR", OTHR_button
    PushButton 75, 30, 25, 10, "REST", REST_button
    PushButton 100, 30, 25, 10, "SECU", SECU_button
    PushButton 125, 30, 25, 10, "TRAN", TRAN_button
    PushButton 190, 20, 25, 10, "MEMB", MEMB_button
    PushButton 215, 20, 25, 10, "MEMI", MEMI_button
    PushButton 240, 20, 25, 10, "REVW", REVW_button
    PushButton 285, 20, 35, 10, "HC", ELIG_HC_button
    PushButton 340, 20, 45, 10, "prev. panel", prev_panel_button
    PushButton 400, 20, 45, 10, "prev. memb", prev_memb_button
    PushButton 340, 30, 45, 10, "next panel", next_panel_button
    PushButton 400, 30, 45, 10, "next memb", next_memb_button
    PushButton 5, 135, 25, 10, "COEX/", COEX_button
    PushButton 30, 135, 25, 10, "DCEX:", DCEX_button
  GroupBox 275, 5, 55, 30, "ELIG panels:"
  GroupBox 335, 5, 115, 40, "STAT-based navigation"
  Text 5, 55, 65, 10, "Recert datestamp:"
  Text 135, 55, 50, 10, "Recert status:"
  Text 280, 55, 35, 10, "HH comp:"
  Text 5, 75, 55, 10, "Earned income:"
  Text 5, 95, 65, 10, "Unearned income:"
  Text 5, 115, 30, 10, "Assets:"
  Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 175, 45, 10, "Other notes:"
  Text 5, 195, 35, 10, "Changes?:"
  Text 5, 215, 50, 10, "Verifs needed:"
  Text 5, 235, 50, 10, "Actions taken:"
  GroupBox 5, 250, 150, 45, "If MA-EPD..."
  Text 10, 265, 50, 10, "New premium:"
  GroupBox 70, 5, 110, 40, "Asset panels"
  Text 335, 255, 65, 10, "Worker signature:"
  GroupBox 5, 5, 60, 40, "Income panels"
  GroupBox 185, 5, 85, 30, "other STAT panels:"
  Text 165, 135, 100, 10, "Cost-effective insa availablity:"
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Grabbing the case number
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog 	
Do		
	err_msg = ""	
	Do	
  		Dialog case_number_and_footer_month_dialog    'initial dialog	
  		If ButtonPressed = 0 then Stopscript    'if cancel is pressed then the script ends	
  		Call check_for_password(are_we_passworded_out)    'function to see if users is password-ed out	
	Loop until are_we_passworded_out = false  	'will loop until user is password-ed back in
	If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."	
	If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."	
	If IsNumeric(MAXIS_case_number) = False or Len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* You must enter a valid case number."	
  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		
LOOP until err_msg = ""		

'initial navigation
Call MAXIS_footer_month_confirmation	'confirming that the footer month/year in the MAXIS panel and the dialog box selected by the user are the same'
'Navigating to STAT, checks for abended cases
call navigate_to_MAXIS_screen("stat", "memb")
'Creating a custom dialog for determining who the HH members are
CALL HH_member_custom_dialog(HH_member_array)

'Autofilling info from STAT
CALL autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
CALL autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
CALL autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
CALL autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
CALL autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
CALL autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
CALL autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "REVW", recert_datestamp)
CALL autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
CALL autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Creating variable for recert_month
recert_month = MAXIS_footer_month & "/" & MAXIS_footer_year

'Showing case note dialog, with navigation and required answers logic
DO
	Do
		Do
			err_msg = ""
			Dialog HC_ER_dialog				'Displays the dialog
			cancel_confirmation				'Asks if we are sure we want to cancel if the cancel button is pressed
			MAXIS_dialog_navigation			'Custom function which contains all of the MAXIS dialog navigation possibilities
			If ButtonPressed = SIR_mail_button then run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"		'Goes to SIR if button is pressed
		Loop until ButtonPressed = -1 		'Loops until OK is selected
		If recert_datestamp = "" or IsDate(recert_datestamp) = False then err_msg = err_msg & vbNewLine & "You need to fill in the datestamp."
		If recert_status = "(select one...)" then err_msg = err_msg & vbNewLine & "* You need to select a recert status."
		If actions_taken = "" then err_msg = err_msg & vbNewLine & "You need to complete the actions taken field."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false	

'The case note----------------------------------------------------------------------------------------------------
call start_a_blank_case_note
CALL write_variable_in_case_note("***" & recert_month & " HC ER received " & recert_datestamp & ": " & recert_status & "***")
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
call write_bullet_and_variable_in_case_note("Assets", assets)
call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
call write_bullet_and_variable_in_case_note("Cost-effective insurance availability",CEI_availability)
call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Changes", changes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
If MMIS_updated_checkbox = 1 then Call write_variable_in_case_note("* MMIS updated.")
call write_variable_in_case_note("---")
call write_bullet_and_variable_in_case_note("MA-EPD premium", MAEPD_premium)
If MADE_check = checked then call write_variable_in_case_note("* Emailed MADE.")
If MAEPD_premium <> "" or MADE_check = checked then call write_variable_in_case_note("---")		'Does this for MAEPD <> blank because if it's blank and there's no MADE_check, it means there's nothing in this section after the ---, and we don't want two in a row now, do we?
call write_variable_in_case_note(worker_signature)

call script_end_procedure("Success! Remember to update MMIS.")
