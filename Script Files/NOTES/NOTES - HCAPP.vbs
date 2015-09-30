'Grabbing stats----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HCAPP.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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


'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

footer_month = datepart("m", date)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = right(datepart("yyyy", next_month), 2)

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog


BeginDialog HCAPP_dialog_01, 0, 0, 446, 300, "HCAPP dialog part 1"
  EditBox 75, 5, 50, 15, HCAPP_datestamp
  EditBox 45, 25, 250, 15, HH_comp
  EditBox 45, 45, 225, 15, cit_id
  EditBox 330, 45, 115, 15, AREP
  EditBox 85, 65, 90, 15, SCHL
  EditBox 235, 65, 210, 15, DISA
  EditBox 35, 85, 160, 15, PREG
  EditBox 235, 85, 210, 15, ABPS
  EditBox 60, 105, 180, 15, retro_request
  EditBox 60, 140, 385, 15, earned_income
  EditBox 70, 160, 375, 15, unearned_income
  EditBox 35, 180, 245, 15, STWK
  EditBox 350, 180, 95, 15, COEX_DCEX
  EditBox 65, 200, 380, 15, notes_on_income
  EditBox 155, 220, 290, 15, is_any_work_temporary
  EditBox 55, 255, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
    PushButton 335, 280, 50, 15, "NEXT", next_page_button
    CancelButton 390, 280, 50, 15
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 275, 50, 25, 10, "AREP/", AREP_button
    PushButton 300, 50, 25, 10, "ALTP:", ALTP_button
    PushButton 5, 70, 25, 10, "SCHL/", SCHL_button
    PushButton 30, 70, 25, 10, "STIN/", STIN_button
    PushButton 55, 70, 25, 10, "STEC:", STEC_button
    PushButton 180, 70, 25, 10, "DISA/", DISA_button
    PushButton 205, 70, 25, 10, "PDED:", PDED_button
    PushButton 5, 90, 25, 10, "PREG:", PREG_button
    PushButton 205, 90, 25, 10, "ABPS:", ABPS_button
    PushButton 5, 185, 25, 10, "STWK:", STWK_button
    PushButton 295, 185, 25, 10, "COEX/", COEX_button
    PushButton 320, 185, 25, 10, "DCEX:", DCEX_button
    PushButton 10, 285, 25, 10, "BUSI", BUSI_button
    PushButton 35, 285, 25, 10, "JOBS", JOBS_button
    PushButton 60, 285, 25, 10, "RBIC", RBIC_button
    PushButton 85, 285, 25, 10, "UNEA", UNEA_button
    PushButton 125, 285, 25, 10, "MEMB", MEMB_button
    PushButton 150, 285, 25, 10, "MEMI", MEMI_button
    PushButton 175, 285, 25, 10, "REVW", REVW_button
    PushButton 215, 285, 35, 10, "ELIG/HC", ELIG_HC_button
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 10, 65, 10, "HCAPP datestamp:"
  Text 5, 30, 35, 10, "HH comp:"
  Text 5, 50, 40, 10, "Cit/ID/imig:"
  Text 5, 110, 50, 10, "Retro request:"
  Text 5, 145, 55, 10, "Earned income:"
  Text 5, 165, 65, 10, "Unearned income:"
  Text 5, 205, 60, 10, "Notes on income:"
  Text 5, 225, 150, 10, "Is any work temporary? If so, explain details:"
  Text 5, 260, 50, 10, "Verifs needed:"
  GroupBox 5, 275, 110, 25, "Income panels"
  GroupBox 120, 275, 85, 25, "other STAT panels:"
EndDialog

BeginDialog HCAPP_dialog_02, 0, 0, 451, 325, "HCAPP dialog part 2"
  EditBox 35, 50, 410, 15, assets
  EditBox 60, 80, 385, 15, INSA
  EditBox 35, 100, 410, 15, ACCI
  EditBox 35, 120, 410, 15, BILS
  EditBox 125, 140, 125, 15, FACI
  CheckBox 255, 145, 80, 10, "Application signed?", application_signed_check
  CheckBox 350, 145, 65, 10, "MMIS updated?", MMIS_updated_check
  CheckBox 20, 160, 115, 10, "Sent forms to AREP?", sent_arep_checkbox
  CheckBox 20, 175, 290, 10, "Check here to have the script update PND2 to show client delay (pending cases only).", client_delay_check
  CheckBox 20, 190, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", TIKL_check
  EditBox 100, 205, 345, 15, FIAT_reasons
  EditBox 55, 225, 215, 15, other_notes
  DropListBox 330, 225, 115, 15, "Select one..."+chr(9)+"incomplete"+chr(9)+"approved"+chr(9)+"denied", HCAPP_status
  EditBox 55, 245, 390, 15, verifs_needed
  EditBox 55, 265, 390, 15, actions_taken
  EditBox 395, 285, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 305, 50, 15
    CancelButton 395, 305, 50, 15
    PushButton 10, 15, 25, 10, "ACCT", ACCT_button
    PushButton 35, 15, 25, 10, "CARS", CARS_button
    PushButton 60, 15, 25, 10, "CASH", CASH_button
    PushButton 85, 15, 25, 10, "OTHR", OTHR_button
    PushButton 10, 25, 25, 10, "REST", REST_button
    PushButton 35, 25, 25, 10, "SECU", SECU_button
    PushButton 60, 25, 25, 10, "TRAN", TRAN_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 85, 25, 10, "INSA/", INSA_button
    PushButton 30, 85, 25, 10, "MEDI:", MEDI_button
    PushButton 5, 105, 25, 10, "ACCI:", ACCI_button
    PushButton 5, 125, 25, 10, "BILS:", BILS_button
    PushButton 5, 145, 25, 10, "FACI/", FACI_button
    PushButton 10, 295, 25, 10, "MEMB", MEMB_button
    PushButton 35, 295, 25, 10, "MEMI", MEMI_button
    PushButton 60, 295, 25, 10, "REVW", REVW_button
    PushButton 95, 295, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 225, 310, 75, 10, "previous page", previous_page_button
  GroupBox 5, 5, 110, 35, "Asset panels"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 55, 30, 10, "Assets:"
  Text 35, 145, 90, 10, "residency/miscellaneous:"
  Text 5, 210, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 230, 45, 10, "Other notes:"
  Text 280, 230, 50, 10, "HCAPP status:"
  Text 5, 250, 50, 10, "Verifs needed:"
  Text 5, 270, 50, 10, "Actions taken:"
  GroupBox 5, 285, 85, 25, "other STAT panels:"
  Text 330, 290, 65, 10, "Worker signature:"
EndDialog


BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog

BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.
application_signed_check = 1 'The script should default to having the application signed.


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to MAXIS
EMConnect ""

'Grabbing the case number
call MAXIS_case_number_finder(case_number)

'Grabbing the footer month/year
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
	footer_month = MAXIS_footer_month
	call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
	If row <> 0 then footer_year = MAXIS_footer_year
End if

footer_month = CStr(footer_month)

'Showing the case number
Do
	Dialog case_number_and_footer_month_dialog
	If ButtonPressed = 0 then stopscript
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
transmit

'Checking to see that we're in MAXIS
call check_for_MAXIS(True)

'Navigating to STAT, grabbing the HH members
call navigate_to_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact a Support Team member.")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling case info
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", HCAPP_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", STWK)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)


'SECTION 07: CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Do
	Do
		Do
			Do
				err_msg = ""
				Dialog HCAPP_dialog_01
				cancel_confirmation
				If HCAPP_datestamp = "" or len(HCAPP_datestamp) > 10 THEN err_msg = "Please enter a valid application datestamp."  'creating err_msg if required items are missing
				If err_msg <> "" THEN Msgbox err_msg
			Loop until ButtonPressed <> no_cancel_button and err_msg = ""
			MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
		Loop until ButtonPressed = next_page_button
		Do
			Do
				err_msg = ""
				Dialog HCAPP_dialog_02
				cancel_confirmation
				If actions_taken = "" THEN err_msg = err_msg & vbCr & "Please complete actions taken section."    'creating err_msg if required items are missing
				If worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature."
				If HCAPP_status = "Select one..." THEN err_msg = err_msg & vbCr & "Please select a CAF Status."
				If err_msg <> "" THEN Msgbox err_msg
			Loop until ButtonPressed <> no_cancel_button and err_msg = ""
			MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
		Loop until ButtonPressed = -1 or ButtonPressed = previous_page_button
		If ButtonPressed = previous_page_button then exit do
	Loop until err_msg = "" 
	If ButtonPressed = -1 then dialog case_note_dialog
	If buttonpressed = yes_case_note_button then exit do
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'PND2 updater------------------------------------------------------------------------------------------------------------------------
If client_delay_check = 1 then 'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
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
'Tikl portion-----------------------------------------------------------------------------------------------------------------------------
If TIKL_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(HCAPP_datestamp, 45, 5, 18) 
	EMSetCursor 9, 3
	EMSendKey "HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out."
	transmit
	PF3
End if

'SECTION 08: THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE

CALL write_variable_in_case_note("***HCAPP received " & HCAPP_datestamp & ": " & HCAPP_status & "***")
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("SCHL/STIN/STEC", SCHL)
CALL write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
CALL write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro request", retro_request)
CALL write_bullet_and_variable_in_CASE_NOTE("ABPS", ABPS)
CALL write_bullet_and_variable_in_CASE_NOTE("EI", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("STWK", STWK)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes on income", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary?", is_any_work_temporary)
CALL write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
CALL write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
CALL write_bullet_and_variable_in_CASE_NOTE("ACCI", ACCI)
CALL write_bullet_and_variable_in_CASE_NOTE("BILS", BILS)
CALL write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
IF application_signed_check = checked THEN 
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
ELSE
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF client_delay_check = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
IF sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
CALL write_bullet_and_variable_in_CASE_NOTE("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
IF MMIS_updated_check = checked then call write_variable_in_case_note("* MMIS updated.")
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")
