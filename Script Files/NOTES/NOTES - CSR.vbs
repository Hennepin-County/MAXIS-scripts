'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CSR.vbs"
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

next_month = dateadd("m", + 1, date)

footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog case_number_dialog, 0, 0, 181, 115, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 95, 60, 30, 10, "HC", HC_checkbox
  CheckBox 10, 80, 100, 10, "Is this an exempt (*) IR?", paperless_checkbox
  ButtonGroup ButtonPressed
    OkButton 35, 95, 50, 15
    CancelButton 95, 95, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs recertifying"
EndDialog

BeginDialog CSR_dialog, 0, 0, 451, 350, "CSR dialog"
  EditBox 65, 15, 50, 15, CSR_datestamp
  DropListBox 170, 15, 75, 15, "select one..."+chr(9)+"complete"+chr(9)+"incomplete", CSR_status
  EditBox 40, 35, 280, 15, HH_comp
  EditBox 65, 55, 380, 15, earned_income
  EditBox 70, 75, 375, 15, unearned_income
  EditBox 65, 95, 380, 15, notes_on_abawd
  EditBox 40, 115, 405, 15, assets
  EditBox 60, 135, 95, 15, SHEL_HEST
  EditBox 220, 135, 95, 15, COEX_DCEX
  EditBox 100, 155, 345, 15, FIAT_reasons
  EditBox 50, 175, 395, 15, other_notes '11
  EditBox 45, 195, 400, 15, changes
  EditBox 60, 215, 385, 15, verifs_needed
  EditBox 60, 235, 385, 15, actions_taken
  EditBox 380, 255, 65, 15, worker_signature
  CheckBox 190, 285, 110, 10, "Send forms to AREP?", sent_arep_checkbox
  CheckBox 190, 300, 175, 10, "Check here to case note grant info from ELIG/FS.", grab_FS_info_checkbox
  CheckBox 190, 315, 210, 10, "Check here if CSR and cash supplement were used as a HRF.", HRF_checkbox
  CheckBox 190, 330, 120, 10, "Check here if an eDRS was sent.", eDRS_sent_checkbox
  EditBox 60, 310, 90, 15, MAEPD_premium
  CheckBox 10, 330, 65, 10, "Emailed MADE?", MADE_checkbox '21
  ButtonGroup ButtonPressed
    OkButton 340, 275, 50, 15
    CancelButton 395, 275, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 140, 25, 10, "SHEL/", SHEL_button '31
    PushButton 30, 140, 25, 10, "HEST:", HEST_button
    PushButton 165, 140, 25, 10, "COEX/", COEX_button
    PushButton 190, 140, 25, 10, "DCEX:", DCEX_button
    PushButton 10, 270, 25, 10, "BUSI", BUSI_button
    PushButton 35, 270, 25, 10, "JOBS", JOBS_button
    PushButton 35, 280, 25, 10, "UNEA", UNEA_button
    PushButton 75, 270, 25, 10, "ACCT", ACCT_button
    PushButton 100, 270, 25, 10, "CARS", CARS_button
    PushButton 125, 270, 25, 10, "CASH", CASH_button '41
    PushButton 150, 270, 25, 10, "OTHR", OTHR_button
    PushButton 75, 280, 25, 10, "REST", REST_button
    PushButton 100, 280, 25, 10, "SECU", SECU_button
    PushButton 125, 280, 25, 10, "TRAN", TRAN_button
    PushButton 190, 270, 25, 10, "MEMB", MEMB_button
    PushButton 215, 270, 25, 10, "MEMI", MEMI_button
    PushButton 240, 270, 25, 10, "REVW", REVW_button
    PushButton 80, 330, 65, 10, "SIR mail", SIR_mail_button
  GroupBox 255, 5, 50, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"   '51
  Text 5, 20, 55, 10, "CSR datestamp:"
  Text 125, 20, 40, 10, "CSR status:"
  Text 5, 40, 35, 10, "HH comp:"
  Text 5, 60, 55, 10, "Earned income:"
  Text 5, 80, 60, 10, "Unearned income:"
  Text 5, 100, 60, 10, "Notes on WREG:"
  Text 5, 120, 30, 10, "Assets:"
  Text 5, 160, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 180, 40, 10, "Other notes:"
  Text 5, 200, 35, 10, "Changes?:"
  Text 5, 220, 50, 10, "Verifs needed:" '61
  Text 5, 240, 50, 10, "Actions taken:"
  GroupBox 5, 260, 175, 35, "Income and asset panels"
  GroupBox 185, 260, 85, 25, "other STAT panels:"
  Text 650, 650, 650, 650, spooky 'getting past 65 item limit
  Text 315, 260, 65, 10, "Worker signature:"
  GroupBox 5, 300, 150, 45, "If MA-EPD..."
  Text 10, 315, 50, 10, "New premium:"
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to MAXIS
EMConnect ""

'Searching for the case_number variable
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Searching for the footer month
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
	footer_month = MAXIS_footer_month
	call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
	If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number dialog
Do
	Dialog case_number_dialog
	If ButtonPressed = 0 then stopscript
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_checkbox = 1 then
	call navigate_to_screen("case", "note")
	PF9
	EMWriteScreen "***Cleared paperless IR for " & footer_month & "/" & footer_year & "***", 4, 3
	EMWriteScreen "---", 5, 3
	worker_signature = InputBox ("Sign your case note:", "worker signature")
	EMWriteScreen worker_signature, 6, 3
	call script_end_procedure("")
End if

'Navigating to STAT/REVW, checking for error prone cases
call navigate_to_screen("stat", "revw")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get into STAT. This case may be in background. Wait a few seconds and try again. If this case is not in background email your script administrator the case number and footer month.")

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

'Autofill DCEX/COEX
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)

'Autofill EI
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)

'Autofill datestamp and UI
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CSR_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'-----------------Creating text for case note
'Programs recertifying case noting info into variable
If cash_checkbox = 1 then programs_recertifying = programs_recertifying & "cash, "
If HC_checkbox = 1 then programs_recertifying = programs_recertifying & "HC, "
If SNAP_checkbox = 1 then programs_recertifying = programs_recertifying & "SNAP, "
programs_recertifying = trim(programs_recertifying)
if right(programs_recertifying, 1) = "," then programs_recertifying = left(programs_recertifying, len(programs_recertifying) - 1)

'Determining the CSR month for header
CSR_month = footer_month & "/" & footer_year

'Showing the case note dialog
Do
	Do
		Do
			Do
				Do
					Dialog CSR_dialog
					If ButtonPressed = 0 then cancel_confirmation
					If ButtonPressed = SIR_mail_button then run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
				Loop until ButtonPressed <> no_cancel_button
				EMReadScreen STAT_check, 4, 20, 21
				If STAT_check = "STAT" then
					If ButtonPressed = prev_panel_button then call panel_navigation_prev
					If ButtonPressed = next_panel_button then call panel_navigation_next
					If ButtonPressed = prev_memb_button then call memb_navigation_prev
					If ButtonPressed = next_memb_button then call memb_navigation_next
				End if
				transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
				EMReadScreen MAXIS_check, 5, 1, 39
				If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
			Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
			If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
			If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
			If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
			If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
			If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
			If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
			If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
			If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
			If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
			If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
			If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
			If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
			If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
			If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
			If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
			If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
			If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
			If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
			If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
			If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
			If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
			If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
		Loop until ButtonPressed = -1
		If (earned_income = "" and unearned_income = "") or actions_taken = "" or CSR_datestamp = "" or worker_signature = "" or CSR_status = "select one..." then MsgBox "You need to fill in the datestamp, income, CSR status, and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
	Loop until (earned_income <> "" or unearned_income <> "") and actions_taken <> "" and CSR_datestamp <> "" and worker_signature <> "" and CSR_status <> "select one..."
	CALL proceed_confirmation(go_to_case_note)	'Asking the worker if they want to proceed to case note.
	If go_to_case_note = True THEN
		If grab_FS_info_checkbox = 1 then
			call navigate_to_screen("elig", "fs")
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
		End if
		call navigate_to_screen("case", "note")
		PF9
		EMReadScreen case_note_check, 17, 2, 33
		EMReadScreen mode_check, 1, 20, 09
		If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
	End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'Writing the case note to MAXIS
call write_variable_in_case_note("***" & CSR_month & " CSR received " & CSR_datestamp & ": " & CSR_status & "***")
call write_bullet_and_variable_in_case_note("Programs recertifying", programs_recertifying)
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
If earned_income <> "" then call write_bullet_and_variable_in_case_note("Earned income", earned_income)
If unearned_income <> "" then call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
If notes_on_abawd <> "" then call write_bullet_and_variable_in_case_note("ABAWD Notes", notes_on_abawd)
If assets <> "" then call write_bullet_and_variable_in_case_note("Assets", assets)
If SHEL_HEST <> "" then call write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL_HEST)
If COEX_DCEX <> "" then call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
if FIAT_reasons <> "" then call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
if other_notes <> "" then call write_bullet_and_variable_in_case_note("Other notes", other_notes)
If changes <> "" then call write_bullet_and_variable_in_case_note("Changes", changes)
If HRF_checkbox = checked then call write_variable_in_case_note("* CSR and cash supplement used as HRF.")
If eDRS_sent_checkbox = checked then call write_variable_in_case_note("* eDRS sent.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
if verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
If MAEPD_premium <> "" then call write_bullet_and_variable_in_case_note("MA-EPD premium", MAEPD_premium)
If MADE_checkbox = checked then call write_variable_in_case_note("* Emailed MADE through DHS-SIR.")
call write_variable_in_case_note("---")
If FSPR_check = "FSPR" then
	call write_variable_in_case_note("   " & FSSM_line_01)
	call write_variable_in_case_note("   " & FSSM_line_02)
	call write_variable_in_case_note("   " & FSSM_line_03)
	call write_variable_in_case_note("---")
End if
call write_variable_in_case_note(worker_signature)

call script_end_procedure("")

