'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - COMBINED AR.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 540          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", 1, date)
MAXIS_footer_month = datepart("m", next_month)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = datepart("yyyy", next_month)
MAXIS_footer_year = "" & MAXIS_footer_year - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 100, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 80, 25, 30, 15, MAXIS_footer_month
  EditBox 120, 25, 30, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "GRH", GRH_check
  CheckBox 50, 60, 30, 10, "MSA", cash_check
  CheckBox 95, 60, 35, 10, "SNAP", SNAP_check
  CheckBox 145, 60, 30, 10, "HC", HC_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 95, 80, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 65, 10, "Footer month/year:"
  GroupBox 5, 45, 170, 30, "Programs recertifying"
EndDialog

BeginDialog Combined_AR_dialog, 0, 0, 441, 335, "Combined AR dialog"
  EditBox 70, 35, 50, 15, recert_datestamp
  EditBox 200, 35, 40, 15, recert_month
  EditBox 60, 55, 50, 15, interview_date
  EditBox 155, 55, 275, 15, HH_comp
  EditBox 50, 75, 380, 15, US_citizen
  EditBox 35, 100, 210, 15, AREP
  EditBox 35, 155, 400, 15, income
  EditBox 35, 175, 400, 15, assets
  EditBox 65, 195, 370, 15, SHEL
  EditBox 100, 215, 335, 15, FIAT_reasons
  EditBox 60, 235, 375, 15, verifs_needed
  EditBox 55, 255, 380, 15, actions_taken
  EditBox 50, 275, 385, 15, other_notes
  CheckBox 5, 300, 65, 10, "R/R explained?", R_R_explained
  CheckBox 80, 300, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
  CheckBox 80, 315, 60, 10, "eDRS checked", eDRS_checked
  DropListBox 230, 295, 60, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete"+chr(9)+"closed", review_status
  EditBox 370, 295, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 330, 315, 50, 15
    CancelButton 385, 315, 50, 15
    PushButton 45, 15, 25, 10, "MEMB", MEMB_button
    PushButton 70, 15, 25, 10, "MEMI", MEMI_button
    PushButton 95, 15, 25, 10, "REVW", REVW_button
    PushButton 185, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 205, 15, 20, 10, "GA", ELIG_GA_button
    PushButton 225, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 245, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 390, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 390, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 100, 25, 10, "AREP:", AREP_button
    PushButton 10, 130, 25, 10, "BUSI", BUSI_button
    PushButton 35, 130, 25, 10, "JOBS", JOBS_button
    PushButton 60, 130, 25, 10, "RBIC", RBIC_button
    PushButton 85, 130, 25, 10, "UNEA", UNEA_button
    PushButton 125, 130, 25, 10, "ACCT", ACCT_button
    PushButton 150, 130, 25, 10, "CARS", CARS_button
    PushButton 175, 130, 25, 10, "CASH", CASH_button
    PushButton 200, 130, 25, 10, "OTHR", OTHR_button
    PushButton 225, 130, 25, 10, "REST", REST_button
    PushButton 250, 130, 25, 10, "SECU", SECU_button
    PushButton 275, 130, 25, 10, "TRAN", TRAN_button
    PushButton 5, 200, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 200, 25, 10, "HEST:", HEST_button
  Text 5, 40, 65, 10, "Recert datestamp:"
  Text 130, 40, 70, 10, "Recert footer month:"
  Text 115, 60, 40, 10, "HH Comp:"
  Text 5, 80, 40, 10, "US citizen?:"
  GroupBox 5, 120, 110, 25, "Income panels"
  GroupBox 120, 120, 185, 25, "Asset panels"
  Text 5, 160, 30, 10, "Income:"
  Text 5, 180, 25, 10, "Assets:"
  Text 5, 220, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 240, 50, 10, "Verifs needed:"
  Text 5, 260, 50, 10, "Actions taken:"
  Text 5, 280, 40, 10, "Other notes:"
  Text 175, 300, 50, 10, "Review status:"
  Text 305, 300, 65, 10, "Sign the case note:"
  GroupBox 180, 5, 90, 25, "ELIG panels:"
  GroupBox 15, 5, 110, 25, "STAT panels:"
  GroupBox 330, 5, 110, 35, "STAT-based navigation"
  Text 5, 60, 55, 10, "Interview Date:"
  ButtonGroup ButtonPressed
    PushButton 20, 15, 25, 10, "HCRE", HCRE_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number & footer month/year
EMConnect ""
call MAXIS_case_number_finder(case_number)
'call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)  removed since typically CARs are run on current month + 1 anyway

MAXIS_footer_month = cstr(MAXIS_footer_month)
MAXIS_footer_year = cstr(MAXIS_footer_year)

'Shows case number dialog
Do
	Dialog case_number_dialog
	If buttonpressed = 0 then StopScript
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checks for an active MAXIS session
call check_for_MAXIS(False)

'Navigates to STAT
call navigate_to_MAXIS_screen("STAT", "REVW")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofill info
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", US_citizen)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", recert_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)
CALL autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL)
CALL autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL)

'Determines recert month
recert_month = MAXIS_footer_month & "/" & MAXIS_footer_year

recert_month = cstr(recert_month)

'Showing the case note dialog
DO
	Do
		DO
			Dialog combined_AR_dialog
			cancel_confirmation
			MAXIS_dialog_navigation
		Loop until ButtonPressed = -1
		If worker_signature = "" or review_status = "Select one..." or actions_taken = "" or recert_datestamp = "" then MsgBox "You must sign your case note and update the datestamp, actions taken, and review status sections."
	Loop until worker_signature <> "" and review_status <> "Select one..." and actions_taken <> "" and recert_datestamp <> ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false			

'The case note----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_case_note("***Combined AR received " & recert_datestamp & " for " & recert_month & ": " & review_status & "***")
CALL write_bullet_and_variable_in_case_note("Interview Date", interview_date)
CALL write_bullet_and_variable_in_case_note("HH comp", HH_comp)
CALL write_bullet_and_variable_in_case_note("Citizenship", US_citizen)
CALL write_bullet_and_variable_in_case_note("AREP", AREP)
CALL write_bullet_and_variable_in_case_note("FACI", FACI)
CALL write_bullet_and_variable_in_case_note("Income", income)
CALL write_bullet_and_variable_in_case_note("Assets", assets)
CALL write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL)
CALL write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
IF R_R_explained = checked THEN CALL write_variable_in_case_note("* R/R explained.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
IF eDRS_checked = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
CALL write_bullet_and_variable_in_case_note("Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")

