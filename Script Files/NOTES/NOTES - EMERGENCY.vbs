'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMERGENCY.vbs"
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
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
footer_month = datepart("m", date) & ""
If len(footer_month) = 1 then footer_month = "0" & footer_month & ""
footer_year = datepart("yyyy", date)
footer_year = "" & footer_year - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 97, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, MAXIS_footer_month
  EditBox 140, 25, 30, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_check
  CheckBox 50, 60, 30, 10, "HC", HC_check
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 95, 80, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Other programs open or applied for:"
EndDialog

'This dialog contains a customized "percent rule" variable, as well as a customized "income days" variable. As such, it can't directly be edited in the dialog editor.
BeginDialog emergency_dialog, 0, 0, 321, 395, "Emergency Dialog"
  EditBox 60, 45, 65, 15, interview_date
  EditBox 170, 45, 150, 15, HH_comp
  CheckBox 25, 75, 40, 10, "Eviction", eviction_check
  CheckBox 75, 75, 70, 10, "Utility disconnect", utility_disconnect_check
  CheckBox 155, 75, 60, 10, "Homelessness", homelessness_check
  CheckBox 230, 75, 65, 10, "Security deposit", security_deposit_check
  EditBox 65, 100, 255, 15, cause_of_crisis
  EditBox 85, 160, 235, 15, income
  EditBox 110, 180, 210, 15, income_under_200_FPG
  EditBox 60, 200, 260, 15, percent_rule_notes
  EditBox 75, 220, 245, 15, monthly_expense
  EditBox 40, 240, 280, 15, assets
  EditBox 60, 260, 260, 15, verifs_needed
  EditBox 75, 280, 245, 15, crisis_resolvable
  EditBox 80, 300, 240, 15, discussion_of_crisis
  EditBox 60, 320, 260, 15, actions_taken
  EditBox 50, 340, 270, 15, referrals
  CheckBox 5, 360, 90, 10, "Sent forms to AREP?", sent_arep_checkbox
  EditBox 75, 375, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 375, 50, 15
    CancelButton 255, 375, 50, 15
    PushButton 10, 15, 25, 10, "ADDR", ADDR_button
    PushButton 35, 15, 25, 10, "MEMB", MEMB_button
    PushButton 60, 15, 25, 10, "MEMI", MEMI_button
    PushButton 10, 25, 25, 10, "PROG", PROG_button
    PushButton 35, 25, 25, 10, "TYPE", TYPE_button
    PushButton 125, 20, 50, 10, "ELIG/EMER", ELIG_EMER_button
    PushButton 210, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 210, 25, 45, 10, "next panel", next_panel_button
    PushButton 270, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 270, 25, 45, 10, "next memb", next_memb_button
    PushButton 75, 130, 25, 10, "BUSI", BUSI_button
    PushButton 100, 130, 25, 10, "JOBS", JOBS_button
    PushButton 75, 140, 25, 10, "RBIC", RBIC_button
    PushButton 100, 140, 25, 10, "UNEA", UNEA_button
    PushButton 150, 130, 25, 10, "ACCT", ACCT_button
    PushButton 175, 130, 25, 10, "CARS", CARS_button
    PushButton 200, 130, 25, 10, "CASH", CASH_button
    PushButton 225, 130, 25, 10, "OTHR", OTHR_button
    PushButton 150, 140, 25, 10, "REST", REST_button
    PushButton 175, 140, 25, 10, "SECU", SECU_button
    PushButton 200, 140, 25, 10, "TRAN", TRAN_button
  GroupBox 5, 5, 85, 35, "other STAT panels:"
  GroupBox 205, 5, 115, 35, "STAT-based navigation"
  Text 5, 50, 50, 10, "Interview date:"
  Text 130, 50, 35, 10, "HH Comp:"
  GroupBox 20, 65, 280, 25, "Crisis (check all that apply):"
  Text 5, 105, 55, 10, "Cause of crisis:"
  GroupBox 70, 120, 60, 35, "Income panels"
  GroupBox 145, 120, 110, 35, "Asset panels"
  Text 5, 165, 75, 10, "Income (past " & emer_number_of_income_days & " days):"
  Text 5, 185, 100, 10, "Is income under 200% FPG?:"
  Text 5, 205, 55, 10, emer_percent_rule_amt & "% rule notes:"
  Text 5, 225, 60, 10, "Monthly expense:"
  Text 5, 245, 30, 10, "Assets:"
  Text 5, 265, 50, 10, "Verifs needed:"
  Text 5, 285, 65, 10, "Crisis resolvable?:"
  Text 5, 305, 75, 10, "Discussion of Crisis:"
  Text 5, 325, 50, 10, "Actions taken:"
  Text 5, 345, 40, 10, "Referrals:"
  Text 5, 380, 65, 10, "Worker signature:"
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
application_signed_check = 1 'The script should default to having the application signed.

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number & footer month/year
EMConnect ""
CALL MAXIS_case_number_finder(case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for an active MAXIS session
Call check_for_MAXIS(False)

'Jumping into STAT
call navigate_to_MAXIS_screen("stat", "hcre")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", monthly_expense) 'Does this last because people like it tacked on to the end, not before. The rest are alphabetical.


'Showing the case note
DO
	Do
		Do
			Dialog emergency_dialog
			MAXIS_dialog_navigation
			cancel_confirmation
		Loop until ButtonPressed = -1
		If ButtonPressed = -1 then dialog case_note_dialog
	    If income = "" or actions_taken = "" or worker_signature = "" then MsgBox "You need to fill in the income and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
	 Loop until income <> "" and actions_taken <> "" and worker_signature <> ""
	 call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
 LOOP UNTIL are_we_passworded_out = false

'Logic to enter what the "crisis" variable is from the checkboxes indicated
If eviction_check = 1 then crisis = crisis & "eviction, "
If utility_disconnect_check = 1 then crisis = crisis & "utility disconnect, "
If homelessness_check = 1 then crisis = crisis & "homelessness, "
If security_deposit_check = 1 then crisis = crisis & "security deposit, "
If eviction_check = 0 and utility_disconnect_check = 0 and homelessness_check = 0 and security_deposit_check = 0 then
  crisis = "no crisis given."
Else
  crisis = trim(crisis)
  crisis = left(crisis, len(crisis) - 1) & "."
End if

'Writing the case note
call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***Emergency app: "& replace(crisis, ".", "") & "***")
If interview_date <> "" then call write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
If HH_comp <> "" then call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
If crisis <> "" then call write_bullet_and_variable_in_CASE_NOTE("Crisis", crisis)
If cause_of_crisis <> "" then call write_bullet_and_variable_in_CASE_NOTE("Cause of crisis", cause_of_crisis)
If income <> "" then call write_bullet_and_variable_in_CASE_NOTE("Income, past " & emer_number_of_income_days & " days", income)
If income_under_200_FPG <> "" then call write_bullet_and_variable_in_CASE_NOTE("Income under 200% FPG", income_under_200_FPG)
If percent_rule_notes <> "" then call write_bullet_and_variable_in_CASE_NOTE(emer_percent_rule_amt & "% rule notes", percent_rule_notes)
If monthly_expense <> "" then call write_bullet_and_variable_in_CASE_NOTE("Monthly expense", monthly_expense)
If assets <> "" then call write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
if verifs_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)
If crisis_resolvable <> "" then call write_bullet_and_variable_in_CASE_NOTE("Crisis resolvable?", crisis_resolvable)
If discussion_of_crisis <> "" then call write_bullet_and_variable_in_CASE_NOTE("Discussion of crisis", discussion_of_crisis)
If actions_taken <> "" then call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If referrals <> "" then call write_bullet_and_variable_in_CASE_NOTE("Referrals", referrals)
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
