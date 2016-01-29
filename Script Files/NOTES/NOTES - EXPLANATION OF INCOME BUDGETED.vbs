'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EXPLANATION OF INCOME BUDGETED.vbs"
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
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 240                  'manual run time in seconds
STATS_denomination = "C"       					'C is for each MEMBER
'END OF stats block==============================================================================================

'DIALOGS---------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog

BeginDialog explanation_of_income_budgeted_dialog, 0, 0, 306, 320, "Explanation Of Income Budgeted Dialog"
  EditBox 105, 40, 65, 15, date_calculated
  EditBox 75, 75, 220, 15, earned_income
  EditBox 75, 95, 220, 15, unea_income
  EditBox 120, 125, 175, 15, explanation_of_income
  EditBox 120, 145, 175, 15, income_not_budgeted_and_why
  DropListBox 120, 165, 175, 15, "Select one..."+chr(9)+"Paystubs"+chr(9)+"Employer's Statement"+chr(9)+"Other (Specified Under 'Other Notes on verifications')", type_of_verification_used
  EditBox 120, 185, 175, 15, other_notes_verifs
  DropListBox 120, 205, 175, 15, "Select one..."+chr(9)+"Most Recent 30 Days"+chr(9)+"30 Days Prior To Date Of Application/Recertification"+chr(9)+"Retrospective Month (MFIP/GA)"+chr(9)+"Other (Specified Under 'Other Notes on income')", time_period_used
  EditBox 120, 225, 175, 15, other_notes_on_income
  CheckBox 5, 250, 195, 10, "Verification/s match what client reported they anticipate.", verifications_match_check
  CheckBox 5, 270, 340, 10, "Agency/client are ''reasonably certain'' budgeted income will continue during the ", reasonably_certain_check
  EditBox 70, 295, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 295, 50, 15
    CancelButton 245, 295, 50, 15
    PushButton 10, 20, 25, 10, "BUSI", BUSI_button
    PushButton 35, 20, 25, 10, "JOBS", JOBS_button
    PushButton 60, 20, 25, 10, "RBIC", RBIC_button
    PushButton 85, 20, 25, 10, "UNEA", UNEA_button
    PushButton 135, 20, 40, 10, "prev panel", prev_panel_button
    PushButton 175, 20, 40, 10, "prev memb", prev_memb_button
    PushButton 215, 20, 40, 10, "next panel", next_panel_button
    PushButton 255, 20, 40, 10, "next memb", next_memb_button
  Text 5, 210, 95, 10, "Time period of income used:"
  Text 5, 170, 85, 10, "Type Of Verification Used:"
  GroupBox 5, 5, 110, 30, "Income panel navigation:"
  Text 5, 300, 60, 10, "Worker signature:"
  Text 5, 230, 80, 10, "Other Notes On Income:"
  Text 5, 45, 95, 10, "Date income was calculated:"
  GroupBox 130, 5, 170, 30, "STAT-based navigation:"
  Text 5, 150, 110, 10, "Income NOT Budgeted and Why:"
  Text 10, 100, 60, 10, "Unearned income:"
  Text 15, 280, 65, 10, "certification period."
  Text 10, 80, 55, 10, "Earned income:"
  GroupBox 5, 65, 295, 55, "Income in this section reflects the income in MAXIS for the selected footer month/year."
  Text 5, 130, 110, 10, "Explanation Of Income Budgeted:"
  Text 5, 190, 105, 10, "Other notes on verfications:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------
'Connect to BlueZone & grabbing the case number & the footer month/year
EMConnect ""
Call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Do
	err_msg = ""
	Do
  	dialog case_number_dialog     'initial dialog
  	If ButtonPressed = 0 then Stopscript    'if cancel is pressed then the script ends
  	Call check_for_password(are_we_passworded_out)    'function to see if users is password-ed out
	Loop until are_we_passworded_out = false  	'will loop until user is password-ed back in
	If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
	If IsNumeric(case_number) = False or Len(case_number) > 8 then err_msg = err_msg & vbNewLine & "* You must enter a valid case number."
  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""

back_to_self        'retuns user back to self menu
MAXIS_background_check    'checks to see if the case is stuck in background
EMWriteScreen MAXIS_footer_month, 20, 43    'writes the selected footer month/year on SELF screen
EMWriteScreen MAXIS_footer_year, 20, 46

'Creating a custom dialog for HH members listed on the case for users to select from
call HH_member_custom_dialog(HH_member_array)

'Now it grabs the income information for the select HH members from the HH array
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unea_income)

date_calculated = date & ""    'setting the date_calculated defaulting to current date for the dialog (must add & "" to turn the variable into a string)

Do
  err_msg = ""
      Do
          Dialog explanation_of_income_budgeted_dialog
          cancel_confirmation
          MAXIS_dialog_navigation
      Loop until ButtonPressed = -1     'Looping until OK button is pressed
  If IsNumeric(case_number) = FALSE or len(case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  If IsDate(date_calculated) = FALSE then err_msg = err_msg & vbNewLine & "* Enter the date income was calculated."
  If explanation_of_income = "" then err_msg = err_msg & vbNewLine & "* Explain the income that is being budgeted."
  If type_of_verification_used = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select the type verification used."
  If time_period_used = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select the time period of the income used."
  If (type_of_verification_used = "Other (Specified Under 'Other Notes on verifications')" AND other_notes_verifs = "") then err_msg = err_msg & vbNewLine & "* You must explain the type of verification used in the ""other notes on verifications"" field."
  If (time_period_used = "Other (Specified Under 'Other Notes on income')" AND other_notes_on_income = "") then err_msg = err_msg & vbNewLine & "* You must explain the time period of income used in the ""other notes on income"" field."
  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""

'cleaning up the variables for the case note
If type_of_verification_used = "Other (Specified Under 'Other Notes on verifications')" then type_of_verification_used = "Other verification"
If time_period_used = "Other (Specified Under 'Other Notes on income')" then time_period_used = "Other time period"

'Writes to the CASE NOTE
Call start_a_blank_case_note    'navigates to, and starts a new case note
Call write_variable_in_CASE_NOTE(">>>Explanation of income budgeted<<<")
Call write_bullet_and_variable_in_CASE_NOTE("Date income was calculated", date_calculated)
Call write_bullet_and_variable_in_CASE_NOTE("Earned income", earned_income)
Call write_bullet_and_variable_in_CASE_NOTE("Unearned income", unea_income)
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Explanation of income budgeted", explanation_of_income)
call write_bullet_and_variable_in_CASE_NOTE("Income not budgeted and why", income_not_budgeted_and_why)
call write_bullet_and_variable_in_CASE_NOTE("Type of verification used", type_of_verification_used)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes on verifications", other_notes_verifs)
call write_bullet_and_variable_in_CASE_NOTE("Time period used", time_period_used)
call write_bullet_and_variable_in_CASE_NOTE("Other notes on income", other_notes_on_income)
Call write_variable_in_CASE_NOTE("---")
If reasonably_certain_check = 1 then call write_variable_in_CASE_NOTE("* Agency/client are 'reasonably certain' budgeted income will continue during      the certification period.")
If verifications_match_check = 1 then call write_variable_in_CASE_NOTE("* Verification/s match what client reported they anticipate.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
