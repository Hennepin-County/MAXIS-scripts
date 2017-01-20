'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EXPLANATION OF INCOME BUDGETED.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 240                  'manual run time in seconds
STATS_denomination = "C"       					'C is for each MEMBER
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS---------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog

BeginDialog explanation_of_income_budgeted_dialog, 0, 0, 306, 370, "Explanation Of Income Budgeted Dialog"
  EditBox 105, 40, 65, 15, date_calculated
  EditBox 75, 75, 220, 15, earned_income
  EditBox 75, 95, 220, 15, self_employment_income
  EditBox 75, 115, 220, 15, unea_income
  EditBox 120, 140, 175, 15, overtime_explanation
  ButtonGroup ButtonPressed
    PushButton 5, 185, 110, 10, "Explanation Of Income Budgeted", income_button
  EditBox 120, 180, 175, 15, explanation_of_income
  EditBox 120, 200, 175, 15, income_not_budgeted_and_why
  DropListBox 120, 220, 175, 15, "Select one..."+chr(9)+"Paystubs"+chr(9)+"Employer's Statement"+chr(9)+"Other (Specified Under 'Other Notes on verifications')", type_of_verification_used
  EditBox 120, 240, 175, 15, other_notes_verifs
  DropListBox 120, 260, 175, 15, "Select one..."+chr(9)+"Most Recent 30 Days"+chr(9)+"30 Days Prior To Date Of Application/Recertification"+chr(9)+"Retrospective Month (MFIP/GA)"+chr(9)+"Other (Specified Under 'Other Notes on income')", time_period_used
  EditBox 120, 280, 175, 15, other_notes_on_income
  CheckBox 5, 300, 195, 10, "Verification/s match what client reported they anticipate.", verifications_match_check
  CheckBox 5, 315, 270, 10, "Agency/client are ''reasonably certain'' budgeted income will continue during the ", reasonably_certain_check
  EditBox 70, 345, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 345, 50, 15
    CancelButton 245, 345, 50, 15
    PushButton 10, 20, 25, 10, "BUSI", BUSI_button
    PushButton 35, 20, 25, 10, "JOBS", JOBS_button
    PushButton 60, 20, 25, 10, "RBIC", RBIC_button
    PushButton 85, 20, 25, 10, "UNEA", UNEA_button
    PushButton 135, 20, 40, 10, "prev panel", prev_panel_button
    PushButton 175, 20, 40, 10, "prev memb", prev_memb_button
    PushButton 215, 20, 40, 10, "next panel", next_panel_button
    PushButton 255, 20, 40, 10, "next memb", next_memb_button
  Text 5, 350, 60, 10, "Worker signature:"
  Text 5, 285, 80, 10, "Other Notes On Income:"
  Text 5, 45, 95, 10, "Date income was calculated:"
  GroupBox 130, 5, 170, 30, "STAT-based navigation:"
  Text 5, 205, 110, 10, "Income NOT Budgeted and Why:"
  Text 10, 120, 60, 10, "Unearned income:"
  Text 15, 330, 65, 10, "certification period."
  Text 10, 80, 55, 10, "Earned income:"
  GroupBox 5, 65, 295, 70, "Income in this section reflects the income in MAXIS for the selected footer month/year."
  Text 195, 45, 110, 10, "Explanation Of Income Budgeted:"
  Text 5, 245, 105, 10, "Other notes on verfications:"
  Text 10, 100, 60, 10, "Self-Employment:"
  Text 5, 225, 85, 10, "Type Of Verification Used:"
  Text 5, 265, 95, 10, "Time period of income used:"
  Text 5, 145, 110, 10, "Overtime income explanation:"
  GroupBox 5, 5, 110, 30, "Income panel navigation:"
  Text 5, 160, 300, 20, "If Overtime was included on any income verification, detail which job, if it was budgeted or excluded, and why."
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

'THE SCRIPT----------------------------------------------------------------------
'Connect to BlueZone & grabbing the case number & the footer month/year
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
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
	If IsNumeric(MAXIS_case_number) = False or Len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* You must enter a valid case number."
  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""

back_to_self        'retuns user back to self menu
MAXIS_background_check    'checks to see if the case is stuck in background
EMWriteScreen MAXIS_footer_month, 20, 43    'writes the selected footer month/year on SELF screen
EMWriteScreen MAXIS_footer_year, 20, 46

'Creating a custom dialog for HH members listed on the case for users to select from
call HH_member_custom_dialog(HH_member_array)

'Now it grabs the income information for the select HH members from the HH array
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", self_employment_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unea_income)

date_calculated = date & ""    'setting the date_calculated defaulting to current date for the dialog (must add & "" to turn the variable into a string)

'If a BUSI panel exists this will display information about the Self Employment policy to guide workers
call navigate_to_MAXIS_screen ("STAT", "BUSI")
IF self_employment_income <> "" Then
	MsgBox("Your case has self employment income on it." & vbNewLine & vbNewLine & "Please be aware of the following:" & vbNewLine & "* Self Employment income can be budgeted using 2 different methods:" & vbNewLine & "     (01) 50% of Gross Income" & vbNewLine & "     (02) Taxable Income - per taxes less than 12 months old" & _
	  vbNewLine & "     Refer to CM 17.15.33.03 for detail" & vbNewLine & vbNewLine &"** SNAP cases with rental income do NOT use these methods" & vbNewLine & "     Refer to CM 17.15.33.30 for correct budgeting for these cases" & vbNewLine & "** Cases with farming income may also use a different budgeting method" & _
	  vbNewLine & "     Refer to CM 17.15.33.24 for more information" & vbNewLine & vbNewLine & "* There are specific rules in regards to clients switching budgeting methods." & vbNewLine & "     If you are documenting a switch, be sure to explain in detail." & vbNewLine & "     Refer to CM 17.15.33.03 for details on changing method")
	Self_employment_message_shown = TRUE
End IF

'Runs the main dialog to detail all the information about income.
DO
	Do
		Do
			err_msg = ""
			Do
				Dialog explanation_of_income_budgeted_dialog
				cancel_confirmation
				MAXIS_dialog_navigation
				If ButtonPressed = income_button Then
				  	Dialog income_notes_dialog
					If ButtonPressed = add_to_notes_button Then
						If jobs_anticipated_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client expects all income from jobs to continue at this amount."
						If new_jobs_checkbox = checked Then explanation_of_income = explanation_of_income & "; This is a new job and actual check stubs have not been received, advised client to provide proof once pay is received if the income received differs significantly."
						If busi_anticipated_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client expects all income from self employment to continue at this amount."
						If busi_method_agree_checkbox = checked Then explanation_of_income = explanation_of_income & "; Explained to client the self employment budgeting methods and client agreed to the method used."
						If rbic_anticipated_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client expects roomer/boarder income to continue at this amount."
						If unea_anticipated_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client expects unearned income to continue at this amount."
						If ui_pending_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client has applied for Unemployment Income recently but request is still pending, will need to be reviewed soon for changes."
						If tikl_for_ui = checked Then explanation_of_income = explanation_of_income & " TIKL set to request an update on Unemployment Income."
						If no_income_checkbox = checked Then explanation_of_income = explanation_of_income & "; Client has reported they have no income and do not expect any changes to this at this time."
						If left(explanation_of_income, 1) = ";" Then explanation_of_income = right(explanation_of_income, len(explanation_of_income) - 1)
					End If
				End If
			Loop until ButtonPressed = -1     'Looping until OK button is pressed
			If IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
			If IsDate(date_calculated) = FALSE then err_msg = err_msg & vbNewLine & "* Enter the date income was calculated."
			If explanation_of_income = "" then err_msg = err_msg & vbNewLine & "* Explain the income that is being budgeted."
			If type_of_verification_used = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select the type verification used."
			If time_period_used = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select the time period of the income used."
			If (type_of_verification_used = "Other (Specified Under 'Other Notes on verifications')" AND other_notes_verifs = "") then err_msg = err_msg & vbNewLine & "* You must explain the type of verification used in the ""other notes on verifications"" field."
			If (time_period_used = "Other (Specified Under 'Other Notes on income')" AND other_notes_on_income = "") then err_msg = err_msg & vbNewLine & "* You must explain the time period of income used in the ""other notes on income"" field."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP until err_msg = ""
		IF self_employment_income <> "" AND Self_employment_message_shown <> TRUE Then  'If self employment information is added to the dialog but was not autofilled from the BUSI panel, the information about Self Employment Policy
			Self_Emp_Info = MsgBox ("Your case has self employment income on it." & vbNewLine & vbNewLine & "Please be aware of the following:" & vbNewLine & "* Self Employment income can be budgeted using 2 different methods:" & vbNewLine & "     (01) 50% of Gross Income" & vbNewLine & "     (02) Taxable Income - per taxes less than 12 months old" & _
			  vbNewLine & "     Refer to CM 17.15.33.03 for detail" & vbNewLine & vbNewLine &"** SNAP cases with rental income do NOT use these methods" & vbNewLine & "     Refer to CM 17.15.33.30 for correct budgeting for these cases" & vbNewLine & "** Cases with farming income may also use a different budgeting method" & _
			  vbNewLine & "     Refer to CM 17.15.33.24 for more information" & vbNewLine & vbNewLine & "* There are specific rules in regards to clients switching budgeting methods." & vbNewLine & "     If you are documenting a switch, be sure to explain in detail." & _
			  vbNewLine & "     Refer to CM 17.15.33.03 for details on changing method" & vbNewLine & vbNewLine & "Ready to case note?" & vbNewLine & "(Click 'No' to return to previous screen to enter more informaiton)", 4, "")
			IF Self_Emp_Info = vbYes Then Exit Do 'If worker selects 'No' on MsgBox, the dialog will loop for worker to add more detail
		End IF
	Loop until Self_Emp_Info = vbYes OR Self_employment_message_shown = TRUE OR self_employment_income = "" 'Will only display the Self Employment policy information once
	Call check_for_password(are_we_passworded_out) 'Handling for Maxis v6
LOOP UNTIL are_we_passworded_out = false

IF tikl_for_ui THEN
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	two_weeks_from_now = DateAdd("d", 14, date)
	call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
	call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
	PF3
END IF

'cleaning up the variables for the case note
If type_of_verification_used = "Other (Specified Under 'Other Notes on verifications')" then type_of_verification_used = "Other verification"
If time_period_used = "Other (Specified Under 'Other Notes on income')" then time_period_used = "Other time period"

'Writes to the CASE NOTE
Call start_a_blank_case_note    'navigates to, and starts a new case note
Call write_variable_in_CASE_NOTE(">>>Explanation of income budgeted<<<")
Call write_bullet_and_variable_in_CASE_NOTE("Date income was calculated", date_calculated)
Call write_bullet_and_variable_in_CASE_NOTE("Earned income", earned_income)
Call write_bullet_and_variable_in_CASE_NOTE("Self Employment income", self_employment_income)
Call write_bullet_and_variable_in_CASE_NOTE("Unearned income", unea_income)
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Explanation of income budgeted", explanation_of_income)
Call write_bullet_and_variable_in_CASE_NOTE("Income reported with overtime hours", overtime_explanation)
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
