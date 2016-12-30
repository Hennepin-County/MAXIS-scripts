'GATHERING STATS============================================================================================================
name_of_script = "ACTIONS - LTC - ICF-DD DEDUCTION FIATER.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denomination = "C"

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

'Dialogs____________________________________________________________________________________________________________
BeginDialog case_number_dialog, 0, 0, 141, 70, "Case number"
  EditBox 80, 5, 55, 15, MAXIS_case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 110, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 30, 50, 50, 15
    CancelButton 85, 50, 50, 15
  Text 25, 10, 45, 10, "Case number:"
  Text 10, 30, 65, 10, "Footer month/year:"
EndDialog

BeginDialog LTC_ICFDD_Fiater_dialog, 0, 0, 286, 190, "LTC-ICF-DD Fiater"
  ButtonGroup ButtonPressed
    OkButton 175, 170, 50, 15
    CancelButton 230, 170, 50, 15
  EditBox 90, 65, 40, 15, special_pers_allow
  EditBox 90, 85, 40, 15, Federal_tax
  EditBox 90, 105, 40, 15, state_tax
  EditBox 90, 125, 40, 15, FICA_witheld
  EditBox 90, 145, 40, 15, transportation_expense
  EditBox 220, 65, 40, 15, meals_expense
  EditBox 220, 85, 40, 15, uniform_expense
  EditBox 220, 105, 40, 15, tools_expense
  EditBox 220, 125, 40, 15, dues_expense
  EditBox 220, 145, 40, 15, other_expense
  ButtonGroup ButtonPressed
    PushButton 150, 20, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 185, 20, 25, 10, "FACI", FACI_button
    PushButton 225, 20, 25, 10, "JOBS", JOBS_button
    PushButton 250, 20, 25, 10, "UNEA", UNEA_button
    PushButton 150, 35, 25, 10, "BILS", BILS_button
    PushButton 175, 35, 25, 10, "INSA", INSA_button
    PushButton 200, 35, 25, 10, "MEDI", MEDI_button
    PushButton 225, 35, 25, 10, "PDED", PDED_button
    PushButton 250, 35, 25, 10, "WKEX", WKEX_button
  Text 10, 90, 45, 10, "Federal Tax:"
  Text 150, 130, 20, 10, "Dues:"
  Text 10, 70, 80, 10, "Special Pers Allowance:"
  Text 150, 150, 30, 10, "Other:"
  Text 10, 110, 40, 10, "State Tax:"
  Text 10, 130, 25, 10, "FICA:"
  Text 150, 90, 35, 10, "Uniforms:"
  Text 10, 150, 55, 10, "Transportation:"
  GroupBox 145, 10, 135, 40, "STAT based navigation"
  Text 150, 110, 25, 10, "Tools:"
  GroupBox 5, 55, 275, 110, "WKEX deductions"
  GroupBox 5, 10, 135, 40, "Unsure of what to deduct?"
  Text 10, 25, 125, 20, "HCPM 23.15.10 covers allowable LTC income deductions."
  Text 150, 70, 25, 10, "Meals:"
EndDialog

'THE SCRIPT========================================================================================
'This connects to Bluezone
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Shows case number dialog
DO
	dialog case_number_dialog
	cancel_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'confirms that case is in the footer month/year selected by the user
Call MAXIS_footer_month_confirmation
MAXIS_background_check

'Enters into STAT for the client
Call navigate_to_MAXIS_screen("STAT", "WKEX")
EMReadScreen WKEX_check, 1, 2, 73
If WKEX_check = "0" then
	script_end_procedure("You do not have a WKEX panel. Please create a WKEX panel for your HC case, and re-run the script.")
Elseif WKEX_check <> "0" then
	'Reads work expenses for MEMB 01, verification codes and impairment related code
	EMReadScreen program_check,          2, 5, 33
	EMReadScreen federal_tax,            8, 7, 57
	EMReadScreen federal_tax_verif_code, 1, 7, 69
	EMReadScreen state_tax,              8, 8, 57
	EMReadScreen state_tax_verif_code, 1, 8, 69
	EMReadScreen FICA_witheld,           8, 9, 57
	EMReadScreen FICA_witheld_verif_code, 1, 9, 69
	EMReadScreen transportation_expense, 8, 10, 57
	EMReadScreen transportation_expense_verif_code, 1, 10, 69
	EMReadScreen transportation_impair, 1, 10, 75
	EMReadScreen meals_expense, 8, 11, 57
	EMReadScreen meals_impair, 1, 11, 75
	EMReadScreen meals_expense_verif_code, 1, 11, 69
	EMReadScreen uniform_expense, 8, 12, 57
	EMReadScreen uniform_expense_verif_code, 1, 12, 69
	EMReadScreen uniform_impair, 1, 12, 75
	EMReadScreen tools_expense, 8, 13, 57
	EMReadScreen tools_expense_verif_code, 1, 13, 69
	EMReadScreen tools_impair, 1, 13, 75
	EMReadScreen dues_expense, 8, 14, 57
	EMReadScreen dues_expense_verif_code, 1, 14, 69
	EMReadScreen dues_impair, 1, 14, 75
	EMReadScreen other_expense, 8, 15,	57
	EMReadScreen other_expense_verif_code, 1, 15, 69
	EMReadScreen other_impair, 1, 15, 75
End IF

'cleaning up the WKEX variables
federal_tax = replace(federal_tax, "_", "")
federal_tax = trim(federal_tax)
state_tax = replace(state_tax, "_", "")
state_tax = trim(state_tax)
FICA_witheld = replace(FICA_witheld, "_", "")
FICA_witheld = trim(FICA_witheld)
transportation_expense = replace(transportation_expense, "_", "")
transportation_expense = trim(transportation_expense)
meals_expense = replace(meals_expense, "_", "")
meals_expense = trim(meals_expense)
uniform_expense = replace(uniform_expense, "_", "")
uniform_expense = trim(uniform_expense)
tools_expense = replace(tools_expense, "_", "")
tools_expense = trim(tools_expense)
dues_expense = replace(dues_expense, "_", "")
dues_expense = trim(dues_expense)
other_expense = replace(other_expense, "_", "")
other_expense = trim(other_expense)

'Gives unverified expenses and blank expenses the value of $0
If federal_tax = "" OR federal_tax_verif_code = "N" then federal_tax = "0"
If state_tax = "" OR state_tax_verif_code = "N" then state_tax = "0"
If FICA_witheld = "" OR FICA_witheld_verif_code = "N" then FICA_witheld = "0"
If transportation_expense = "" OR transportation_expense_verif_code = "N" OR _
	transportation_impair =  "_" OR transportation_impair = "N" then transportation_expense = "0"
If meals_expense = "" OR meals_expense_verif_code = "N" OR _
	meals_impair = "_" OR meals_impair = "N" then meals_expense = "0"
If uniform_expense = "" OR uniform_expense_verif_code = "N" OR _
	uniform_impair = "_" OR uniform_impair = "N" then uniform_expense = "0"
If tools_expense = "" OR tools_expense_verif_code = "N" OR _
	tools_impair = "_" OR tools_impair = "N" then tools_expense = "0"
If dues_expense = "" OR dues_expense_verif_code = "N" OR _
	dues_impair = "_" OR dues_impair = "N" then dues_expense = "0"
If other_expense = "" OR other_expense_verif_code = "N" OR _
	other_impair = "_" OR other_impair = "N" then other_expense = "0"

'Checks PDED other expenses, will need to add PDED and WKEX other expenses together
Call navigate_to_MAXIS_screen("STAT", "PDED")
EMReadScreen other_earned_income_PDED, 8, 11, 62

'cleaning up PDED variables
other_earned_income_PDED = replace(other_earned_income_PDED, "_", "")
other_earned_income_PDED = trim(other_earned_income_PDED)

'Gives blank expenses the value of $0
If other_earned_income_PDED = "" then other_earned_income_PDED = "0"

'creating new variables for input of deductions that don't have their own expense field in the HC budget
Total_taxes = abs(federal_tax) + abs(state_tax)
Total_employment_expense = abs(meals_expenses) + abs(uniform_expense) + abs(tools_expense) + abs(dues_expenses)
total_other_expenses = abs(other_expenses) + abs(other_earned_income_PDED)

'Determining if earned income is less than $80
Call navigate_to_MAXIS_screen ("STAT", "JOBS")
EMReadScreen JOBS_panel_income, 7, 17, 68
JOBS_panel_income = trim(JOBS_panel_income)
If abs(JOBS_panel_income) < 80 then
	special_pers_allow = JOBS_panel_income	'if less then $80 deduction is earned income amount
ELSE
	special_pers_allow = "80.00"		'otherwise deduction is $80
END IF

'Shows the LTC_ICFDD_Fiater_dialog
DO
	DO
		dialog LTC_ICFDD_Fiater_dialog
		cancel_confirmation
		MAXIS_Dialog_navigation
	Loop until ButtonPressed = -1 							' - 1 is OK button
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Navigates to ELIG/HC.
Call navigate_to_MAXIS_screen("ELIG", "HC")
EMReadScreen person_check, 1, 8, 26
If person_check <> "_" then script_end_procedure("Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results.")
'navigates to the HC summary screen
EMWriteScreen "x", 8, 26
Transmit

'Turns on FIAT mode, checks to make sure it worked and sets the reason as "06".
PF9
EMReadScreen FIAT_check, 4, 10, 20
If FIAT_check = "FIAT" then
  EMWriteScreen "06", 11, 26
  transmit
End IF

'Defining the variables for the following search
ELIG_HC_row = 6
ELIG_HC_col = 1

'Determining the col variable based on the indicated footer month/year
EMSearch MAXIS_footer_month & "/" & MAXIS_footer_year, ELIG_HC_row, ELIG_HC_col
If ELIG_HC_col = 0 then script_end_procedure("Requested footer month not found. You may have entered ELIG/HC in an invalid footer month, or results haven't been generated for that month. Check these out and try again.")
col = ELIG_HC_col + 1

'setting the variable for the next do...loop
budget_months = 0

'Fills all budget months with "x's", so that the script will go into each one in succession.
Do
  EMReadScreen budget_check, 1, 12, col
  If budget_check = "/" then
    budget_months = budget_months + 1
    EMWriteScreen "x", 9, col + 1
  End if
  col = col + 11
Loop until col > 75
'Jumps into the budget screen LBUD
transmit

For amt_of_months_to_do = 1 to budget_months
	'Checks to see if this is an LBUD case. It'll stop if it's neither.
	EMReadScreen LBUD_check, 4, 3, 45
	If LBUD_check <> "LBUD" then script_end_procedure("This is not a method L. This script is only for use in Method L cases.")

 	'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This gathers the current month to see if we've moved past it later.
  	EMReadScreen starting_bdgt_month, 5, 6, 14
	'Transmit to the next screen after putting an "x" on the Countable Earned Income (LBUD) screen
	EMWriteScreen "x", 9, 3
	Transmit
	EMWriteScreen "x", 6, 7
	transmit
	EMWriteScreen "___________", 8, 43			'enters into the gross earned income field as this needs to carry over through the whole budget period
	EMWriteScreen JOBS_panel_income, 8, 43
	EMWriteScreen "N", 8, 59					'excluded income "N"
	transmit
	transmit 									'must transmit twice to go back to deductions
	EMWriteScreen "__________", 7, 42
	EMWriteScreen special_pers_allow, 7, 42
	EMWriteScreen "__________", 8, 42
	EMWriteScreen FICA_witheld, 8, 42
	EMWriteScreen "__________", 9, 42
	EMWriteScreen transportation_expense, 9, 42
	EMWriteScreen "__________", 10, 42
	EMWriteScreen Total_employment_expense, 10, 42
	EMWriteScreen "__________", 11, 42
	EMWriteScreen Total_taxes, 11, 42
	EMWriteScreen "__________", 12, 42
	EMWriteScreen total_other_expenses, 12, 42
	transmit
	transmit

	'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This is a fix for that.
 	EMReadScreen ending_bdgt_month, 5, 6, 14
 	If starting_bdgt_month = ending_bdgt_month then transmit

 	'Resets the variables to check on the next month.
 	LBUD_check = ""
NEXT

script_end_procedure("Success, the budget has been updated to reflect earned income disregards. Please review the case prior to approval. Use NOTES - LTC - MA APPROVAL to case note the approval.")
