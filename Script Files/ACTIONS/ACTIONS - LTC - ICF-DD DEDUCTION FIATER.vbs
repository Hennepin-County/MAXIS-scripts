'GATHERING STATS============================================================================================================
name_of_script = "ACTIONS - LTC-ICF-DD DEDUCTION FIATER.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denomination = "C"

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
'END FUNCTIONS LIBRARY BLOCK============================================================================================================

'Dialog____________________________________________________________________________________________________________
BeginDialog LTC_ICFDD_Fiater_dialog, 0, 0, 146, 255, "LTC-ICF-DD Fiater"
  EditBox 90, 15, 40, 15, special_pers_allow
  EditBox 90, 30, 40, 15, Federal_tax
  EditBox 90, 50, 40, 15, state_tax
  EditBox 90, 65, 40, 15, FICA_witheld
  EditBox 90, 85, 40, 15, transportation_expense
  EditBox 90, 105, 40, 15, meal_expenses
  EditBox 90, 125, 40, 15, uniform_expense
  EditBox 90, 145, 40, 15, tool_expense
  EditBox 90, 165, 40, 15, due_expense
  EditBox 90, 185, 40, 15, other_expenses
  ButtonGroup ButtonPressed
    PushButton 10, 220, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 45, 220, 25, 10, "WKEX", WKEX_button
    PushButton 90, 220, 25, 10, "JOBS", JOBS_button
    PushButton 70, 220, 25, 10, "FACI", FACI_button
    PushButton 115, 220, 25, 10, "UNEA", UNEA_button
    OkButton 35, 235, 50, 15
    CancelButton 90, 235, 50, 15
  GroupBox 0, 210, 145, 25, "STAT based navigation"
  GroupBox 0, 0, 135, 205, "Deductions"
  Text 10, 70, 70, 10, "FICA:"
  Text 10, 90, 70, 10, "Transportation:"
  Text 10, 130, 50, 10, "Uniforms:"
  Text 10, 55, 50, 10, "State Tax:"
  Text 10, 150, 50, 10, "Tools:"
  Text 10, 35, 55, 10, "Federal Tax:"
  Text 10, 170, 50, 10, "Dues:"
  Text 10, 20, 80, 10, "Special Pers Allowance:"
  Text 10, 190, 50, 10, "Other:"
  Text 10, 110, 70, 10, "Meals:"
EndDialog



BeginDialog case_number_dialog, 0, 0, 161, 70, "Case number"
  EditBox 75, 5, 75, 15, case_number
  EditBox 75, 25, 25, 15, MAXIS_footer_month
  EditBox 125, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 45, 50, 50, 15
    CancelButton 100, 50, 50, 15
  Text 25, 10, 45, 10, "Case number:"
  Text 5, 30, 70, 10, "MAXIS footer month:"
  Text 105, 30, 20, 10, "Year:"
EndDialog



'THE SCRIPT========================================================================================

'This connects to Bluezone
EMConnect ""
call MAXIS_case_number_finder(case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'Shows case number dialog
dialog case_number_dialog
cancel_confirmation

'Enters into STAT for the client
Call navigate_to_MAXIS_screen("STAT", "WKEX")

EMReadScreen WKEX_check, 1,2,73
IF WKEX_check = "1" then 

'Checks for work expenses for MEMB01
	EMReadScreen program_check,          2, 5,33
	EMReadScreen federal_tax,            8, 7,57
	EMReadScreen state_tax,              8, 8,57
	EMReadScreen FICA_witheld,           8, 9,57
	EMReadScreen transportation_expense, 8, 10,57
	EMReadScreen meals_expense,          8, 11,57
	EMReadScreen uniform_expense,        8, 12,57		
	EMReadScreen tools_expense,          8, 13,57
	EMReadScreen dues_expenses,          8, 14,57
	EMReadScreen other_expenses,         8, 15,57
End IF

'Adds work expenses into dialog
If federal_tax = "________" then federal_tax = "0"
If state_tax = "________" then state_tax = "0"
If meals_expense = "________" then meals_expense = "0"
If uniform_expense = "________" then uniform_expense = "0"
If tools_expense = "________" then tools_expense = "0"
If dues_expenses = "________" then dues_expenses = "0"
If other_expenses = "________" then other_expenses = "0"


Total_taxes = abs(federal_tax) & abs(state_tax)
MsgBox "Make sure to load WKEX panel!"

Total_employment_expense = abs(meals_expenses) & abs(uniform_expense) & abs(tools_expense) & abs(dues_expenses) 
MsgBox "Check Total expenses!"

Call navigate_to_MAXIS_screen ("STAT", "JOBS"
EMReadScreen JOBS_panel_income, 7, 17, 68
JOBS_panel_income = trim(JOBS_panel_income)
JOBS_panel_income = round(JOBS_panel_income)

If abs(JOBS_panel_income) < "80" then 
	special_pers_allow = JOBS_panel_income
ELSE
	special_pers_allow = "80"
END IF 

'Shows the LTC_ICFDD_Fiater_dialog
dialog LTC_ICFDD_Fiater_dialog
cancel_confirmation

'Navigates to ELIG/HC.
Call navigate_to_MAXIS_screen("ELIG", "HC")


'Checks to see if MEMB 01 has HC, and puts an "x" there.  
EMReadScreen person_check, 4, 8, 26
EMWriteScreen "x", 8, 26
Transmit

'Turns on FIAT mode, checks to make sure it worked and sets the reason as "06".
PF9
EMReadScreen FIAT_check, 4, 10,20
If FIAT_check = "FIAT" then
  EMWriteScreen "06", 11,26
  transmit
End IF

'Defining the variables for the following search
ELIG_HC_row = 6
ELIG_HC_col = 1

'Determining the col variable based on the indicated footer month/year
EMSearch spousal_allocation_footer_month & "/" & spousal_allocation_footer_year, ELIG_HC_row, ELIG_HC_col
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

'Jumps into the budget screen (LBUD or SBUD)
transmit

For amt_of_months_to_do = 1 to budget_months

  'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This gathers the current month to see if we've moved past it later.
  EMReadScreen starting_bdgt_month, 5, 6, 14

  'Checks to see if this is an LBUD or SBUD. It'll stop if it's neither.
  EMReadScreen LBUD_check, 4, 3, 45
  If LBUD_check <> "LBUD" then
    EMReadScreen SBUD_check, 4, 3, 44
    If SBUD_check <> "SBUD" then script_end_procedure("This is not a method L or method S case.")
  End if

'Transmit to the next screen after putting an "x" on the Countable Earned Income (LBUD) screen 
 
	EMWriteScreen "x", 9,3
	Transmit
	EMWriteScreen "__________", 			7,42
	EMWriteScreen special_pers_allow,         7,42
	EMWriteScreen "__________",			8,42
	EMWriteScreen FICA_witheld,               8,42
	EMWriteScreen "__________",			9,42
	EMWriteScreen transportation_expense,     9,42
	EMWriteScreen "__________",			10,42
	EMWriteScreen Total_employment_expense,   10,42
	EMWriteScreen "__________",			11,42
	EMWriteScreen Total_taxes,                11,42
	EMWriteScreen "__________",			12,42
	EMWriteScreen other_expenses,             12,42
	TRANSMIT
	PF3
NEXT


call script_end_procedure("Sucess! The script is complete.")
