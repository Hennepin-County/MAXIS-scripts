'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EXPEDITED DETERMINATION.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("03/05/2020", "Added enhanced handling for the month the script will use to look at information. The best informaiton is provided in the month of application.", "Casey Love, Hennepin County")
call changelog_update("05/28/2019", "Updates to read the Expedited Screening case note.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
'connecting to MAXIS & searches for the case number
EMConnect ""

Call check_for_MAXIS(false)
call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Call find_user_name(worker_name)
If MAXIS_case_number <> "" Then



End If

'dialog to gather the Case Number and such
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 291, 95, "SNAP EXP Determination - Case Information"
  EditBox 85, 5, 60, 15, MAXIS_case_number
  DropListBox 85, 25, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", maxis_updated_yn
  EditBox 85, 55, 200, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 75, 50, 15
    CancelButton 235, 75, 50, 15
  Text 30, 10, 50, 10, "Case Number:"
  Text 20, 30, 65, 10, "MAXIS Updated?"
  Text 85, 40, 200, 10, "(All income, asset, and expense information entered in STAT)"
  Text 10, 60, 70, 10, "Sign your case note:"
EndDialog


Do
	Do
		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your worker signature"
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If maxis_updated_yn = "?" Then err_msg = err_msg & vbCr & "* Indicate if MAXIS has been updated with the known information about income, assets, and expenses"
		IF err_msg <> "" THEN MsgBox err_msg & vbCr & vbCr & "Please resolve this to continue"
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false				'loops until user passwords back in

exp_screening_note_found = False
snap_elig_results_read = False
do_we_have_applicant_id = False
developer_mode = False

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If
If MX_region = "TRAINING" Then developer_mode = True

' Call navigate_to_MAXIS_screen("STAT", "PROG")
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
If is_this_priv = True Then Call script_end_procedure("This case is PRIVILEGED and cannot be accessed. Request access to the case first and retry the script once you have access to the case.")

EMReadScreen case_pw, 7, 21, 17

EMReadScreen date_of_application, 8, 10, 33
EMReadScreen interview_date, 8, 10, 33

date_of_application = replace(date_of_application, " ", "/")
interview_date = replace(interview_date, " ", "/")
If interview_date = "__/__/__" Then interview_date = ""





Do
	Do
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 156, 70, "SNAP EXP Determination - Application Information"
		  EditBox 90, 5, 60, 15, date_of_application
		  EditBox 90, 25, 60, 15, interview_date
		  Text 20, 10, 65, 10, "Date of Application:"
		  Text 25, 30, 60, 10, "Date of Interview:"
		  ButtonGroup ButtonPressed
		    OkButton 45, 50, 50, 15
		    CancelButton 100, 50, 50, 15
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		If IsDate(date_of_application) = False Then
			err_msg = err_msg & vbCr & "* The date of application needs to be entered as a valid date."
		Else
			If DateDiff("d", interview_date, date) < 0 Then err_msg = err_msg & vbCr & "* The Application Date cannot be a Future date."
		End If
		If IsDate(interview_date) = False Then
			err_msg = err_msg & vbCr & "* The interview date needs to be entered as a valid date. An Expedited Determination cannot be completed without the interview."
		Else
			If DateDiff("d", interview_date, date) < 0 Then err_msg = err_msg & vbCr & "* The Interview Date cannot be a Future date."
		End If
		If IsDate(date_of_application) = True AND IsDate(interview_date) = True Then
			' MsgBox DateDiff("d", interview_date, date_of_application)
			If DateDiff("d", interview_date, date_of_application) > 0 Then err_msg = err_msg & vbCr & "* The Interview Date Cannot be before the Application Date."
		End If
		IF err_msg <> "" THEN MsgBox err_msg & vbCr & vbCr & "Please resolve this to continue"
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false

MAXIS_footer_month = DatePart("m", date_of_application)
MAXIS_footer_month = right("0"&MAXIS_footer_month, 2)

MAXIS_footer_year = right(DatePart("yyyy", date_of_application), 2)

expedited_package = MAXIS_footer_month & "/" & MAXIS_footer_year
If DatePart("d", date_of_application) > 15 Then
	second_month_of_exp_package = DateAdd("m", 1, date_of_application)
	NEXT_footer_month = DatePart("m", second_month_of_exp_package)
	NEXT_footer_month = right("0"&NEXT_footer_month, 2)

	NEXT_footer_year = right(DatePart("yyyy", second_month_of_exp_package), 2)
	expedited_package = expedited_package & " and " & NEXT_footer_month & "/" & NEXT_footer_year
End If
original_expedited_package = expedited_package

'Script is going to find information that was writen in an Expedited Screening case note using scripts
navigate_to_MAXIS_screen "CASE", "NOTE"

row = 1
col = 1
EMSearch "Received", row, col
IF row <> 0 THEN
	exp_screening_note_found = TRUE
	For look_for_right_note = 57 to 72
		EMReadScreen xfs_screen_note, 18, row, look_for_right_note
        xfs_screen_note = UCase(xfs_screen_note)
		IF xfs_screen_note = "CLIENT APPEARS EXP" or xfs_screen_note = "CLIENT DOES NOT AP" THEN
			exp_screening_note_found = TRUE	'IF the script found a case note with the NOTES - Expedited Screening format - it can find the information used
			IF look_for_right_note = 57 or look_for_right_note = 65 THEN
				EMReadScreen xfs_screening, 32, row, 42
			ElseIf look_for_right_note = 64 OR look_for_right_note = 72 THEN
				EMReadScreen xfs_screening, 31, row, 49
			End If
			EMWriteScreen "x", row, 3
			transmit
			Exit For
		END If
	Next
END IF

'Script is gathering the income/asset/expense information from the XFS Screening note
IF exp_screening_note_found = TRUE THEN
    EMReadScreen xfs_screening, 40, 4, 36
    xfs_screening = replace(xfs_screening, "~", "")
    xfs_screening = trim(xfs_screening)
	xfs_screening = UCase(xfs_screening)
	xfs_screening_display = xfs_screening & ""
	row = 1
	col = 1
	EMSearch "CAF 1 income", row, col
	EMReadScreen caf_one_income, 8, row, 42
	IF IsNumeric(caf_one_income) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_income = abs(caf_one_income)
	Else
		caf_one_income = 0
	End If

	row = 1
	col = 1
	EMSearch "CAF 1 liquid assets", row, col
	EMReadScreen caf_one_assets, 8, row, 42
	If IsNumeric(caf_one_assets)= True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_assets = caf_one_assets * 1
	Else
		caf_one_assets = 0
	End If

	caf_one_resources = caf_one_income + caf_one_assets	'Totaling the amounts for the case note

	row = 1
	col = 1
	EMSearch "CAF 1 rent", row, col
	EMReadScreen caf_one_rent, 8, row, 42
	IF IsNumeric(caf_one_rent) = True Then 		'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_rent = abs(caf_one_rent)
	Else
		caf_one_rent = 0
	End If

	row = 1
	col = 1
	EMSearch "Utilities (AMT", row, col
	EMReadScreen caf_one_utilities, 8, row, 42
	If IsNumeric(caf_one_utilities) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_utilities = abs(caf_one_utilities)
	Else
		caf_one_utilities = 0
	End If

	caf_one_expenses = caf_one_rent + caf_one_utilities		'Totaling the amounts for a case note

	'The script not adjusts the format so it looks nice
	caf_one_income = FormatCurrency(caf_one_income)
	caf_one_assets = FormatCurrency(caf_one_assets)
	caf_one_rent = FormatCurrency(caf_one_rent)
	caf_one_utilities = FormatCurrency(caf_one_utilities)
	caf_one_resources = FormatCurrency(caf_one_resources)
	caf_one_expenses = FormatCurrency(caf_one_expenses)
	PF3
End IF

determined_utilities = ""
If maxis_updated_yn = "No" Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)

If maxis_updated_yn = "Yes" Then
	'Script now goes to ELIG to find what the income/expesnse that are being used are to autofill the dialog
	navigate_to_MAXIS_screen "ELIG", "FS"
	EMReadScreen elig_screen_check, 4, 3, 48
	IF elig_screen_check = "FSPR" Then
		snap_elig_results_read = True
		transmit
		EMReadScreen is_elig_XFS, 17, 4, 3
		IF is_elig_XFS = "EXPEDITED SERVICE" THEN 	'Determines if MAXIS thinks the case is Expedited
			is_elig_XFS = TRUE
		ELSE
			is_elig_XFS = FALSE
		END IF
		is_elig_XFS = is_elig_XFS & ""
		'MsgBox is_elig_XFS

		transmit		'Finding Income and formating it
		EMReadScreen elig_gross_income, 9, 7, 72
		elig_gross_income = trim(elig_gross_income)
		elig_gross_income = abs(elig_gross_income)
		transmit

		'Finding the shelter and utility expenses and combining them and formating them
		EMReadScreen elig_heat, 3, 9, 31
		IF elig_heat = "   " THEN elig_heat = 0
		elig_heat = trim(elig_heat)
		elig_heat = abs(elig_heat)

		EMReadScreen elig_electric, 3, 8, 31
		IF elig_electric = "   " THEN elig_electric = 0
		elig_electric = trim(elig_electric)
		elig_electric = abs(elig_electric)

		EMReadScreen elig_phone, 2, 11, 32
		IF elig_phone = "  " THEN elig_phone = 0
		elig_phone = trim(elig_phone)
		elig_phone = abs(elig_phone)

		EMReadScreen elig_rent, 5, 5, 29
		IF elig_rent = "     " THEN elig_rent = 0
		elig_rent = trim(elig_rent)
		elig_rent = abs(elig_rent)

		EMReadScreen elig_tax, 5, 6, 29
		IF elig_tax = "     " THEN elig_tax = 0
		elig_tax = trim(elig_tax)
		elig_tax = abs(elig_tax)

		EMReadScreen elig_ins, 5, 7, 29
		IF elig_ins = "     " THEN elig_ins = 0
		elig_ins = trim(elig_ins)
		elig_ins = abs(elig_ins)

		EMReadScreen elig_other_exp, 5, 12, 29
		IF elig_other_exp = "     " THEN elig_other_exp = 0
		elig_other_exp = trim(elig_other_exp)
		elig_other_exp = abs(elig_other_exp)

		IF elig_heat <> 0 THEN
			elig_util = elig_heat
		ELSE
			elig_util = elig_electric + elig_phone
		END IF

		elig_shel = elig_rent + elig_tax + elig_ins + elig_other_exp
	End If

	'Going to STAT for asset information
	navigate_to_MAXIS_screen "STAT", "PNLR"
	For pnlr_row = 3 to 19
		EMReadScreen asset_panel_type, 4, pnlr_row, 5
		IF asset_panel_type = "CASH" THEN
			EMReadScreen asset_listed, 6, pnlr_row, 26
		ELSEIF asset_panel_type = "ACCT" THEN
			EMReadScreen asset_listed, 6, pnlr_row, 31
		Else
			asset_listed = 0
		End If
		asset_amount = asset_amount + abs(trim(asset_listed))
	Next
End If

'Prepping variables to fill in the edit boxes
' determined_income = elig_gross_income & ""
' determined_assets = asset_amount & ""
' determined_shel = elig_shel & ""
' determined_utilities = elig_util & ""

'-------------------------------------------------------------------------------------------------DIALOG
next_btn = 2
finish_btn = 3

amounts_btn 		= 10
determination_btn 	= 20
review_btn 			= 30

income_calc_btn								= 100
asset_calc_btn								= 110
housing_calc_btn							= 120
utility_calc_btn							= 130
snap_active_in_another_state_btn			= 140
case_previously_had_postponed_verifs_btn	= 150
household_in_a_facility_btn					= 160

knowledge_now_support_btn		= 500
te_02_10_01_btn					= 510

hsr_manual_expedited_snap_btn 	= 1000
hsr_snap_applications_btn		= 1100
ryb_exp_identity_btn			= 1200
ryb_exp_timeliness_btn			= 1300
sir_exp_flowchart_btn			= 1400
cm_04_04_btn					= 1500
cm_04_06_btn					= 1600
ht_id_in_solq_btn				= 1700
cm_04_12_btn					= 1800
temp_prog_changes_ebt_card_btn 	= 1900

const account_type_const	= 0
const account_owner_const	= 1
const bank_name_const		= 2
const account_amount_const	= 3
const account_notes_const 	= 4

Dim ACCOUNTS_ARRAY
ReDim ACCOUNTS_ARRAY(account_notes_const, 0)

function app_month_income_detail(determined_income)
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Determination of Income in Month of Application"
	  ButtonGroup ButtonPressed
		Text 10, 5, 435, 10, "These questions will help you to guide the resident through understanding what income we need to count for the month of application."
		Text 10, 20, 150, 10, "FIRST - Explain to the resident these things:"
		Text 25, 30, 410, 10, "- Income in the App Month is used to determine if we can get your some SNAP benefits right away - an EXPEDITED Issuance."
		Text 25, 40, 410, 10, "- We just need a best estimate of this income - it doesn't have to be exact. There is no penalty for getting this detail incorrect."
		Text 25, 50, 410, 10, "- I can help you walk through your income sources."
		Text 25, 60, 350, 10, "-  We need you to answer these questions to complete the interview for your application for SNAP benefits."
		GroupBox 5, 75, 440, 105, "JOBS Income: For every Job in the Household"
		Text 15, 90, 200, 10, "How many paychecks have you received in MM/YY so far?"
		Text 30, 105, 170, 10, "How much were all of the checks for, before taxes?"
		Text 15, 120, 215, 10, "How many paychecks do you still expect to receive in MM/YY?"
		Text 30, 135, 225, 10, "How many hours a week did you or will you work for these checks?"
		Text 30, 150, 120, 10, "What is your rate of pay per hour?"
		Text 30, 165, 255, 10, "Do you get tips/commission/bonuses? How much do you expect those to be?"
		GroupBox 5, 185, 440, 90, "BUSI Income: For each self employment in the Household"
		Text 15, 200, 235, 10, "How much do you typically receive in a month of this self employment?"
		Text 15, 215, 275, 10, "Is your self employment based on a contract or contracts? And how are they paid?"
		Text 15, 230, 305, 10, "If this is hard to determine, how much to you make in any other period (year, week, quarter)?"
		Text 30, 245, 200, 10, "Is this consistent over the period or from period to period?"
		Text 30, 260, 115, 10, "If it is not, what are the variations?"
		GroupBox 5, 280, 440, 45, "UNEA Income: For each other source of income in the Household"
		Text 15, 295, 200, 10, "How often and how much do you receive from each source?"
		Text 15, 310, 230, 10, "If this is irregular, what have you gotten for the past couple months?"
		Text 5, 330, 380, 10, "After calculating all of these income questions, repeat the amount and each source and confirm that it seems close."
		PushButton 395, 330, 50, 15, "Return", return_btn
	EndDialog

	dialog Dialog1

	determined_income = determined_income & ""
	ButtonPressed = income_calc_btn
end function
function app_month_asset_detail(determined_assets, cash_amount_yn, bank_account_yn, ACCOUNTS_ARRAY)
	return_btn = 5001
	enter_btn = 5002
	add_another_btn = 5003
	remove_one = 5004

	determined_assets = 0
	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 271, 135, "Determination of Assets in Month of Application"
		  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
		  GroupBox 10, 25, 255, 40, "Cash"
		  Text 25, 45, 155, 10, "Does the household have any Cash Savings?"
		  DropListBox 180, 40, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", cash_amount_yn
		  GroupBox 10, 70, 255, 40, "Accounts"
		  Text 20, 90, 190, 10, "Does anyone in the household have any Bank Accounts?"
		  DropListBox 210, 85, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", bank_account_yn
		  ButtonGroup ButtonPressed
		    PushButton 215, 115, 50, 15, "Enter", enter_btn
		EndDialog

		dialog Dialog1

		If cash_amount_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has CASH."
		If bank_account_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has A BANK ACCOUNT."

		If prvt_err_msg <> "" Then MsgBox prvt_err_msg
	Loop until prvt_err_msg = ""

	Do
		prvt_err_msg = ""

		If cash_amount_yn = "No" Then cash_grp_len = 30
		If cash_amount_yn = "Yes" Then cash_grp_len = 50
		If bank_account_yn = "No" Then acct_grp_len = 30
		If bank_account_yn = "Yes" Then acct_grp_len = 60 + (UBound(ACCOUNTS_ARRAY, 2) + 1) * 20
		dlg_len = 55 + cash_grp_len + acct_grp_len

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 351, dlg_len, "Determination of Assets in Month of Application"
		  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
		  GroupBox 10, 25, 220, cash_grp_len, "Cash"
		  If cash_amount_yn = "Yes" Then
			  Text 20, 40, 155, 10, "This household HAS Cash Savings."
			  Text 20, 55, 150, 10, "How much in total does the household have?"
			  EditBox 175, 50, 45, 15, cash_amount
			  y_pos = 80
		  Else
			  Text 20, 40, 155, 10, "This household does NOT have Cash."
			  y_pos = 60
		  End If
		  GroupBox 10, y_pos, 335, acct_grp_len, "Accounts"
		  y_pos = y_pos + 15
		  If bank_account_yn = "Yes" Then
			  Text 20, y_pos, 190, 10, "This household HAS Bank Accounts."
			  y_pos = y_pos + 15
			  Text 20, y_pos, 50, 10, "Account Type"
			  Text 90, y_pos, 70, 10, "Owner of Account"
			  Text 180, y_pos, 45, 10, "Bank Name"
			  Text 285, y_pos, 35, 10, "Amount"
			  y_pos = y_pos + 15

			  For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
				  ACCOUNTS_ARRAY(account_amount_const, the_acct) = ACCOUNTS_ARRAY(account_amount_const, the_acct) & ""
				  DropListBox 20, y_pos, 60, 45, "Select One..."+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Other", ACCOUNTS_ARRAY(account_type_const, the_acct)
				  EditBox 90, y_pos, 85, 15, ACCOUNTS_ARRAY(account_owner_const, the_acct)
				  EditBox 180, y_pos, 100, 15, ACCOUNTS_ARRAY(bank_name_const, the_acct)
				  EditBox 285, y_pos, 50, 15, ACCOUNTS_ARRAY(account_amount_const, the_acct)
				  y_pos = y_pos + 20
			  Next
		  Else
		  	  Text 20, y_pos, 155, 10, "This household does NOT have Bank Accounts."
		  End If
		  ButtonGroup ButtonPressed
		    If bank_account_yn = "Yes" Then PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_btn
		    If bank_account_yn = "Yes" Then PushButton 275, y_pos, 60, 10, "REMOVE ONE", remove_one
			PushButton 295, dlg_len - 20, 50, 15, "Return", return_btn
		EndDialog

		dialog Dialog1

		last_acct_tiem = UBound(ACCOUNTS_ARRAY, 2)
		If ButtonPressed = add_another_btn Then
			last_acct_tiem = last_acct_tiem + 1
			ReDim Preserve ACCOUNTS_ARRAY(account_notes_const, last_acct_tiem)
		End If
		If ButtonPressed = remove_one Then
			last_acct_tiem = last_acct_tiem - 1
			ReDim Preserve ACCOUNTS_ARRAY(account_notes_const, last_acct_tiem)
		End If

		cash_amount = trim(cash_amount)
		If cash_amount <> "" And IsNumeric(cash_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Cash Amount as a number."

		For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
			ACCOUNTS_ARRAY(account_amount_const, the_acct) = trim(ACCOUNTS_ARRAY(account_amount_const, the_acct))
			If ACCOUNTS_ARRAY(account_amount_const, the_acct) <> "" And IsNumeric(ACCOUNTS_ARRAY(account_amount_const, the_acct)) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Bank Account amounts as a member."
			If ACCOUNTS_ARRAY(account_type_const, the_acct)	= "Select One..." Then prvt_err_msg = prvt_err_msg & vbCr & "* Select the Bank Account type."
		Next
	Loop Until ButtonPressed = return_btn

	If cash_amount = "" Then cash_amount = 0
	cash_amount = cash_amount * 1
	For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
		If ACCOUNTS_ARRAY(account_amount_const, the_acct) = "" Then ACCOUNTS_ARRAY(account_amount_const, the_acct) = 0
		ACCOUNTS_ARRAY(account_amount_const, the_acct) = ACCOUNTS_ARRAY(account_amount_const, the_acct) * 1
		determined_assets = determined_assets + ACCOUNTS_ARRAY(account_amount_const, the_acct)
	Next
	determined_assets = determined_assets + cash_amount

	determined_assets = determined_assets & ""
	ButtonPressed = asset_calc_btn
end function
function app_month_housing_detail(determined_shel)
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Determination of Housing Cost in Month of Application"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1


	determined_shel = determined_shel & ""
	ButtonPressed = housing_calc_btn
end function
function app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
	calculate_btn = 5000
	return_btn = 5001
	determined_utilities = 0
	If heat_expense = True then heat_checkbox = checked
	If ac_expense = True then ac_checkbox = checked
	If electric_expense = True then electric_checkbox = checked
	If phone_expense = True then phone_checkbox = checked
	If none_expense = True then none_checkbox = checked

	Do
		current_utilities = all_utilities

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 246, 175, "Determination of Utilities in Month of Application"
		  CheckBox 30, 45, 50, 10, "Heat", heat_checkbox
		  CheckBox 30, 60, 65, 10, "Air Conditioning", ac_checkbox
		  CheckBox 30, 75, 50, 10, "Electric", electric_checkbox
		  CheckBox 30, 90, 50, 10, "Phone", phone_checkbox
		  CheckBox 30, 105, 50, 10, "NONE", none_checkbox
		  ButtonGroup ButtonPressed
		    PushButton 170, 105, 65, 15, "Calculate", calculate_btn
		    PushButton 170, 155, 65, 15, "Return", return_btn
		  Text 10, 10, 235, 10, "Check the boxes for each utility the household is responsible to pay:"
		  GroupBox 15, 30, 225, 95, "Utilities"
		  Text 150, 45, 50, 10, "$ " & determined_utilities
		  Text 150, 60, 35, 35, all_utilities
		  Text 15, 135, 225, 20, "Remember, this expense could be shared, they are still considered responsible to pay and we count the WHOLE standard."
		EndDialog

		dialog Dialog1

		some_vs_none_discrepancy = False
		If (heat_checkbox = checked OR ac_checkbox = checked OR electric_checkbox = checked OR phone_checkbox = checked) AND none_checkbox = checked Then some_vs_none_discrepancy = True
		If some_vs_none_discrepancy = True Then MsgBox "Attention:" & vbCr & vbCr & "You have selected NONE and selected at least one other utility expense. If it is NONE, then no other utilities should be checked."

		all_utilities = ""
		If heat_checkbox = checked Then all_utilities = all_utilities & ", Heat"
		If ac_checkbox = checked Then all_utilities = all_utilities & ", AC"
		If electric_checkbox = checked Then all_utilities = all_utilities & ", Electric"
		If phone_checkbox = checked Then all_utilities = all_utilities & ", Phone"
		If none_checkbox = checked Then all_utilities = all_utilities & ", None"
		If left(all_utilities, 2) = ", " Then all_utilities = right(all_utilities, len(all_utilities) - 2)

		If all_utilities = current_utilities AND ButtonPressed = -1 Then ButtonPressed = return_btn

		determined_utilities = 0
		If heat_checkbox = checked OR ac_checkbox = checked Then
			determined_utilities = determined_utilities + 496
		Else
			If electric_checkbox = checked Then determined_utilities = determined_utilities + 154
			If phone_checkbox = checked Then determined_utilities = determined_utilities + 56
		End If

	Loop Until ButtonPressed = return_btn And some_vs_none_discrepancy = False

	heat_expense = False
	ac_expense = False
	electric_expense = False
	phone_expense = False
	none_expense = False

	If heat_checkbox = checked Then heat_expense = True
	If ac_checkbox = checked Then ac_expense = True
	If electric_checkbox = checked Then electric_expense = True
	If phone_checkbox = checked Then phone_expense = True
	If none_checkbox = checked Then none_expense = True

	ButtonPressed = utility_calc_btn
end function
function determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
	determined_income = trim(determined_income)
	If determined_income = "" Then determined_income = 0
	determined_income = determined_income * 1

	determined_assets = trim(determined_assets)
	If determined_assets = "" Then determined_assets = 0
	determined_assets = determined_assets * 1

	determined_shel = trim(determined_shel)
	If determined_shel = "" Then determined_shel = 0
	determined_shel = determined_shel * 1

	determined_utilities = trim(determined_utilities)
	If determined_utilities = "" Then determined_utilities = 0
	determined_utilities = determined_utilities * 1

	calculated_resources = determined_income + determined_assets
	calculated_expenses = determined_shel + determined_utilities

	calculated_low_income_asset_test = False
	calculated_resources_less_than_expenses_test = False
	is_elig_XFS = False

	If determined_income < 150 AND determined_assets <= 100 Then calculated_low_income_asset_test = True
	If calculated_resources < calculated_expenses Then calculated_resources_less_than_expenses_test = True

	If calculated_low_income_asset_test = True OR calculated_resources_less_than_expenses_test = True Then is_elig_XFS = True

	determined_income = determined_income & ""
	determined_assets = determined_assets & ""
	determined_shel = determined_shel & ""
	determined_utilities = determined_utilities & ""
end function

' BeginDialog Dialog1, 0, 0, 381, 295, "Case Received SNAP in Another State"
'   DropListBox 255, 55, 110, 45, "", other_snap_state
'   EditBox 255, 75, 60, 15, other_state_reported_benefit_end_date
'   CheckBox 40, 95, 320, 10, "Check here is resident reports the benefits are NOT ended or it is UKNOWN if they are ended.", other_state_benefits_not_ended_checkbox
'   DropListBox 255, 110, 60, 45, "", List2
'   EditBox 255, 130, 60, 15, other_state_verified_benefit_end_date
'   ButtonGroup ButtonPressed
'     PushButton 325, 170, 50, 15, "Calculate", calc_other_state_benefit_btn
'   Text 10, 10, 365, 10, "If a Household has received SNAP in another state, we may still be able to issue Expedited SNAP in Minnesota. "
'   Text 10, 25, 320, 10, "Complete the following information to get guidance on handling cases with SNAP in another State:"
'   GroupBox 10, 45, 365, 120, "Other State Benefits"
'   Text 20, 60, 235, 10, "What State is the Household / Resident receiving SNAP benefits from?"
'   Text 40, 80, 215, 10, "When is the resident REPORTING benefits ending in this state?"
'   Text 20, 115, 230, 10, "Have you called the other state to confirm / discover the SNAP status?"
'   Text 20, 135, 230, 10, "What end date has been confirmed / verified for the other state SNAP?"
'   GroupBox 10, 190, 365, 80, "Resolution"
'   Text 20, 205, 205, 20, "SNAP should be denied as the other state end date is AFTER the 30 day processing period of the application in MN."
'   Text 245, 205, 120, 10, "Date of Application:"
'   Text 255, 215, 110, 10, "End Of Benefits:"
'   Text 30, 230, 120, 10, "SNAP Denial Date: "
'   Text 30, 245, 335, 10, "Denial Reason:"
'   ButtonGroup ButtonPressed
'     PushButton 325, 275, 50, 15, "Return", Button3
' EndDialog


function snap_in_another_state_detail(date_of_application, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
	original_snap_denial_date = snap_denial_date
	original_snap_denial_reason = snap_denial_explain
	calculation_done = False
	other_state_benefits_openended = False
	action_due_to_out_of_state_benefits = ""
	' other_snap_state = "MN - Minnesota"
	day_30_from_application = DateAdd("d", 30, date_of_application)
	calculate_btn = 5000
	return_btn = 5001

	Do
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			If calculation_done = False Then BeginDialog Dialog1, 0, 0, 381, 190, "Case Received SNAP in Another State"
			If calculation_done = True Then BeginDialog Dialog1, 0, 0, 381, 295, "Case Received SNAP in Another State"
			  DropListBox 255, 55, 110, 45, "Select One..."+chr(9)+state_list, other_snap_state
			  EditBox 255, 75, 60, 15, other_state_reported_benefit_end_date
			  CheckBox 40, 95, 320, 10, "Check here is resident reports the benefits are NOT ended or it is UKNOWN if they are ended.", other_state_benefits_not_ended_checkbox
			  DropListBox 255, 110, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", other_state_contact_yn
			  EditBox 255, 130, 60, 15, other_state_verified_benefit_end_date
			  ButtonGroup ButtonPressed
			    PushButton 325, 170, 50, 15, "Calculate", calculate_btn
			  Text 10, 10, 365, 10, "If a Household has received SNAP in another state, we may still be able to issue Expedited SNAP in Minnesota. "
			  Text 10, 25, 320, 10, "Complete the following information to get guidance on handling cases with SNAP in another State:"
			  GroupBox 10, 45, 365, 120, "Other State Benefits"
			  Text 20, 60, 235, 10, "What State is the Household / Resident receiving SNAP benefits from?"
			  Text 40, 80, 215, 10, "When is the resident REPORTING benefits ending in this state?"
			  Text 20, 115, 230, 10, "Have you called the other state to confirm / discover the SNAP status?"
			  Text 20, 135, 230, 10, "What end date has been confirmed / verified for the other state SNAP?"

			  If calculation_done = True Then
				  GroupBox 10, 190, 365, 80, "Resolution"
				  If action_due_to_out_of_state_benefits = "DENY" Then
					  Text 20, 205, 205, 20, "SNAP should be denied as the other state end date is AFTER the 30 day processing period of the application in MN."
					  Text 245, 205, 120, 10, "Date of Application: " & date_of_application
					  If IsDate(other_state_verified_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_verified_benefit_end_date
					  ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_reported_benefit_end_date
					  End If
					  Text 30, 230, 120, 10, "SNAP Denial Date: " & snap_denial_date
					  Text 30, 240, 335, 30, "Denial Reason: " & snap_denial_explain
				  ElseIf action_due_to_out_of_state_benefits = "APPROVE" Then
					  Text 20, 205, 205, 20, "SNAP should be APPROVEED "
					  Text 245, 205, 120, 10, "Date of Application: " & date_of_application
					  Text 25, 215, 175, 10, "Eligibility can start in MN as of " & mn_elig_begin_date
					  If other_state_contact_yn <> "Yes" Then
					  	Text 20, 230, 340, 10, "Verification of out of state eligibility end can be postponed "
						Text 20, 240, 340, 10, "We should make reasonable efforts to obtain verification so, "
						Text 20, 250, 340, 10, "it is best to attempt a call to the other state right away for verification."
					  End If
				  ElseIf action_due_to_out_of_state_benefits = "FOLLOW UP" Then
					  Text 20, 205, 205, 20, "You must connect with the other state to determine when the benefits have ended or IF the benefits will end."
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 325, 275, 50, 15, "Return", return_btn
			  End If
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then Exit Do

			If IsDate(other_state_reported_benefit_end_date) = False AND other_state_benefits_not_ended_checkbox = unchecked Then prvt_err_msg = prvt_err_msg & vbCr & "* We cannot complete the calculation if a reported end date has not been entered."
			If IsDate(other_state_reported_benefit_end_date) = True AND other_state_benefits_not_ended_checkbox = checked Then prvt_err_msg = prvt_err_msg & vbCr & "* You have entered an end date AND indicated the benefits have not ended by checking the box. Please select only one."

			If IsDate(other_state_reported_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_reported_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the REPORTED end date."
			End If
			If IsDate(other_state_verified_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_verified_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the VERIFIED end date."
			End If
			If prvt_err_msg <> "" Then
				MsgBox prvt_err_msg
				calculation_done = False
			End If

		Loop until prvt_err_msg = ""

		If ButtonPressed = 0 Then
			calculation_done = False
			Exit Do
		End If

		calculation_done = True
		If other_snap_state = "NB - MN Newborn" OR other_snap_state = "MN - Minnesota" OR other_snap_state = "Select One..." OR other_snap_state = "FC - Foreign Country" OR other_snap_state = "UN - Unknown" Then other_snap_state = ""
		If IsDate(other_state_verified_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_verified_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
				snap_denial_date = date
				If other_snap_state = "" Then snap_denial_explain = snap_denial_explain & "; Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days."
				If other_snap_state <> "" Then snap_denial_explain = snap_denial_explain & "; Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days."
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_verified_benefit_end_date)
				If DateDiff("d", mn_elig_begin_date, date_of_application) > 0 Then
					mn_elig_begin_date = date_of_application
					expedited_package = original_expedited_package
				Else
					MN_elig_month = DatePart("m", mn_elig_begin_date)
					MN_elig_month = right("0"&MN_elig_month, 2)
					MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
					expedited_package = MN_elig_month & "/" & MN_elig_year
				End If
			End If
		ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_reported_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
				snap_denial_date = date
				If other_snap_state = "" Then snap_denial_explain = snap_denial_explain & "; Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days."
				If other_snap_state <> "" Then snap_denial_explain = snap_denial_explain & "; Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days."
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_reported_benefit_end_date)
				If DateDiff("d", mn_elig_begin_date, date_of_application) > 0 Then
					mn_elig_begin_date = date_of_application
					expedited_package = original_expedited_package
				Else
					MN_elig_month = DatePart("m", mn_elig_begin_date)
					MN_elig_month = right("0"&MN_elig_month, 2)
					MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
					expedited_package = MN_elig_month & "/" & MN_elig_year
				End If
			End If
		ElseIf other_state_benefits_not_ended_checkbox = checked Then
			action_due_to_out_of_state_benefits = "FOLLOW UP"
			other_state_benefits_openended = True
		End If
		If action_due_to_out_of_state_benefits <> "DENY" Then
			snap_denial_date = original_snap_denial_date
			snap_denial_explain = original_snap_denial_reason
		End If
		If action_due_to_out_of_state_benefits <> "APPROVE" Then expedited_package = original_expedited_package
	Loop until ButtonPressed = return_btn
	ButtonPressed = snap_active_in_another_state_btn
end function

function previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
	review_btn = 5005
	return_btn = 5001
	prev_post_verif_assessment_done = True
	case_has_previously_postponed_verifs_that_prevent_exp_snap = False

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 446, 160, "Case Previously Received EXP SNAP with Postponed Verifications"
		  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
		  Text 10, 30, 125, 10, "This does not apply to cases where:"
		  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
		  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
		  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
		  Text 15, 85, 175, 15, "What is the DATE OF APPLICATION for the Expedited Approval that had Postponed Verifications?"
		  EditBox 195, 85, 50, 15, previous_date_of_application
		  Text 275, 110, 115, 10, "Are these verifications mandatory?"
		  DropListBox 400, 105, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_verifs_mandatory_yn
		  Text 15, 110, 175, 10, "List the verifications that were previously postponed:"
		  EditBox 15, 120, 425, 15, prev_verif_list
		  Text 15, 145, 220, 10, "Does the case have Postponed Verifications for THIS Application?"
		  DropListBox 235, 140, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", curr_verifs_postponed_yn
		  ButtonGroup ButtonPressed
		    PushButton 390, 140, 50, 15, "Review", review_btn
		EndDialog

		dialog Dialog1

		If ButtonPressed = 0 Then
			prev_post_verif_assessment_done = False
			Exit Do
		End If

		prev_verif_list = trim(prev_verif_list)
		If IsDate(previous_date_of_application) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the date of application from the last time this case received an Expedited SNAP approval WITH Postponed Verifications."
		If prev_verifs_mandatory_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* You must review the verifications that were previously postponed and enter them here."
		If prev_verif_list = "" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the verifications that were previously postponed and indicate if any of them were mandatory."
		If curr_verifs_postponed_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Indicate if the CURRENT application has verifications required that would need to be postponed to approve the Expedited SNAP."

		If prvt_err_msg <> "" Then MsgBox prvt_err_msg
	Loop until prvt_err_msg = ""

	PREVIOUS_footer_month = DatePart("m", previous_date_of_application)
	PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

	PREVIOUS_footer_year = right(DatePart("yyyy", previous_date_of_application), 2)

	If DatePart("d", previous_date_of_application) > 15 Then
		second_month_of_previous_exp_package = DateAdd("m", 1, previous_date_of_application)
		PREVIOUS_footer_month = DatePart("m", second_month_of_previous_exp_package)
		PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

		PREVIOUS_footer_year = right(DatePart("yyyy", second_month_of_previous_exp_package), 2)
	End If
	previous_expedited_package = PREVIOUS_footer_month & "/" & PREVIOUS_footer_year

	ask_more_questions = False
	If IsDate(previous_date_of_application) = True AND prev_verifs_mandatory_yn = "Yes" AND curr_verifs_postponed_yn = "Yes" Then ask_more_questions = True
	If ask_more_questions = True Then
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 436, 110, "Case Previously Received EXP SNAP with Postponed Verifications"
			  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
			  Text 10, 30, 125, 10, "This does not apply to cases where:"
			  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
			  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
			  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
			  Text 10, 80, 180, 10, "Did the case get approved for any SNAP after " & previous_expedited_package & "?"
			  DropListBox 195, 75, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", ongoing_snap_approved_yn
			  Text 20, 95, 170, 10, "Check ECF, are the postponed verifications on file?"
			  DropListBox 195, 90, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_post_verifs_recvd_yn
			  ButtonGroup ButtonPressed
			    PushButton 380, 90, 50, 15, "Review", review_btn

			  Text 10, 270, 280, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
			  Text 15, 295, 210, 10, "- Provide all verifications that were postponed and mandatory."
			  Text 15, 305, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
			  Text 20, 315, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then
				prev_post_verif_assessment_done = False
				Exit Do
			End If

			If ongoing_snap_approved_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review MAXIS and determine if SNAP was approved after the last month of the expedited package (" & previous_expedited_package & "). If it was, the case met all requirements to gain SNAP eligibility."
			If prev_post_verifs_recvd_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the ECF case file and see if the mandatory postponed verifications were ever received, even if SNAP was not approved."

			If prvt_err_msg <> "" Then MsgBox prvt_err_msg
		Loop until prvt_err_msg = ""
	End If

	If ask_more_questions = False OR ongoing_snap_approved_yn = "Yes" OR prev_post_verifs_recvd_yn = "Yes" Then
		Dialog1 = ""
		y_pos = 85

		BeginDialog Dialog1, 0, 0, 436, 120, "Case Previously Received EXP SNAP with Postponed Verifications"
		  GroupBox 10, 10, 415, 55, "EXPEDITED CAN BE APPROVED"
		  Text 25, 25, 100, 10, "Based on this case situation"
		  Text 30, 35, 325, 10, "This case CAN be approved for Expedited without a delay due to Previous Postponed Verifications."
		  Text 35, 45, 285, 10, "(There may be another reason for delay, complete the rest of the review to determine.)"
		  Text 15, 75, 45, 10, "Explanation:"
		  If prev_verifs_mandatory_yn = "No" Then
			  Text 15, y_pos, 350, 10, "The previously postponed verifications were not mandatory, so case met all SNAP eligibility criteria."
			  y_pos = y_pos + 10
		  End If
		  If curr_verifs_postponed_yn = "No" Then
			  Text 15, y_pos, 350, 10, "There are no verifications that are required and being postponed now, so case meets all SNAP eligibility criteria."
			  y_pos = y_pos + 10
		  End If
		  If ongoing_snap_approved_yn = "Yes" Then
			  Text 15, y_pos, 350, 10, "Case was approved regular SNAP after the expedited package time, so case met all SNAP eligibility criteria."
			  y_pos = y_pos + 10
		  End If
		  If prev_post_verifs_recvd_yn = "Yes" Then
			  Text 50, y_pos, 350, 10, "The postponed verifications have been received, which meets the requirement to receive another posponed verification approval package."
			  y_pos = y_pos + 10
		  End If
		  ButtonGroup ButtonPressed
		    PushButton 380, 100, 50, 15, "Update", update_btn
		EndDialog

		dialog Dialog1

	End If

	If ask_more_questions = True AND ongoing_snap_approved_yn = "No" AND prev_post_verifs_recvd_yn = "No" Then
		case_has_previously_postponed_verifs_that_prevent_exp_snap = True

		BeginDialog Dialog1, 0, 0, 291, 145, "Case Previously Received EXP SNAP with Postponed Verifications"
		  GroupBox 5, 5, 280, 60, "EXPEDITED APPROVAL MUST BE DELAYED"
		  Text 20, 20, 100, 10, "Based on this case situation"
		  Text 25, 30, 195, 10, "This case CANNOT be approved for Expedited at this time."
		  Text 30, 40, 235, 20, "The case would require postponing verifications when we already have allowed for postponed verifications that have not been received."
		  Text 10, 70, 275, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
		  Text 15, 95, 210, 10, "- Provide all verifications that were postponed and mandatory."
		  Text 15, 105, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
		  Text 20, 115, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
		  ButtonGroup ButtonPressed
		    PushButton 235, 125, 50, 15, "Update", update_btn
		EndDialog

		dialog Dialog1
	End If

	If case_has_previously_postponed_verifs_that_prevent_exp_snap = False Then delay_explanation = replace(delay_explanation, "Approval cannot be completed as case has postponed verifications when postpone verifications were previously allowed and not provided, nor has the case meet 'ongoing SNAP' eligibility", "")
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then delay_explanation = delay_explanation & "; Approval cannot be completed as case has postponed verifications when postpone verifications were previously allowed and not provided, nor has the case meet 'ongoing SNAP' eligibility."

	ButtonPressed = case_previously_had_postponed_verifs_btn
end function

function household_in_a_facility_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Case Previously Received EXP SNAP with Postponed Verifications"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

	ButtonPressed = household_in_a_facility_btn
end function

function send_support_email_to_KN()

	email_subject = "Assistance with Case at SNAP Application - Possible EXP"
	If developer_mode = True Then email_subject = "TESTING RUN - " & email_subject & " - can be deleted"

	email_body = "I am completing a SNAP Expedited Determination." & vbCr & vbCr
	email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr & vbCr
	email_body = email_body & "Amounts currently entered at the Determination:" & vbCr
	email_body = email_body & "Income: $ " & determined_income & vbCr
	email_body = email_body & "Assets: $ " & determined_assets & vbCr
	email_body = email_body & "Housing: $ " & determined_shel & vbCr
	email_body = email_body & "Utilities: $ " & determined_utilities & vbCr & vbCr
	email_body = email_body & "Script Calculations:" & vbCr
	If is_elig_XFS = True Then email_body = email_body & "Case appears EXPEDITED." & vbCr
	If is_elig_XFS = False Then email_body = email_body & "Case does NOT appear Expedtied." & vbCr
	email_body = email_body & "Unit has less than $150 monthly Gross Income AND $100 or less in assets: " & calculated_low_income_asset_test & vbCr
	email_body = email_body & "Unit's combined resources are less than housing expense: " & calculated_resources_less_than_expenses_test & vbCr & vbCr
	email_body = email_body & "Case Dates/Timelines:" & vbCr
	email_body = email_body & "Date of Application: " & date_of_application & vbCr
	email_body = email_body & "Date of Interview: " & interview_date & vbCr
	email_body = email_body & "Date of Approval: " & approval_date & " (or planned date of approval)" & vbCr
	email_body = email_body & "Processing Delay Explanation: " & delay_explanation & vbCr
	email_body = email_body & "SNAP Denial Date: " & snap_denial_date & vbCr
	email_body = email_body & "Denial Explanation: " & snap_denial_explain & vbCr & vbCr
	email_body = email_body & "Other Information:" & vbCr
	If applicant_id_on_file_yn <> "" AND applicant_id_on_file_yn <> "?" Then email_body = email_body & "Is there an ID on file for the applicant? " & applicant_id_on_file_yn & vbCr
	If applicant_id_through_SOLQ <> "" AND applicant_id_through_SOLQ <> "?" Then email_body = email_body & "Can the Identity of the applicant be cleard through SOLQ/SMI? " & applicant_id_through_SOLQ & vbCr
	If postponed_verifs_yn <> "" AND postponed_verifs_yn <> "?" Then email_body = email_body & "Are there Postponed Verifications for this case? " & postponed_verifs_yn & vbCr
	If trim(list_postponed_verifs) <> "" Then email_body = email_body & "Postponed Verifications: " & list_postponed_verifs & vbCr
	If action_due_to_out_of_state_benefits <> "" Then
		email_body = email_body & "Other SNAP State: " & other_snap_state & vbCr
		email_body = email_body & "Reported End Date: " & other_state_reported_benefit_end_date & vbCr
		If other_state_benefits_openended = True Then email_body = email_body & "End date of SNAP in other state not determined." & vbCr
		email_body = email_body & "Has other State End Date been Confirmed/Verified: " & other_state_contact_yn & vbCr
		email_body = email_body & "Verified End Date: " & other_state_verified_benefit_end_date & vbCr
		email_body = email_body & "Action recommended by script based on information provided: " & action_due_to_out_of_state_benefits & vbCr
	End If
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then email_body = email_body & "It appears this case has postponed verifications from a previous EXP SNAP package that prevent approval of a new Expedited Package." & vbCr & vbCr

	email_body = email_body & "---" & vbCr
	If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

	email_body = "~~This email is generated from wihtin the 'Expedited Determination' Script.~~" & vbCr & vbCr & email_body
	call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", True)
	' call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", False)
	' create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
end function

function view_poli_temp(temp_one, temp_two, temp_three, temp_four)
	call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
	EMWriteScreen "TEMP", 5, 40     'Writes TEMP

	'Writes the panel_title selection
	Call write_value_and_transmit("TABLE", 21, 71)

	If temp_one <> "" Then temp_one = right("00" & temp_one, 2)
	If len(temp_two) = 1 Then temp_two = right("00" & temp_two, 2)
	If len(temp_three) = 1 Then temp_three = right("00" & temp_three, 2)
	If len(temp_four) = 1 Then temp_four = right("00" & temp_four, 2)

	total_code = "TE" & temp_one & "." & temp_two
	If temp_three <> "" Then total_code = total_code & "." & temp_three
	If temp_four <> "" Then total_code = total_code & "." & temp_four

	EMWriteScreen total_code, 3, 21
	transmit

	EMWriteScreen "X", 6, 4
	transmit
end function

show_pg_amounts = 1
show_pg_determination = 2
show_pg_review = 3

page_display = show_pg_amounts


Do
	Do
		err_msg = ""
		If page_display = show_pg_determination Then Call determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)

		BeginDialog Dialog1, 0, 0, 555, 385, "Full Expedited Determination"
		  ButtonGroup ButtonPressed
		  	If page_display = show_pg_amounts then
				Text 504, 12, 65, 10, "Amounts"

				GroupBox 5, 5, 390, 75, "Expedited Screening"
				If exp_screening_note_found = True Then
					Text 10, 20, 145, 10, "Information pulled from previous case note."
					Text 20, 35, 65, 10, "Income from CAF1: "
					Text 115, 35, 80, 10, caf_one_income
					Text 195, 35, 60, 10, "Assets from CAF1: "
					Text 270, 35, 75, 10, caf_one_assets
					Text 20, 50, 90, 10, "Rent/Mortgage from CAF1: "
					Text 115, 50, 65, 10, caf_one_rent
					Text 195, 50, 65, 10, "Utilities from CAF1: "
					Text 270, 50, 75, 10, caf_one_utilities
					Text 15, 65, 160, 10, xfs_screening
				End If
				If exp_screening_note_found = False Then
					Text 10, 20, 350, 10, "CASE:NOTE for Expedited Screening could not be found. No information to Display."
					Text 10, 30, 350, 10, "Review Application for screening answers"
				End If
				Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
				GroupBox 5, 105, 390, 125, "Information from SNAP/ELIG"
				Text 15, 125, 60, 10, "Gross Income:    $"
				EditBox 75, 120, 155, 15, determined_income
				Text 15, 145, 35, 10, "Assets:   $"
				EditBox 50, 140, 180, 15, determined_assets
				Text 15, 165, 70, 10, "Shelter Expense:    $"
				EditBox 85, 160, 145, 15, determined_shel
				Text 15, 185, 60, 10, "Utilities Expense:"
				Text 77, 185, 145, 15, "$  " & determined_utilities
				PushButton 255, 120, 120, 13, "Calculate Income", income_calc_btn
				PushButton 255, 140, 120, 13, "Calculate Assets", asset_calc_btn
				PushButton 255, 160, 120, 13, "Calculate Housing Cost", housing_calc_btn
			    PushButton 255, 180, 120, 13, "Calculate Utilities", utility_calc_btn
				If snap_elig_results_read = True Then Text 55, 200, 180, 10, "Autofilled information based on current STAT and ELIG panels"
				Text 15, 215, 250, 10, "Blank amounts will be defaulted to ZERO."
				' GroupBox 5, 220, 390, 100, "Supports"
				' Text 15, 235, 260, 10, "If you need support in handling for expedited, please access these resources:"
			    ' PushButton 25, 250, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
				' PushButton 25, 265, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
				' PushButton 25, 280, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
				' PushButton 25, 295, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			    ' PushButton 180, 250, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			    ' PushButton 180, 265, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			    ' PushButton 180, 280, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn
			End If
			If page_display = show_pg_determination then
				Text 495, 27, 65, 10, "Determination"

				GroupBox 5, 5, 470, 130, "Expedited Determination"
				Text 15, 20, 120, 10, "Determination Amounts Entered:"
				Text 140, 20, 85, 10, "Total App Month Income:"
				Text 230, 20, 35, 10, "$ " & determined_income
				Text 140, 30, 85, 10, "Total App Month Assets:"
				Text 230, 30, 35, 10, "$ " & determined_assets
				Text 140, 40, 85, 10, "Total App Month Housing:"
				Text 230, 40, 35, 10, "$ " & determined_shel
				Text 140, 50, 85, 10, "Total App Month Utility:"
				Text 230, 50, 35, 10, "$ " & determined_utilities
				Text 295, 20, 135, 10, "Combined Resources (Income + Assets):"
				Text 435, 20, 35, 10, "$ " & calculated_resources
				Text 330, 40, 100, 10, "Combined Housing Expense:"
				Text 435, 40, 35, 10, "$ " & calculated_expenses
				Text 295, 75, 125, 20, "Unit has less than $150 monthly Gross Income AND $100 or less in assets:"
				Text 430, 85, 35, 10, calculated_low_income_asset_test
				Text 295, 100, 125, 20, "Unit's combined resources are less than housing expense:"
				Text 430, 105, 35, 10, calculated_resources_less_than_expenses_test
				If is_elig_XFS = True Then Text 15, 75, 200, 10, "This case APPEARS EXPEDITED based on this above critera."
				If is_elig_XFS = False Then Text 15, 75, 250, 10, "This case does NOT appear to be expedited based on this above critera."
				Text 30, 95, 60, 10, "Date of Approval:"
				EditBox 90, 90, 60, 15, approval_date
				Text 155, 95, 75, 10, "(or planned approval)"
				Text 23, 110, 65, 10, "Date of Application:"
				Text 90, 110, 60, 10, date_of_application
				Text 30, 120, 60, 10, "Date of Interview:"
				Text 90, 120, 60, 10, interview_date

				GroupBox 5, 135, 470, 155, "Possible Approval Delays"
			    Text 95, 150, 205, 10, "Is there a document for proof of identity of the applicant on file?"
			    DropListBox 300, 145, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_on_file_yn
			    Text 95, 165, 200, 10, "Can the Identity of the applicant be cleard through SOLQ/SMI?"
			    DropListBox 300, 160, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_through_SOLQ
			    PushButton 350, 160, 120, 13, "HOT TOPIC - Using SOLQ for ID", ht_id_in_solq_btn
			    Text 10, 185, 85, 10, "Explain Approval Delays:"
			    EditBox 95, 180, 375, 15, delay_explanation
			    Text 175, 200, 80, 10, "Specifc case situations:"
			    PushButton 255, 200, 215, 15, "SNAP is Active in Another State in " & MAXIS_footer_month & "/" & MAXIS_footer_year, snap_active_in_another_state_btn
			    PushButton 255, 215, 215, 15, "Expedited Approved Previously with Postponed Verifications", case_previously_had_postponed_verifs_btn
			    PushButton 255, 230, 215, 15, "Household is Currently in a Facility", household_in_a_facility_btn
				Text 15, 255, 330, 10, "If it is already determined that SNAP should be denied, enter a denial date and explanation of denial."
				Text 355, 255, 65, 10, "SNAP Denial Date:"
				EditBox 420, 250, 50, 15, snap_denial_date
				Text 30, 275, 65, 10, "Denial Explanation:"
				EditBox 95, 270, 375, 15, snap_denial_explain
			End If
			If page_display = show_pg_review then
				Text 507, 42, 65, 10, "Review"

				GroupBox 5, 5, 470, 100, "Actions to Take"
				Text 20, 40, 45, 10, "Next Steps:"

				If IsDate(snap_denial_date) = True Then
					Text 15, 20, 280, 20, "DENIAL has been determined - Case does not meet 'All Other Eligibility Criteria' and Expedited Determination is not needed"

					Text 25, 55, 435, 10, "Update MAXIS STAT panels correctly to general results to Deny the Application"
					Text 25, 70, 435, 10, "Complete the DENIAL and enter a full, detailed CASE/NOTE of the Denial Action and Reasons."
					Text 25, 85, 435, 10, "Complete ALL PROCESSING before moving on to your next tast. Contact Knowledge Now if you are unsure of a Denial."
				ElseIf is_elig_XFS = True Then
			    	If IsDate(approval_date) = True Then
						Text 15, 20, 205, 10, "Case appears EXPEDITED and there are NO Delay reasons"

					    Text 25, 55, 435, 10, "Update MAXIS STAT panels to generate EXPEDITED SNAP Eligibility Results"
					    Text 25, 70, 435, 10, "Expedited Package includes " & expedited_package
					    Text 25, 85, 435, 10, "Approve SNAP Expedited package before moving on to the next task"
					Else
					End If
				ElseIf is_elig_XFS = False Then
				End If

				GroupBox 5, 105, 470, 100, "Postponed Verifications"
				If is_elig_XFS = True AND IsDate(snap_denial_date) = False Then
					Text 15, 125, 160, 10, "Are there Postponed Verifications for this case?"
					DropListBox 180, 120, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", postponed_verifs_yn
					Text 20, 145, 80, 10, "Postponed Verifications:"
					EditBox 105, 140, 360, 15, list_postponed_verifs
				End If
				Text 310, 15, 100, 10, "For help with the next steps:"
			    PushButton 310, 25, 155, 13, "Request Support from Knowledge Now", knowledge_now_support_btn
				If is_elig_XFS = True AND IsDate(snap_denial_date) = False Then
				    PushButton 320, 120, 145, 13, "TE 02.10.01 EXP w/ Pending Verifs", te_02_10_01_btn
				    Text 20, 165, 120, 10, "Can I postpone Verifications for ..."
				    Text 145, 165, 70, 10, "Immigration - YES."
				    Text 225, 165, 55, 10, "Sponsor - YES."
				    Text 300, 165, 125, 10, "anything OTHER than ID - YES. "
					Text 30, 180, 300, 10, "Appplicant's identity is the ONLY required verification to approve Expedited SNAP."
				    PushButton 320, 177, 145, 13, "CM 04.12 Verification Requirement for EXP", cm_04_12_btn
				End If
				If is_elig_XFS = False Then
					Text 15, 125, 450, 10, "We cannot postpone any verifications for a case that does not meet Expedited criteria."
				End If
				If IsDate(snap_denial_date) = True Then
					Text 15, 125, 450, 10, "Additional verifications are not needed if a Denial has already been determined."
				End If

			    GroupBox 5, 205, 470, 70, "EBT Information"
				If IsDate(snap_denial_date) = True Then
					Text 15, 220, 415, 10, "Advise resident to keep track of an EBT card they have received, even though the application is being denied."
					Text 20, 235, 415, 10, "If the case ever reapplies, or is determined eligible, the EBT card remains connected to the case and getting benefits will be easier."
				Else
					Text 15, 220, 335, 10, "Do not delay in approving SNAP benefits due to if the household does or does not have an EBT card."
				    Text 20, 235, 415, 10, "If there has never been a card issued for a case, approving the benefit with an REI will prevent a card from being sent via mail."
				    Text 20, 245, 305, 10, "If a case needs the first card mailed, do NOT REI benefits as they will not receive their card."
				End If
				Text 15, 260, 255, 10, "EBT Card issues can be complicated. Refer to the EBT Card Information here:"
			    PushButton 270, 257, 195, 13, "Temporary Program Changes - EBT Cards ", temp_prog_changes_ebt_card_btn

			End If
			GroupBox 5, 295, 470, 60, "If you need support in handling for expedited, please access these resources:"
			PushButton 15, 305, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
			PushButton 15, 320, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
			PushButton 15, 335, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			PushButton 165, 305, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
			PushButton 165, 320, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			PushButton 315, 305, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			PushButton 315, 320, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn

		    If page_display <> show_pg_amounts then PushButton 485, 10, 65, 13, "Amounts", amounts_btn
		    If page_display <> show_pg_determination then PushButton 485, 25, 65, 13, "Determination", determination_btn
		    If page_display <> show_pg_review then PushButton 485, 40, 65, 13, "Review", review_btn
		    If page_display <> show_pg_review then PushButton 445, 365, 50, 15, "Next", next_btn
			If page_display = show_pg_review then PushButton 445, 365, 50, 15, "Finish", finish_btn
		    CancelButton 500, 365, 50, 15
		    ' OkButton 500, 350, 50, 15
		EndDialog

		Dialog Dialog1
		cancel_confirmation

		If ButtonPressed = -1 Then
			If page_display <> show_pg_review then ButtonPressed = next_btn
			If page_display = show_pg_review then ButtonPressed = finish_btn
		End If

		If ButtonPressed = income_calc_btn Then Call app_month_income_detail(determined_income)
		If ButtonPressed = asset_calc_btn Then Call app_month_asset_detail(determined_assets, cash_amount_yn, bank_account_yn, ACCOUNTS_ARRAY)
		If ButtonPressed = housing_calc_btn Then Call app_month_housing_detail(determined_shel)
		If ButtonPressed = utility_calc_btn Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
		If ButtonPressed = snap_active_in_another_state_btn Then
			If IsDate(date_of_application) = False Then MsgBox "Attention:" & vbCr & vbCr & "The funcationality to determine actions if a household is reporting benefits in another state cannot be run if a valid application date has not been entered."
			If IsDate(date_of_application) = True Then Call snap_in_another_state_detail(date_of_application, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
		End If
		If ButtonPressed = case_previously_had_postponed_verifs_btn Then Call previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
		If ButtonPressed = household_in_a_facility_btn Then Call household_in_a_facility_detail

		If ButtonPressed = knowledge_now_support_btn Then Call send_support_email_to_KN
		If ButtonPressed = te_02_10_01_btn Then Call view_poli_temp("02", "10", "01", "")

		If page_display = show_pg_amounts Then

		End If
		If page_display = show_pg_determination Then
			delay_due_to_interview = False
			do_we_have_applicant_id = "UNKNOWN"
			If applicant_id_on_file_yn = "Yes" OR applicant_id_through_SOLQ = "Yes" Then do_we_have_applicant_id = True
			If applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "No" Then do_we_have_applicant_id = False

			' If IsDate(date_of_application) = False Then err_msg = err_msg & vbCr & "* The date of application needs to be entered as a valid date."
			' If IsDate(interview_date) = False Then err_msg = err_msg & vbCr & "* The interview date needs to be entered as a valid date. An Expedited Determination cannot be completed without the interview."
			If IsDate(snap_denial_date) = True Then
				If DateDiff("d", date, snap_denial_date) > 0 Then err_msg = err_msg & vbCr & "* Future Date denials or 'Possible' denials are not what the 'SNAP Denial Date' field is for." & vbCr &_
																						  "* Only indicate a denial if you already have enough information to determine that the SNAP application should be denied." & vbCr &_
																						  "* If this is the determination, review the date in the SNAP Denial Field as it appears to be a future date."
				snap_denial_explain = trim(snap_denial_explain)
				If len(snap_denial_explain) < 20 then err_msg = err_msg & vbCr & "* Since this SNAP case is to be denied, explain the reason for denial in detail."
			Else
				If is_elig_XFS = True Then
					If IsDate(approval_date) = True Then
						If DateDiff("d", date, approval_date) > 0 Then err_msg = err_msg & vbCr & "* Approvals should happen the same day an Expedited Determination is completed if the case is Expedited. Since the Income, Assets, and Expenses indicate this case is expedited AND we appear to be ready to approve, this should be completed today."
						' If DateDiff("d", interview_date, date) < 0 Then
					End If
					If applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Indicate if we have identity of the applicant on file or available through SOLQ"
					ElseIf applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Since there is no identity found in the file for the applicant, check SOLQ/SMI to verify identity."
					ElseIf applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "No" Then
						err_msg = err_msg & vbCr & "* Since the applicant's identity cannot be cleared through SOLQ/SMI, check the case file and person file for documents that can be used to verify identity. Remember that SNAP does NOT require a Photo ID or Official Government ID."
					End If

					'Defaulting Delay Explanation
					If IsDate(approval_date) = True AND IsDate(interview_date) = True AND IsDate(date_of_application) = True Then
						If DateDiff("d", date_of_application, approval_date) > 7 Then
							If DateDiff("d", interview_date, approval_date) = 0 Then delay_due_to_interview = True
						End If
					End If
					If delay_due_to_interview = True AND InStr(delay_explanation, "Approval of Expedited delayed until completion of Interview") = 0 Then
						delay_explanation = delay_explanation & "; Approval of Expedited delayed until completion of Interview."
					End If
					If delay_due_to_interview = False then
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview.", "")
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview", "")
					End If
					If do_we_have_applicant_id = False AND InStr(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant") = 0 Then
						delay_explanation = delay_explanation & "; Approval cannot be completed as we have NO Proof of Identity for the Applicant."
					End If
					If do_we_have_applicant_id <> False Then
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant.", "")
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant", "")
					End If


					delay_explanation = trim(delay_explanation)
					Do while Instr(delay_explanation, "; ;") <> 0
						delay_explanation = replace(delay_explanation, "; ;", "; ")
					Loop
					Do while Instr(delay_explanation, ";;") <> 0
						delay_explanation = replace(delay_explanation, ";;", "; ")
					Loop
					Do while Instr(delay_explanation, "  ") <> 0
						delay_explanation = replace(delay_explanation, "  ", " ")
					Loop
					delay_explanation = trim(delay_explanation)
					If left(delay_explanation, 1) = ";" Then delay_explanation = right(delay_explanation, len(delay_explanation) - 1)
					If right(delay_explanation, 1) = ";" Then delay_explanation = left(delay_explanation, len(delay_explanation) - 1)
					delay_explanation = trim(delay_explanation)

					expedited_approval_delayed = False
					If IsDate(approval_date) = False Then expedited_approval_delayed = True
					If IsDate(approval_date) = True  AND IsDate(date_of_application) = True Then
						If DateDiff("d", date_of_application, approval_date) > 7 Then expedited_approval_delayed = True
					End If
					If expedited_approval_delayed = True AND len(delay_explanation) < 20 Then err_msg = err_msg & vbCR & "* The approval of the Expedited SNAP is or has been delayed. Provide a detailed explaination of the reason for delay or complete the approval."

				End If
				If is_elig_XFS = False Then

				End If
			End If

		End If
		If page_display = show_pg_review Then
			If postponed_verifs_yn = "Yes" AND trim(list_postponed_verifs) = "" Then err_msg = err_msg * vbCr & "* Since you have Postponed Verifications indicated, list what they are for the NOTE."
		End If



		If ButtonPressed = next_btn AND err_msg = "" Then page_display = page_display + 1
		If ButtonPressed = amounts_btn Then page_display = show_pg_amounts
		If ButtonPressed = determination_btn AND err_msg = "" Then page_display = show_pg_determination
		If ButtonPressed = review_btn AND err_msg = "" AND page_display <> show_pg_amounts Then page_display = show_pg_review
		If ButtonPressed = review_btn AND err_msg = "" AND page_display = show_pg_amounts Then page_display = show_pg_determination

		If err_msg <> "" And ButtonPressed < 100 AND page_display <> show_pg_amounts Then MsgBox err_msg

		If ButtonPressed <> finish_btn Then err_msg = "LOOP"

		If ButtonPressed >= 1000 Then
			If ButtonPressed = hsr_manual_expedited_snap_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Expedited_SNAP.aspx"
			If ButtonPressed = hsr_snap_applications_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Applications.aspx"
			If ButtonPressed = ryb_exp_identity_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%201%20-%20Identity.mp4"
			If ButtonPressed = ryb_exp_timeliness_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%202%20-%20Timeliness.mp4"
			If ButtonPressed = sir_exp_flowchart_btn Then resource_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP%20Expedited%20Service%20Flowchart.pdf"
			If ButtonPressed = cm_04_04_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000404"
			If ButtonPressed = cm_04_06_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000406"
			If ButtonPressed = ht_id_in_solq_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/How-to-use-SMI-SOLQ-to-verify-ID-for-SNAP.aspx"
			If ButtonPressed = cm_04_12_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000412"
			If ButtonPressed = temp_prog_changes_ebt_card_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Temporary-Program-Changes--EBT-cards,-checks,-bus-cards.aspx"

			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & resource_URL
		End If



	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false


'
' ' 'Running the Dialog asking for all the detail and explanations
' ' DO
' ' 	Do
' ' 		Dialog Dialog1
' ' 		cancel_confirmation
' ' 		err_msg = ""
' ' 		IF is_elig_XFS = "FALSE" AND out_of_state_explanation = "" AND previous_xfs_explanation = "" AND other_explanation = "" AND abawd_explanation = "" THEN err_msg = err_msg & vbCr & "You have determined this case to NOT be Expedited but have provided no detail explanation" & vbCr & "Please complete at least one of the explanation boxes."
' ' 		IF id_check = checked AND other_explanation = "" THEN err_msg = err_msg & vbCr & "Please provided detail about no ID, remember that this is ONLY for the applicant and does NOT need to be a photo ID"
' ' 		IF err_msg <> "" Then MsgBox err_msg
' ' 	Loop until err_msg = ""
' ' 	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
' ' LOOP UNTIL are_we_passworded_out = false
'
' 'Formating the information from the edit boxes
' ' If determined_income = "" Then determined_income = 0
' ' If determined_assets = "" Then determined_assets = 0
' ' If determined_shel = "" Then determined_shel = 0
' ' If determined_utilities = "" Then determined_utilities = 0
' ' determined_resources = abs(determined_income) + (determined_assets * 1)
' ' determined_expenses = abs(determined_shel) + abs(determined_utilities)
' ' determined_assets = FormatCurrency(determined_assets)
' ' determined_expenses = FormatCurrency(determined_expenses)
' ' determined_income = FormatCurrency(determined_income)
' ' determined_resources = FormatCurrency(determined_resources)
' ' determined_shel = FormatCurrency(determined_shel)
' ' determined_utilities = FormatCurrency(determined_utilities)
'
' ' 'Converting String entries to Boolean
' ' IF is_elig_XFS = "TRUE" Then is_elig_XFS = TRUE
' ' IF is_elig_XFS = "FALSE" Then is_elig_XFS = FALSE
'
' '-------------------------------------------------------------------------------------------------DIALOG
' Dialog1 = "" 'Blanking out previous dialog detail
' BeginDialog Dialog1, 0, 0, 196, 120, "Expedited Timeliness"
'   EditBox 80, 5, 110, 15, date_of_application
'   EditBox 80, 25, 110, 15, interview_date
'   EditBox 80, 45, 110, 15, approval_date
'   EditBox 10, 80, 180, 15, delay_explanation
'   ButtonGroup ButtonPressed
'     OkButton 85, 100, 50, 15
'     CancelButton 140, 100, 50, 15
'   Text 10, 50, 60, 10, "Date of Approval"
'   Text 10, 10, 65, 10, "Date of Application"
'   Text 10, 65, 85, 10, "Explain any delays here"
'   Text 10, 30, 50, 10, "Interview Date"
' EndDialog
' 'Dialog about timeliness will run if case is determined to be expedited
' IF is_elig_XFS = TRUE Then
' 	Do
' 		Do
' 			Do
' 				Dialog Dialog1
' 				cancel_confirmation
' 				err_msg = ""
' 				IF date_of_application = "" OR IsDate(date_of_application) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid date of application."
' 				IF interview_date = "" OR IsDate(interview_date) = FALSE Then err_msg = err_msg & vbCr & "Pleaes enter a valid Interview Date."
' 				IF approval_date = "" OR IsDate(approval_date) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid Date of Approval."
' 				IF err_msg <> "" Then MsgBox err_msg
' 			Loop until err_msg = ""
' 			days_delayed = DateDiff ("d", date_of_application, approval_date)
' 			IF days_delayed > 7 AND delay_explanation = "" Then err_msg = err_msg & vbCr & "Your approval is more than 7 days from the date of application." & vbCr & "Please provide an explanation for the delay."
' 			If err_msg <> "" Then MsgBox err_msg
' 		Loop until err_msg = ""
' 		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
' 	Loop until are_we_passworded_out = false					'loops until user passwords back in
' End If



txt_file_name = "expedited_determination_detail_" & MAXIS_case_number & "_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
exp_info_file_path = t_drive &"\Eligibility Support\Assignments\Expedited Information"  & txt_file_name
MsgBox exp_info_file_path

With objFSO
	'Creating an object for the stream of text which we'll use frequently
	Dim objTextStream

	Set objTextStream = .OpenTextFile(exp_info_file_path, ForWriting, true)

	objTextStream.WriteLine ""

	objTextStream.WriteLine "CASE NUMBER ^*^*^" & MAXIS_case_number
	objTextStream.WriteLine "WORKER NAME ^*^*^" & worker_name
	objTextStream.WriteLine "CASE X NUMBER  ^*^*^" & case_pw
	objTextStream.WriteLine "DATE OF APPLICATION ^*^*^" & date_of_application
	objTextStream.WriteLine "DATE OF INTERVIEW ^*^*^" & interview_date
	objTextStream.WriteLine "EXPEDITED SCREENING STATUS ^*^*^" & xfs_screening
	objTextStream.WriteLine "EXPEDITED DETERMINATION STATUS ^*^*^" & is_elig_XFS
	objTextStream.WriteLine "DATE OF APPROVAL ^*^*^" & approval_date
	objTextStream.WriteLine "SNAP DENIAL DATE ^*^*^" & snap_denial_date
	objTextStream.WriteLine "SNAP DENIAL REASON ^*^*^" & snap_denial_explain
	objTextStream.WriteLine "ID ON FILE ^*^*^" & do_we_have_applicant_id
	objTextStream.WriteLine "END DATE OF SNAP IN ANOTHER STATE ^*^*^" & other_state_reported_benefit_end_date
	objTextStream.WriteLine "EXPEDITED APPROVE PREVIOUSLY POSTPONED ^*^*^" & case_has_previously_postponed_verifs_that_prevent_exp_snap				'(Boolean)
	objTextStream.WriteLine "EXPLAIN APPROVAL DELAYS  ^*^*^" & delay_explanation								'(all of them)
	objTextStream.WriteLine "POSTPONED VERIFICATIONS ^*^*^" & postponed_verifs_yn
	objTextStream.WriteLine "WHAT ARE THE POSTPONED VERIFICATIONS ^*^*^" & list_postponed_verifs
	objTextStream.WriteLine "DATE OF SCRIPT RUN ^*^*^" & date

	'Close the object so it can be opened again shortly
	objTextStream.Close

End With




MsgBox "STOP HERE"


'creating a custom header: this is read by BULK - EXP SNAP REVIEW script so don't mess this please :)
IF is_elig_XFS = true then
	case_note_header_text = "Expedited Determination: SNAP appears expedited"
ELSEIF is_elig_XFS = False then
	case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
END IF

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------
navigate_to_MAXIS_screen "CASE", "NOTE"
Call start_a_blank_case_note
Call write_variable_in_case_note (case_note_header_text)
IF exp_screening_note_found = TRUE Then
	Call write_bullet_and_variable_in_case_note ("Expedited Screening found", xfs_screening)
	Call write_variable_in_case_note ("*   Based on: Income: " & caf_one_income & ",   Assets: " & caf_one_assets & ",  Totaling: " & caf_one_resources)
	Call write_variable_in_case_note ("*             Shelter: " & caf_one_rent & ", Utilities: " & caf_one_utilities & ", Totaling: " & caf_one_expenses)
	Call write_variable_in_case_note ("---")
End If
If interview_date <> "" Then Call write_variable_in_case_note ("* Interview completed on: " & interview_date & " and full Expedited Determination Done")
IF is_elig_XFS = TRUE Then Call write_variable_in_case_note ("* Case is determined to meet criteria and Expedited SNAP can be approved.")
IF is_elig_XFS = FALSE Then Call write_variable_in_case_note ("* Expedited SNAP cannot be approved as case does not meet all criteria")
Call write_variable_in_case_note ("*   Based on: Income: " & determined_income & ",   Assets: " & determined_assets & ",   Totaling: " & determined_resources)
Call write_variable_in_case_note ("*             Shelter: " & determined_shel & ", Utilities: " & determined_utilities & ",  Totaling: " & determined_expenses)
IF id_check = checked Then Call write_variable_in_case_note ("* Applicant has not provided proof of ID.")
IF out_of_state_explanation <> "" Then
	Call write_variable_in_case_note ("* SNAP benefits have been received in another state")
	Call write_variable_in_case_note ("*    " & out_of_state_explanation)
End If
If previous_xfs_explanation <> "" Then
	Call write_variable_in_case_note ("* Expedited SNAP was the last approval and delayed verifs were not received")
	Call write_variable_in_case_note ("*    " & previous_xfs_explanation)
End If
Call write_bullet_and_variable_in_case_note("ABAWD info/explanation", abawd_explanation)
Call write_bullet_and_variable_in_case_note ("Other Notes", other_explanation)
Call write_variable_in_case_note ("---")
IF is_elig_XFS = TRUE Then
	Call write_bullet_and_variable_in_case_note ("Date of Application", date_of_application)
	Call write_bullet_and_variable_in_case_note ("Date of Interview", interview_date)
	Call write_bullet_and_variable_in_case_note ("Date of Approval", approval_date)
	Call write_bullet_and_variable_in_case_note ("Reason for Delay", delay_explanation)
	Call write_variable_in_case_note ("---")
End If
Call write_variable_in_case_note(worker_signature)

script_end_procedure ("Success! The script is complete. Case note has been entered detailing your Expedited Determination.")
