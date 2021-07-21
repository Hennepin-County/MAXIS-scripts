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
call MAXIS_case_number_finder(MAXIS_case_number)
If MAXIS_case_number <> "" Then
    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr

    Call navigate_to_MAXIS_screen("STAT", "PROG")

    EMReadScreen fs_appl_date, 8, 10, 33
    fs_appl_date = replace(fs_appl_date, " ", "/")

    If IsDate(fs_appl_date) = TRUE Then
        MAXIS_footer_month = DatePart("m", fs_appl_date)
        MAXIS_footer_month = right("0"&MAXIS_footer_month, 2)

        MAXIS_footer_year = right(DatePart("yyyy", fs_appl_date), 2)
    Else
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
    End If

End If

'dialog to gather the Case Number and such
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 291, 90, "SNAP EXP Determination - Case Information"
  EditBox 85, 5, 60, 15, MAXIS_case_number
  EditBox 230, 5, 25, 15, MAXIS_footer_month
  EditBox 260, 5, 25, 15, MAXIS_footer_year
  EditBox 85, 25, 200, 15, worker_signature
  DropListBox 165, 50, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", maxis_updated_yn
  ButtonGroup ButtonPressed
    OkButton 180, 70, 50, 15
    CancelButton 235, 70, 50, 15
  Text 30, 10, 50, 10, "Case Number:"
  Text 150, 10, 80, 10, "Application month/year:"
  Text 10, 30, 70, 10, "Sign your case note:"
  Text 10, 50, 145, 20, "Have you updated MAXIS STAT panels with all income, asset, and expense information?"
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
determined_income = elig_gross_income & ""
determined_assets = asset_amount & ""
determined_shel = elig_shel & ""
determined_utilities = elig_util & ""

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'THIS DIALOG IS DEFINED HERE BECAUSE OTHERWISE THE SCRIPT DOES NOT AUTOFILL THE TEXT FIELDS THAT ARE VAIRABLES
BeginDialog Dialog1, 0, 0, 401, 325, "Expedited Determination"
  EditBox 265, 50, 120, 15, determined_income
  EditBox 240, 70, 145, 15, determined_assets
  EditBox 275, 90, 110, 15, determined_shel
  EditBox 275, 110, 110, 15, determined_utilities
  DropListBox 300, 130, 85, 20, "TRUE"+chr(9)+"FALSE", is_elig_XFS
  CheckBox 200, 150, 150, 10, "Check here if APPLICANT has no form of ID", id_check
  EditBox 15, 180, 370, 15, out_of_state_explanation
  EditBox 15, 215, 370, 15, previous_xfs_explanation
  EditBox 15, 250, 370, 15, abawd_explanation
  EditBox 15, 285, 370, 15, other_explanation
  ButtonGroup ButtonPressed
    OkButton 290, 305, 50, 15
    CancelButton 345, 305, 50, 15
  Text 205, 55, 50, 10, "Gross Income:"
  Text 205, 75, 25, 10, "Assets:"
  Text 205, 95, 60, 10, "Shelter Expense:"
  Text 205, 115, 60, 10, "Utilities Expense:"
  Text 200, 135, 85, 10, "Client appears Expedited:"
  Text 10, 95, 160, 10, xfs_screening
  GroupBox 195, 5, 195, 140, "Information from SNAP/ELIG"
  GroupBox 5, 5, 180, 105, "Expedited Screening"
  Text 10, 20, 145, 10, "Information pulled from previous case note."
  Text 15, 50, 60, 10, "Assets from CAF1:"
  Text 15, 65, 90, 10, "Rent/Mortgage from CAF1:"
  Text 15, 80, 65, 10, "Utilities from CAF1:"
  Text 110, 35, 80, 10, caf_one_income
  Text 110, 50, 75, 10, caf_one_assets
  Text 110, 65, 65, 10, caf_one_rent
  Text 200, 35, 180, 10, "This information can be altered for the case note."
  Text 15, 35, 65, 10, "Income from CAF1: "
  Text 10, 115, 170, 35, "If the Expedited Determination for screening and elig do not match, detail the information that changed the determination from what is on CAF1 to the final determination."
  Text 10, 270, 95, 10, "Other detail needed to clarify"
  Text 110, 80, 75, 10, caf_one_utilities
  Text 10, 165, 255, 10, "If client received SNAP benefits out of state that impact eligibility, explain here"
  Text 200, 20, 180, 10, "Information based on current STAT and ELIG panels"
  Text 10, 200, 395, 10, "If the last issuance client received was Expedited and delayed verifications were not provided, explain in detail here:"
  Text 10, 235, 140, 10, "If client is an ABAWD, provide detail here:"
EndDialog
' BeginDialog Dialog1, 0, 0, 555, 385, "Expedited Determination"
'   GroupBox 5, 5, 390, 75, "Expedited Screening"
'   Text 10, 20, 145, 10, "Information pulled from previous case note."
'   Text 20, 35, 65, 10, "Income from CAF1: "
'   Text 115, 35, 80, 10, "caf_one_income"
'   Text 195, 35, 60, 10, "Assets from CAF1:"
'   Text 270, 35, 75, 10, "caf_one_assets"
'   Text 20, 50, 90, 10, "Rent/Mortgage from CAF1:"
'   Text 115, 50, 65, 10, "caf_one_rent"
'   Text 195, 50, 65, 10, "Utilities from CAF1:"
'   Text 270, 50, 75, 10, "caf_one_utilities"
'   Text 15, 65, 160, 10, "xfs_screening"
'   Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
'   GroupBox 5, 105, 390, 110, "Information from SNAP/ELIG"
'   Text 15, 125, 50, 10, "Gross Income:"
'   EditBox 75, 120, 155, 15, determined_income
'   ButtonGroup ButtonPressed
'     PushButton 255, 120, 120, 15, "Calculate Income", income_calc_btn
'   Text 15, 145, 25, 10, "Assets:"
'   EditBox 50, 140, 180, 15, determined_assets
'   ButtonGroup ButtonPressed
'     PushButton 255, 140, 120, 15, "Calculate Assets", asset_calc_btn
'   Text 15, 165, 60, 10, "Shelter Expense:"
'   EditBox 85, 160, 145, 15, determined_shel
'   ButtonGroup ButtonPressed
'     PushButton 255, 160, 120, 15, "Calculate Housing Cost", housing_calc_btn
'   Text 15, 185, 60, 10, "Utilities Expense:"
'   EditBox 85, 180, 145, 15, determined_utilities
'   ButtonGroup ButtonPressed
'     PushButton 255, 180, 120, 15, "Calculate Utilities", utility_calc_btn
'   Text 55, 200, 180, 10, "Autofilled information based on current STAT and ELIG panels"
'   GroupBox 5, 220, 390, 100, "Supports"
'   Text 15, 235, 260, 10, "If you need support in handling for expedited, please access these resources:"
'   ButtonGroup ButtonPressed
'     PushButton 25, 250, 150, 15, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
'     PushButton 180, 250, 150, 15, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
'     PushButton 25, 265, 150, 15, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
'     PushButton 180, 265, 150, 15, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
'     PushButton 25, 280, 150, 15, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
'     PushButton 180, 280, 150, 15, "CM 04.06 - 1st Mont Processing", cm_04_06_btn
'     PushButton 25, 295, 150, 15, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
'     PushButton 485, 10, 65, 15, "Amounts", amounts_btn
'     PushButton 485, 25, 65, 15, "Determination", determination_btn
'     PushButton 485, 40, 65, 15, "Review", review_btn
'     PushButton 445, 365, 50, 15, "Next", next_btn
'     CancelButton 500, 365, 50, 15
'     OkButton 500, 350, 50, 15
' EndDialog
'
'
' BeginDialog Dialog1, 0, 0, 555, 385, "Expedited Determination"
'   GroupBox 5, 5, 470, 130, "Expedited Determination"
'   Text 15, 20, 120, 10, "Determination Amounts Entered:"
'   Text 140, 20, 85, 10, "Total App Month Income:"
'   Text 230, 20, 35, 10, "$ XXXX"
'   Text 140, 30, 85, 10, "Total App Month Assets:"
'   Text 230, 30, 35, 10, "$ XXXX"
'   Text 140, 40, 85, 10, "Total App Month Housing:"
'   Text 230, 40, 35, 10, "$ XXXX"
'   Text 140, 50, 85, 10, "Total App Month Utility:"
'   Text 230, 50, 35, 10, "$ XXXX"
'   Text 295, 20, 135, 10, "Combined Resources (Income + Assets):"
'   Text 435, 20, 35, 10, "$ XXXX"
'   Text 330, 40, 100, 10, "Combined Housing Expense:"
'   Text 435, 40, 35, 10, "$ XXXX"
'   Text 185, 65, 250, 10, "Unit has less than $150 monthly Gross Income AND $100 or less in assets:"
'   Text 440, 65, 35, 10, "TRUE"
'   Text 235, 80, 195, 10, "Unit's combined resources are less than housing expense:"
'   Text 440, 80, 35, 10, "TRUE"
'   Text 15, 90, 315, 10, "This case APPEARS EXPEDITED based on this above critera."
'   Text 25, 110, 60, 10, "Date of Approval:"
'   EditBox 85, 105, 60, 15, Edit5
'   Text 150, 110, 75, 10, "(or planned approval)"
'   Text 330, 100, 65, 10, "Date of Application:"
'   EditBox 400, 95, 60, 15, Edit3
'   Text 335, 120, 60, 10, "Date of Interview:"
'   EditBox 400, 115, 60, 15, Edit4
'   GroupBox 5, 140, 470, 130, "Possible Approval Delays"
'   Text 15, 155, 330, 10, "If it is already determined that SNAP should be denied, enter a denial date and explanation of denial."
'   Text 20, 170, 65, 10, "SNAP Denial Date:"
'   EditBox 85, 165, 50, 15, Edit1
'   Text 145, 170, 45, 10, "Explanation:"
'   EditBox 190, 165, 280, 15, Edit2
'   Text 15, 190, 130, 10, "Is there an ID on file for the applicant?"
'   DropListBox 145, 185, 40, 45, "", List1
'   Text 195, 190, 200, 10, "Can the Identity of the applicant be cleard through SOLQ/SMI?"
'   DropListBox 400, 185, 70, 45, "", List2
'   ButtonGroup ButtonPressed
'     PushButton 300, 200, 170, 15, "HOT TOPIC - Using SOLQ/SMI for ID", ht_id_in_solq_btn
'   Text 15, 215, 120, 10, "Document specifc case situations:"
'   ButtonGroup ButtonPressed
'     PushButton 15, 225, 160, 15, "SNAP is Active in Another State in MM/YY", snap_active_in_another_state_btn
'     PushButton 180, 225, 210, 15, "Expedited Approved Previously with Postponed Verifications", Button18
'   Text 15, 250, 90, 10, "Explain Approval Delays:"
'   EditBox 105, 245, 365, 15, Edit6
'   GroupBox 5, 275, 470, 80, "Supports"
'   Text 15, 290, 260, 10, "If you need support in handling for expedited, please access these resources:"
'   ButtonGroup ButtonPressed
'     PushButton 20, 305, 150, 15, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
'     PushButton 20, 320, 150, 15, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
'     PushButton 20, 335, 150, 15, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
'     PushButton 170, 305, 150, 15, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
'     PushButton 170, 320, 150, 15, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
'     PushButton 315, 305, 150, 15, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
'     PushButton 315, 320, 150, 15, "CM 04.06 - 1st Mont Processing", cm_04_06_btn
'     PushButton 445, 365, 50, 15, "Next", next_btn
'     CancelButton 500, 365, 50, 15
'     OkButton 500, 350, 50, 15
'     PushButton 485, 25, 65, 15, "Determination", determination_btn
'     PushButton 485, 10, 65, 15, "Amounts", amounts_btn
'     PushButton 485, 40, 65, 15, "Review", review_btn
' EndDialog

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

hsr_manual_expedited_snap_btn 	= 1000
hsr_snap_applications_btn		= 1100
ryb_exp_identity_btn			= 1200
ryb_exp_timeliness_btn			= 1300
sir_exp_flowchart_btn			= 1400
cm_04_04_btn					= 1500
cm_04_06_btn					= 1600
ht_id_in_solq_btn				= 1700


function app_month_income_detail()
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
end function
function app_month_asset_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Determination of Assets in Month of Application"
	  ButtonGroup ButtonPressed
	  	Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

end function
function app_month_housing_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Determination of Housing Cost in Month of Application"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

end function
function app_month_utility_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Determination of Utilities in Month of Application"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

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
function snap_in_another_state_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Case Received SNAP in Another State"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

end function
function previous_postponed_verifs_detail()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Case Previously Received EXP SNAP with Postponed Verifications"
	  ButtonGroup ButtonPressed
	    Text 10, 5, 435, 10, "FUNCTIONALITY TO BE FILLED IN HERE"
	EndDialog

	dialog Dialog1

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
					Text 195, 35, 60, 10, "Assets from CAF1:"
					Text 270, 35, 75, 10, caf_one_assets
					Text 20, 50, 90, 10, "Rent/Mortgage from CAF1:"
					Text 115, 50, 65, 10, caf_one_rent
					Text 195, 50, 65, 10, "Utilities from CAF1:"
					Text 270, 50, 75, 10, caf_one_utilities
					Text 15, 65, 160, 10, xfs_screening
				End If
				If exp_screening_note_found = False Then
					Text 10, 20, 350, 10, "CASE:NOTE for Expedited Screening could not be found. No information to Display."
					Text 10, 30, 350, 10, "Review Application for screening answers"
				End If
				Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
				GroupBox 5, 105, 390, 110, "Information from SNAP/ELIG"
				Text 15, 125, 50, 10, "Gross Income:"
				EditBox 75, 120, 155, 15, determined_income
			    PushButton 255, 120, 120, 15, "Calculate Income", income_calc_btn
				Text 15, 145, 25, 10, "Assets:"
				EditBox 50, 140, 180, 15, determined_assets
			    PushButton 255, 140, 120, 15, "Calculate Assets", asset_calc_btn
				Text 15, 165, 60, 10, "Shelter Expense:"
				EditBox 85, 160, 145, 15, determined_shel
			    PushButton 255, 160, 120, 15, "Calculate Housing Cost", housing_calc_btn
				Text 15, 185, 60, 10, "Utilities Expense:"
				EditBox 85, 180, 145, 15, determined_utilities
			    PushButton 255, 180, 120, 15, "Calculate Utilities", utility_calc_btn
				If snap_elig_results_read = True Then Text 55, 200, 180, 10, "Autofilled information based on current STAT and ELIG panels"
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
				Text 185, 65, 250, 10, "Unit has less than $150 monthly Gross Income AND $100 or less in assets:"
				Text 440, 65, 35, 10, calculated_low_income_asset_test
				Text 235, 80, 195, 10, "Unit's combined resources are less than housing expense:"
				Text 440, 80, 35, 10, calculated_resources_less_than_expenses_test
				If is_elig_XFS = True Then Text 15, 90, 315, 10, "This case APPEARS EXPEDITED based on this above critera."
				If is_elig_XFS = False Then Text 15, 90, 315, 10, "This case does NOT appear to be expedited based on this above critera."
				Text 25, 110, 60, 10, "Date of Approval:"
				EditBox 85, 105, 60, 15, approval_date
				Text 150, 110, 75, 10, "(or planned approval)"
				Text 330, 100, 65, 10, "Date of Application:"
				EditBox 400, 95, 60, 15, date_of_application
				Text 335, 120, 60, 10, "Date of Interview:"
				EditBox 400, 115, 60, 15, interview_date
				GroupBox 5, 140, 470, 130, "Possible Approval Delays"
				Text 15, 155, 330, 10, "If it is already determined that SNAP should be denied, enter a denial date and explanation of denial."
				Text 20, 170, 65, 10, "SNAP Denial Date:"
				EditBox 85, 165, 50, 15, snap_denial_date
				Text 145, 170, 45, 10, "Explanation:"
				EditBox 190, 165, 280, 15, snap_denial_explain
				Text 15, 190, 130, 10, "Is there an ID on file for the applicant?"
				DropListBox 145, 185, 40, 45, "", applicant_id_on_file_yn
				Text 195, 190, 200, 10, "Can the Identity of the applicant be cleard through SOLQ/SMI?"
				DropListBox 400, 185, 70, 45, "", applicant_id_through_SOLQ
				PushButton 300, 200, 170, 15, "HOT TOPIC - Using SOLQ/SMI for ID", ht_id_in_solq_btn
				Text 15, 225, 80, 10, "Specifc case situations:"
			    PushButton 100, 220, 160, 15, "SNAP is Active in Another State in MM/YY", snap_active_in_another_state_btn
			    PushButton 260, 220, 210, 15, "Expedited Approved Previously with Postponed Verifications", case_previously_had_postponed_verifs_btn
				Text 15, 250, 90, 10, "Explain Approval Delays:"
				EditBox 105, 245, 365, 15, delay_explanation
			End If
			If page_display = show_pg_review then
				Text 507, 42, 65, 10, "Review"
			End If
			GroupBox 5, 275, 470, 80, "Supports"
			Text 15, 290, 260, 10, "If you need support in handling for expedited, please access these resources:"
			PushButton 20, 305, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
			PushButton 20, 320, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
			PushButton 20, 335, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			PushButton 170, 305, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
			PushButton 170, 320, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
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

		If ButtonPressed = next_btn AND err_msg = "" Then page_display = page_display + 1
		If ButtonPressed = amounts_btn AND err_msg = "" Then page_display = show_pg_amounts
		If ButtonPressed = determination_btn AND err_msg = "" Then page_display = show_pg_determination
		If ButtonPressed = review_btn AND err_msg = "" Then page_display = show_pg_review

		If ButtonPressed <> finish_btn Then err_msg = "LOOP"

		If ButtonPressed >= 100 Then
			If ButtonPressed = income_calc_btn Then Call app_month_income_detail
			If ButtonPressed = asset_calc_btn Then Call app_month_asset_detail
			If ButtonPressed = housing_calc_btn Then Call app_month_housing_detail
			If ButtonPressed = utility_calc_btn Then Call app_month_utility_detail
			If ButtonPressed = snap_active_in_another_state_btn Then Call snap_in_another_state_detail
			If ButtonPressed = case_previously_had_postponed_verifs_btn Then Call previous_postponed_verifs_detail

			If ButtonPressed >= 1000 Then
				If ButtonPressed = hsr_manual_expedited_snap_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Expedited_SNAP.aspx"
				If ButtonPressed = hsr_snap_applications_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Applications.aspx"
				If ButtonPressed = ryb_exp_identity_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%201%20-%20Identity.mp4"
				If ButtonPressed = ryb_exp_timeliness_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%202%20-%20Timeliness.mp4"
				If ButtonPressed = sir_exp_flowchart_btn Then resource_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP%20Expedited%20Service%20Flowchart.pdf"
				If ButtonPressed = cm_04_04_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000404"
				If ButtonPressed = cm_04_06_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000406"
				If ButtonPressed = ht_id_in_solq_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/How-to-use-SMI-SOLQ-to-verify-ID-for-SNAP.aspx"

				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & resource_URL
			End If
		End If

		If err_msg <> "" AND err_msg <> "LOOP" Then MsgBox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
MsgBox "STOP HERE"

' 'Running the Dialog asking for all the detail and explanations
' DO
' 	Do
' 		Dialog Dialog1
' 		cancel_confirmation
' 		err_msg = ""
' 		IF is_elig_XFS = "FALSE" AND out_of_state_explanation = "" AND previous_xfs_explanation = "" AND other_explanation = "" AND abawd_explanation = "" THEN err_msg = err_msg & vbCr & "You have determined this case to NOT be Expedited but have provided no detail explanation" & vbCr & "Please complete at least one of the explanation boxes."
' 		IF id_check = checked AND other_explanation = "" THEN err_msg = err_msg & vbCr & "Please provided detail about no ID, remember that this is ONLY for the applicant and does NOT need to be a photo ID"
' 		IF err_msg <> "" Then MsgBox err_msg
' 	Loop until err_msg = ""
' 	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
' LOOP UNTIL are_we_passworded_out = false

'Formating the information from the edit boxes
If determined_income = "" Then determined_income = 0
If determined_assets = "" Then determined_assets = 0
If determined_shel = "" Then determined_shel = 0
If determined_utilities = "" Then determined_utilities = 0
determined_resources = abs(determined_income) + (determined_assets * 1)
determined_expenses = abs(determined_shel) + abs(determined_utilities)
determined_assets = FormatCurrency(determined_assets)
determined_expenses = FormatCurrency(determined_expenses)
determined_income = FormatCurrency(determined_income)
determined_resources = FormatCurrency(determined_resources)
determined_shel = FormatCurrency(determined_shel)
determined_utilities = FormatCurrency(determined_utilities)

' 'Converting String entries to Boolean
' IF is_elig_XFS = "TRUE" Then is_elig_XFS = TRUE
' IF is_elig_XFS = "FALSE" Then is_elig_XFS = FALSE

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 196, 120, "Expedited Timeliness"
  EditBox 80, 5, 110, 15, date_of_application
  EditBox 80, 25, 110, 15, interview_date
  EditBox 80, 45, 110, 15, approval_date
  EditBox 10, 80, 180, 15, delay_explanation
  ButtonGroup ButtonPressed
    OkButton 85, 100, 50, 15
    CancelButton 140, 100, 50, 15
  Text 10, 50, 60, 10, "Date of Approval"
  Text 10, 10, 65, 10, "Date of Application"
  Text 10, 65, 85, 10, "Explain any delays here"
  Text 10, 30, 50, 10, "Interview Date"
EndDialog
'Dialog about timeliness will run if case is determined to be expedited
IF is_elig_XFS = TRUE Then
	Do
		Do
			Do
				Dialog Dialog1
				cancel_confirmation
				err_msg = ""
				IF date_of_application = "" OR IsDate(date_of_application) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid date of application."
				IF interview_date = "" OR IsDate(interview_date) = FALSE Then err_msg = err_msg & vbCr & "Pleaes enter a valid Interview Date."
				IF approval_date = "" OR IsDate(approval_date) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid Date of Approval."
				IF err_msg <> "" Then MsgBox err_msg
			Loop until err_msg = ""
			days_delayed = DateDiff ("d", date_of_application, approval_date)
			IF days_delayed > 7 AND delay_explanation = "" Then err_msg = err_msg & vbCr & "Your approval is more than 7 days from the date of application." & vbCr & "Please provide an explanation for the delay."
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
End If

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
