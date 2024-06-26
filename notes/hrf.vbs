'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HRF.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
Call changelog_update("03/27/2024", "Added a checkbox option to indicate that a future month HRF has not been received when processing a HRF for the current month. This adds a line to the CASE/NOTE indicating this future HRF is not received.", "Casey Love, Hennepin County")
Call changelog_update("02/27/2024", "Removed eligibility details from case note. Please use NOTES-Eligibility Summary to document this information.", "Megan Geissler, Hennepin County")
Call changelog_update("06/26/2023", "Added handling to support selection of specific programs for HRF processing.", "Ilse Ferris, Hennepin County")
Call changelog_update("07/10/2019", "Fixed a bug that prevented the script from reading the grant amount if Significant Change was applied on MFIP. Additionally added functionality to copy significant change information into the casenote if ELIG/MF is read.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
call changelog_update("04/23/2018", "Added NOTES on INCOME field and some preselected options to input on NOTES on INCOME field for more detailed case notes.", "Casey Love, Hennepin County")
call changelog_update("02/23/2018", "Added closing message to reminder to workers to accept all work items upon processing HRF's.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2016", "Added seperate functionality for LTC HRF cases.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect "" 'Connecting to BlueZone
Call MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing case number & footer month/year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Do
    Do
        '-------------------------------------------------------------------------------------------------DIALOG
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 181, 110, "HRF Case Number"
          EditBox 80, 5, 70, 15, MAXIS_case_number
          EditBox 80, 25, 18, 15, MAXIS_footer_month
          EditBox 106, 25, 18, 15, MAXIS_footer_year
          CheckBox 10, 70, 30, 10, "MFIP", MFIP_check
          CheckBox 45, 70, 30, 10, "SNAP", SNAP_check
          CheckBox 85, 70, 20, 10, "HC", HC_check
          CheckBox 115, 70, 25, 10, "GA", GA_check
          CheckBox 145, 70, 50, 10, "MSA", MSA_check
          ButtonGroup ButtonPressed
            OkButton 35, 90, 50, 15
            CancelButton 95, 90, 50, 15
          Text 30, 10, 50, 10, "Case number:"
          Text 5, 30, 75, 10, "Footer month (MM/YY):"
		  Text 101, 30, 4, 10, "/"
          Text 80, 40, 75, 10, "(benefit month)"
          GroupBox 5, 55, 170, 30, "Programs Recertifying"
        EndDialog

        err_msg = ""
      	Dialog Dialog1
      	cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        If (MFIP_check = 0 and SNAP_check = 0 and HC_check = 0 and GA_check = 0 and MSA_check = 0) then err_msg = err_msg & "* Select all applicable programs at monthly report."

		'Checking for PRIV cases.
		EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
		IF priv_check = "PRIVIL" THEN script_end_procedure("This case is a privliged case. You do not have access to this case.")

        'Checking to ensure the case is actually at a HRF
        Call check_for_MAXIS(False)
        Call navigate_to_MAXIS_screen("STAT", "MONT")
        EmReadscreen HRF_panels, 1, 2, 73
        If HRF_panels = 0 then
            script_end_procedure_with_error_report("This case is not subject to monthly reporting. The script will now end.")
        Else
            cash_hrf = False    'defaulting programs to false for determination
            hc_hrf = False
            snap_hrf = False

            'setting boolean here since there are more than one cash programs
            If MFIP_check = 1 or GA_check = 1 or MSA_check = 1 then
                cash_progs = True
            else
                cash_progs = False
            End if

            EmReadscreen cash_code, 1, 11, 43
            EmReadscreen snap_code, 1, 11, 53
            EmReadscreen HC_code, 1, 11, 63

            If cash_code <> "_" then cash_HRF = True
            If snap_code <> "_" then snap_HRF = True
            If HC_code <> "_" then hc_HRF = True

            'program selected in dialog, not open available as HRF process
            If HC_check = 1 and hc_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process health care, but you cannot use a HRF for health care programs on this case. Update your program selections." & vbcr
            If SNAP_check = 1 and snap_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process SNAP, but you cannot use a HRF for SNAP on this case. Update your program selections." & vbcr
            If cash_progs = True and cash_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process cash programs, but you cannot use a HRF for cash programs on this case. Update your program selection." & vbcr

            'program listed on MONT page, but NOT in program selection in dialog
            If HC_check = 0 and hc_hrf = True then err_msg = err_msg & vbcr & "* Health Care is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
            If SNAP_check = 0 and snap_hrf = True then err_msg = err_msg & vbcr & "* SNAP is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
            If cash_progs = false and cash_hrf = True then err_msg = err_msg & vbcr & "* Cash is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
        End if
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'NAV to STAT
call navigate_to_MAXIS_screen("STAT", "MEMB")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling info for case note
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up info for case note
HRF_computer_friendly_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
retro_month_name = monthname(datepart("m", (dateadd("m", -2, HRF_computer_friendly_month))))
pro_month_name = monthname(datepart("m", (HRF_computer_friendly_month)))
HRF_month = retro_month_name & "/" & pro_month_name
next_month_hrf_not_received_checkbox = unchecked
next_retro_month_name = monthname(datepart("m", (dateadd("m", -1, HRF_computer_friendly_month))))
next_month_name = monthname(datepart("m", DateAdd("m", 1, HRF_computer_friendly_month)))
next_HRF_month = next_retro_month_name & "/" & next_month_name

'If a HRF is being run for a HC case, script will ask if this is a LTC case
If HC_check = checked Then
	'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
	LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have different fields in their dialog.", vbYesNoCancel)
	If LTC_case = vbCancel then stopscript
Else
	LTC_case = vbNo
End If

'If workers answers yes to this is a LTC case - script runs this specific functionality
If LTC_case = vbYes then
	'LTC cases should not have these programs active
	If MFIP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* MFIP will be removed."
	If SNAP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* SNAP will be removed."
	If GA_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* GA will be removed."

	'Alerting worker that these programs will be unchecked.
	If uncheck_msg <> "" Then MsgBox "You have checked programs that should not be active with LTC. These programs will not be added to the note." & vbNewLine & uncheck_msg

	MFIP_check = unchecked
	SNAP_check = unchecked
	GA_check = unchecked

	'Getting some additional information for the dialog to be autofilled
	call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
	call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
	call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)

	'Going to find the current facility to autofil the dialog
	Call navigate_to_MAXIS_screen ("STAT", "FACI")

	'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
	Do
		EMReadScreen FACI_current_panel, 1, 2, 73
		EMReadScreen FACI_total_check, 1, 2, 78
		EMReadScreen in_year_check_01, 4, 14, 53
		EMReadScreen in_year_check_02, 4, 15, 53
		EMReadScreen in_year_check_03, 4, 16, 53
		EMReadScreen in_year_check_04, 4, 17, 53
		EMReadScreen in_year_check_05, 4, 18, 53
		EMReadScreen out_year_check_01, 4, 14, 77
		EMReadScreen out_year_check_02, 4, 15, 77
		EMReadScreen out_year_check_03, 4, 16, 77
		EMReadScreen out_year_check_04, 4, 17, 77
		EMReadScreen out_year_check_05, 4, 18, 77
		If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or (in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
			currently_in_FACI = True
			exit do
		Elseif FACI_current_panel = FACI_total_check then
			currently_in_FACI = False
			exit do
		Else
			transmit
		End if
	Loop until FACI_current_panel = FACI_total_check

	If in_year_check_01 <> "____" and out_year_check_01 = "____" Then EMReadScreen date_in, 10, 14, 47
	If in_year_check_02 <> "____" and out_year_check_02 = "____" Then EMReadScreen date_in, 10, 15, 47
	If in_year_check_03 <> "____" and out_year_check_03 = "____" Then EMReadScreen date_in, 10, 16, 47
	If in_year_check_04 <> "____" and out_year_check_04 = "____" Then EMReadScreen date_in, 10, 17, 47
	If in_year_check_05 <> "____" and out_year_check_05 = "____" Then EMReadScreen date_in, 10, 18, 47

	admit_date = replace(date_in, " ", "/")

	'Gets Facility name and admit in date and enters it into the dialog
	If currently_in_FACI = True then
		EMReadScreen FACI_name, 30, 6, 43
		facility_info = trim(replace(FACI_name, "_", ""))
	End if

	'confirms that case is in the footer month/year selected by the user
	Call MAXIS_footer_month_confirmation
	Call MAXIS_background_check

	'Goes to STAT WKEX to get deductions and possible FIAT reasons to autofil the dialog
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

	'Gives unverified expenses and blank expenses the value of $0 and adds non-zero amounts to the dialog for autofil
	If federal_tax = "" OR federal_tax_verif_code = "N" then
		federal_tax = "0"
	Else
		hc_deductions = hc_deductions & "; Federal Tax - $" & federal_tax
	End if
	If state_tax = "" OR state_tax_verif_code = "N" then
		state_tax = "0"
	Else
		hc_deductions = hc_deductions & "; State Tax - $" & state_tax
	End if
	If FICA_witheld = "" OR FICA_witheld_verif_code = "N" then
		FICA_witheld = "0"
	Else
		hc_deductions = hc_deductions & "; FICA - $" & FICA_witheld
	End if
	If transportation_expense = "" OR transportation_expense_verif_code = "N" OR transportation_impair =  "_" OR transportation_impair = "N" then
		transportation_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Transportation Expense - $" & transportation_expense
	End if
	If meals_expense = "" OR meals_expense_verif_code = "N" OR meals_impair = "_" OR meals_impair = "N" then
		meals_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Meals Expense - $" & meals_expense
	End if
	If uniform_expense = "" OR uniform_expense_verif_code = "N" OR uniform_impair = "_" OR uniform_impair = "N" then
		uniform_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Uniform Expense - $" & uniform_expense
	End if
	If tools_expense = "" OR tools_expense_verif_code = "N" OR tools_impair = "_" OR tools_impair = "N" then
		tools_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Tools Expense - $" & tools_expense
	End if
	If dues_expense = "" OR dues_expense_verif_code = "N" OR dues_impair = "_" OR dues_impair = "N" then
		dues_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Dues Expense - $" & dues_expense
	End if
	If other_expense = "" OR other_expense_verif_code = "N" OR other_impair = "_" OR other_impair = "N" then
		other_expense = "0"
	Else
		hc_deductions = hc_deductions & "; Other Expense - $" & other_expense
	End if

	'Checks PDED other expenses, will need to add PDED and WKEX other expenses together
	Call navigate_to_MAXIS_screen("STAT", "PDED")
	EMReadScreen other_earned_income_PDED, 8, 11, 62

	'cleaning up PDED variables
	other_earned_income_PDED = replace(other_earned_income_PDED, "_", "")
	other_earned_income_PDED = trim(other_earned_income_PDED)

	'Gives blank expenses the value of $0
	If other_earned_income_PDED = "" then
		other_earned_income_PDED = "0"
	Else
		hc_deductions = hc_deductions & "; Other Earned Income Deductions - $" & other_earned_income_PDED
	End If

	'Determining if earned income is less than $80
	Call navigate_to_MAXIS_screen ("STAT", "JOBS")
	EMReadScreen JOBS_panel_income, 7, 17, 68
	JOBS_panel_income = trim(JOBS_panel_income)
	If IsNumeric(JOBS_panel_income) = TRUE Then
		If abs(JOBS_panel_income) < 80 then
			special_pers_allow = JOBS_panel_income	'if less then $80 deduction is earned income amount
		ELSE
			special_pers_allow = "80.00"		'otherwise deduction is $80
		END IF
	Else
		JOBS_panel_income = ""
	End If

	If JOBS_panel_income <> "" Then hc_deductions = hc_deductions & "; Special Allowance - $" & special_pers_allow

	'All of the deductions found to this point need to be FIATed. Added these to the FIAT varirable.
	FIAT_reasons = hc_deductions

	'Going to see if there is a deduction on MEDI. (This does not have to be FIATED)
	Call navigate_to_MAXIS_screen ("STAT", "MEDI")
	EMReadScreen medi_panel_exists, 1, 2, 78
	If medi_panel_exists = "1" Then
		EMReadScreen part_b_premium, 9, 7, 72
		part_b_premium = trim(part_b_premium)
		If part_b_premium <> "________" Then hc_deductions = hc_deductions & "; Medicare Premium - $" & part_b_premium
	End If

	'Formatting the variables for the dialog
	hc_deductions = right(hc_deductions, len(hc_deductions) - 2)
	FIAT_reasons = right(FIAT_reasons, len(FIAT_reasons) - 2)

	'The case note dialog, complete with panel navigation, reading the ELIG/MSA or ELIG/HC screen, and navigation to case note, as well as logic for certain sections to be required.
	DO
		DO
			Do
			    '-------------------------------------------------------------------------------------------------DIALOG
			    Dialog1 = "" 'Blanking out previous dialog detail
			    BeginDialog Dialog1, 0, 0, 451, 295, "HRF for LTC Cases"
			      EditBox 65, 10, 85, 15, HRF_datestamp
			      DropListBox 240, 10, 80, 15, "complete"+chr(9)+"incomplete", HRF_status
			      EditBox 50, 30, 165, 15, facility_info
			      EditBox 280, 30, 55, 15, admit_date
			      CheckBox 350, 5, 80, 10, "Sent 3050 to Facility", sent_3050_checkbox
			      CheckBox 350, 20, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
			      CheckBox 350, 35, 80, 10, "Next HRF Released", HRF_release_checkbox
			      EditBox 65, 50, 380, 15, earned_income
			      EditBox 70, 70, 375, 15, unearned_income
			      EditBox 110, 90, 335, 15, notes_on_income
			      EditBox 40, 110, 405, 15, assets
			      EditBox 50, 130, 395, 15, hc_deductions
			      EditBox 100, 150, 345, 15, FIAT_reasons
			      EditBox 50, 170, 395, 15, other_notes
			      EditBox 235, 190, 210, 15, verifs_needed
			      EditBox 235, 210, 210, 15, actions_taken
			      EditBox 165, 275, 105, 15, worker_signature
			      ButtonGroup ButtonPressed
			    	OkButton 340, 275, 50, 15
			    	CancelButton 390, 275, 50, 15
			    	PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
			    	PushButton 10, 205, 25, 10, "BUSI", BUSI_button
			    	PushButton 35, 205, 25, 10, "JOBS", JOBS_button
			    	PushButton 10, 215, 25, 10, "RBIC", RBIC_button
			    	PushButton 35, 215, 25, 10, "UNEA", UNEA_button
			    	PushButton 75, 205, 25, 10, "ACCT", ACCT_button
			    	PushButton 100, 205, 25, 10, "CARS", CARS_button
			    	PushButton 125, 205, 25, 10, "CASH", CASH_button
			    	PushButton 150, 205, 25, 10, "OTHR", OTHR_button
			    	PushButton 75, 215, 25, 10, "REST", REST_button
			    	PushButton 100, 215, 25, 10, "SECU", SECU_button
			    	PushButton 125, 215, 25, 10, "TRAN", TRAN_button
			    	PushButton 10, 250, 25, 10, "MEMB", MEMB_button
			    	PushButton 35, 250, 25, 10, "MEMI", MEMI_button
			    	PushButton 60, 250, 25, 10, "MONT", MONT_button
			    	PushButton 10, 260, 25, 10, "PARE", PARE_button
			    	PushButton 35, 260, 25, 10, "SANC", SANC_button
			    	PushButton 60, 260, 25, 10, "TIME", TIME_button
			    	PushButton 295, 245, 20, 10, "HC", ELIG_HC_button
			    	PushButton 295, 255, 20, 10, "MSA", ELIG_MSA_button
			    	PushButton 345, 245, 45, 10, "prev. panel", prev_panel_button
			    	PushButton 390, 245, 45, 10, "prev. memb", prev_memb_button
			    	PushButton 345, 255, 45, 10, "next panel", next_panel_button
			    	PushButton 390, 255, 45, 10, "next memb", next_memb_button
			      Text 5, 15, 55, 10, "HRF datestamp:"
			      Text 195, 15, 40, 10, "HRF status:"
			      Text 5, 35, 45, 10, "Facility Info:"
			      Text 230, 35, 50, 10, "Admit In Date:"
			      Text 5, 55, 55, 10, "Earned income:"
			      Text 5, 75, 60, 10, "Unearned income:"
			      Text 5, 115, 30, 10, "Assets:"
			      Text 5, 135, 40, 10, "Deductions:"
			      Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
			      Text 5, 175, 45, 10, "Other notes:"
			      GroupBox 5, 190, 60, 40, "Income panels"
			      GroupBox 70, 190, 110, 40, "Asset panels"
			      GroupBox 5, 235, 85, 40, "other STAT panels:"
			      Text 185, 195, 50, 10, "Verifs needed:"
			      Text 185, 215, 50, 10, "Actions taken:"
			      GroupBox 280, 230, 50, 40, "ELIG panels:"
			      GroupBox 340, 230, 100, 40, "STAT-based navigation"
			      Text 100, 280, 60, 10, "Worker signature:"
				  If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
					GroupBox 90, 235, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
					Text 95, 245, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
					CheckBox 95, 255, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
				  End If
			    EndDialog
				err_msg = ""
				Dialog Dialog1
				cancel_confirmation
				Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
				MAXIS_dialog_navigation
                If ButtonPressed = income_notes_button Then
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
				    BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
				      CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
				      CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
				      CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
				      CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
				      CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
				      CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
				      CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
				      CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
				      CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
				      CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
				      CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
				      ButtonGroup ButtonPressed
				        PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
				        CancelButton 295, 205, 50, 15
				      Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
				      GroupBox 5, 60, 340, 45, "JOBS Income"
				      GroupBox 5, 110, 340, 45, "BUSI Income"
				      GroupBox 5, 160, 340, 40, "UNEA Income"
				    EndDialog
                    Dialog Dialog1
                    If ButtonPressed = add_to_notes_button Then
                        If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
                        If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
                        If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
                        If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
                        If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
                        If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
                        If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
                        If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
                        If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
                        If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
                        If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
                        If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
                    End If
                End If
				IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
				IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
				IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
				IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
				IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
				IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
			case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
			IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
		LOOP UNTIL case_note_confirmation = vbYes
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false

	'Setting up some variables for the case note
	programs_list = "HC"
	If MSA_check = checked Then programs_list = programs_list & " & MSA"
	If admit_date <> "" then facility_info = facility_info & ". Admit Date: " & admit_date


	'Enters the case note-----------------------------------------------------------------------------------------------
	start_a_blank_CASE_NOTE
	Call write_variable_in_case_note("***" & MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
	call write_bullet_and_variable_in_case_note("Programs", programs_list)
	call write_bullet_and_variable_in_case_note("Facility", facility_info)
	call write_bullet_and_variable_in_case_note("Earned income", earned_income)
	call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
    call write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
	call write_bullet_and_variable_in_case_note("Assets", assets)
	call write_bullet_and_variable_in_case_note("Deductions", hc_deductions)
	call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
	call write_bullet_and_variable_in_case_note("Other notes", other_notes)
	If sent_3050_checkbox = 1 then call write_variable_in_CASE_NOTE("* Sent 3050 to Facility")
	If HRF_release_checkbox = 1 then call write_variable_in_CASE_NOTE("* Released HRF in MAXIS for next month.")
	IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
	call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
	call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
	If next_month_hrf_not_received_checkbox = checked Then
		call write_variable_in_CASE_NOTE("* HRF for next month (" & CM_plus_1_mo & "/" & CM_plus_1_yr & ") has not been received.")
		call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
	End If

	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = "Success! Your HRF for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " on a LTC case has been case noted."
ElseIf LTC_case = vbNo then							'Shows dialog if not LTC
	'The case note dialog, complete with panel navigation, reading the ELIG/MFIP screen, and navigation to case note, as well as logic for certain sections to be required.
	DO
		DO
			Do
				err_msg = ""
				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
			    BeginDialog Dialog1, 0, 0, 451, 285, MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF dialog"
			      EditBox 65, 30, 50, 15, HRF_datestamp
			      DropListBox 170, 30, 75, 15, "complete"+chr(9)+"incomplete", HRF_status
			      EditBox 65, 50, 380, 15, earned_income
			      EditBox 70, 70, 375, 15, unearned_income
			      EditBox 110, 90, 335, 15, notes_on_income
			      EditBox 30, 110, 90, 15, YTD
			      EditBox 170, 110, 275, 15, changes
			      EditBox 30, 130, 415, 15, EMPS
			      EditBox 100, 150, 345, 15, FIAT_reasons
			      EditBox 50, 170, 395, 15, other_notes
			      CheckBox 190, 190, 60, 10, "10% sanction?", ten_percent_sanction_check
			      CheckBox 265, 190, 60, 10, "30% sanction?", thirty_percent_sanction_check
			      CheckBox 330, 190, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
			      EditBox 235, 205, 210, 15, verifs_needed
			      EditBox 235, 225, 210, 15, actions_taken
			      EditBox 340, 245, 105, 15, worker_signature
			      ButtonGroup ButtonPressed
			        OkButton 340, 265, 50, 15
			        CancelButton 395, 265, 50, 15
			        PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
			        PushButton 260, 20, 20, 10, "FS", ELIG_FS_button
			        PushButton 280, 20, 20, 10, "HC", ELIG_HC_button
			        PushButton 300, 20, 20, 10, "MFIP", ELIG_MFIP_button
			        PushButton 260, 30, 20, 10, "GA", ELIG_GA_button
			        PushButton 335, 20, 45, 10, "prev. panel", prev_panel_button
			        PushButton 395, 20, 45, 10, "prev. memb", prev_memb_button
			        PushButton 335, 30, 45, 10, "next panel", next_panel_button
			        PushButton 395, 30, 45, 10, "next memb", next_memb_button
			        PushButton 5, 135, 25, 10, "EMPS", EMPS_button
			        PushButton 10, 205, 25, 10, "BUSI", BUSI_button
			        PushButton 35, 205, 25, 10, "JOBS", JOBS_button
			        PushButton 10, 215, 25, 10, "RBIC", RBIC_button
			        PushButton 35, 215, 25, 10, "UNEA", UNEA_button
			        PushButton 75, 205, 25, 10, "ACCT", ACCT_button
			        PushButton 100, 205, 25, 10, "CARS", CARS_button
			        PushButton 125, 205, 25, 10, "CASH", CASH_button
			        PushButton 150, 205, 25, 10, "OTHR", OTHR_button
			        PushButton 75, 215, 25, 10, "REST", REST_button
			        PushButton 100, 215, 25, 10, "SECU", SECU_button
			        PushButton 125, 215, 25, 10, "TRAN", TRAN_button
			        PushButton 10, 250, 25, 10, "MEMB", MEMB_button
			        PushButton 35, 250, 25, 10, "MEMI", MEMI_button
			        PushButton 60, 250, 25, 10, "MONT", MONT_button
			        PushButton 10, 260, 25, 10, "PARE", PARE_button
			        PushButton 35, 260, 25, 10, "SANC", SANC_button
			        PushButton 60, 260, 25, 10, "TIME", TIME_button
			      Text 5, 115, 20, 10, "YTD:"
			      Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
			      Text 5, 175, 45, 10, "Other notes:"
			      GroupBox 5, 190, 60, 40, "Income panels"
			      GroupBox 70, 190, 110, 40, "Asset panels"
			      Text 280, 250, 60, 10, "Worker signature:"
			      Text 185, 230, 50, 10, "Actions taken:"
			      GroupBox 5, 235, 85, 40, "other STAT panels:"
			      Text 185, 210, 50, 10, "Verifs needed:"
			      Text 125, 35, 40, 10, "HRF status:"
			      Text 130, 115, 35, 10, "Changes?:"
			      GroupBox 330, 5, 115, 40, "STAT-based navigation"
			      Text 5, 35, 55, 10, "HRF datestamp:"
			      Text 5, 55, 55, 10, "Earned income:"
			      Text 5, 75, 60, 10, "Unearned income:"
			      GroupBox 255, 5, 70, 40, "ELIG panels:"
				  If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
					GroupBox 90, 245, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
					Text 95, 255, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
					CheckBox 95, 265, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
				  End If

			    EndDialog
				Dialog Dialog1
				cancel_confirmation
				Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
				MAXIS_dialog_navigation
                If ButtonPressed = income_notes_button Then
				    '-------------------------------------------------------------------------------------------------DIALOG
				    Dialog1 = "" 'Blanking out previous dialog detail
				    BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
				      CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
				      CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
				      CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
				      CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
				      CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
				      CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
				      CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
				      CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
				      CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
				      CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
				      CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
				      ButtonGroup ButtonPressed
				    	PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
				    	CancelButton 295, 205, 50, 15
				      Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
				      GroupBox 5, 60, 340, 45, "JOBS Income"
				      GroupBox 5, 110, 340, 45, "BUSI Income"
				      GroupBox 5, 160, 340, 40, "UNEA Income"
				    EndDialog
                    Dialog dialog1
                    If ButtonPressed = add_to_notes_button Then
                        If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
                        If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
                        If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
                        If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
                        If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
                        If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
                        If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
                        If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
                        If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
                        If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
                        If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
                        If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
                    End If
                End If
				IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
				IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
				IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
				IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
				IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
				IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
			case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
			IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
		LOOP UNTIL case_note_confirmation = vbYes
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false

	'Creating program list---------------------------------------------------------------------------------------------
	If MFIP_check = 1 Then programs_list = "MFIP "
	If SNAP_check = 1 Then programs_list = programs_list & "SNAP "
	If HC_check = 1 Then programs_list = programs_list & "HC "
	If GA_check = 1 Then programs_list = programs_list & "GA "
	If MSA_check = 1 Then programs_list = programs_list & "MSA "

	'Enters the case note-----------------------------------------------------------------------------------------------
	start_a_blank_CASE_NOTE
	Call write_variable_in_case_note("***" & HRF_month & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
	call write_bullet_and_variable_in_case_note("Programs", programs_list)
	call write_bullet_and_variable_in_case_note("Earned income", earned_income)
	call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
    CALL write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
	call write_bullet_and_variable_in_case_note("YTD", YTD)
	call write_bullet_and_variable_in_case_note("Changes", changes)
	call write_bullet_and_variable_in_case_note("EMPS", EMPS)
	call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
	call write_bullet_and_variable_in_case_note("Other notes", other_notes)
	If ten_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 10% sanction.")
	If thirty_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 30% sanction.")
	IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
	call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
	call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
	If next_month_hrf_not_received_checkbox = checked Then
		call write_variable_in_CASE_NOTE("* HRF for next month (" & next_HRF_month & ") has not been received.")
		call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
	End If
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)



	end_msg = "Success! Your HRF for " & HRF_month & " has been case noted."

End If

script_end_procedure_with_error_report(end_msg & vbcr & "Please make sure to accept the Work items in ECF associated with this HRF. Thank you!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------02/27/2024
'--Tab orders reviewed & confirmed----------------------------------------------02/27/2024
'--Mandatory fields all present & Reviewed--------------------------------------02/27/2024
'--All variables in dialog match mandatory fields-------------------------------02/27/2024
'Review dialog names for content and content fit in dialog----------------------02/27/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------02/27/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------02/27/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------02/27/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------02/27/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------02/27/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------02/27/2024
'--PRIV Case handling reviewed -------------------------------------------------02/27/2024
'--Out-of-County handling reviewed----------------------------------------------NA
'--script_end_procedures (w/ or w/o error messaging)----------------------------02/27/2024
'--BULK - review output of statistics and run time/count (if applicable)--------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------02/27/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------02/27/2024
'--Incrementors reviewed (if necessary)-----------------------------------------02/27/2024
'--Denomination reviewed -------------------------------------------------------02/27/2024
'--Script name reviewed---------------------------------------------------------02/27/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------02/27/2024
'--comment Code-----------------------------------------------------------------02/27/2024
'--Update Changelog for release/update------------------------------------------02/27/2024
'--Remove testing message boxes-------------------------------------------------02/27/2024
'--Remove testing code/unnecessary code-----------------------------------------02/27/2024
'--Review/update SharePoint instructions----------------------------------------02/27/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------02/27/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------02/27/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------02/27/2024
'--Complete misc. documentation (if applicable)---------------------------------02/27/2024
'--Update project team/issue contact (if applicable)----------------------------02/27/2024