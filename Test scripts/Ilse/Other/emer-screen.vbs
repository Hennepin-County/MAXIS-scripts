'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMERGENCY SCREENING.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'Defining some constants to make array life easier
const clt_name = 0
const clt_ref  = 1
const include_clt = 2
const clt_a_c = 3
const clt_asset_total = 4
const clt_sav_acct = 5
const clt_chk_acct = 6
const clt_asset_other_type = 7
const clt_asset_other_bal = 8
const asset_verif = 9
const clt_ei_gross = 10
const clt_ei_net = 11
const clt_ssi_income = 12
const clt_rsdi_income = 13
const clt_ssa_verif = 14
const clt_other_unea_1_type = 15
const clt_other_unea_1_amt = 16
const clt_other_unea_1_verif = 17
const clt_other_unea_2_type = 18
const clt_other_unea_2_amt = 19
const clt_other_unea_2_verif = 20

'FUNCTIONS==============================================================================================

function MEMB_NUMBER_BUTTON_PRESSED
	MEMB_function
	HH_member_array = ""
	FOR i = 0 to total_clients
		IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
			IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
				'msgbox all_clients_
				HH_member_array = Right(all_clients_array(i, 0), len(all_clients_array(i, 0))   ) & ", " & HH_member_array
			END IF
		END IF
	NEXT
	hh_size_split = Len(HH_member_array) - Len(Replace(HH_member_array,",",""))
	hh_size = CStr(hh_size_split)
end function

function SELECT_EMERGENCY_BUTTON_PRESSED

	BeginDialog btn_dialog, 0, 0, 256, 140, "Shelter Information Calc"
	  CheckBox 25, 25, 145, 10, "Eviction - past due rent", eviction_type
	  DropListBox 175, 25, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", eviction_verification
	  CheckBox 25, 50, 145, 10, "Damage deposit for new/affordable place", Damage_deposit_type
	  DropListBox 175, 50, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", dd_verification
	  CheckBox 25, 75, 105, 10, "Utility disconnection/shut off", utility_type
	  DropListBox 175, 75, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", disconnection_verification
	  ButtonGroup ButtonPressed
		OkButton 100, 115, 50, 15
	  GroupBox 5, 5, 245, 100, "What Type of Emergency?"
	  GroupBox 15, 15, 225, 30, ""
	  GroupBox 15, 40, 225, 30, ""
	  GroupBox 15, 65, 225, 30, ""
	EndDialog
	Do
		err_msg = ""
		MsgBox "Select Emergency Dialog"
		Dialog btn_dialog + vbSystemModal
		If eviction_type = unchecked AND damage_deposit_type = unchecked AND utility_type = unchecked Then err_msg = err_msg & vbnewLine & "You must pick and emergency type. Please select at least 1."
		If err_msg <> "" Then msgbox "*** What is the emergency for? ***" & vbnewLine & err_msg
	Loop Until err_msg = ""

	EMERGENCY_NEED_BUTTON_PRESSED

end function

function EMERGENCY_NEED_BUTTON_PRESSED

	BeginDialog btn_dialog, 0, 0, 186, 130, "Emergency Need Information Calc"
	  EditBox 65, 10, 50, 15, rent_due
	  DropListBox 120, 10, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", rent_due_verification
	  EditBox 65, 30, 50, 15, late_fees
	  DropListBox 120, 30, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", late_fees_verification
	  EditBox 65, 50, 50, 15, damage_dep
	  DropListBox 120, 50, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", dd_verification
	  EditBox 65, 70, 50, 15, court_fees
	  DropListBox 120, 70, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", court_fees_verification
	  EditBox 65, 90, 50, 15, hest_due
	  DropListBox 120, 90, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", utility_verification
	  ButtonGroup ButtonPressed
		OkButton 130, 110, 50, 15
	  Text 5, 35, 35, 10, "Late fees:"
	  Text 5, 75, 35, 10, "Court fees:"
	  Text 5, 15, 50, 10, "Past Due Rent:"
	  Text 5, 95, 25, 10, "Utility:"
	  Text 5, 55, 60, 10, "Damage Deposit:"
	EndDialog

	MsgBox "Emergency Need dialog - 2"
	Dialog btn_dialog + vbSystemModal

	rent_due = rent_due * 1
	late_fees = late_fees * 1
	damage_dep = damage_dep * 1
	court_fees = court_fees * 1
	hest_due = hest_due * 1

end function

function EARNED_INCOME_BUTTON_PRESSED

	const employee      = 0
	const employer      = 1
	const how_many_chck = 2
	const check_1_date  = 3
	const check_1_gross = 4
	const check_1_net   = 5
	const check_2_date  = 6
	const check_2_gross = 7
	const check_2_net   = 8
	const check_3_date  = 9
	const check_3_gross = 10
	const check_3_net   = 11
	const check_4_date  = 12
	const check_4_gross = 13
	const check_4_net   = 14
	const check_5_date  = 15
	const check_5_gross = 16
	const check_5_net   = 17
	const job_gross     = 18
	const job_net       = 19
	const job_verif     = 20

	For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
		FULL_EMER_ARRAY(clt_ei_gross, all_clts) = 0
		FULL_EMER_ARRAY(clt_ei_net, all_clts) = 0
	Next

	Do
		add_to_len = 0
		For every_one = 0 to UBound(EI_ARRAY, 2)
			add_to_len = add_to_len + 20
			add_to_len = add_to_len + (20 * EI_ARRAY(how_many_chck, every_one))
		Next

		BeginDialog btn_dialog, 0, 0, 340, 40 + add_to_len, "Earned Income"
			y_pos = 0
			For job_in_case = 0 to UBound(EI_ARRAY, 2)
				DropListBox 5, 20 + y_pos, 105, 45, HH_Memb_DropDown, EI_ARRAY(employee, job_in_case)
				EditBox 120, 20 + y_pos, 130, 15, EI_ARRAY(employer, job_in_case)
				DropListBox 260, 20 + y_pos, 35, 45, " "+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5", EI_ARRAY(how_many_chck, job_in_case)
				ButtonGroup ButtonPressed
			  	  PushButton 310, 20 + y_pos, 25, 15, "Enter", job_enter
				array_counter = 3
				If EI_ARRAY(how_many_chck, job_in_case) <> "" AND EI_ARRAY(how_many_chck, job_in_case) <> " " Then
					'If EI_ARRAY(job_verif, job_in_case) = "" Then EI_ARRAY(job_verif, job_in_case) = "Verifications?"
					Text 35, 40 + y_pos, 20, 10, "Date"
					Text 120, 40 + y_pos, 50, 10, "Gross Amount"
					Text 200, 40 + y_pos, 50, 10, "Net Amount"
					DropListBox 260, 40 + y_pos, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", EI_ARRAY(job_verif, job_in_case)
					For checks_to_enter = 1 to EI_ARRAY(how_many_chck, job_in_case)
						EditBox 35, 55 + y_pos, 50, 15, EI_ARRAY(array_counter, job_in_case)
						EditBox 120, 55 + y_pos, 50, 15, EI_ARRAY(array_counter + 1, job_in_case)
						EditBox 200, 55 + y_pos, 50, 15, EI_ARRAY(array_counter + 2, job_in_case)
						array_counter = array_counter + 3
						y_pos = y_pos + 15
					Next
					y_pos = y_pos + 10
				Else
					y_pos = y_pos + 0
				End If
			Next
			ButtonGroup ButtonPressed
			  PushButton 5, 40 + y_pos, 10, 15, "+", plus_button
			  PushButton 15, 40 + y_pos, 10, 15, "-", minus_button
			  OkButton 285, 40 + y_pos, 50, 15
			Text 5, 5, 45, 10, "HH Member"
			Text 120, 5, 40, 10, "Employer"
			Text 260, 5, 40, 10, "# of Checks"
		EndDialog

		Dialog btn_dialog
		If ButtonPressed = plus_button Then
		 	add_another = Ubound(EI_ARRAY, 2) + 1
			ReDim Preserve EI_ARRAY (19, add_another)
		End If
	Loop Until ButtonPressed = -1

	case_ei_gross = 0
	case_ei_net = 0
	For each_job = 0 to UBOUND(EI_ARRAY, 2)
		EI_ARRAY(job_gross, each_job) = EI_ARRAY(check_1_gross, each_job) + EI_ARRAY(check_2_gross, each_job) + EI_ARRAY(check_3_gross, each_job) + EI_ARRAY(check_4_gross, each_job) + EI_ARRAY(check_5_gross, each_job)
		EI_ARRAY(job_net, each_job)   = EI_ARRAY(check_1_net, each_job) + EI_ARRAY(check_2_net, each_job) + EI_ARRAY(check_3_net, each_job) + EI_ARRAY(check_4_net, each_job) + EI_ARRAY(check_5_net, each_job)

		For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
			If Left(EI_ARRAY(employee, each_job), 2) = FULL_EMER_ARRAY(clt_ref, all_clts) Then
				FULL_EMER_ARRAY(clt_ei_gross, all_clts) = FULL_EMER_ARRAY(clt_ei_gross, all_clts) + EI_ARRAY(job_gross, each_job)
				FULL_EMER_ARRAY(clt_ei_net, all_clts) = FULL_EMER_ARRAY(clt_ei_net, all_clts) + EI_ARRAY(job_net, each_job)
			End If
		Next

		case_ei_gross = case_ei_gross + EI_ARRAY(job_gross, each_job)
		case_ei_net = case_ei_net + EI_ARRAY(job_net, each_job)
	Next
	case_ei_gross = case_ei_gross *1
	case_ei_net = case_ei_net *1
end function

function ASSETS_BUTTON_PRESSED

	BeginDialog btn_dialog, 0, 0, 400, 60 + (20 * UBOUND(FULL_EMER_ARRAY, 2)), "Assets"
	  Text 5, 5, 50, 10, "Person"
	  Text 110, 5, 35, 10, "Checking"
	  Text 165, 5, 30, 10, "Savings"
	  Text 220, 5, 50, 10, "Other"
	  For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
	  	If FULL_EMER_ARRAY(include_clt, all_clts) = checked Then
		  	Text 5, 20 + (20 * all_clts), 90, 10, FULL_EMER_ARRAY(clt_ref, all_clts) & " - " & FULL_EMER_ARRAY(clt_name, all_clts)
		  	EditBox 110, 20 + (20 * all_clts), 40, 15, FULL_EMER_ARRAY(clt_chk_acct, all_clts)
		  	EditBox 165, 20 + (20 * all_clts), 40, 15, FULL_EMER_ARRAY(clt_sav_acct, all_clts)
		  	ComboBox 220, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Debit Card"+chr(9)+"Cash", FULL_EMER_ARRAY(clt_asset_other_type, all_clts)
		  	EditBox 285, 20 + (20 * all_clts), 40, 15, FULL_EMER_ARRAY(clt_asset_other_bal, all_clts)
		  	DropListBox 335, 20 + (20 * all_clts), 60, 45, "Verification?"+chr(9)+"Requested"+chr(9)+"Received", FULL_EMER_ARRAY(asset_verif, all_clts)
		End If
	  Next
	  ButtonGroup ButtonPressed
	    OkButton 345, 40 + (20 * UBOUND(FULL_EMER_ARRAY, 2)), 50, 15
	EndDialog

	Dialog btn_dialog

	total_case_assets = 0
	For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
		If FULL_EMER_ARRAY(clt_chk_acct, all_clts)        = "" Then FULL_EMER_ARRAY(clt_chk_acct, all_clts) = 0
		If FULL_EMER_ARRAY(clt_sav_acct, all_clts)        = "" Then FULL_EMER_ARRAY(clt_sav_acct, all_clts) = 0
		If FULL_EMER_ARRAY(clt_asset_other_bal, all_clts) = "" Then FULL_EMER_ARRAY(clt_asset_other_bal, all_clts) = 0

		FULL_EMER_ARRAY(clt_asset_total, all_clts) = FULL_EMER_ARRAY(clt_chk_acct, all_clts) + FULL_EMER_ARRAY(clt_sav_acct, all_clts) + FULL_EMER_ARRAY(clt_asset_other_bal, all_clts)
		total_case_assets = total_case_assets + FULL_EMER_ARRAY(clt_asset_total, all_clts)
	Next

end function

function UNEA_BUTTON_PRESSED

	BeginDialog btn_dialog, 0, 0, 556, 60 + (20 * UBOUND(FULL_EMER_ARRAY, 2)), "Unearned Income"
	  Text 5, 5, 50, 10, "Person"
	  Text 110, 5, 35, 10, "RSDI"
	  Text 145, 5, 30, 10, "SSI"
	  Text 235, 5, 50, 10, "Other - 1"
	  Text 395, 5, 50, 10, "Other - 2"
	  For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
	  	  If FULL_EMER_ARRAY(include_clt, all_clts) = checked Then
		  	  Text 5, 20 + (20 * all_clts), 90, 10, FULL_EMER_ARRAY(clt_ref, all_clts) & " - " & FULL_EMER_ARRAY(clt_name, all_clts)
			  EditBox 110, 20 + (20 * all_clts), 25, 15, FULL_EMER_ARRAY(clt_rsdi_income, all_clts)
			  EditBox 145, 20 + (20 * all_clts), 25, 15, FULL_EMER_ARRAY(clt_ssi_income, all_clts)
			  DropListBox 175, 20 + (20 * all_clts), 45, 45, "Verification?"+chr(9)+"Requested"+chr(9)+"Received", FULL_EMER_ARRAY(clt_ssa_verif, all_clts)
			  ComboBox 235, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", FULL_EMER_ARRAY(clt_other_unea_1_type, all_clts)

			  EditBox 300, 20 + (20 * all_clts), 40, 15, FULL_EMER_ARRAY(clt_other_unea_1_amt, all_clts)
			  DropListBox 345, 20 + (20 * all_clts), 45, 45, "Verification?"+chr(9)+"Requested"+chr(9)+"Received", FULL_EMER_ARRAY(clt_other_unea_1_verif, all_clts)
			  ComboBox 395, 20 + (20 * all_clts), 60, 45, ""+chr(9)+"Other"+chr(9)+"Child Support"+chr(9)+"SSI"+chr(9)+"RSDI"+chr(9)+"Non-MN PA"+chr(9)+"VA Disability Benefit"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"VA Aid & Attendance"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Req FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Req FS"+chr(9)+"Dividends"+chr(9)+"Interest"+chr(9)+"Cnt Gifts Or Prizes"+chr(9)+"Strike Benefit 27 Contract For Deed"+chr(9)+"Illegal Income"+chr(9)+"Infrequent <30 Not Counted"+chr(9)+"Other FS Only"+chr(9)+"Infreq <= $20 MSA Exclusion"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Spousal Sup"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"County 88 Gaming", FULL_EMER_ARRAY(clt_other_unea_2_type, all_clts)

			  EditBox 460, 20 + (20 * all_clts), 40, 15, FULL_EMER_ARRAY(clt_other_unea_2_amt, all_clts)
			  DropListBox 505, 20 + (20 * all_clts), 45, 45, "Verification?"+chr(9)+"Requested"+chr(9)+"Received", FULL_EMER_ARRAY(clt_other_unea_2_verif, all_clts)
		  End If
	  Next
	  ButtonGroup ButtonPressed
	    OkButton 500, 40 + (20 * UBOUND(FULL_EMER_ARRAY, 2)), 50, 15
	EndDialog

	Dialog btn_dialog

	total_case_unea = 0
	For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
		If FULL_EMER_ARRAY(clt_rsdi_income, all_clts)      = "" Then FULL_EMER_ARRAY(clt_rsdi_income, all_clts) = 0
		If FULL_EMER_ARRAY(clt_ssi_income, all_clts)       = "" Then FULL_EMER_ARRAY(clt_ssi_income, all_clts) = 0
		If FULL_EMER_ARRAY(clt_other_unea_1_amt, all_clts) = "" Then FULL_EMER_ARRAY(clt_other_unea_1_amt, all_clts) = 0
		If FULL_EMER_ARRAY(clt_other_unea_2_amt, all_clts) = "" Then FULL_EMER_ARRAY(clt_other_unea_2_amt, all_clts) = 0

		total_case_unea = total_case_unea + FULL_EMER_ARRAY(clt_rsdi_income, all_clts) + FULL_EMER_ARRAY(clt_ssi_income, all_clts) + FULL_EMER_ARRAY(clt_other_unea_1_amt, all_clts) + FULL_EMER_ARRAY(clt_other_unea_2_amt, all_clts)
	Next

end function

function SHELTER_BUTTON_PRESSED
	BeginDialog btn_dialog, 0, 0, 216, 70, "Shelter Information Calc"
	  EditBox 95, 10, 50, 15, rent_portion
	  DropListBox 150, 10, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", rent_verification
	  EditBox 95, 30, 50, 15, other_fees
	  DropListBox 150, 30, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", Other_fees_verification
	  ButtonGroup ButtonPressed
	    OkButton 160, 50, 50, 15
	  Text 10, 15, 85, 10, "Clt Portion Monthly Rent:"
	  Text 10, 35, 75, 10, "Other fees(garage,etc):"
	EndDialog

	Dialog btn_dialog

	rent_portion = rent_portion * 1
	other_fees = other_fees * 1
end function

function EXPENSE_BUTTON_PRESSED
	Do
		err_msg = ""

		food_allotment_expense = "Food Allotment($" & thrifty_food & ") - FS/MF-FS issued ($" & fs_mf_total & ") = $" & fs_expense

		BeginDialog btn_dialog, 0, 0, 261, 195, "Living Expense Paid from:" & chr(9) & dateadd("d", -30, app_date) & chr(9) & " To:" & chr(9) & dateadd("d", -1, app_date) & chr(9)
		  EditBox 120, 45, 50, 15, shel_paid
		  DropListBox 175, 45, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", shel_verification
		  EditBox 120, 65, 50, 15, hest_paid
		  DropListBox 175, 65, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", hest_verification
		  DropListBox 120, 85, 50, 45, "Select One"+chr(9)+"Yes"+chr(9)+"No", flat_living_expense
		  EditBox 120, 105, 50, 15, actual_paid
		  DropListBox 175, 105, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", actual_verification
		  EditBox 120, 125, 50, 15, other_paid
		  DropListBox 175, 125, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", other_paid_verification
		  ButtonGroup ButtonPressed
			OkButton 105, 160, 50, 15
		  Text 80, 70, 40, 10, "Utility Paid:"
		  Text 100, 130, 20, 10, "Other:"
		  Text 45, 50, 75, 10, "Shelter Expense Paid:"
		  Text 35, 90, 85, 10, "Flat $500 Living Expense:"
		  Text 45, 110, 75, 10, "Actual Living Expense:"
		  Text 30, 30, 205, 10, food_allotment_expense
		  GroupBox 5, 10, 250, 135, "Living Expense Paid from:" & chr(9) & dateadd("d", -30, app_date) & chr(9) & " To:" & chr(9) & dateadd("d", -1, app_date) & chr(9)
		EndDialog

		Dialog btn_dialog
		If shel_paid = "" then shel_paid = "0"
		If hest_paid = "" then hest_paid = "0"
		If actual_paid = "" then actual_paid = "0"
		If other_paid = "" then other_paid = "0"
		If actual_paid <> "0" and flat_living_expense = "Yes" then err_msg = "You selected 'Yes' for Flat $500 Living Expense, you cannot list amounts in 'Actual Living Expense field.' Please correct this."
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	Loop until err_msg = ""
end function

function CLT_PYMT_BUTTON_PRESSED
	BeginDialog btn_dialog, 0, 0, 281, 135, "Enter Client Payments"
	  ButtonGroup ButtonPressed
	    OkButton 220, 115, 50, 15
	  Text 10, 30, 170, 10, "Reported assets of $" & FormatNumber(total_assets)
	  Text 10, 50, 170, 10, percent_test
	  Text 10, 70, 170, 10, shel_max_test
	  Text 10, 90, 170, 10, hest_due_test
	  Text 10, 10, 165, 10, "Case tests"
	  Text 220, 10, 60, 10, "Payment Amount"
	  Text 10, 105, 150, 30, "Entering a payment amount will resolve any failing test with the condition the payment is received prior to approval of EA funds."
	  EditBox 220, 25, 50, 15, clt_portion_assets
	  EditBox 220, 45, 50, 15, clt_portion_percent
	  EditBox 220, 65, 50, 15, clt_portion_shel
	  EditBox 220, 85, 50, 15, clt_portion_hest
	EndDialog

	Dialog btn_dialog

	If clt_portion_assets = "" Then clt_portion_assets = 0
	If clt_portion_percent = "" Then clt_portion_percent = 0
	If clt_portion_shel = "" Then clt_portion_shel = 0
	If clt_portion_hest = "" then clt_portion_hest = 0
	client_payment = clt_portion_assets + clt_portion_percent + clt_portion_shel + clt_portion_hest

	If clt_portion_percent <> 0 Then
	percent_test = ":: 50% test: PASSED! *pending other payment of $" & clt_portion_percent
	test_pass3 = true
	End If

	If clt_portion_shel <> 0 Then
	shel_max_test = ":: Under Shelter Max: PASSED! *pending other payment of $" & clt_portion_shel
	test_pass5 = true
	End if

	If clt_portion_hest <> 0 Then
	hest_due_test = ":: Under Utilities Max: PASSED! *pending other payment of $" & clt_portion_hest
	test_pass6 = true
	End if
end function

function STAT_NAV
	'This part works with the prev/next buttons on several of our dialogs. You need to name your buttons prev_panel_button, next_panel_button, prev_memb_button, and next_memb_button in order to use them.
	EMReadScreen STAT_check, 4, 20, 21
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		End if
	End if
	If ButtonPressed = ADDR_button then call navigate_to_MAXIS_screen("stat", "ADDR")
	 If ButtonPressed = SHEL_button then call navigate_to_MAXIS_screen("stat", "SHEL")
	If ButtonPressed = BUSI_button then call navigate_to_MAXIS_screen("stat", "BUSI")
	If ButtonPressed = JOBS_button then call navigate_to_MAXIS_screen("stat", "JOBS")
	If ButtonPressed = MEMB_button then call navigate_to_MAXIS_screen("stat", "MEMB")
	If ButtonPressed = TYPE_button then call navigate_to_MAXIS_screen("stat", "TYPE")
	If ButtonPressed = PROG_button then call navigate_to_MAXIS_screen("stat", "PROG")
	If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("stat", "REVW")
	If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("stat", "UNEA")
	If ButtonPressed = CURR_button then call navigate_to_MAXIS_screen("case", "CURR")
	If ButtonPressed = INQX_button then
		Call navigate_to_MAXIS_screen("MONY", "INQX")
		EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
		EMWriteScreen begin_search_year, 6, 41
		EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
		EMWriteScreen MAXIS_footer_year, 6, 56
		transmit
	End If
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		End if
	End if
end function

'this function creates the hh member dynamic dialog
function MEMB_function

	BeginDialog btn_dialog, 0,  0, 256, (35 + (UBound(FULL_EMER_ARRAY, 2) + 1) * 15), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	  Text 10, 5, 145, 10, "Who is applying?:"
	  FOR all_clts = 0 to UBound(FULL_EMER_ARRAY, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
	  	  Checkbox 10, (20 + (all_clts * 15)), 150, 10, FULL_EMER_ARRAY(clt_ref, all_clts) & " - " &  FULL_EMER_ARRAY(clt_name, all_clts) & "  " & FULL_EMER_ARRAY(clt_a_c, all_clts), FULL_EMER_ARRAY(include_clt, all_clts)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	  NEXT
	  ButtonGroup ButtonPressed
	  OkButton 200, 20, 50, 15
	EndDialog

	Dialog btn_dialog

	HH_size = 0
	For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
		If FULL_EMER_ARRAY(include_clt, all_clts) = checked Then HH_size = HH_size + 1
	Next

	FIND_FPG_THRIFTY_FOOD
end function

function FIND_FPG_THRIFTY_FOOD
'FPG and Thrifty standards'
	If HH_size = 0 then
		FPG_size = "$0.00"
		thrifty_food = "0"
	ElseIf HH_size = 1 then
		FPG_size = "$1962"
		thrifty_food = "194"
	ElseIf HH_size = 2 then
		FPG_size = "$2655"
		thrifty_food = "357"
	ElseIf HH_size = 3 then
		FPG_size = "$3348"
		thrifty_food = "511"
	ElseIf HH_size = 4 then
		FPG_size = "$4042"
		thrifty_food = "649"
	ElseIf HH_size = 5 then
		FPG_size = "$4735"
		thrifty_food = "771"
	ElseIf HH_size = 6 then
		FPG_size = "$5428"
		thrifty_food = "925"
	ElseIf HH_size = 7 then
		FPG_size = "$6122"
		thrifty_food = "1022"
	ElseIf HH_size = 8 then
		FPG_size = "$6815"
		thrifty_food = "1169"
	ElseIf HH_size = 9 then
		FPG_size = "$7508"
		thrifty_food = "1315"
	ElseIf HH_size = 10 then
		FPG_size = "$8202"
		thrifty_food = "1461"
	ElseIf HH_size = 11 then
		FPG_size = "$8895"
		thrifty_food = "1607"
	ElseIf HH_size = 12 then
		FPG_size = "$9588"
		thrifty_food = "1753"
	ElseIf HH_size = 13 then
		FPG_size = "$10281"
		thrifty_food = "1899"
	ElseIf HH_size = 14 then
		FPG_size = "$10974"
		thrifty_food = "2045"
	ElseIf HH_size = 15 then
		FPG_size = "$11667"
		thrifty_food = "2191"
	ElseIf HH_size = 16 then
		FPG_size = "$12360"
		thrifty_food = "2337"
	ElseIf HH_size = 17 then
		FPG_size = "$13053"
		thrifty_food = "2483"
	ElseIf HH_size = 18 then
		FPG_size = "$13746"
		thrifty_food = "2629"
	ElseIf HH_size = 19 then
		FPG_size = "$14439"
		thrifty_food = "2775"
	ElseIf HH_size = 20 then
		FPG_size = "$15132"
		thrifty_food = "2921"
	End If
end function


'=======================================================================================================
EMConnect ""

'Declaring our arrays - because life is easier with arrays
Dim FULL_EMER_ARRAY ()
ReDim FULL_EMER_ARRAY (20, 0)
Dim EI_ARRAY ()
ReDim  EI_ARRAY (20, 0)



call check_for_MAXIS(False)	'checking for an active MAXIS session

call MAXIS_case_number_finder(MAXIS_case_number)
'Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'formats default date'
If len(datepart("m", date())) = 1 then
	m = "0" & datepart("m", date())
Else
	m = datepart("m",date())
End IF
If len(datepart("d", date())) = 1 then
	d = "0" & datepart("d", date())
Else
	d = datepart("d",date())
End IF
If len(datepart("m", date()-30)) = 1 then
	ea_eval_m = "0" & datepart("m", date()+10)
Else
	ea_eval_m = datepart("m",date()-30)
End IF
If len(datepart("d", date()-30)) = 1 then
	ea_eval_d = "0" & datepart("d", date()+10)
Else
	ea_eval_d = datepart("d",date()-30)
End IF
app_date= m & "/" & d & "/" & right(datepart("yyyy", date()), 2)
'determines EA Eval Period'
ea_eval_date = ea_eval_m & "/" & ea_eval_d & "/" & right(datepart("yyyy", date()-30), 2)

BeginDialog case_number_dialog, 0, 0, 281, 85, "EA/EGA Screening"
  EditBox 65, 5, 80, 15, MAXIS_case_number
  CheckBox 170, 10, 110, 10, "Check here for a Quick Screen", quick_screen_checkbox
  EditBox 85, 25, 60, 15, app_date
  DropListBox 215, 25, 60, 45, "Select One"+chr(9)+"EA"+chr(9)+"EGA", prog_type_case_dialog
  EditBox 75, 45, 200, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 170, 65, 50, 15
    CancelButton 225, 65, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 75, 10, "Date of app (xx/xx/xx):"
  Text 175, 30, 35, 10, "EA/EGA?:"
  Text 5, 50, 65, 10, "Worker's Signature:"
EndDialog

'The Script'
Do
	err_msg = ""
	Dialog case_number_dialog
	cancel_confirmation
	If MAXIS_case_number = "" then err_msg = err_msg & vbCr & "You must have a case number to continue."
	If len(MAXIS_case_number) > 8 then err_msg = err_msg & vbCr & "Your case number need to be 8 digits or less."
	If prog_type_case_dialog = "Select One" then err_msg = err_msg & vbCr & "You must choose a program type."
	If DateValue(app_date) > Date() then err_msg = err_msg & vbCr & "You cannot enter a future application date."
	If err_msg <> "" then Msgbox err_msg
	call check_for_password (are_we_passworded_out) 'adding functionality for MAXIS v.6 Password Out issue'
Loop until err_msg = ""


'reformats App Date Again'
If len(datepart("m", app_date)) = 1 then
	m = "0" & datepart("m", app_date)
Else
	m = datepart("m", app_date)
End IF

MAXIS_footer_month = m
MAXIS_footer_year = right(datepart("yyyy", app_date), 2)
back_to_self

'DATE CALCULATIONS From Ilse Hennepin County however Ramsey County goes by date of EA/EGA payment issuance
'creating month variable 13 months prior to current footer month/year to search for EMER programs issued
begin_search_month = dateadd("m", -13, app_date)
begin_search_year = datepart("yyyy", begin_search_month)
begin_search_year = right(begin_search_year, 2)
begin_search_month = datepart("m", begin_search_month)
If len(begin_search_month) = 1 then begin_search_month = "0" & begin_search_month
'End of date calculations----------------------------------------------------------------------------------------------

'Creating the dropdown of HH Members for use in dialogs
Call Generate_Client_List(HH_Memb_DropDown, "Select One...")

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_search_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 9, 50		'selecting EA
EMWriteScreen "x", 11, 50		'selecting EGA
transmit

'searching for EA/EG issued on the INQD screen
DO
	row = 6
	DO
		EMReadScreen emer_issued, 1, row, 16		'searching for EMER programs as they start with E
		IF emer_issued = "E" then
			'reading the EMER information for EMER issuance
			EMReadScreen EMER_type, 2, row, 16
			EMReadScreen EMER_amt_issued, 7, row, 39
			EMReadScreen EMER_elig_start_date, 8, row, 7
			'EMReadScreen EMER_elig_end_date, 8, row, 73
			exit do
		ELSE
			row = row + 1
		END IF
	Loop until row = 18				'repeats until the end of the page
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

'creating variables and conditions for EMER screening
New_EMER_year = dateadd("YYYY", 1, EMER_elig_start_date)
EMER_available_date = dateadd("d", 1, New_EMER_year)	'creating emer available date that is 1 day & 1 year past the EMER_elig_end_date
EMER_last_used_dates = EMER_elig_start_date ''& " - " & EMER_elig_end_date	'combining dates into new variable

If emer_issued <> "E" or datevalue(app_date) > datevalue(EMER_available_date) then	'creating variables for cases that have not had EMER issued in current 13 months
 	EMER_last_used_dates = "n/a"
	EMER_available_date = "Currently available"
END IF

'Declares a variable from EA Evaluation start date to be use for inqx search programs'
begin_eval_day = dateadd("d", -30, app_date)
begin_eval_month = datepart("m", begin_eval_day)
begin_eval_year = datepart("yyyy", begin_eval_day)
begin_eval_year = right(begin_eval_year, 2)
If len(begin_eval_month)= 1 then begin_eval_month = "0" & begin_eval_month

'Screen FS Prog'
Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_eval_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_eval_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 9, 5

transmit

fs_amt_total = 0
For maxis_row = 6 to 18
	EMReadScreen issued_date, 8, maxis_row, 7
	If issued_date <> "        " then
    	If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then
			EMReadScreen snap_grant, 6, maxis_row, 39
			snap_grant = replace(snap_grant, " ","")
			snap_grant = snap_grant * 1
			fs_amt_total = fs_amt_total + snap_grant
			fs_prog = true
		End If
	End If
Next

'Screen MFIP Prog'

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_eval_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_eval_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 10, 5

transmit

mf_amt_total = 0
mf_fs_amt_total = 0
mf_hg_amt_total = 0

For maxis_row = 6 to 18
	EMReadScreen issued_date, 8, maxis_row, 7
	EMReadScreen prog_type, 5, maxis_row, 16
	If issued_date <> "        " then
		If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then

			If prog_type = "MF-MF" then
				EMReadScreen mf_mf_amt, 6, maxis_row, 39
				mf_mf_amt = replace(mf_mf_amt, " ","")
				mf_mf_amt = mf_mf_amt * 1
            	mf_amt_total = mf_amt_total + mf_mf_amt
				mf_prog = true
			End If

			If prog_type = "MF-FS" then
				EMReadScreen mf_fs_amt_issued, 6, maxis_row, 39
				mf_fs_amt_issued = replace(mf_fs_amt_issued, " ","")
				mf_fs_amt_issued = mf_fs_amt_issued * 1
				mf_fs_amt_total = mf_fs_amt_issued
				mf_fs_prog = true
			End If

			If prog_type = "MF-HG" then
				EMReadScreen mf_hg_amt_issued, 6, maxis_row, 39
				mf_hg_amt_issued = replace(mf_hg_amt_issued, " ","")
				mf_hg_amt_issued = mf_hg_amt_issued * 1
				mf_hg_amt_total = mf_hg_amt_issued
				mf_hg_prog = true
			End If
		End If
	End If
Next

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

memb_row = 5
people_counter = 0
Call navigate_to_MAXIS_screen ("STAT", "MEMB")
Do
	EMReadScreen ref_numb, 2, memb_row, 3
	If ref_numb = "  " Then Exit Do
	ReDim Preserve FULL_EMER_ARRAY(20, people_counter)
	EMWriteScreen ref_numb, 20, 76
	transmit
	EMReadScreen first_name, 12, 6, 63
	EMReadScreen last_name, 25, 6, 30
	EMReadscreen client_age, 2, 8, 76
	client_age = replace(client_age, " ", "")
	client_age = client_age * 1
	If client_age >= 20 then
		client_is = "(ADULT)"
	Else
		client_is = "(CHILD)"
	End If

	FULL_EMER_ARRAY (clt_name, people_counter) = ref_numb
	FULL_EMER_ARRAY (clt_ref, people_counter) = replace(first_name, "_", "") & " " & replace(last_name, "_", "")
	FULL_EMER_ARRAY (include_clt, people_counter) = unchecked
	FULL_EMER_ARRAY (clt_a_c, people_counter) = client_is

	memb_row = memb_row + 1
	people_counter = people_counter + 1
Loop until memb_row = 20

HH_size = people_counter

If quick_screen_checkbox = checked Then

Else

	'THis is for the EA Dialog'
	If prog_type_case_dialog = "EA" then
		MEMB_function

		Do
			err_msg = ""
			HH_size = 0
			For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
				If FULL_EMER_ARRAY(include_clt, all_clts) = checked Then HH_size = HH_size + 1
			Next
			FIND_FPG_THRIFTY_FOOD

			'Program type'
			Item_Counter = 1
			eviction_msg = ""
			damage_deposit_msg = ""
			utility_msg = ""
			If eviction_type = 1 then
				eviction_msg = Item_Counter & ") For eviction / past due rent  "
				Item_Counter = Item_Counter + 1
			End If
			If damage_deposit_type = 1 then
				damage_deposit_msg = Item_Counter & ") Damage deposit for new/affordable place  "
				Item_Counter = Item_Counter + 1
			End If
			If utility_type = 1 then
				utility_msg = Item_Counter & ") Utility disconnection/shut off  "
				Item_Counter = Item_Counter + 1
			End If
			If eviction_msg = "" and damage_deposit_msg = "" and utility_msg = "" then
				ea_type_msg = "NONE"
			Else
				ea_type_msg = eviction_msg & damage_deposit_msg & utility_msg
			End If

			'totals expenses'
			If rent_portion = "" then rent_portion = "0"
			If other_fees = "" then other_fees = "0"
			If rent_due = "" then rent_due = "0"
			If late_fees = "" then late_fees = "0"
			If court_fees = "" then court_fees = "0"
			If hest_due = "" then hest_due = "0"
			If damage_dep = "" then damage_dep = "0"
			rent_mo = rent_portion + other_fees
			shel_due = rent_due + late_fees + court_fees + hest_due + damage_dep

			'generating verif request list'
			verif_counter = 1
			Verif_request_list = ""
			If rent_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") monthly/rent cost, "
				verif_counter = verif_counter + 1
			End If
			If Other_fees_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") other monthly fees, "
				verif_counter = verif_counter + 1
			End If
			If rent_due_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") rent due balance, "
				verif_counter = verif_counter + 1
			End If
			If late_fees_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") late fees, "
				verif_counter = verif_counter + 1
			End If
			If dd_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") damage deposit fee, "
				verif_counter = verif_counter + 1
			End If
			If court_fees_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") court fees, "
				verif_counter = verif_counter + 1
			End If
			If utility_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") utility bills, "
				verif_counter = verif_counter + 1
			End If
			If eviction_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") eviction notice, "
				verif_counter = verif_counter + 1
			End If
			If disconnection_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") disconnection notice, "
				verif_counter = verif_counter + 1
			End If
			'expense verif'
			If shel_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") shelter expenses paid, "
				verif_counter = verif_counter + 1
			End If
			If hest_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") utilities paid, "
				verif_counter = verif_counter + 1
			End If
			If actual_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") actual living expenses paid, "
				verif_counter = verif_counter + 1
			End If
			If other_paid_verification = "Requested" then
				Verif_request_list = Verif_request_list & verif_counter & ") other expenses paid, "
				verif_counter = verif_counter + 1
			End If

			For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
				If FULL_EMER_ARRAY(asset_verif, all_clts) = "Requested" then
					Verif_request_list = Verif_request_list & verif_counter &  ") Asset balance/bank statements belonging to: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					verif_counter = verif_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_ssa_verif, all_clts) = "Requested" then
					Verif_request_list = Verif_request_list & verif_counter &  ") Social Security Income for : " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					verif_counter = verif_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_other_unea_1_verif, all_clts) = "Requested" then
					Verif_request_list = Verif_request_list & verif_counter &  ") " & FULL_EMER_ARRAY(clt_other_unea_1_type, all_clts) & " income for: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					verif_counter = verif_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_other_unea_2_verif, all_clts) = "Requested" then
					Verif_request_list = Verif_request_list & verif_counter &  ") " & FULL_EMER_ARRAY(clt_other_unea_2_type, all_clts) & " income for: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					verif_counter = verif_counter + 1
				End If
			Next

			For each_job = 0 to UBound(EI_ARRAY, 2)
				If EI_ARRAY(job_verif, each_job) = "Requested" Then
					Verif_request_list = Verif_request_list & verif_counter & ") Paystubs for: " & right(EI_ARRAY(employee, each_job), len(EI_ARRAY(employee, each_job)-2))
					If EI_ARRAY(employer, each_job) <> "" OR UCASE(EI_ARRAY(employer, each_job)) <> "UNKNOWN" Then Verif_request_list = Verif_request_list & " from: " & EI_ARRAY(employer, each_job)
					verif_counter = verif_counter + 1
				End If
			Next


			'generating verif received list'
			docs_counter = 1
			docs_received_list = ""
			If rent_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") monthly/rent cost, "
				docs_counter = docs_counter + 1
			End If
			If Other_fees_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") other monthly fees, "
				docs_counter = docs_counter + 1
			End If
			If rent_due_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") rent due balance, "
				docs_counter = docs_counter + 1
			End If
			If late_fees_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") late fees, "
				docs_counter = docs_counter + 1
			End If
			If dd_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") damage deposit fee, "
				docs_counter = docs_counter + 1
			End If
			If court_fees_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") court fees, "
				docs_counter = docs_counter + 1
			End If
			If utility_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") utility bills, "
				docs_counter = docs_counter + 1
			End If
			If eviction_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") eviction notice, "
				docs_counter = docs_counter + 1
			End If
			If disconnection_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") disconnection notice, "
				docs_counter = docs_counter + 1
			End If
			'expense verif'
			If shel_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") shelter expenses paid, "
				docs_counter = docs_counter + 1
			End If
			If hest_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") utilities paid, "
				docs_counter = docs_counter + 1
			End If
			If actual_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") actual living expenses paid, "
				docs_counter = docs_counter + 1
			End If
			If other_paid_verification = "Received" then
				docs_received_list = docs_received_list & docs_counter & ") other expenses paid, "
				docs_counter = docs_counter + 1
			End If

			For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
				If FULL_EMER_ARRAY(asset_verif, all_clts) = "Received" then
					docs_received_list = docs_received_list & docs_counter &  ") Asset balance/bank statements belonging to: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					docs_counter = docs_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_ssa_verif, all_clts) = "Received" then
					docs_received_list = docs_received_list & docs_counter &  ") Social Security Income for : " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					docs_counter = docs_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_other_unea_1_verif, all_clts) = "Received" then
					docs_received_list = docs_received_list & docs_counter &  ") " & FULL_EMER_ARRAY(clt_other_unea_1_type, all_clts) & " income for: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					docs_counter = docs_counter + 1
				End If

				If FULL_EMER_ARRAY(clt_other_unea_2_verif, all_clts) = "Received" then
					docs_received_list = docs_received_list & docs_counter &  ") " & FULL_EMER_ARRAY(clt_other_unea_2_type, all_clts) & " income for: " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
					docs_counter = docs_counter + 1
				End If
			Next

			For each_job = 0 to UBound(EI_ARRAY, 2)
				If EI_ARRAY(job_verif, each_job) = "Received" Then
					docs_received_list = docs_received_list & docs_counter & ") Paystubs for: " & right(EI_ARRAY(employee, each_job), len(EI_ARRAY(employee, each_job)-2))
					If EI_ARRAY(employer, each_job) <> "" OR UCASE(EI_ARRAY(employer, each_job)) <> "UNKNOWN" Then docs_received_list = docs_received_list & " from: " & EI_ARRAY(employer, each_job)
					docs_counter = docs_counter + 1
				End If
			Next

			'Getting total of adults and ratio responsibility
			adults_applying = 0
			adults_not_applying = 0
			clt_applying = 0
			clt_not_applying = 0
			not_applying_adults_list = ""
			For all_clts = 0 to UBOUND(FULL_EMER_ARRAY, 2)
				If FULL_EMER_ARRAY(include_clt, all_clts) = checked Then
					If FULL_EMER_ARRAY(clt_a_c, all_clts) = "(ADULT)" Then adults_applying = adults_applying + 1
				ElseIf FULL_EMER_ARRAY(include_clt, all_clts) = unchecked Then
					If FULL_EMER_ARRAY(clt_a_c, all_clts) = "(ADULT)" Then adults_not_applying = adults_not_applying + 1
					not_applying_adults_list = not_applying_adults_list & "MEMB " & FULL_EMER_ARRAY(clt_ref, all_clts) & " - " & FULL_EMER_ARRAY(clt_name, all_clts) & ", "
				End If
			Next

			number_of_adults_hh = adults_applying + adults_not_applying
			ratio_responsibility = adults_not_applying/number_of_adults_hh
			adult_not_applying_portion_of_due = Left((shel_due * ratio_responsibility), 7)
			adult_not_applying_each_portion_of_due = shel_due/number_of_adults_hh
			If adults_not_applying <> 0 then
				If shel_due <> "0" then
					hh_msg = "Not applying: " & not_applying_adults_list & " The bal/ratio is split by " & number_of_adults_hh & " adults in the household. $" & FormatNumber(adult_not_applying_portion_of_due) & " must be paid first to pass cost/eff test. Emergency programs will only be approved for $" & FormatNumber(shel_due - adult_not_applying_portion_of_due) & "."
					'test_pass7 = false
				End If
			Else
				hh_msg = ""
				'test_pass7 = true
			End If

			'Programs into dialog'
			fs_results = ""
			mf_results = ""
			mf_fs_results = ""
			mf_hg_results = ""

			If fs_prog = true then fs_results = "FS: $" & fs_amt_total & "   "
			If mf_prog = true then mf_results = "MFIP: $" & mf_amt_total & "   "
			If mf_fs_prog = true then mf_fs_results = "MF-FS: $" & mf_fs_amt_total & "   "
			If mf_hg_prog = true then mf_hg_results = "MF-HG: $" & mf_hg_amt_total & "   "

			'Food allotment - fs grant received'
			fs_mf_total = mf_fs_amt_total + fs_amt_total
			fs_expense = thrifty_food - fs_mf_total
			if fs_expense < 0 then fs_expense = "0"

			'living expense total'
			If flat_living_expense = "Yes" then
				flat_living_expense_amt = "500"
			Else
				flat_living_expense_amt = "0"
			End If

			total_expense = shel_paid + hest_paid + actual_paid + other_paid + fs_expense + flat_living_expense_amt

			'%50 test'
			pa_grants_total = mf_hg_amt_total + mf_amt_total
			total_net_income_for_test = case_ei_net + total_case_unea + pa_grants_total
			half_total_net_income = total_net_income_for_test/2

			shel_max = (rent_mo * 2) + court_fees + late_fees
			shel_max_allowed = rent_due + late_fees + court_fees + damage_dep

			'EA Tests'
			'12 months
			If EMER_available_date = "Currently available" then
			   month_test = ":: 12 month test: PASSED!"
			   test_pass_emer_avail = true
			Else
			   month_test = ":: 12 month test: FAILED!"
			   test_pass_emer_avail = false
			End If
			'FPG test
			If total_gross_income <= FPG_size then
			   FPG_test = ":: FPG test: PASSED!"
			   test_pass_fpg = true
			Else
			   FPG_test = ":: FPG test: FAILED! :: Over by $" & total_gross_income - FPG_size
			   test_pass_fpg = false
			End If
			'50% test
			If half_total_net_income <= total_expense then
			   percent_test = ":: 50% test: PASSED!"
			   test_pass_50perc = true
			Else
			   percent_test = ":: 50% test: FAILED! :: short by $" & half_total_net_income - total_expense
			   test_pass_50perc = false
			End If
			'CostEff test
			If total_net_income_for_test >= rent_mo then
			   cost_eff_test = ":: Cost-Eff: PASSED!"
			   test_pass_cost_eff = true
			Else
			   cost_eff_test = ":: Cost-Eff: FAILED! :: rent over net by $" & rent_mo - total_net_income_for_test
			   test_pass_cost_eff = false
			End If
			'Under Shelter Maximum
			If shel_max >= shel_max_allowed then
			   shel_max_test = ":: Under Shelter Max: PASSED!"
			   test_pass_shel_max = true
			Else
			   shel_max_test = ":: Under Shelter Max: FAILED! :: MAX is: $" & shel_max
			   test_pass_shel_max = false
			End If

			'Under Utilities Maximum
			If hest_due <= 1800 then
				hest_due_test = ":: Under Utilities Max: PASSED!"
			   test_pass_util_max = true
			Else
			   hest_due_test = ":: Under Utilities Max: FAILED! :: over $" & (hest_due - 1800)
			   test_pass_util_max = false
			End If

			'Potential Elig'
			If test_pass_emer_avail = true and test_pass_fpg = true and test_pass_50perc = true and test_pass_cost_eff = true and test_pass_shel_max = true and test_pass_util_max = true then
			  Potential_Elig = "Potential Eligibility?:  ::YES::"
			Else
			  Potential_Elig =  "Potential Eligibility?:  ::NO::" & vbNewLine & "Please resolve the 'FAILED!' tests above to be eligible"
			End If

			If client_payment <> 0 Then
				other_payment_test = "***Payment required prior to EA Approval in the amount of $" & client_payment
			Else
				other_payment_test = ""
			End If

			BeginDialog MAIN_dialog, 0, 0, 470, 435, "Emergency Assistance Screening"
			  'Case Information Section
			  GroupBox 5, 5, 315, 110, "Case Information"
			  CheckBox 210, 35, 90, 10, "Check Here if Same Day", same_day_checkbox
			  Text 15, 20, 65, 10, "Date of Application:"
			  Text 115, 20, 50, 10, date_of_app
			  Text 210, 20, 45, 10, "Active Case:"
			  DropListBox 260, 15, 35, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", active_case
			  'About the emer application
			  Text 15, 35, 90, 10, "Emergency Program Type:"
			  Text 115, 35, 30, 10, "EA/EGA"
			  Text 15, 50, 95, 10, "30 Day Period of evaluation:"
			  Text 115, 50, 110, 10, "From: mm/dd/yy   To: mm/dd/yy"
			  'HH Comp information
			  GroupBox 175, 65, 140, 45, "Household Composition"
			  ButtonGroup ButtonPressed
			    PushButton 275, 75, 35, 10, "HH Memb", HH_memb_button
			  Text 185, 80, 30, 10, "Size: " & HH_size
			  Text 185, 95, 80, 10, "200% FPG = $" & FPG_size
			  'About the 12 month limit
			  GroupBox 15, 65, 150, 45, "Emer Disbursement/ 12 Month History"
			  Text 25, 80, 40, 10, "Last Used:"
			  Text 70, 80, 50, 10, EMER_last_used_dates
			  Text 25, 95, 35, 10, "Available:"
			  Text 65, 95, 80, 10, EMER_available_date

			  'What is the emergency
			  GroupBox 5, 120, 315, 95, "Emergency Information"
			  ButtonGroup ButtonPressed
			    PushButton 20, 140, 35, 15, "Select", select_emergency_button
			  GroupBox 15, 130, 300, 35, ""
			  Text 65, 140, 235, 20, ea_type_msg
			  Text 65, 170, 105, 10, "Rent due: $" & (rent_due + late_fees + court_fees) & " (includes fees)"
			  Text 65, 185, 100, 10, "Damage Deposit: $" & damage_dep
			  Text 65, 200, 80, 10, "Utilities Due: $" & hest_due
			  ButtonGroup ButtonPressed
			    PushButton 20, 170, 25, 15, "Calc", emergency_need_calc_button
			  Text 155, 200, 30, 10, "Landlord:"
			  EditBox 190, 195, 125, 15, land_lord_info
			  ButtonGroup ButtonPressed
			    PushButton 190, 170, 125, 15, "Enter Client Portion Responsibility", clt_pymt_button

			  'info from eval pd
			  GroupBox 5, 220, 315, 135, "Information from 30 Day Eval Period (mm/dd/yy to mm/dd/yy)"
			  GroupBox 10, 230, 165, 40, "Earned Income"
			  GroupBox 185, 230, 130, 40, "Unearned Income"
			  ButtonGroup ButtonPressed
			    PushButton 15, 245, 25, 15, "Calc", earned_income_button
			  Text 50, 245, 100, 10, "Total Gross: $" & case_ei_gross		'GROSS EARNED INCOME
			  Text 50, 255, 100, 10, "Total Net: $" & case_ei_net			'NET EARNED INCOME
			  ButtonGroup ButtonPressed
			    PushButton 195, 245, 25, 15, "Calc", unea_button
			  Text 235, 250, 50, 10, "Total: $" & total_case_unea
			  GroupBox 10, 275, 165, 40, "Expenses Paid"
			  ButtonGroup ButtonPressed
			    PushButton 15, 290, 25, 15, "Calc", expenses_paid_button
			  Text 50, 290, 45, 10, "Total: $"
			  Text 110, 290, 45, 10, "Rent: $"
			  Text 110, 300, 50, 10, "Utilities: $"
			  GroupBox 185, 275, 130, 40, "Available Assets"
			  ButtonGroup ButtonPressed
			    PushButton 195, 290, 25, 15, "Calc", assets_button
			  Text 235, 295, 50, 10, "Total: $" & total_case_assets
			  GroupBox 10, 320, 305, 30, "Benefits Issued in Eval Period"
			  Text 20, 335, 285, 10, fs_results & mf_results & mf_hg_results & mf_fs_results

			  Text 330, 10, 120, 10, Potential_Elig

			  GroupBox 325, 25, 140, 155, "EA TESTS"
			  Text 335, 40, 125, 150, month_test & vbNewLine & "  EMER Last Used: " & EMER_last_used_dates & vbNewLine & FPG_test & vbNewLine & "  200% FPG: " & FPG_size & ", Total Gross: " & (case_ei_gross + total_case_unea) & vbNewLine & percent_test & vbNewLine & "  50% of Net Income: $" & half_total_net_income & ", Expenses Paid: $" & total_expense & vbNewLine & cost_eff_test & vbNewLine & " Net Income: $" & (case_ei_net + total_case_unea) & ", Ongoing Expense: $" & vbNewLine & shel_max_test & vbNewLine & hest_due_test & vbNewLine & hh_msg & vbNewLine &  other_payment_test

			  ButtonGroup ButtonPressed
				PushButton 5, 360, 65, 15, "Monthly Expenses", monthly_expenses_button
			  Text 80, 365, 90, 10, "Monthly Expenses: $" & total_expense

			  GroupBox 325, 180, 140, 90, "Verification Received"
			  Text 330, 190, 130, 75, Verif_received_list
			  GroupBox 325, 265, 140, 100, "Verification Requested"
			  Text 330, 275, 130, 85, Verif_request_list

			  CheckBox 175, 360, 140, 10, "Application Completed and Signed", apllication_complete_checkbox
			  Text 5, 400, 40, 10, "Other notes:"
			  Text 5, 380, 105, 10, "Additional Verification Request:"
			  Text 5, 420, 45, 10, "Action taken:"
			  EditBox 50, 395, 290, 15, other_notes
			  EditBox 50, 415, 290, 15, action_taken
			  EditBox 110, 375, 230, 15, verification_request
			  Text 365, 385, 20, 10, "STAT"
			  ButtonGroup ButtonPressed
			    PushButton 385, 370, 25, 10, "CURR", CURR_button
			    PushButton 385, 380, 25, 10, "MEMB", MEMB_button
			    PushButton 410, 370, 25, 10, "TYPE", TYPE_button
			    PushButton 435, 370, 25, 10, "PROG", PROG_button
			    PushButton 435, 380, 25, 10, "SHEL", SHEL_button
			    PushButton 410, 380, 25, 10, "ADDR", ADDR_button
			    PushButton 410, 400, 25, 10, "prev", prev_panel_button
			    PushButton 435, 400, 25, 10, "next", next_panel_button
			    PushButton 360, 370, 25, 10, "INQX", INQX_button
			    PushButton 435, 390, 25, 10, "BUSI", BUSI_button
			    PushButton 385, 390, 25, 10, "JOBS", JOBS_button
			    PushButton 410, 390, 25, 10, "UNEA", UNEA_button
				CancelButton 415, 410, 50, 15
				OkButton 365, 410, 50, 15
			EndDialog

			Dialog MAIN_dialog
			cancel_confirmation

			clt_portion_assets = clt_portion_assets
			clt_portion_percent = clt_portion_percent
			clt_portion_shel = clt_portion_shel
			clt_portion_hest = clt_portion_hest
			client_payment = client_payment

			STAT_NAV
			If ButtonPressed = HH_memb_button				Then MEMB_function
			If ButtonPressed = select_emergency_button		Then SELECT_EMERGENCY_BUTTON_PRESSED
			If ButtonPressed = emergency_need_calc_button	Then EMERGENCY_NEED_BUTTON_PRESSED
			If ButtonPressed = clt_pymt_button				Then CLT_PYMT_BUTTON_PRESSED
			If ButtonPressed = earned_income_button			Then EARNED_INCOME_BUTTON_PRESSED
			If ButtonPressed = unea_button					Then UNEA_BUTTON_PRESSED
			If ButtonPressed = expenses_paid_button			Then EXPENSE_BUTTON_PRESSED
			If ButtonPressed = assets_button				Then ASSETS_BUTTON_PRESSED
			If ButtonPressed = monthly_expenses_button		Then SHELTER_BUTTON_PRESSED
		Loop Until ButtonPressed = ok_button




	ElseIf prog_type_case_dialog = "EGA" then
		MEMB_function




	End If



End If
