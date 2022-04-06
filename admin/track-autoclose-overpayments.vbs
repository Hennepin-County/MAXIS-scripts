'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - OVERPAYMENT ON AUTOCLOSE PAUSE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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

'COLUMNS'
const det_case_numb_col 		=  1
const det_process_col 			=  2
const det_issued_fs_f_col 		=  3
const det_issued_fs_s_col 		=  4
const det_issued_mf_mf_col 		=  5
const det_issued_mf_fs_f_col 	=  6
const det_issued_mf_fs_s_col 	=  7
const det_issued_mf_hg_col 		=  8
const det_form_col 				=  9
const det_form_date_col 		= 10
const det_intv_col 				= 11
const det_intv_date_col 		= 12
const det_verifs_col 			= 13
const det_process_complete_col 	= 14
const det_op_fs_f_col 			= 15
const det_op_fs_s_col 			= 16
const det_op_mf_mf_col 			= 17
const det_op_mf_fs_f_col 		= 18
const det_op_mf_fs_s_col 		= 19
const det_op_mf_hg_col 			= 20
const det_supp_fs_f_col 		= 21
const det_supp_fs_s_col 		= 22
const det_supp_mf_mf_col 		= 23
const det_supp_mf_fs_f_col 		= 24
const det_supp_mf_fs_s_col 		= 25
const det_supp_mf_hg_col 		= 26
const det_orig_earned_income_col 		= 27
const det_orig_unearned_income_col 		= 28
const det_orig_total_income_col 		= 29
const det_orig_total_ded_col 			= 30
const det_orig_net_income_col 			= 31
const det_orig_housing_cost_col 		= 32
const det_orig_utility_cost_col 		= 33
const det_orig_total_shel_cost_col 		= 34
const det_orig_net_adj_income_col 		= 35
const det_orig_hh_size_col 				= 36
const det_orig_snap_benefit_col 		= 37
const det_correct_earned_income_col 	= 38
const det_correct_unearned_income_col 	= 39
const det_correct_total_income_col 		= 40
const det_correct_total_ded_col 		= 41
const det_correct_net_income_col 		= 42
const det_correct_housing_cost_col 		= 43
const det_correct_utility_cost_col 		= 44
const det_correct_total_shel_cost_col 	= 45
const det_correct_net_adj_income_col 	= 46
const det_correct_hh_size_col 			= 47
const det_snap_proration_col			= 48
const det_correct_snap_benefit_col 		= 49
const det_pdf_link_col 					= 50


const rept_case_numb_col 		= 1
const rept_process_col 			= 2
const rept_issued_fs_f_col 		= 3
const rept_issued_fs_s_col 		= 4
const rept_issued_mf_fs_f_col 	= 5
const rept_issued_mf_fs_s_col 	= 6
const rept_op_fs_f_col 			= 7
const rept_op_fs_s_col 			= 8
const rept_op_mf_fs_f_col 		= 9
const rept_op_mf_fs_s_col 		= 10

'Array parameters'
const ref_number							= 0
const last_name_const						= 1
const first_name_const						= 2
const age_const								= 3
const full_name_const						= 4
const memb_droplist_const					= 5
const earned_income_exists_const			= 6
const unearned_income_exists_const			= 7
const mfip_elig								= 8
const earned_inc_budgeted_const				= 9
const earned_inc_disregard_budgeted_const	= 10
const avail_earned_inc_budgeted_const		= 11
const allocation_budgeted_const				= 12
const child_support_cost_budgeted_const		= 13
const counted_earned_inc_budgeted_const		= 14
const unearned_inc_budgeted_const			= 15
const allocation_bal_budgeted_const			= 16
const child_support_cost_bal_budgeted_const	= 17
const counted_unearned_inc_budgeted_const	= 18
const earned_inc_correct_const				= 19
const earned_inc_disregard_correct_const	= 20
const avail_earned_inc_correct_const		= 21
const allocation_correct_const				= 22
const child_support_cost_correct_const		= 23
const counted_earned_inc_correct_const		= 24
const unearned_inc_correct_const			= 25
const allocation_bal_correct_const			= 26
const child_support_cost_bal_correct_const	= 27
const counted_unearned_inc_correct_const	= 28
const last_const 							= 40

Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(last_const, 0)

Const end_of_doc = 6

function ensure_variable_is_a_number(variable_here, decimal_places)
	If variable_here = "" Then variable_here = 0
	If IsNumeric(variable_here) = False Then variable_here = 0
	variable_here = FormatNumber(variable_here, decimal_places, -1, 0, 0)
	variable_here = variable_here *1
end function

function budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, output_type)
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(earned_income_correct_amt, 2)
	Call ensure_variable_is_a_number(unearned_correct_amt, 2)

	'TODO - To make this global we need to handle for if the earned income deduction is allowed or not
	total_income_correct_amt = earned_income_correct_amt + unearned_correct_amt
	earned_deduction_correct_amt = 0.2 * earned_income_correct_amt

	total_income_correct_amt = FormatNumber(total_income_correct_amt, 2, -1, 0, -1)
	earned_deduction_correct_amt = FormatNumber(earned_deduction_correct_amt, 2, -1, 0, -1)

	If UCase(output_type) = "STRING" Then
		earned_income_correct_amt = earned_income_correct_amt & ""
		unearned_correct_amt = unearned_correct_amt & ""
		total_income_correct_amt = total_income_correct_amt & ""
		earned_deduction_correct_amt = earned_deduction_correct_amt & ""
	End If
end function

function budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, output_type)
	' disa_household - True/False
	' cat_elig - True/False
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(correct_hh_size, 0)

	'TODO - To make this global we need a variable to get the right month for the trifty food plan
	'TODO - To make this global we need to handle for proration
	'THRIFTY FOOD PLAN - CM 22.12.01
	If correct_hh_size = 0 Then
		max_snap_benefit = 0
		standard_deduction_correct_amt = 0
		max_net_adj_income_correct_amt = 0
		max_gross_income_correct_amt = 0
	Elseif correct_hh_size = 1 Then
		max_snap_benefit = 250
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1074
		If cat_elig = True Then max_gross_income_correct_amt = 1771
		If cat_elig = False Then max_gross_income_correct_amt = 1396
	ElseIf correct_hh_size = 2 Then
		max_snap_benefit = 459
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1452
		If cat_elig = True Then max_gross_income_correct_amt = 2396
		If cat_elig = False Then max_gross_income_correct_amt = 1888
	ElseIf correct_hh_size = 3 Then
		max_snap_benefit = 658
		standard_deduction_correct_amt = 177
		max_net_adj_income_correct_amt = 1830
		If cat_elig = True Then max_gross_income_correct_amt = 3020
		If cat_elig = False Then max_gross_income_correct_amt = 2379
	ElseIf correct_hh_size = 4 Then
		max_snap_benefit = 835
		standard_deduction_correct_amt = 184
		max_net_adj_income_correct_amt = 2209
		If cat_elig = True Then max_gross_income_correct_amt = 3644
		If cat_elig = False Then max_gross_income_correct_amt = 2871
	ElseIf correct_hh_size = 5 Then
		max_snap_benefit = 992
		standard_deduction_correct_amt = 215
		max_net_adj_income_correct_amt = 2587
		If cat_elig = True Then max_gross_income_correct_amt = 4268
		If cat_elig = False Then max_gross_income_correct_amt = 3363
	ElseIf correct_hh_size = 6 Then
		max_snap_benefit = 1190
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 2965
		If cat_elig = True Then max_gross_income_correct_amt = 4893
		If cat_elig = False Then max_gross_income_correct_amt = 3855
	ElseIf correct_hh_size = 7 Then
		max_snap_benefit = 1316
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 3344
		If cat_elig = True Then max_gross_income_correct_amt = 5517
		If cat_elig = False Then max_gross_income_correct_amt = 4347
	ElseIf correct_hh_size = 8 Then
		max_snap_benefit = 1504
		standard_deduction_correct_amt = 246
		max_net_adj_income_correct_amt = 3722
		If cat_elig = True Then max_gross_income_correct_amt = 6141
		If cat_elig = False Then max_gross_income_correct_amt = 4839
	Else
		max_snap_benefit = 1504 + ((correct_hh_size-8) * 188)
		max_net_adj_income_correct_amt = 3722 + ((correct_hh_size-8) * 379)
		standard_deduction_correct_amt = 246
		If cat_elig = True Then max_gross_income_correct_amt = 6141 + ((correct_hh_size-8) * 625)
		If cat_elig = False Then max_gross_income_correct_amt = 4839 + ((correct_hh_size-8) * 492)
	End If

	If disa_household = True Then max_shelter_cost_correct_amt = 0
	If disa_household = False Then max_shelter_cost_correct_amt = 597


	standard_deduction_correct_amt = FormatNumber(standard_deduction_correct_amt, 2, -1, 0, -1)
	max_shelter_cost_correct_amt = FormatNumber(max_shelter_cost_correct_amt, 2, -1, 0, -1)
	max_gross_income_correct_amt = FormatNumber(max_gross_income_correct_amt, 2, -1, 0, -1)
	max_net_adj_income_correct_amt = FormatNumber(max_net_adj_income_correct_amt, 2, -1, 0, -1)
	max_snap_benefit = FormatNumber(max_snap_benefit, 2, -1, 0, -1)

	If UCase(output_type) = "STRING" Then
		correct_hh_size = correct_hh_size & ""
		standard_deduction_correct_amt = standard_deduction_correct_amt & ""
		max_shelter_cost_correct_amt = max_shelter_cost_correct_amt & ""
		max_gross_income_correct_amt = max_gross_income_correct_amt & ""
		max_net_adj_income_correct_amt = max_net_adj_income_correct_amt & ""
		max_snap_benefit = max_snap_benefit & ""
	End If
end function

' function budget_calculate_deductions()
' function budget_calculate_deductions(earned_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
function budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
	' output_type - "STRING" or "NUMBER"
	Call ensure_variable_is_a_number(earned_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(medical_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(dependent_care_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(child_support_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(standard_deduction_correct_amt, 2)
	Call ensure_variable_is_a_number(total_income_correct_amt, 2)

	total_deduction_correct_amt = earned_deduction_correct_amt + medical_deduction_correct_amt + dependent_care_deduction_correct_amt + child_support_deduction_correct_amt + standard_deduction_correct_amt
	net_income_correct_amt = total_income_correct_amt - total_deduction_correct_amt
	fifty_perc_net_income_correct_amt = net_income_correct_amt * 0.5
	Call ensure_variable_is_a_number(fifty_perc_net_income_correct_amt, 2)

	If UCase(output_type) = "STRING" Then
		total_deduction_correct_amt = total_deduction_correct_amt & ""
		earned_deduction_correct_amt = earned_deduction_correct_amt & ""
		medical_deduction_correct_amt = medical_deduction_correct_amt & ""
		dependent_care_deduction_correct_amt = dependent_care_deduction_correct_amt & ""
		child_support_deduction_correct_amt = child_support_deduction_correct_amt & ""
		standard_deduction_correct_amt = standard_deduction_correct_amt & ""

		net_income_correct_amt = net_income_correct_amt & ""
		fifty_perc_net_income_correct_amt = fifty_perc_net_income_correct_amt & ""
	End If
end function

function budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, output_type)
	Call ensure_variable_is_a_number(rent_mortgage_correct_amt, 2)
	Call ensure_variable_is_a_number(tax_correct_amt, 2)
	Call ensure_variable_is_a_number(insurance_correct_amt, 2)
	Call ensure_variable_is_a_number(other_cost_correct_amt, 2)
	Call ensure_variable_is_a_number(utilities_correct_amt, 2)
	Call ensure_variable_is_a_number(fifty_perc_net_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_shelter_cost_correct_amt, 2)
	Call ensure_variable_is_a_number(net_income_correct_amt, 2)

	total_shelter_cost_correct_amt = rent_mortgage_correct_amt + tax_correct_amt + insurance_correct_amt + other_cost_correct_amt + utilities_correct_amt
	adj_shelter_cost_correct_amt = total_shelter_cost_correct_amt - fifty_perc_net_income_correct_amt
	If adj_shelter_cost_correct_amt < 0 Then adj_shelter_cost_correct_amt = 0
	If max_shelter_cost_correct_amt = 0 Then max_shelter_cost_correct_amt = adj_shelter_cost_correct_amt
	If adj_shelter_cost_correct_amt > max_shelter_cost_correct_amt Then
		counted_shelter_cost_correct_amt = max_shelter_cost_correct_amt
	Else
		counted_shelter_cost_correct_amt = adj_shelter_cost_correct_amt
	End If
	net_adj_income_correct_amt = net_income_correct_amt - counted_shelter_cost_correct_amt
	If net_adj_income_correct_amt < 0 Then net_adj_income_correct_amt = 0

	If UCase(output_type) = "STRING" Then
		rent_mortgage_correct_amt = rent_mortgage_correct_amt & ""
		tax_correct_amt = tax_correct_amt & ""
		insurance_correct_amt = insurance_correct_amt & ""
		other_cost_correct_amt = other_cost_correct_amt & ""
		utilities_correct_amt = utilities_correct_amt & ""
		total_shelter_cost_correct_amt = total_shelter_cost_correct_amt & ""

		fifty_perc_net_income_correct_amt = fifty_perc_net_income_correct_amt & ""
		adj_shelter_cost_correct_amt = adj_shelter_cost_correct_amt & ""
		counted_shelter_cost_correct_amt = counted_shelter_cost_correct_amt & ""
		max_shelter_cost_correct_amt = max_shelter_cost_correct_amt & ""
		net_income_correct_amt = net_income_correct_amt & ""
		net_adj_income_correct_amt = net_adj_income_correct_amt & ""
	End If
end function

function budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, sanction_recoupment_correct_amt, snap_correct_amt, snap_issued_amt, snap_overpayment_exists, snap_supplement_exists, snap_proration_date, snap_overpayment_amt, snap_supplement_amt, output_type)
	' cat_elig - True/Fals

	Call ensure_variable_is_a_number(total_income_correct_amt, 2)
	Call ensure_variable_is_a_number(net_adj_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_net_adj_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_gross_income_correct_amt, 2)
	Call ensure_variable_is_a_number(max_snap_benefit, 2)
	Call ensure_variable_is_a_number(snap_issued_amt, 2)
	Call ensure_variable_is_a_number(sanction_recoupment_correct_amt, 2)
	If IsDate(snap_proration_date) = False Then
		snap_proration_date = #2/1/2022#
	Else
		snap_proration_date = DateAdd("d", 0, snap_proration_date)
	End If

	proration_day = DatePart("d", snap_proration_date)
	proration_percentage = 1.00
	If proration_day = 2 Then proration_percentage = .9643
	If proration_day = 3 Then proration_percentage = .9286
	If proration_day = 4 Then proration_percentage = .8929
	If proration_day = 5 Then proration_percentage = .8571
	If proration_day = 6 Then proration_percentage = .8214
	If proration_day = 7 Then proration_percentage = .7857
	If proration_day = 8 Then proration_percentage = .7500
	If proration_day = 9 Then proration_percentage = .7143
	If proration_day = 10 Then proration_percentage = .6786
	If proration_day = 11 Then proration_percentage = .6429
	If proration_day = 12 Then proration_percentage = .6071
	If proration_day = 13 Then proration_percentage = .5714
	If proration_day = 14 Then proration_percentage = .5357
	If proration_day = 15 Then proration_percentage = .5000
	If proration_day = 16 Then proration_percentage = .4643
	If proration_day = 17 Then proration_percentage = .4286
	If proration_day = 18 Then proration_percentage = .3929
	If proration_day = 19 Then proration_percentage = .3571
	If proration_day = 20 Then proration_percentage = .3214
	If proration_day = 21 Then proration_percentage = .2857
	If proration_day = 22 Then proration_percentage = .2500
	If proration_day = 23 Then proration_percentage = .2143
	If proration_day = 24 Then proration_percentage = .1786
	If proration_day = 25 Then proration_percentage = .1429
	If proration_day = 26 Then proration_percentage = .1071
	If proration_day = 27 Then proration_percentage = .0714
	If proration_day = 28 Then proration_percentage = .0357
	snap_proration_date = snap_proration_date & ""

	snap_overpayment_exists = False
	snap_supplement_exists = False
	mfip_overpayment_exists = False
	mfip_supplement_exists = False
	income_exceeded = False
	snap_overpayment_amt = 0
	snap_supplement_amt = 0

	If cat_elig = True Then
		If total_income_correct_amt > max_gross_income_correct_amt Then income_exceeded = True
	Else
		If net_adj_income_correct_amt > max_net_adj_income_correct_amt Then income_exceeded = True
	End If

	If income_exceeded = False Then
		thirty_perc_of_net_income = 0.3 * net_adj_income_correct_amt
		monthly_snap_benefit_correct_amt = max_snap_benefit - thirty_perc_of_net_income
		monthly_snap_benefit_correct_amt = Int(monthly_snap_benefit_correct_amt)

		monthly_snap_benefit_correct_amt = monthly_snap_benefit_correct_amt * proration_percentage
		monthly_snap_benefit_correct_amt = Int(monthly_snap_benefit_correct_amt)

		Call ensure_variable_is_a_number(monthly_snap_benefit_correct_amt, 2)
		monthly_snap_benefit_correct_amt = monthly_snap_benefit_correct_amt * 1
		sanction_recoupment_correct_amt = sanction_recoupment_correct_amt * 1
		snap_correct_amt = monthly_snap_benefit_correct_amt - sanction_recoupment_correct_amt
	End If
	If monthly_snap_benefit_correct_amt > snap_issued_amt Then
		snap_supplement_exists = True
		snap_supplement_amt = monthly_snap_benefit_correct_amt - snap_issued_amt
	End If
	If monthly_snap_benefit_correct_amt < snap_issued_amt Then
		snap_overpayment_exists = True
		snap_overpayment_amt = snap_issued_amt - monthly_snap_benefit_correct_amt
	End If

	If UCase(output_type) = "STRING" Then
		total_income_correct_amt = total_income_correct_amt & ""
		net_adj_income_correct_amt = net_adj_income_correct_amt & ""
		max_net_adj_income_correct_amt = max_net_adj_income_correct_amt & ""
		max_gross_income_correct_amt = max_gross_income_correct_amt & ""
		max_snap_benefit = max_snap_benefit & ""
		snap_issued_amt = snap_issued_amt & ""
		rent_mortgage_correct_amt = rent_mortgage_correct_amt & ""
		monthly_snap_benefit_correct_amt = monthly_snap_benefit_correct_amt & ""
		snap_overpayment_amt = snap_overpayment_amt & ""
		snap_supplement_amt = snap_supplement_amt & ""
	End if

end function

' function budget_calculate_mfip(ARRAY_NAME, , output_type)
'
' end function
'
' function budget_calculate_mfip(, output_type)
'
' end function
function read_amount_from_MAXIS(variable_here, length, row, col)
	EMReadScreen variable_here, length, row, col
	variable_here = trim(variable_here)
	If variable_here = "" Then variable_here = 0
	If IsNumeric(variable_here) = False Then variable_here = 0
	variable_here = FormatNumber(variable_here, decimal_places, -1, 0, 0)
	' variable_here = variable_here *1
end function

'Connecting to MAXIS
EMConnect ""
'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = "02"
MAXIS_footer_year = "22"

pdf_doc_path = ""

calc_btn = 500
snap_claculation_done_btn = 501
assignment_instructions_btn = 505
script_instructions_btn = 506

cat_elig = True
disa_household = False

excel_details_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Tracking Details - Testing.xlsx"
excel_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Overpayments Report - Testing.xlsx"


If user_ID_for_validation = "WFM207" Then
	user_name = "Mandora Young"
	excel_details_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Tracking Details - Mandora.xlsx"
	excel_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Overpayments Report - Mandora.xlsx"
End If
If user_ID_for_validation = "YEYA001" Then
	user_name = "Yeng Yang"
	excel_details_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Tracking Details - Yeng.xlsx"
	excel_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Overpayments Report - Yeng.xlsx"
End If

'Grabbing the case number
call MAXIS_case_number_finder(MAXIS_case_number)
back_to_self 'to ensure we are not in edit mode'
EMWriteScreen MAXIS_case_number, 18, 43

'case number dialog
Do
	err_msg = ""
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 166, 140, "Case Number Dialog"
	  EditBox 90, 10, 70, 15, MAXIS_case_number
	  ButtonGroup ButtonPressed
	    OkButton 55, 120, 50, 15
	    CancelButton 110, 120, 50, 15
	    PushButton 10, 80, 110, 15, "Assignment Instructions", assignment_instructions_btn
	    PushButton 10, 95, 110, 15, "Script Instructions", script_instructions_btn
	  Text 10, 15, 80, 10, "Enter the Case Number:"
	  Text 10, 35, 150, 45, "This script is specific to the detailed review of the cases impacted by the Autoclose Pause that happened in 02/22 and does not take any MAXIS action or create CASE/NOTEs as this process is handled external from MAXIS."
	EndDialog

	dialog Dialog1
	cancel_without_confirmation

	Call validate_MAXIS_case_number(err_msg, "*")

	If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbCr & err_msg

	If ButtonPressed = assignment_instructions_btn or ButtonPressed = script_instructions_btn Then
		err_msg = "LOOP"
		If ButtonPressed = assignment_instructions_btn Then Call word_doc_open(t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Tier Two Auto-Close Review Instructions.docx", objWord, objDoc)
		If ButtonPressed = script_instructions_btn Then Call word_doc_open(t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Script - ADMIN - Track Autoclose Overpayments Instructions.docx", objWord, objDoc)
	End If
 Loop until err_msg = ""

'read for MFIP/SNAP in 02/22
SNAP_active = False
MFIP_active = False
Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen cash_1_prog, 2, 6, 67
EMReadScreen cash_1_stat, 4, 6, 74
EMReadScreen cash_2_prog, 2, 7, 67
EMReadScreen cash_2_stat, 4, 7, 74
EMReadScreen snap_prog, 2, 10, 67
EMReadScreen snap_stat, 4, 10, 74

If cash_1_stat = "ACTV" and cash_1_prog = "MF" Then MFIP_active = True
If cash_2_stat = "ACTV" and cash_2_prog = "MF" Then MFIP_active = True
If snap_stat = "ACTV" Then SNAP_active = True
call back_to_self

If MFIP_active = True Then Call script_end_procedure("MFIP was active in 02/22. MFIP cases are not able to be handled at this time.")
If SNAP_active = False Then Call script_end_procedure("This case does not appear to have been active SNAP in 02/22 and thes script cannot continue.")
If SNAP_active = False and MFIP_active = False Then script_end_procedure("This case does not appear to have been active SNAP or MFIP in 02/22 and thes script cannot continue.")

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
	If access_denied_check = "ACCESS DENIED" Then
		PF10
	End If
	If client_array <> "" Then client_array = client_array & "|" & ref_nbr
	If client_array = "" Then client_array = client_array & ref_nbr
	transmit      'Going to the next MEMB panel
	Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
	member_count = member_count + 1
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
' MsgBox client_array
client_array = split(client_array, "|")
memb_droplist = ""

clt_count = 0

For each hh_clt in client_array
	ReDim Preserve HH_MEMB_ARRAY(last_const, clt_count)
	HH_MEMB_ARRAY(ref_number, clt_count) = hh_clt
	' HH_MEMB_ARRAY(define_the_member, clt_count)

	Call navigate_to_MAXIS_screen("STAT", "MEMB")		'===============================================================================================
	EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
	' MsgBox "1"
	transmit
	' MsgBox "2"

	EMReadscreen HH_MEMB_ARRAY(last_name_const, clt_count), 25, 6, 30
	EMReadscreen HH_MEMB_ARRAY(first_name_const, clt_count), 12, 6, 63
	EMReadScreen HH_MEMB_ARRAY(age_const, clt_count), 3, 8, 76

	HH_MEMB_ARRAY(age_const, clt_count) = trim(HH_MEMB_ARRAY(age_const, clt_count))
	If HH_MEMB_ARRAY(age_const, clt_count) = "" Then HH_MEMB_ARRAY(age_const, clt_count) = 0
	HH_MEMB_ARRAY(age_const, clt_count) = HH_MEMB_ARRAY(age_const, clt_count) * 1
	If HH_MEMB_ARRAY(age_const, clt_count) >=60 Then disa_household = True

	HH_MEMB_ARRAY(last_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(last_name_const, clt_count), "_", ""))
	HH_MEMB_ARRAY(first_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(first_name_const, clt_count), "_", ""))
	HH_MEMB_ARRAY(full_name_const, clt_count) = HH_MEMB_ARRAY(last_name_const, clt_count) & ", " & HH_MEMB_ARRAY(first_name_const, clt_count)

	HH_MEMB_ARRAY(earned_income_exists_const, clt_count) = False
	HH_MEMB_ARRAY(unearned_income_exists_const, clt_count) = False
	HH_MEMB_ARRAY(memb_droplist_const, clt_count) = HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)
	memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)
	' MsgBox HH_MEMB_ARRAY(full_name_const, clt_count)
	If disa_household = False Then
		Call navigate_to_MAXIS_screen("STAT", "DISA")
		EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
		transmit

		EMReadScreen FS_disa_status, 2, 12, 59
		If FS_disa_status = "01" Then disa_household = True
		If FS_disa_status = "02" Then disa_household = True
		If FS_disa_status = "03" Then disa_household = True
		If FS_disa_status = "04" Then disa_household = True
		If FS_disa_status = "08" Then disa_household = True
		If FS_disa_status = "10" Then disa_household = True
		If FS_disa_status = "11" Then disa_household = True
		If FS_disa_status = "12" Then disa_household = True
		If FS_disa_status = "13" Then disa_household = True
		If FS_disa_status = "14" Then disa_household = True
	End If
Next

'Read ELIG for 02/22
If SNAP_active = True Then
	Call navigate_to_MAXIS_screen("ELIG", "FS__")
	EMWriteScreen "99", 19, 78
	transmit
	elig_row = 7
	version_numb = ""
	Do
		EMReadScreen approval_status, 8, elig_row, 50
		If approval_status = "APPROVED" Then
			EMReadScreen version_numb, 2, elig_row, 22
			version_numb = trim(version_numb)
			version_numb = right("00" & version_numb, 2)
			Exit Do
		End If
		elig_row = elig_row + 1
	Loop until approval_status = "        "
	transmit
	EMWriteScreen version_numb, 19, 78
	transmit

	write_value_and_transmit "FSCR", 19, 70

	write_value_and_transmit "FSB1", 19, 70
	Call read_amount_from_MAXIS(earned_income_budgeted_amt, 10, 8, 31)
	Call read_amount_from_MAXIS(unearned_budgeted_amt, 10, 18, 31)
	Call read_amount_from_MAXIS(total_income_budgeted_amt, 10, 7, 71)

	Call read_amount_from_MAXIS(standard_deduction_budgeted_amt, 10, 10, 71)
	Call read_amount_from_MAXIS(earned_deduction_budgeted_amt, 10, 11, 71)
	Call read_amount_from_MAXIS(medical_deduction_budgeted_amt, 10, 12, 71)
	Call read_amount_from_MAXIS(dependent_care_deduction_budgeted_amt, 10, 13, 71)
	Call read_amount_from_MAXIS(child_support_deduction_budgeted_amt, 10, 14, 71)
	Call read_amount_from_MAXIS(total_deduction_budgeted_amt, 10, 16, 71)

	Call read_amount_from_MAXIS(net_income_budgeted_amt, 10, 18, 71)


	write_value_and_transmit "FSB2", 19, 70

	Call read_amount_from_MAXIS(rent_mortgage_budgeted_amt, 10, 5, 27)
	Call read_amount_from_MAXIS(tax_budgeted_amt, 10, 6, 27)
	Call read_amount_from_MAXIS(insurance_budgeted_amt, 10, 7, 27)
	Call read_amount_from_MAXIS(other_cost_budgeted_amt, 10, 12, 27)
	rent_mortgage_budgeted_amt = rent_mortgage_budgeted_amt* 1
	tax_budgeted_amt = tax_budgeted_amt* 1
	insurance_budgeted_amt = insurance_budgeted_amt* 1
	other_cost_budgeted_amt = other_cost_budgeted_amt* 1
	total_housing_cost_budgeted_amt = rent_mortgage_budgeted_amt + tax_budgeted_amt + insurance_budgeted_amt + other_cost_budgeted_amt

	Call read_amount_from_MAXIS(net_adj_income_budgeted_amt, 10, 7, 27)
	Call read_amount_from_MAXIS(electricity, 10, 8, 27)
	Call read_amount_from_MAXIS(heat_air, 10, 9, 27)
	Call read_amount_from_MAXIS(phone, 10, 11, 27)
	electricity = electricity*1
	heat_air = heat_air*1
	phone = phone*1
	utilities_budgeted_amt = heat_air + electricity + phone
	utilities_budgeted_amt = utilities_budgeted_amt & ""
	Call read_amount_from_MAXIS(total_shelter_cost_budgeted_amt, 10, 14, 27)

	Call read_amount_from_MAXIS(snap_issued_amt, 10, 10, 71)
	Call read_amount_from_MAXIS(state_benefit_amt, 10, 17, 71)
	Call read_amount_from_MAXIS(fed_benefit_amt, 10, 18, 71)
	snap_issued_amt = snap_issued_amt*1
	fed_benefit_amt = fed_benefit_amt*1
	state_benefit_amt = state_benefit_amt*1
	If fed_benefit_amt = 0 AND state_benefit_amt = 0 Then fed_benefit_amt = snap_issued_amt

	fed_percent = fed_benefit_amt/snap_issued_amt
	state_percent = state_benefit_amt/snap_issued_amt
	' MsgBox "State Percent - " & fed_percent & vbCr & "Fed Percent - " & state_percent

	write_value_and_transmit "FSSM", 20, 70

	EMReadScreen budgeted_hh_size, 2, 13, 31
	budgeted_hh_size = trim(budgeted_hh_size)

	earned_income_correct_amt = earned_income_budgeted_amt
	unearned_correct_amt = unearned_budgeted_amt
	correct_hh_size = budgeted_hh_size
	medical_deduction_correct_amt = medical_deduction_budgeted_amt
	dependent_care_deduction_correct_amt = dependent_care_deduction_budgeted_amt
	child_support_deduction_correct_amt = child_support_deduction_budgeted_amt
	rent_mortgage_correct_amt = rent_mortgage_budgeted_amt
	tax_correct_amt = tax_budgeted_amt
	insurance_correct_amt = insurance_budgeted_amt
	other_cost_correct_amt = other_cost_budgeted_amt
	utilities_correct_amt = utilities_budgeted_amt



	' 978321
	'
	' snap_issued_amt = 1316
	call back_to_self
End If

If MFIP_active = True Then
	Call navigate_to_MAXIS_screen("ELIG", "MFIP")
	EMWriteScreen "99", 20, 79
	transmit
	elig_row = 7
	version_numb = ""
	Do
		EMReadScreen approval_status, 8, elig_row, 50
		If approval_status = "APPROVED" Then
			EMReadScreen version_numb, 2, elig_row, 22
			version_numb = trim(version_numb)
			version_numb = right("00" & version_numb, 2)
			Exit Do
		End If
		elig_row = elig_row + 1
	Loop until approval_status = "        "
	transmit
	EMWriteScreen version_numb, 20, 79
	transmit

	For hh_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		elig_row = 7
		Do
			EMReadScreen elig_ref_numb, 2, elig_row, 6
			If elig_ref_numb = HH_MEMB_ARRAY(ref_number, hh_memb) Then
				EmReadScreen memb_code, 1, elig_row, 36
				HH_MEMB_ARRAY(mfip_elig, hh_memb) = False
				If memb_code = "A" Then HH_MEMB_ARRAY(mfip_elig, hh_memb) = True
			End If
			elig_row = elig_row + 1
		Loop until elig_ref_numb = "  "
	Next

	EMWriteScreen "MFB1", 20, 71
	transmit

	For hh_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		EMWriteScreen "X", 6, 3
		transmit
		Do
			EMReadScreen elig_person, 40, 8, 28
			elig_person = trim(elig_person)
			MsgBox "ARRAY NAME - " & HH_MEMB_ARRAY(full_name_const, hh_memb) & vbCr & "Elig name - " & elig_person & vbCr & "EARNED"
			If HH_MEMB_ARRAY(full_name_const, hh_memb) = elig_person Then
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(earned_inc_budgeted_const, hh_memb), 9, 13, 54)
				Call read_amount_from_MAXIS(disregard, 9, 16, 54)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(avail_earned_inc_budgeted_const, hh_memb), 9, 17, 54)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(allocation_budgeted_const, hh_memb), 9, 18, 54)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(child_support_cost_budgeted_const, hh_memb), 9, 19, 54)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(counted_earned_inc_budgeted_const, hh_memb), 9, 20, 54)
				If HH_MEMB_ARRAY(counted_earned_inc_budgeted_const, hh_memb) <> "0.00" Then HH_MEMB_ARRAY(earned_income_exists_const, hh_memb) = True

				disregard = disregard * 1
				HH_MEMB_ARRAY(earned_inc_disregard_budgeted_const, hh_memb) = disregard + 65

				HH_MEMB_ARRAY(earned_inc_correct_const, hh_memb) = HH_MEMB_ARRAY(earned_inc_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(earned_inc_disregard_correct_const, hh_memb) = HH_MEMB_ARRAY(earned_inc_disregard_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(avail_earned_inc_correct_const, hh_memb) = HH_MEMB_ARRAY(avail_earned_inc_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(allocation_correct_const, hh_memb) = HH_MEMB_ARRAY(allocation_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(child_support_cost_correct_const, hh_memb) = HH_MEMB_ARRAY(child_support_cost_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(counted_earned_inc_correct_const, hh_memb) = HH_MEMB_ARRAY(counted_earned_inc_budgeted_const, hh_memb)
				MsgBox "PERSON - " & HH_MEMB_ARRAY(full_name_const, hh_memb) & vbCr & HH_MEMB_ARRAY(counted_earned_inc_budgeted_const, hh_memb)
				PF3
				Exit Do
			End If
			transmit
			EMReadScreen read_all_persons, 12, 5, 32
		Loop until read_all_persons <> "Maxis Person"

		EMWriteScreen "X", 11, 3
		transmit
		Do
			EMReadScreen elig_person, 29, 8, 34
			elig_person = trim(elig_person)
			MsgBox "ARRAY NAME - " & HH_MEMB_ARRAY(full_name_const, hh_memb) & vbCr & "Elig name - " & elig_person & vbCr & "UNEA"

			If HH_MEMB_ARRAY(full_name_const, hh_memb) = elig_person Then
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(unearned_inc_budgeted_const, hh_memb), 9, 11, 49)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(allocation_bal_budgeted_const, hh_memb), 9, 12, 49)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(child_support_cost_bal_budgeted_const, hh_memb), 9, 13, 49)
				Call read_amount_from_MAXIS(HH_MEMB_ARRAY(counted_unearned_inc_budgeted_const, hh_memb), 9, 14, 49)
				If HH_MEMB_ARRAY(counted_unearned_inc_budgeted_const, hh_memb) <> "0.00" Then HH_MEMB_ARRAY(unearned_income_exists_const, hh_memb) = True

				HH_MEMB_ARRAY(unearned_inc_correct_const, hh_memb) = HH_MEMB_ARRAY(unearned_inc_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(allocation_bal_correct_const, hh_memb) = HH_MEMB_ARRAY(allocation_bal_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(child_support_cost_bal_correct_const, hh_memb) = HH_MEMB_ARRAY(child_support_cost_bal_budgeted_const, hh_memb)
				HH_MEMB_ARRAY(counted_unearned_inc_correct_const, hh_memb) = HH_MEMB_ARRAY(counted_unearned_inc_budgeted_const, hh_memb)
				MsgBox "PERSON - " & HH_MEMB_ARRAY(full_name_const, hh_memb) & vbCr & HH_MEMB_ARRAY(counted_unearned_inc_budgeted_const, hh_memb)
				PF3
				Exit Do
			End If
			transmit
			EMReadScreen read_all_persons, 15, 6, 34
		Loop until read_all_persons <> "Unearned Income"

		call read_amount_from_MAXIS(deemed_inc_budgeted_amt, 10, 12, 32)
		call read_amount_from_MAXIS(cses_exclusion_budgeted_amt, 10, 13, 32)


	Next

	Call read_amount_from_MAXIS(mfip_total_issued_amt, 10, 9, 71)

	write_value_and_transmit "MFB2", 19, 70
	Call read_amount_from_MAXIS(mfip_MF_MF_issued_amt, 10, 12, 32)

	Call read_amount_from_MAXIS(mfip_MF_FS_issued_amt, 10, 7, 32)
	Call read_amount_from_MAXIS(mfip_MF_FS_S_issued_amt, 10, 15, 45)
	mfip_MF_FS_issued_amt = mfip_MF_FS_issued_amt*1
	mfip_MF_FS_S_issued_amt = mfip_MF_FS_S_issued_amt*1
	mfip_MF_FS_F_issued_amt = mfip_MF_FS_issued_amt - mfip_MF_FS_S_issued_amt

	Call read_amount_from_MAXIS(mfip_MF_HG_issued_amt, 10, 17, 32)

	write_value_and_transmit "MFSM", 19, 70

	EMReadScreen mfip_budgeted_caregivers, 3, 7, 73
	EMReadScreen mfip_budgeted_children, 3, 8, 73
	mfip_budgeted_caregivers = trim(mfip_budgeted_caregivers)
	mfip_budgeted_children = trim(mfip_budgeted_children)
	' MsgBox "caregivers: " & mfip_budgeted_caregivers & vbCr & "children: " & mfip_budgeted_children

	call back_to_self

	selected_memb = 0
End If

' START A LOOP HERE
recalculation_confirmed = False
snap_overpayment_exists = False
snap_supplement_exists = False
calculation_needed = True
snap_proration_date = "2/1/2022"
Do
	'Determine what happened with the review/mont process by dialog
	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 316, 105, "02/22 Report Process Information"
		  DropListBox 180, 10, 60, 45, "Select One..."+chr(9)+"ER"+chr(9)+"SR"+chr(9)+"HRF", feb_process
		  DropListBox 260, 25, 50, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", process_complete
		  DropListBox 65, 45, 90, 45, "Select One..."+chr(9)+"None Received"+chr(9)+"CAF"+chr(9)+"HRF"+chr(9)+"HUF"+chr(9)+"MNbenefits"+chr(9)+"CSR"+chr(9)+"Combined AR", form_received
		  EditBox 260, 45, 50, 15, form_received_date
		  DropListBox 65, 65, 90, 45, "Select One..."+chr(9)+"Not Required"+chr(9)+"Completed"+chr(9)+"Incomplete"+chr(9)+"N/A", interview_information
		  EditBox 260, 65, 50, 15, interview_date
		  DropListBox 65, 85, 60, 45, "Select One..."+chr(9)+"None Needed"+chr(9)+"Partial"+chr(9)+"Complete"+chr(9)+"None Received"+chr(9)+"N/A", verifs_received
		  ButtonGroup ButtonPressed
		    OkButton 205, 85, 50, 15
		    CancelButton 260, 85, 50, 15
		  Text 15, 15, 50, 10, "Case Number:"
		  Text 70, 15, 50, 10, MAXIS_case_number
		  Text 125, 15, 55, 10, "02/22 Process:"
		  Text 15, 30, 245, 10, "Was the MONT/REVW completed, with all required forms and verifications?"
		  Text 10, 50, 55, 10, "Form Received:"
		  Text 185, 50, 70, 10, "Date Form Recieved:"
		  Text 30, 70, 35, 10, "Interview:"
		  Text 205, 70, 50, 10, "Interview Date:"
		  Text 20, 90, 45, 10, "Verifications:"
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If form_received = "None Received" Then
			interview_information = "N/A"
			verifs_received = "N/A"
			err_msg = "LOOP"
		End If

		If feb_process = "Select One..." Then err_msg = err_msg & vbCr & "* Select the process that was due for 02/22."
		If process_complete = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the process was completed and case would have been able to be processedd and 'APP'd with the everything on file."
		If form_received = "Select One..." Then err_msg = err_msg & vbCr & "* Select which form was submitted or indicate that no form was received."
		If form_received <> "Select One..." and form_received <> "None Received" Then
			If IsDate(form_received_date) = False Then err_msg = err_msg & vbCr & "* Since a form was received, enter a valid date for the date the form was received."
			If interview_information = "N/A" Then err_msg = err_msg & vbCr & "* Interview cannot be 'N/A' if the form was received, identify if the interview was complete, incomplete, or not reqquired."
			If verifs_received = "N/A" Then err_msg = err_msg & vbCr & "* Verifications cannot be 'N/A' if the form was received, identify if verifications were complete, partial, none received, or not needed."
		End If
		If interview_information = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate hwat happened with the interview process."
		If interview_information = "Completed" and IsDate(interview_date) = False Then  err_msg = err_msg & vbCr & "* Since the interview was completed, enter a valid date for the date the interview was completed."
		If verifs_received = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the status of the verifications for this case in the 02/22 report process."
		If process_complete = "Yes" and form_received = "None Received" Then err_msg = err_msg & vbCr & "* If the process is complete, The form received should not be 'None Received' - enter the form name."
		If process_complete = "Yes" and interview_information = "Incomplete" Then err_msg = err_msg & vbCr & "* If the process is complete, the interview should not be listed as 'Incomplete' - it should either be 'Not Required' or 'Completed'."
		If process_complete = "Yes" and verifs_received = "Partial" Then err_msg = err_msg & vbCr & "* If the process is complete, verifications received should not be 'Partial' - they should either be 'Complete' or 'None Needed'."

		If err_msg <> "" and left(err_msg, 4) <> "LOOP" then MsgBox "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""

	If IsDate(form_received_date) = True Then snap_proration_date = form_received_date
	Call budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, "STRING")
	Call budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, "STRING")
	Call budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, "STRING")
	Call budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, "STRING")
	Call budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, sanction_recoupment_correct_amt, snap_correct_amt, snap_issued_amt, snap_overpayment_exists, snap_supplement_exists, snap_proration_date, snap_overpayment_amt, snap_supplement_amt, "STRING")

	If process_complete = "No" Then calculation_needed = False

	'dialog for OP calculation
	If MFIP_active = True Then
		If calculation_needed = True Then
			income_selection_person = HH_MEMB_ARRAY(memb_droplist_const, 0)
			Do
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 556, 385, "02/22 MFIP Incorrect Payment Calculation"
				  'ISSUANCE
				  GroupBox 10, 5, 200, 35, "Benefit Issued for 02/22"
				  Text 15, 15, 30, 10, "MF-Cash:"
				  Text 50, 15, 30, 10, "$ " & mfip_MF_MF_issued_amt
				  Text 80, 15, 30, 10, "MF-Food:"
				  Text 110, 15, 30, 10, "$ " & mfip_MF_FS_issued_amt
				  Text 145, 15, 25, 10, "MF-HG:"
				  Text 170, 15, 30, 10, "$ " & mfip_MF_HG_issued_amt
				  Text 25, 25, 25, 10, "SNAP:"
				  Text 50, 25, 30, 10, "$ " & snap_issued_amt
				  'Earned Income
				  GroupBox 10, 40, 185, 245, "Person Based Income Calculation"
				  Text 30, 60, 30, 10, "Person:"
				  DropListBox 65, 55, 125, 45, memb_droplist, income_selection_person
				  Text 30, 80, 55, 10, "Earned Income:"
				  Text 90, 80, 30, 10, "$ " & HH_MEMB_ARRAY(earned_inc_budgeted_const, selected_memb)
				  EditBox 140, 75, 50, 15, HH_MEMB_ARRAY(earned_inc_correct_const, selected_memb)
				  Text 35, 100, 45, 10, "EI Disregard:"
				  Text 90, 100, 30, 10, "$ " & HH_MEMB_ARRAY(earned_inc_disregard_budgeted_const, selected_memb)
				  Text 140, 100, 30, 10, "$ " & HH_MEMB_ARRAY(earned_inc_disregard_correct_const, selected_memb)
				  Text 15, 115, 70, 10, "Available Earned Inc:"
				  Text 90, 115, 30, 10, "$ " & HH_MEMB_ARRAY(avail_earned_inc_budgeted_const, selected_memb)
				  Text 140, 115, 30, 10, "$ " & HH_MEMB_ARRAY(avail_earned_inc_correct_const, selected_memb)
				  Text 45, 135, 40, 10, "Allocation:"
				  Text 90, 135, 30, 10, "$ " & HH_MEMB_ARRAY(allocation_budgeted_const, selected_memb)
				  EditBox 140, 130, 50, 15, HH_MEMB_ARRAY(allocation_correct_const, selected_memb)
				  Text 20, 155, 65, 10, "Child Support Cost:"
				  Text 90, 155, 30, 10, "$ " & HH_MEMB_ARRAY(child_support_cost_budgeted_const, selected_memb)
				  EditBox 140, 150, 50, 15, HH_MEMB_ARRAY(child_support_cost_correct_const, selected_memb)
				  GroupBox 10, 250, 185, 15, ""
				  Text 15, 175, 65, 10, "Counted Earned Inc:"
				  Text 90, 175, 30, 10, "$ " & HH_MEMB_ARRAY(counted_earned_inc_budgeted_const, selected_memb)
				  Text 140, 175, 30, 10, "$ " & HH_MEMB_ARRAY(counted_earned_inc_correct_const, selected_memb)
				  Text 25, 200, 60, 10, "Unearned Income:"
				  Text 90, 200, 30, 10, "$ " & HH_MEMB_ARRAY(unearned_inc_budgeted_const, selected_memb)
				  EditBox 140, 195, 50, 15, HH_MEMB_ARRAY(unearned_inc_correct_const, selected_memb)
				  Text 35, 220, 50, 10, "Allocation Bal:"
				  Text 90, 220, 30, 10, "$ " & HH_MEMB_ARRAY(allocation_bal_budgeted_const, selected_memb)
				  Text 140, 220, 30, 10, "$ " & HH_MEMB_ARRAY(allocation_bal_correct_const, selected_memb)
				  Text 20, 235, 60, 10, "Child Support Bal:"
				  Text 90, 235, 30, 10, "$ " & HH_MEMB_ARRAY(child_support_cost_bal_budgeted_const, selected_memb)
				  Text 140, 235, 30, 10, "$ " & HH_MEMB_ARRAY(child_support_cost_bal_correct_const, selected_memb)
				  GroupBox 10, 170, 185, 15, ""
				  Text 30, 255, 55, 10, "Counted UNEA:"
				  Text 90, 255, 30, 10, "$ " & HH_MEMB_ARRAY(counted_unearned_inc_budgeted_const, selected_memb)
				  Text 140, 255, 30, 10, "$ " & HH_MEMB_ARRAY(counted_unearned_inc_correct_const, selected_memb)
				  ButtonGroup ButtonPressed
				    PushButton 160, 270, 30, 10, "CALC", calc_btn
				  'Unearned Income
				  grp_len = 205 + UBound(HH_MEMB_ARRAY, 2)*20
				  GroupBox 200, 40, 135, grp_len, "List of Income"
				  Text 205, 50, 55, 10, "Earned Income:"
				  Text 210, 60, 55, 10, "Total Budgeted:"
				  Text 275, 60, 30, 10, "$ " & earned_income_budgeted_amt
				  y_pos = 75
				  For hh_memb = 0 to UBOUND(HH_MEMB_ARRAY, 2)
					  Text 210, y_pos, 55, 10, "MEMB " & HH_MEMB_ARRAY(ref_number, hh_memb)
					  Text 270, y_pos, 30, 10, "$ " & HH_MEMB_ARRAY(counted_earned_inc_correct_const, hh_memb)
					  y_pos = y_pos + 10
				  Next
				  y_pos = y_pos + 10
				  ' Text 210, 75, 55, 10, "MEMB 01"
				  ' Text 210, 85, 55, 10, "MEMB 01"
				  ' Text 270, 85, 30, 10, "$ " & earned_income_budgeted_amt
				  ' Text 210, 95, 55, 10, "MEMB 01"
				  ' Text 270, 95, 30, 10, "$ " & earned_income_budgeted_amt
				  ' Text 215, 105, 55, 10, "Total Earned"
				  ' Text 275, 105, 30, 10, "$ " & earned_income_budgeted_amt
				  Text 205, y_pos, 65, 10, "Unearned Income:"
				  y_pos = y_pos + 10
				  Text 210, y_pos, 55, 10, "Total Budgeted:"
				  Text 275, y_pos, 30, 10, "$ " & earned_income_budgeted_amt
				  y_pos = y_pos + 10
				  For hh_memb = 0 to UBOUND(HH_MEMB_ARRAY, 2)
					  Text 210, y_pos, 55, 10, "MEMB " & HH_MEMB_ARRAY(ref_number, hh_memb)
					  Text 270, y_pos, 30, 10, "$ " & HH_MEMB_ARRAY(counted_unearned_inc_correct_const, hh_memb)
					  y_pos = y_pos + 10
				  Next
				  y_pos = y_pos + 10

				  ' Text 210, 155, 55, 10, "MEMB 01"
				  ' Text 270, 155, 30, 10, "$ " & earned_income_budgeted_amt
				  ' Text 210, 165, 55, 10, "MEMB 01"
				  ' Text 270, 165, 30, 10, "$ " & earned_income_budgeted_amt
				  ' Text 210, 175, 55, 10, "MEMB 01"
				  ' Text 270, 175, 30, 10, "$ " & earned_income_budgeted_amt
				  ' Text 215, 185, 55, 10, "Total Earned"
				  ' Text 275, 185, 30, 10, "$ " & earned_income_budgeted_amt
				  Text 205, y_pos, 65, 10, "Deemed Income:"
				  y_pos = y_pos + 10
				  Text 210, y_pos, 55, 10, "Total Budgeted:"
				  Text 275, y_pos, 30, 10, "$ " & earned_income_budgeted_amt
				  y_pos = y_pos + 15
				  Text 220, y_pos + 5, 45, 10, "Total Correct:"
				  EditBox 275, y_pos, 50, 15, Edit1
				  ButtonGroup ButtonPressed
				    PushButton 295, 270, 30, 10, "CALC", calc_btn
				  'HH Comp
				  GroupBox 10, 285, 185, 70, "Household Composition"
				  Text 15, 295, 95, 10, "Budgeted Assistance Unit:"
				  Text 25, 310, 40, 10, "Caregivers:"
				  Text 70, 310, 10, 10, mfip_budgeted_caregivers
				  Text 100, 310, 30, 10, "Children:"
				  Text 135, 310, 10, 10, mfip_budgeted_children
				  Text 15, 325, 85, 10, "Correct Assistance Unit:"
				  Text 25, 340, 40, 10, "Caregivers:"
				  EditBox 70, 335, 20, 15, correct_caregiver
				  Text 100, 340, 30, 10, "Children:"
				  EditBox 135, 335, 20, 15, correct_children
				  ButtonGroup ButtonPressed
				    PushButton 160, 340, 30, 10, "CALC", Button3
				  'Budget
				  GroupBox 340, 5, 210, 355, "Corrected Budget"
				  Text 375, 15, 55, 10, "Earned Income:"
				  Text 435, 15, 30, 10, "$ " & hold_var
				  Text 355, 25, 75, 10, "Earned Inc Disregard:"
				  Text 445, 25, 30, 10, "- $ " & hold_var
				  Text 365, 35, 65, 10, "Child Support Ded:"
				  Text 445, 35, 30, 10, "- $ " & hold_var
				  Text 365, 45, 65, 10, "Net Earned Income:"
				  Text 445, 45, 30, 10, "$ " & hold_var
				  Text 360, 60, 65, 10, "Family Wage Level:"
				  Text 435, 60, 30, 10, "$ " & hold_var
				  Text 365, 70, 65, 10, "Net Earned Income:"
				  Text 445, 70, 30, 10, "- $ " & hold_var
				  Text 390, 80, 40, 10, "Difference:"
				  Text 445, 80, 30, 10, "$ " & hold_var
				  Text 355, 90, 75, 10, "Transitional Standard:"
				  Text 435, 90, 30, 10, "$ " & hold_var
				  GroupBox 340, 95, 210, 20, ""
				  Text 355, 105, 120, 10, "Difference or Transitional Standard:"
				  Text 475, 105, 30, 10, "$ " & hold_var
				  Text 370, 120, 60, 10, "Unearned Income:"
				  Text 435, 120, 30, 10, "$ " & hold_var
				  Text 370, 130, 65, 10, "Child Support Ded:"
				  Text 445, 130, 30, 10, "- $ " & hold_var
				  Text 350, 140, 80, 10, "Child Support Exclusion:"
				  Text 445, 140, 30, 10, "- $ " & hold_var
				  Text 375, 150, 55, 10, "Deemed Income:"
				  Text 435, 150, 30, 10, "$ " & hold_var
				  Text 350, 165, 125, 10, "Net Difference Transitional Standard:"
				  Text 475, 165, 30, 10, "$ " & hold_var
				  Text 425, 175, 50, 10, " Cash Portion:"
				  Text 475, 175, 30, 10, "$ " & hold_var
				  Text 425, 185, 45, 10, "Food Portion:"
				  Text 475, 185, 30, 10, "$ " & hold_var
				  Text 380, 200, 50, 10, "Subsidy/Tribal:"
				  Text 445, 200, 30, 10, "- $ " & hold_var
				  Text 375, 210, 60, 10, "Net Cash Portion:"
				  Text 435, 210, 30, 10, "$ " & hold_var
				  Text 355, 220, 80, 10, "Tribal Counted Income:"
				  Text 445, 220, 30, 10, "- $ " & hold_var
				  Text 375, 230, 60, 10, "Net Food Portion:"
				  Text 435, 230, 30, 10, "$ " & hold_var
				  Text 380, 240, 95, 10, "Total Cash and Food Portion:"
				  Text 475, 240, 30, 10, "$ " & hold_var
				  Text 375, 250, 60, 10, "Sanction Amount:"
				  Text 435, 250, 30, 10, "$ " & hold_var
				  Text 370, 265, 65, 10, "Correct MFIP Grant:"
				  Text 435, 265, 30, 10, "$ " & hold_var
				  Text 360, 275, 75, 10, "Correct Housing Grant:"
				  Text 435, 275, 30, 10, "$ " & hold_var
				  Text 400, 290, 75, 10, "MFIP Grant Received:"
				  Text 480, 290, 30, 10, "$ " & mfip_total_issued_amt
				  Text 425, 300, 50, 10, "HG Received:"
				  Text 480, 300, 30, 10, "$ " & mfip_MF_HG_issued_amt
				  If mfip_overpayment_exists = True Then
					  Text 455, 315, 50, 10, "Overpayment:"
					  Text 510, 335, 30, 10, "$ " & mfip_total_overpayment_amt
					  Text 455, 325, 45, 10, "Cash Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_cash_overpayment_amt
					  Text 455, 335, 50, 10, " Food Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_food_overpayment_amt
					  Text 465, 345, 40, 10, "HG Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_hg_overpayment_amt
				  End If
				  If mfip_supplement_exists = True Then
					  Text 455, 315, 50, 10, "Supplement:"
					  Text 510, 335, 30, 10, "$ " & mfip_total_supplement_amt
					  Text 455, 325, 45, 10, "Cash Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_cash_supplement_amt
					  Text 455, 335, 50, 10, " Food Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_food_supplement_amt
					  Text 465, 345, 40, 10, "HG Portion:"
					  Text 510, 335, 30, 10, "$ " & mfip_hg_supplement_amt
				  End If
				  If mfip_overpayment_exists = False And mfip_supplement_exists = False Then
					 Text 400, 315, 100, 10, "02/22 Issuance was Correct"
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 385, 365, 165, 15, "MFIP Budget is Complete", mfip_claculation_done_btn

				EndDialog

				dialog Dialog1
				cancel_confirmation

				' If ButtonPressed = -1 Then ButtonPressed = calc_btn
				output_type = "STRING"
				If ButtonPressed = mfip_claculation_done_btn Then output_type = "NUMBER"


				For hh_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
					If income_selection_person = HH_MEMB_ARRAY(memb_droplist_const, hh_memb) Then selected_memb = hh_memb
				Next
			Loop until ButtonPressed = mfip_claculation_done_btn
		End If
	End If

	If SNAP_active = True Then
		If calculation_needed = True Then
			Do
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 556, 385, "02/22 SNAP Incorrect Payment Calculation"

				  GroupBox 10, 5, 200, 35, "Benefit Issued for 02/22"
				  Text 15, 15, 30, 10, "MF-Cash:"
				  Text 50, 15, 30, 10, "$ " & mfip_MF_MF_issued_amt
				  Text 80, 15, 30, 10, "MF-Food:"
				  Text 110, 15, 30, 10, "$ " & mfip_MF_FS_issued_amt
				  Text 145, 15, 25, 10, "MF-HG:"
				  Text 170, 15, 30, 10, "$ " & mfip_MF_HG_issued_amt
				  Text 25, 25, 25, 10, "SNAP:"
				  Text 50, 25, 30, 10, "$ " & snap_issued_amt

				  GroupBox 10, 40, 200, 75, "Income"
				  Text 90, 40, 35, 10, "Budgeted"
				  Text 140, 40, 35, 10, "Correct"
				  Text 30, 60, 55, 10, "Earned Income:"
				  Text 90, 60, 30, 10, "$ " & earned_income_budgeted_amt
				  EditBox 140, 55, 50, 15, earned_income_correct_amt
				  Text 20, 80, 65, 10, "Unearned Income:"
				  Text 90, 80, 30, 10, "$ " & unearned_budgeted_amt
				  EditBox 140, 75, 50, 15, unearned_correct_amt
				  Text 60, 100, 20, 10, "Total:"
				  Text 90, 100, 30, 10, "$ " & total_income_budgeted_amt
				  Text 140, 100, 30, 10, "$ " & total_income_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 100, 30, 10, "CALC", calc_btn

				  GroupBox 10, 120, 200, 115, "Deductions"
				  Text 90, 120, 35, 10, "Budgeted"
				  Text 140, 120, 35, 10, "Correct"
				  Text 35, 140, 50, 10, "Standard Ded:"
				  Text 90, 140, 30, 10, "$ " & standard_deduction_budgeted_amt
				  Text 140, 140, 30, 15, "$ " & standard_deduction_correct_amt
				  Text 15, 150, 70, 10, "Earned Income Ded:"
				  Text 90, 150, 30, 10, "$ " & earned_deduction_budgeted_amt
				  Text 140, 150, 30, 15, "$ " & earned_deduction_correct_amt
				  Text 40, 165, 50, 10, "Medical Ded:"
				  Text 90, 165, 30, 10, "$ " & medical_deduction_budgeted_amt
				  EditBox 140, 160, 50, 15, medical_deduction_correct_amt
				  Text 15, 185, 70, 10, "Dependent Care Ded:"
				  Text 90, 185, 30, 10, "$ " & dependent_care_deduction_budgeted_amt
				  EditBox 140, 180, 50, 15, dependent_care_deduction_correct_amt
				  Text 37, 205, 50, 10, "Child Support:"
				  Text 90, 205, 30, 10, "$ " & child_support_deduction_budgeted_amt
				  EditBox 140, 200, 50, 15, child_support_deduction_correct_amt
				  Text 35, 220, 20, 10, "Total:"
				  Text 90, 220, 30, 10, "$ " & total_deduction_budgeted_amt
				  Text 140, 220, 30, 10, "$ " & total_deduction_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 220, 30, 10, "CALC", calc_btn

				  'SHELTER '
				  GroupBox 10, 240, 200, 140, "Shelter Costs"
				  Text 90, 240, 35, 10, "Budgeted"
				  Text 140, 240, 35, 10, "Correct"
				  Text 30, 260, 55, 10, "Rent/Mortgage:"
				  Text 90, 260, 30, 10, "$ " & rent_mortgage_budgeted_amt
				  EditBox 140, 255, 50, 15, rent_mortgage_correct_amt
				  Text 35, 280, 45, 10, "Property Tax:"
				  Text 90, 280, 30, 10, "$ " & tax_budgeted_amt
				  EditBox 140, 275, 50, 15, tax_correct_amt
				  Text 25, 300, 60, 10, "Home Insurance:"
				  Text 90, 300, 30, 10, "$ " & insurance_budgeted_amt
				  EditBox 140, 295, 50, 15, insurance_correct_amt
				  Text 15, 320, 20, 10, "Other:"
				  EditBox 40, 315, 45, 15, other_cost_detail
				  Text 90, 320, 30, 10, "$ " & other_cost_budgeted_amt
				  EditBox 140, 315, 50, 15, other_cost_correct_amt
				  Text 55, 340, 30, 10, "Utilities:"
				  Text 90, 340, 30, 10, "$ " & utilities_budgeted_amt
				  DropListBox 140, 335, 50, 15, ""+chr(9)+"488"+chr(9)+"205"+chr(9)+"149"+chr(9)+"56"+chr(9)+"0", utilities_correct_amt
				  Text 60, 360, 20, 10, "Total:"
				  Text 90, 360, 30, 10, "$ " & total_shelter_cost_budgeted_amt
				  Text 140, 360, 30, 10, "$ " & total_shelter_cost_correct_amt
				  ButtonGroup ButtonPressed
				    PushButton 175, 360, 30, 10, "CALC", calc_btn

				  GroupBox 215, 5, 120, 110, "HH Composition"
				  Text 230, 20, 65, 10, "Budgeted HH Size:"
				  Text 300, 20, 15, 10, budgeted_hh_size
				  Text 240, 40, 55, 10, "Correct HH Size:"
				  EditBox 300, 35, 25, 15, correct_hh_size
				  ButtonGroup ButtonPressed
				    PushButton 295, 55, 30, 10, "CALC", calc_btn
				  Text 230, 70, 75, 10, "Standard Deduction:"
				  Text 260, 85, 35, 10, "Budgeted:"
				  Text 300, 85, 25, 10, "$ " & standard_deduction_budgeted_amt
				  Text 270, 100, 30, 10, "Correct:"
				  Text 300, 100, 25, 10, "$ " & standard_deduction_correct_amt
				  Text 220, 125, 50, 10, "Proration Date:"
				  EditBox 275, 120, 60, 15, snap_proration_date
				  'BUTTON
				  GroupBox 340, 5, 210, 345, "Corrected Budget"
				  Text 360, 20, 55, 10, " Earned Income:"
				  Text 425, 20, 30, 10, "$ " & earned_income_correct_amt
				  Text 355, 30, 60, 10, "Unearned Income:"
				  Text 425, 30, 30, 10, "$ " & unearned_correct_amt
				  Text 385, 40, 50, 10, " Total Income:"
				  Text 440, 40, 30, 10, "$ " & total_income_correct_amt

				  Text 350, 55, 70, 10, " Earned Income Ded:"
				  Text 425, 55, 30, 10, "$ " & earned_deduction_correct_amt
				  Text 370, 65, 50, 10, " Standard Ded:"
				  Text 425, 65, 30, 10, "$ " & standard_deduction_correct_amt
				  Text 375, 75, 45, 10, "Medical Ded:"
				  Text 425, 75, 30, 10, "$ " & medical_deduction_correct_amt
				  Text 350, 85, 70, 10, "Dependent Care Ded:"
				  Text 425, 85, 30, 10, "$ " & dependent_care_deduction_correct_amt
				  Text 375, 95, 50, 10, "Child Support:"
				  Text 425, 95, 30, 10, "$ " & child_support_deduction_correct_amt
				  Text 375, 105, 60, 10, " Total Deductions:"
				  Text 440, 105, 30, 10, "$ " & total_deduction_correct_amt

				  Text 390, 120, 40, 10, "Net Income:"
				  Text 440, 120, 30, 10, "$ " & net_income_correct_amt

				  Text 370, 135, 50, 10, "Rent/Mortgage:"
				  Text 425, 135, 30, 10, "$ " & rent_mortgage_correct_amt
				  Text 375, 145, 45, 10, "Property Tax:"
				  Text 425, 145, 30, 10, "$ " & tax_correct_amt
				  Text 360, 155, 60, 10, " House Insurance:"
				  Text 425, 155, 30, 10, "$ " & insurance_correct_amt
				  Text 390, 165, 30, 10, " Utilities:"
				  Text 425, 165, 30, 10, "$ " & utilities_correct_amt
				  Text 355, 175, 70, 10, "Other (" & other_cost_detail & "):"
				  Text 425, 175, 30, 10, "$ " & other_cost_correct_amt
				  Text 365, 185, 70, 10, " Total Shelter Costs:"
				  Text 440, 185, 30, 10, "$ " & total_shelter_cost_correct_amt

				  Text 360, 200, 65, 10, "50% of Net Income:"
				  Text 425, 200, 30, 10, "$ " & fifty_perc_net_income_correct_amt
				  Text 345, 210, 80, 10, "Adjusted Shelter Costs:"
				  Text 425, 210, 30, 10, "$ " & adj_shelter_cost_correct_amt
				  Text 360, 220, 65, 10, " Max Allow Shelter:"
				  Text 425, 220, 30, 10, "$ " & max_shelter_cost_correct_amt
				  Text 345, 230, 90, 10, " Counted Shelter Expense:"
				  Text 440, 230, 30, 10, "$ " & counted_shelter_cost_correct_amt
				  Text 350, 245, 70, 10, "Net Adjusted Income:"
				  Text 425, 245, 30, 10, "$ " & net_adj_income_correct_amt
				  Text 365, 255, 55, 10, "Household Size:"
				  Text 425, 255, 30, 10, correct_hh_size
				  Text 350, 265, 70, 10, " Max Net Adj Income:"
				  Text 425, 265, 30, 10, "$ " & max_net_adj_income_correct_amt
				  Text 360, 275, 75, 10, "Monthly SNAP benefit:"
				  Text 440, 275, 30, 10, "$ " & monthly_snap_benefit_correct_amt
				  Text 360, 285, 75, 10, "Sanction/Recoupment:"
				  Text 440, 285, 30, 10, "$ " & sanction_recoupment_correct_amt
				  Text 405, 305, 100, 10, "Correct SNAP Benefit Amount:"
				  Text 510, 305, 30, 10, "$ " & snap_correct_amt
				  Text 425, 320, 80, 10, "Benefit amount issued:"
				  Text 510, 320, 30, 10, "$ " & snap_issued_amt
				  If snap_overpayment_exists = True Then
					  Text 455, 335, 50, 10, "Overpayment:"
					  Text 510, 335, 30, 10, "$ " & snap_overpayment_amt
				  End If
				  If snap_supplement_exists = True Then
					  Text 455, 335, 50, 10, "Supplement:"
					  Text 510, 335, 30, 10, "$ " & snap_supplement_amt
				  End If
				  If snap_overpayment_exists = False And snap_supplement_exists = False Then
				  	  Text 400, 335, 100, 10, "02/22 Issuance was Correct"
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 480, 10, 65, 12, "CALCULATE", calc_btn
					PushButton 385, 365, 165, 15, "SNAP Budget is Complete", snap_claculation_done_btn

				EndDialog

				dialog Dialog1
				cancel_confirmation

				' If ButtonPressed = -1 Then ButtonPressed = calc_btn
				output_type = "STRING"
				If ButtonPressed = snap_claculation_done_btn Then output_type = "NUMBER"

				Call budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, output_type)
				Call budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, output_type)
				Call budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, output_type)
				Call budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, output_type)
				Call budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, sanction_recoupment_correct_amt, snap_correct_amt, snap_issued_amt, snap_overpayment_exists, snap_supplement_exists, snap_proration_date, snap_overpayment_amt, snap_supplement_amt, output_type)

			Loop until ButtonPressed = snap_claculation_done_btn
			rent_mortgage_correct_amt = rent_mortgage_correct_amt * 1
			tax_correct_amt = tax_correct_amt * 1
			insurance_correct_amt = insurance_correct_amt * 1
			other_cost_correct_amt = other_cost_correct_amt * 1
			total_housing_cost_correct_amt = rent_mortgage_correct_amt + tax_correct_amt + insurance_correct_amt + other_cost_correct_amt
		Else
			snap_correct_amt = 0
			monthly_snap_benefit_correct_amt = 0
			snap_overpayment_amt = snap_issued_amt
			snap_overpayment_exists = True

			earned_income_correct_amt = ""
			unearned_correct_amt = ""
			total_income_correct_amt = ""
			total_deduction_correct_amt = ""
			net_income_correct_amt = ""
			total_housing_cost_correct_amt = ""
			utilities_correct_amt = ""
			total_shelter_cost_correct_amt = ""
			net_adj_income_correct_amt = ""
			correct_hh_size = ""
			earned_deduction_correct_amt = ""
			standard_deduction_correct_amt = ""
			medical_deduction_correct_amt = ""
			dependent_care_deduction_correct_amt = ""
			child_support_deduction_correct_amt = ""
			rent_mortgage_correct_amt = ""
			tax_correct_amt = ""
			insurance_correct_amt = ""
			utilities_budgeted_amt = ""
			other_cost_correct_amt = ""
			fifty_perc_net_income_correct_amt = ""
			adj_shelter_cost_correct_amt = ""
			max_shelter_cost_correct_amt = ""
			counted_shelter_cost_correct_amt = ""
			max_net_adj_income_correct_amt = 0
			sanction_recoupment_correct_amt = ""

		End If

		SNAP_fed_correct_amt = snap_correct_amt * fed_percent
		SNAP_state_correct_amt = snap_correct_amt * state_percent
		If snap_overpayment_exists = True Then
			SNAP_fed_op = snap_overpayment_amt * fed_percent
			SNAP_state_op = snap_overpayment_amt * state_percent
		End If

		If snap_supplement_exists = True Then
			SNAP_fed_supp = snap_supplement_amt * fed_percent
			SNAP_state_supp = snap_supplement_amt * state_percent
		End If


	End If

	' MsgBox "DONE"
	SNAP_confirmation_answer = "Select One..."
	'dialog with calculation and ready for confirmation
	Do
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 206, 205, "Confirm Budget Calculation"
		  If SNAP_active = True Then
			  GroupBox 5, 5, 195, 135, "SNAP"
			  Text 20, 20, 85, 10, "Original SNAP Issuance:"
			  Text 105, 20, 40, 10, "$ " & snap_issued_amt
			  Text 115, 30, 40, 10, "$ " & fed_benefit_amt
			  Text 155, 30, 30, 10, "Federal "
			  Text 115, 40, 40, 10, "$ " & state_benefit_amt
			  Text 155, 40, 20, 10, "State"
			  Text 20, 55, 75, 10, "SNAP Recalculation:"
			  If snap_overpayment_exists = True Then
				  Text 35, 65, 60, 10, "Overpayment"
				  Text 35, 75, 30, 10, "Amount:"
				  Text 70, 75, 40, 10, "$ " & snap_overpayment_amt
				  Text 80, 85, 40, 10, "$ " & SNAP_fed_op
				  Text 120, 85, 30, 10, "Federal "
				  Text 80, 95, 40, 10, "$ " & SNAP_state_op
				  Text 120, 95, 20, 10, "State"
			  End If
			  If snap_supplement_exists = True Then
				  Text 35, 65, 60, 10, "Supplement"
				  Text 35, 75, 30, 10, "Amount:"
				  Text 70, 75, 40, 10, "$ " & snap_supplement_amt
				  Text 80, 85, 40, 10, "$ " & SNAP_fed_supp
				  Text 120, 85, 30, 10, "Federal "
				  Text 80, 95, 40, 10, "$ " & SNAP_state_supp
				  Text 120, 95, 20, 10, "State"
			  End If
			  If snap_overpayment_exists = False And snap_supplement_exists = False Then
				  Text 35, 65, 100, 10, "02/22 Issuance was Correct"
			  End If
			  Text 15, 110, 90, 10, "Is this calculation Correct?"
			  DropListBox 15, 120, 180, 45, "Select One..."+chr(9)+"Yes - this recalculation is correct"+chr(9)+"No - something needs to be updated", SNAP_confirmation_answer
			  Text 5, 145, 193, 35, "Once this is confirmed, the script will update documentation. It will appear that nothing is happening. Leave the computer to process for a minute and the script will alert you once it is done. Do not multitask at this time."
		  End If
		  ButtonGroup ButtonPressed
		    PushButton 5, 185, 195, 15, "Enter Calculation Information to Tracking Spreadsheet", Button3
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If SNAP_confirmation_answer = "Select One..." Then Msgbox "Please Resolve to Continue:" & vbCr & vbCr & "* Please indicate if this calculation is correct or not."
	Loop until SNAP_confirmation_answer <> "Select One..."

	If SNAP_confirmation_answer = "Yes - this recalculation is correct" Then recalculation_confirmed = True
	If SNAP_confirmation_answer = "No - something needs to be updated" Then
		Call budget_calculate_income(earned_income_correct_amt, unearned_correct_amt, earned_deduction_correct_amt, total_income_correct_amt, "STRING")
		Call budget_calculate_household(correct_hh_size, disa_household, cat_elig, standard_deduction_correct_amt, max_shelter_cost_correct_amt, max_gross_income_correct_amt, max_net_adj_income_correct_amt, max_snap_benefit, "STRING")
		Call budget_calculate_deductions(earned_deduction_correct_amt, medical_deduction_correct_amt, dependent_care_deduction_correct_amt, child_support_deduction_correct_amt, standard_deduction_correct_amt, total_deduction_correct_amt, total_income_correct_amt, net_income_correct_amt, fifty_perc_net_income_correct_amt, "STRING")
		Call budget_calculate_shelter_costs(rent_mortgage_correct_amt, tax_correct_amt, insurance_correct_amt, other_cost_correct_amt, utilities_correct_amt, total_shelter_cost_correct_amt, adj_shelter_cost_correct_amt, max_shelter_cost_correct_amt, counted_shelter_cost_correct_amt, fifty_perc_net_income_correct_amt, net_income_correct_amt, net_adj_income_correct_amt, "STRING")
		Call budget_calculate_benefit_details(cat_elig, total_income_correct_amt, net_adj_income_correct_amt, max_net_adj_income_correct_amt, max_gross_income_correct_amt, max_snap_benefit, monthly_snap_benefit_correct_amt, sanction_recoupment_correct_amt, snap_correct_amt, snap_issued_amt, snap_overpayment_exists, snap_supplement_exists, snap_proration_date, snap_overpayment_amt, snap_supplement_amt, "STRING")
	End If

	' LOOP UNTIL THIS IS CONFIRMED
Loop until recalculation_confirmed = True


'Create PDF and save
If snap_overpayment_exists = True Then
	Set objWord = CreateObject("Word.Application")

	'Adding all of the information in the dialogs into a Word Document
	objWord.Caption = "AutoClose Pause OP Calculation - CASE #" & MAXIS_case_number			'Title of the document
	' objWord.Visible = True														'Let the worker see the document
	objWord.Visible = False														'Let the worker see the document

	Set objDoc = objWord.Documents.Add()										'Start a new document
	Set objSelection = objWord.Selection

	objSelection.PageSetup.TopMargin = 36
	objSelection.PageSetup.BottomMargin = 36
	objSelection.ParagraphFormat.SpaceAfter = 0

	objSelection.Font.Name = "Arial"											'Setting the font before typing
	objSelection.Font.Size = "16"
	objSelection.Font.Bold = TRUE
	objSelection.TypeText "SNAP Overpayment - Case # " & MAXIS_case_number
	objSelection.TypeParagraph()
	objSelection.Font.Size = "12"
	objSelection.Font.Bold = FALSE

	objSelection.TypeText "Details about the AutoClose Process that was Paused and Follow Up Review"
	objSelection.TypeText vbCr

	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, 4, 4					'This sets the rows and columns needed row then column
	'This table starts with 1 column - other columns are added after we split some of the cells
	set process_info = objDoc.Tables(1)		'Creates the table with the specific index'

	for row = 1 to 4
		process_info.Cell(row, 1).SetHeight 15, 2			'setting the heights of the rows
	Next
	process_info.Columns(1).SetWidth 150, 2
	process_info.Columns(2).SetWidth 100, 2
	process_info.Columns(3).SetWidth 150, 2
	process_info.Columns(4).SetWidth 100, 2

	process_info.Cell(1, 1).Range.Text = "Process"
	process_info.Cell(1, 2).Range.Text = feb_process
	process_info.Cell(1, 3).Range.Text = "The " & feb_process & " was completed"
	process_info.Cell(1, 4).Range.Text = process_complete
	process_info.Cell(2, 1).Range.Text = "Form Received"
	process_info.Cell(2, 2).Range.Text = form_received
	If form_received <> "None Received" Then
		process_info.Cell(2, 3).Range.Text = "Form Date"
		process_info.Cell(2, 4).Range.Text = form_received_date
	End If
	process_info.Cell(3, 1).Range.Text = "Interview"
	process_info.Cell(3, 2).Range.Text = interview_information
	If interview_date <> "" Then
		process_info.Cell(3, 3).Range.Text = "Interview Date"
		process_info.Cell(3, 4).Range.Text = interview_date
	End If
	process_info.Cell(4, 1).Range.Text = "Verifications"
	process_info.Cell(4, 2).Range.Text = verifs_received
	process_info.Cell(4, 3).Range.Text = "Proration Date"
	process_info.Cell(4, 4).Range.Text = snap_proration_date

	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()


	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, 33, 1					'This sets the rows and columns needed row then column
	'This table starts with 1 column - other columns are added after we split some of the cells
	set snap_op_table = objDoc.Tables(2)		'Creates the table with the specific index'
	snap_op_table.AutoFormat(16)							'This adds the borders to the table and formats it

	for row = 1 to 33
		snap_op_table.Cell(row, 1).SetHeight 15, 2			'setting the heights of the rows
	Next

	for row = 1 to 2
		snap_op_table.Rows(row).Cells.Split 1, 2, TRUE
		snap_op_table.Cell(row, 1).SetWidth 250, 2
		snap_op_table.Cell(row, 2).SetWidth 85, 2
	Next
	snap_op_table.Cell(3, 1).SetWidth 335, 2
	snap_op_table.Cell(3, 1).Range.Font.Bold = TRUE
	for row = 4 to 6
		snap_op_table.Rows(row).Cells.Split 1, 2, TRUE
		snap_op_table.Cell(row, 1).SetWidth 250, 2
		snap_op_table.Cell(row, 2).SetWidth 85, 2
	Next
	snap_op_table.Cell(7, 1).SetWidth 335, 2
	snap_op_table.Cell(7, 1).Range.Font.Bold = TRUE
	for row = 8 to 14
		snap_op_table.Rows(row).Cells.Split 1, 2, TRUE
		snap_op_table.Cell(row, 1).SetWidth 250, 2
		snap_op_table.Cell(row, 2).SetWidth 85, 2
	Next
	snap_op_table.Cell(15, 1).SetWidth 335, 2
	snap_op_table.Cell(15, 1).Range.Font.Bold = TRUE
	for row = 16 to 33
		snap_op_table.Rows(row).Cells.Split 1, 2, TRUE
		snap_op_table.Cell(row, 1).SetWidth 250, 2
		snap_op_table.Cell(row, 2).SetWidth 85, 2
	Next

	snap_op_table.Cell(1, 1).Range.Text = "Income Month/Year"
	snap_op_table.Cell(1, 2).Range.Text = "02/22"
	snap_op_table.Cell(2, 1).Range.Text = "Benefit Month/Year"
	snap_op_table.Cell(2, 2).Range.Text = "02/22"
	snap_op_table.Cell(3, 1).Range.Text = "Income"
	snap_op_table.Cell(4, 1).Range.Text = chr(9) & chr(9) & "Earned Income"
	snap_op_table.Cell(4, 2).Range.Text = "$  " & earned_income_correct_amt
	snap_op_table.Cell(5, 1).Range.Text = chr(9) & chr(9) & "Unearned Income"
	snap_op_table.Cell(5, 2).Range.Text = "$  " & unearned_correct_amt
	snap_op_table.Cell(6, 1).Range.Text = chr(9) & "Total Income"
	snap_op_table.Cell(6, 2).Range.Text = "$  " & total_income_correct_amt
	snap_op_table.Cell(7, 1).Range.Text = "Deductions for SNAP"
	snap_op_table.Cell(8, 1).Range.Text = chr(9) & chr(9) & "Earned income Deduction"
	snap_op_table.Cell(8, 2).Range.Text = "$  " & earned_deduction_correct_amt
	snap_op_table.Cell(9, 1).Range.Text = chr(9) & chr(9) & "Standard Deduction"
	snap_op_table.Cell(9, 2).Range.Text = "$  " & standard_deduction_correct_amt
	snap_op_table.Cell(10, 1).Range.Text = chr(9) & chr(9) & "Medical Deduction"
	snap_op_table.Cell(10, 2).Range.Text = "$  " & medical_deduction_correct_amt
	snap_op_table.Cell(11, 1).Range.Text = chr(9) & chr(9) & "Dependent Care Deduction"
	snap_op_table.Cell(11, 2).Range.Text = "$  " & dependent_care_deduction_correct_amt
	snap_op_table.Cell(12, 1).Range.Text = chr(9) & chr(9) & "Child Support"
	snap_op_table.Cell(12, 2).Range.Text = "$  " & child_support_deduction_correct_amt
	snap_op_table.Cell(13, 1).Range.Text = chr(9) & "Total Deductions"
	snap_op_table.Cell(13, 2).Range.Text = "$  " & total_deduction_correct_amt
	snap_op_table.Cell(14, 1).Range.Text = chr(9) & "Net Income"
	snap_op_table.Cell(14, 2).Range.Text = "$  " & net_income_correct_amt
	snap_op_table.Cell(15, 1).Range.Text = "Shelter Costs"
	snap_op_table.Cell(16, 1).Range.Text = chr(9) & chr(9) & "Rent/Mortgage"
	snap_op_table.Cell(16, 2).Range.Text = "$  " & rent_mortgage_correct_amt
	snap_op_table.Cell(17, 1).Range.Text = chr(9) & chr(9) & "Property Tax"
	snap_op_table.Cell(17, 2).Range.Text = "$  " & tax_correct_amt
	snap_op_table.Cell(18, 1).Range.Text = chr(9) & chr(9) & "House Insurance"
	snap_op_table.Cell(18, 2).Range.Text = "$  " & insurance_correct_amt
	snap_op_table.Cell(19, 1).Range.Text = chr(9) & chr(9) & "Utilities"
	snap_op_table.Cell(19, 2).Range.Text = "$  " & utilities_correct_amt
	snap_op_table.Cell(20, 1).Range.Text = chr(9) & chr(9) & "Other " & other_cost_detail
	snap_op_table.Cell(20, 2).Range.Text = "$  " & other_cost_correct_amt
	snap_op_table.Cell(21, 1).Range.Text = chr(9) & "Total Shelter Costs"
	snap_op_table.Cell(21, 2).Range.Text = "$  " & total_shelter_cost_correct_amt
	snap_op_table.Cell(22, 1).Range.Text = chr(9) & chr(9) & "50% of Net Income"
	snap_op_table.Cell(22, 2).Range.Text = "$  " & fifty_perc_net_income_correct_amt
	snap_op_table.Cell(23, 1).Range.Text = chr(9) & chr(9) & "Adjusted Shelter Costs"
	snap_op_table.Cell(23, 2).Range.Text = "$  " & adj_shelter_cost_correct_amt
	snap_op_table.Cell(24, 1).Range.Text = chr(9) & chr(9) & "Max Allow Shelter"
	snap_op_table.Cell(24, 2).Range.Text = "$  " & max_shelter_cost_correct_amt
	snap_op_table.Cell(25, 1).Range.Text = chr(9) & "Shelter Expense"
	snap_op_table.Cell(25, 2).Range.Text = "$  " & counted_shelter_cost_correct_amt
	snap_op_table.Cell(26, 1).Range.Text = chr(9) & "Net Adusted Income"
	snap_op_table.Cell(26, 2).Range.Text = "$  " & net_adj_income_correct_amt
	snap_op_table.Cell(27, 1).Range.Text = chr(9) & "Household Size"
	snap_op_table.Cell(27, 2).Range.Text = "   " & correct_hh_size
	snap_op_table.Cell(28, 1).Range.Text = chr(9) & "Max Net Adjusted Income"
	snap_op_table.Cell(28, 2).Range.Text = "$  " & max_net_adj_income_correct_amt
	snap_op_table.Cell(29, 1).Range.Text = "Monthly SNAP Benefit"
	snap_op_table.Cell(29, 2).Range.Text = "$  " & monthly_snap_benefit_correct_amt
	snap_op_table.Cell(30, 1).Range.Text = chr(9) & "Drug felon sanction/Recoupment"
	snap_op_table.Cell(30, 2).Range.Text = "$  " & sanction_recoupment_correct_amt
	snap_op_table.Cell(31, 1).Range.Text = "Correct SNAP Benefit Amount"
	snap_op_table.Cell(31, 2).Range.Text = "$  " & snap_correct_amt
	snap_op_table.Cell(32, 1).Range.Text = "Benefit Amount Issued"
	snap_op_table.Cell(32, 2).Range.Text = "$  " & snap_issued_amt
	snap_op_table.Cell(33, 1).Range.Text = "Overpayment"
	snap_op_table.Cell(33, 2).Range.Text = "$  " & snap_overpayment_amt

	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

	objSelection.TypeText vbCr
	objSelection.TypeText "Overpayment calculation is being completed in response to the AutoClose process being paused in 02/22 for REVW and MONT panels in Hennepin County. These overpayments do not follow the typical process for responsibility and will not be entered in CCOL or CASE/NOTE."
	objSelection.TypeText vbCr
	objSelection.TypeText "Calculation completed by: " & user_name

	'Here we are creating the file path and saving the file
	file_safe_date = replace(date, "/", "-")		'dates cannot have / for a file name so we change it to a -

	'We set the file path and name based on case number and date. We can add other criteria if important.
	'This MUST have the 'pdf' file extension to work
	pdf_doc_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Overpayment Calculation Documents" & "\AutoClose Pause OP Calculation - CASE " & MAXIS_case_number & ".pdf"
	' If developer_mode = True Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Archive\TRAINING REGION Interviews - NOT for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"

	'Now we save the document.
	'MS Word allows us to save directly as a PDF instead of a DOC.
	'the file path must be PDF
	'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
	objDoc.SaveAs pdf_doc_path, 17

	objDoc.Close wdDoNotSaveChanges
	objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)

	Call excel_open(excel_report_file_path, False, False, ObjReportExcel, objReportWorkbook)
	rept_excel_row = 1
	Do
		rept_excel_row = rept_excel_row + 1
		this_case_number = trim(ObjReportExcel.Cells(rept_excel_row, 1).Value)
	Loop Until this_case_number = ""												'if the case number is blank then the row is blank

	ObjReportExcel.Cells(rept_excel_row, rept_case_numb_col).Value  		= MAXIS_case_number
	ObjReportExcel.Cells(rept_excel_row, rept_process_col).Value  			= feb_process
	ObjReportExcel.Cells(rept_excel_row, rept_issued_fs_f_col).Value  		= fed_benefit_amt
	ObjReportExcel.Cells(rept_excel_row, rept_issued_fs_s_col).Value  		= state_benefit_amt
	ObjReportExcel.Cells(rept_excel_row, rept_issued_mf_fs_f_col).Value  	= mfip_MF_FS_F_issued_amt
	ObjReportExcel.Cells(rept_excel_row, rept_issued_mf_fs_s_col).Value  	= mfip_MF_FS_S_issued_amt
	ObjReportExcel.Cells(rept_excel_row, rept_op_fs_f_col).Value  			= SNAP_fed_op
	ObjReportExcel.Cells(rept_excel_row, rept_op_fs_s_col).Value  			= SNAP_state_op
	ObjReportExcel.Cells(rept_excel_row, rept_op_mf_fs_f_col).Value  		= ""
	ObjReportExcel.Cells(rept_excel_row, rept_op_mf_fs_s_col).Value  		= ""

	objReportWorkbook.Save()		'saving the excel
	ObjReportExcel.ActiveWorkbook.Close

	ObjReportExcel.Application.Quit
	ObjReportExcel.Quit
End If
EMWaitReady 1, 1
' MsgBox "Pause Here"



'Open Excel and add information to the excel
Call excel_open(excel_details_file_path, False, False, ObjDetailsExcel, objDetailWorkbook)

total_excel_row = 1
Do
	total_excel_row = total_excel_row + 1
	this_case_number = trim(ObjDetailsExcel.Cells(total_excel_row, 1).Value)
Loop Until this_case_number = ""												'if the case number is blank then the row is blank

ObjDetailsExcel.Cells(total_excel_row, det_case_numb_col).Value 			= MAXIS_case_number
ObjDetailsExcel.Cells(total_excel_row, det_process_col).Value 				= feb_process
ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_f_col).Value 			= fed_benefit_amt
ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_s_col).Value 			= state_benefit_amt
ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_mf_col).Value 			= mfip_MF_MF_issued_amt
ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_f_col).Value 		= mfip_MF_FS_F_issued_amt
ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_s_col).Value 		= mfip_MF_FS_S_issued_amt
ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_hg_col).Value 			= mfip_MF_HG_issued_amt
ObjDetailsExcel.Cells(total_excel_row, det_form_col).Value 					= form_received
ObjDetailsExcel.Cells(total_excel_row, det_form_date_col).Value 			= form_received_date
ObjDetailsExcel.Cells(total_excel_row, det_intv_col).Value 					= interview_information
ObjDetailsExcel.Cells(total_excel_row, det_intv_date_col).Value 			= interview_date
ObjDetailsExcel.Cells(total_excel_row, det_verifs_col).Value 				= verifs_received
ObjDetailsExcel.Cells(total_excel_row, det_process_complete_col).Value 		= process_complete
ObjDetailsExcel.Cells(total_excel_row, det_op_fs_f_col).Value 				= SNAP_fed_op
ObjDetailsExcel.Cells(total_excel_row, det_op_fs_s_col).Value 				= SNAP_state_op
ObjDetailsExcel.Cells(total_excel_row, det_op_mf_mf_col).Value 				= ""
ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_f_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_s_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_op_mf_hg_col).Value 				= ""
ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_f_col).Value 			= SNAP_fed_supp
ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_s_col).Value 			= SNAP_state_supp
ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_mf_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_f_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_s_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_hg_col).Value 			= ""
ObjDetailsExcel.Cells(total_excel_row, det_orig_earned_income_col).Value 	= earned_income_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_unearned_income_col).Value 	= unearned_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_total_income_col).Value 	= total_income_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_total_ded_col).Value 		= total_deduction_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_net_income_col).Value 		= net_income_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_housing_cost_col).Value 	= total_housing_cost_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_utility_cost_col).Value 	= utilities_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_total_shel_cost_col).Value 	= total_shelter_cost_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_net_adj_income_col).Value 	= net_adj_income_budgeted_amt
ObjDetailsExcel.Cells(total_excel_row, det_orig_hh_size_col).Value 			= budgeted_hh_size
ObjDetailsExcel.Cells(total_excel_row, det_orig_snap_benefit_col).Value 		= snap_issued_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_earned_income_col).Value 	= earned_income_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_unearned_income_col).Value 	= unearned_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_total_income_col).Value 		= total_income_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_total_ded_col).Value 		= total_deduction_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_net_income_col).Value 		= net_income_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_housing_cost_col).Value 		= total_housing_cost_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_utility_cost_col).Value 		= utilities_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_total_shel_cost_col).Value 	= total_shelter_cost_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_net_adj_income_col).Value 	= net_adj_income_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_correct_hh_size_col).Value 			= correct_hh_size
If calculation_needed = True Then ObjDetailsExcel.Cells(total_excel_row, det_snap_proration_col).Value 			= snap_proration_date
ObjDetailsExcel.Cells(total_excel_row, det_correct_snap_benefit_col).Value 		= snap_correct_amt
ObjDetailsExcel.Cells(total_excel_row, det_pdf_link_col).Value					= ""
If pdf_doc_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & pdf_doc_path & chr(34) & ", " & chr(34) & "AutoClose Pause OP Calculation - CASE " & MAXIS_case_number & ".pdf" & chr(34) & ")"

objDetailWorkbook.Save()		'saving the excel
ObjDetailsExcel.ActiveWorkbook.Close

ObjDetailsExcel.Application.Quit
ObjDetailsExcel.Quit

Call script_end_procedure("Information has been saved.")
