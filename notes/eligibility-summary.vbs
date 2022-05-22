'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
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

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result)
	Call write_value_and_transmit("99", cmd_row, cmd_col)

	row = 7
	Do
		EMReadScreen elig_version, 2, row, 22
		EmReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do

		row = row + 1
	Loop until approval_status = ""
	Call clear_line_of_text(18, 54)

	Call write_value_and_transmit(elig_version, 18, 54)
	version_number = "0" & elig_version
	version_date = elig_date
	version_result = elig_result
end function

Function read_SNAP_elig(footer_month, footer_year, version_number, version_date, elig_result, fs_expedited, exp_package_includes_month_one, exp_package_includes_month_two, fs_prorated, earned_income_budgeted, unearned_income_budgeted, shel_costs_budgeted, hest_costs_budgeted, categorical_eligibility, case_appl_withdrawn, case_applct_elig, case_comdty, case_disq, case_dupl_assist, case_eligible_person, case_fail_coop, case_fail_file, case_prosp_gross_inc, case_prosp_net_inc, case_recert, case_residence, case_resource, case_retro_gross_inc, case_retro_net_inc, case_strike, case_xfer_resource_inc, case_verif, case_voltry_quit, case_work_reg, fail_file_hrf, fail_file_sr, resource_cash, resource_acct, resource_secu, resource_cars, resource_rest, resource_other, resource_burial, resource_spon, resource_total, resource_max, fsb1_gross_wages, fsb1_self_emp, fsb1_total_earned_inc, fsb1_pa_grant_inc, fsb1_rsdi_inc, fsb1_ssi_inc, fsb1_va_inc, fsb1_uc_wc_inc, fsb1_cses_inc, fsb1_other_unea_inc, fsb1_total_unea_inc, fsb1_schl_inc, fsb1_farm_ofset, fsb1_total_gross_inc, fsb1_max_gross_inc, fsb1_deduct_standard, fsb1_deduct_earned, fsb1_deduct_medical, fsb1_deduct_depndt_care, fsb1_cses, fsb1_total_deduct, fsb1_net_inc, fsb2_shel_rent_mort, fsb2_shel_prop_tax, fsb2_shel_home_ins, fsb2_shel_electricity, fsb2_shel_heat_ac, fsb2_shel_water_garbage, fsb2_shel_phone, fsb2_shel_other, fsb2_shel_total, fsb2_50_perc_net_inc, fsb2_adj_shel_costs, fsb2_max_allow_shel, fsb2_shel_expenses, fsb2_net_adj_inc, fsb2_monthly_fs_allot, fsb2_drug_felon_sanc_amt, fsb2_recoup_amount, fsb2_benefit_amount, fsb2_state_food_amt, fsb2_fed_food_amt, recoup_from_fed_fs, recoup_from_state_fs, fssm_approved_date, fssm_date_last_approval, fssm_curr_prog_status, fssm_elig_result, fssm_reporting_status, fssm_info_source, fssm_benefit, fssm_elig_revw_date, fssm_budget_cycle, fssm_numb_in_assist_unit, fssm_total_resources, fssm_max_resources, fssm_net_adj_inc, fssm_monthly_fs_allotment, fssm_prorated_amt, fssm_prorated_date, fssm_benefit_amt, exp_criteria_migrant_destitute, exp_criteria_resource_100_income_150, exp_criteria_resource_income_less_shelter, exp_verif_status_postponed, exp_verif_status_out_of_state, exp_verif_status_all_provided, fssm_worker_message_one, fssm_worker_message_two, MEMBER_ARRAY, ref_numb_const, request_yn_const, memb_code_const, memb_status_info_const, memb_counted_const, memb_state_food_const, memb_elig_status_const, memb_begin_date_const, memb_budg_cycle_const, memb_abawd_const, memb_absence_const, memb_roomer_const, memb_boarder_const, memb_citizenship_const, memb_citizenship_coop_const, memb_cmdty_const, memb_disq_const, memb_dupl_assist_const, memb_fraud_const, memb_eligible_student_const, memb_institution_const, memb_mfip_elig_const, memb_non_applcnt_const, memb_residence_const, memb_ssn_coop_const, memb_unit_memb_const, memb_work_reg_const, memb_drug_felon_test_const)

	fs_expedited = False
	exp_package_includes_month_one = False
	exp_package_includes_month_two = False
	fs_prorated = False
	earned_income_budgeted = False
	unearned_income_budgeted = False
	shel_costs_budgeted = False
	hest_costs_budgeted = False
	categorical_eligibility = ""

	call navigate_to_MAXIS_screen("ELIG", "FS  ")
	EMWriteScreen footer_month, 19, 54
	EMWriteScreen footer_year, 19, 57
	Call find_last_approved_ELIG_version(19, 78, version_number, version_date, elig_result)

	row = 7
	Do
		EMReadScreen ref_numb, 2, row, 10

		For case_memb = 0 to UBound(MEMBER_ARRAY, 2)

			If MEMBER_ARRAY(ref_numb_const, case_memb) = ref_numb Then
				EMReadScreen request_yn, 1, row, 32
				EMReadScreen memb_code, 1, row, 35
				EMReadScreen memb_count, 11, row, 39
				EMReadScreen memb_state_food, 1, row, 50
				EMReadScreen memb_elig, 10, row, 57
				EMReadScreen memb_begin_date, 8, row, 68
				EMReadScreen memb_budg_cycle, 1, row, 78

				MEMBER_ARRAY(request_yn_const, case_memb) = request_yn
				MEMBER_ARRAY(memb_code_const, case_memb) = memb_code
				If memb_code = "A" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Eligible"
				If memb_code = "C" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Citizenship"
				If memb_code = "F" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Fraud, DISQ, Work Reg"
				If memb_code = "D" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Duplicate Assistance"
				If memb_code = "I" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Ineligible"
				If memb_code = "N" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Unit Member"
				If memb_code = "S" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Ineligible Student"
				If memb_code = "U" Then MEMBER_ARRAY(memb_status_info_const, case_memb) = "Unknown"
				MEMBER_ARRAY(memb_counted_const, case_memb) = trim(memb_count)
				If memb_state_food = "Y" Then MEMBER_ARRAY(memb_state_food_const, case_memb) = True
				If memb_state_food = "N" Then MEMBER_ARRAY(memb_state_food_const, case_memb) = False
				MEMBER_ARRAY(memb_elig_status_const, case_memb) = trim(memb_elig)
				MEMBER_ARRAY(memb_begin_date_const, case_memb) = memb_begin_date
				If memb_budg_cycle = "P" Then MEMBER_ARRAY(memb_budg_cycle_const, case_memb) = "Prospective"
				If memb_budg_cycle = "R" Then MEMBER_ARRAY(memb_budg_cycle_const, case_memb) = "Retrospective"

				Call write_value_and_transmit("X", row, 5)

				EMReadScreen memb_abawd, 			6, 6, 20
				EMReadScreen memb_absence, 			6, 7, 20
				EMReadScreen memb_roomer, 			6, 8, 20
				EMReadScreen memb_boarder, 			6, 9, 20
				EMReadScreen memb_citizenship, 		6, 10, 20
				EMReadScreen memb_citizenship_coop, 6, 11, 20
				EMReadScreen memb_cmdty, 			6, 12, 20
				EMReadScreen memb_disq,				6, 13, 20
				EMReadScreen memb_dupl_assist, 		6, 14, 20

				MEMBER_ARRAY(memb_abawd_const, case_memb) = trim(memb_abawd)
				MEMBER_ARRAY(memb_absence_const, case_memb) = trim(memb_absence)
				MEMBER_ARRAY(memb_roomer_const, case_memb) = trim(memb_roomer)
				MEMBER_ARRAY(memb_boarder_const, case_memb) = trim(memb_boarder)
				MEMBER_ARRAY(memb_citizenship_const, case_memb) = trim(memb_citizenship)
				MEMBER_ARRAY(memb_citizenship_coop_const, case_memb) = trim(memb_citizenship_coop)
				MEMBER_ARRAY(memb_cmdty_const, case_memb) = trim(memb_cmdty)
				MEMBER_ARRAY(memb_disq_const, case_memb) = trim(memb_disq)
				MEMBER_ARRAY(memb_dupl_assist_const, case_memb) = trim(memb_dupl_assist)

				EMReadScreen memb_fraud, 			6, 6, 54
				EMReadScreen memb_eligible_student, 6, 7, 54
				EMReadScreen memb_institution, 		6, 8, 54
				EMReadScreen memb_mfip_elig, 		6, 9, 54
				EMReadScreen memb_non_applcnt, 		6, 10, 54
				EMReadScreen memb_residence, 		6, 11, 54
				EMReadScreen memb_ssn_coop, 		6, 12, 54
				EMReadScreen memb_unit_memb, 		6, 13, 54
				EMReadScreen memb_work_reg, 		6, 14, 54

				MEMBER_ARRAY(memb_fraud_const, case_memb) = trim(memb_fraud)
				MEMBER_ARRAY(memb_eligible_student_const, case_memb) = trim(memb_eligible_student)
				MEMBER_ARRAY(memb_institution_const, case_memb) = trim(memb_institution)
				MEMBER_ARRAY(memb_mfip_elig_const, case_memb) = trim(memb_mfip_elig)
				MEMBER_ARRAY(memb_non_applcnt_const, case_memb) = trim(memb_non_applcnt)
				MEMBER_ARRAY(memb_residence_const, case_memb) = trim(memb_residence)
				MEMBER_ARRAY(memb_ssn_coop_const, case_memb) = trim(memb_ssn_coop)
				MEMBER_ARRAY(memb_unit_memb_const, case_memb) = trim(memb_unit_memb)
				MEMBER_ARRAY(memb_work_reg_const, case_memb) = trim(memb_work_reg)
				transmit
			End If
		Next

		row = row + 1
		EMReadScreen next_ref_numb, 2, row, 10
	Loop until next_ref_numb = "  "

	transmit 		'FSCR
	EmReadScreen case_expedited_indicator, 9, 4, 3
	If case_expedited_indicator = "EXPEDITED" Then fs_expedited = True

	EMReadScreen case_appl_withdrawn, 	6, 7, 9
	EMReadScreen case_applct_elig, 		6, 8, 9
	EMReadScreen case_comdty, 			6, 9, 9
	EMReadScreen case_disq, 			6, 10, 9
	EMReadScreen case_dupl_assist, 		6, 11, 9
	EMReadScreen case_eligible_person, 	6, 12, 9
	EMReadScreen case_fail_coop, 		6, 13, 9
	EMReadScreen case_fail_file, 		6, 14, 9
	EMReadScreen case_prosp_gross_inc, 	6, 15, 9
	EMReadScreen case_prosp_net_inc, 	6, 16, 9
	case_appl_withdrawn = trim(case_appl_withdrawn)
	case_applct_elig = trim(case_applct_elig)
	case_comdty = trim(case_comdty)
	case_disq = trim(case_disq)
	case_dupl_assist = trim(case_dupl_assist)
	case_eligible_person = trim(case_eligible_person)
	case_fail_coop = trim(case_fail_coop)
	case_fail_file = trim(case_fail_file)
	case_prosp_gross_inc = trim(case_prosp_gross_inc)
	case_prosp_net_inc = trim(case_prosp_net_inc)

	EMReadScreen case_recert, 			6, 7, 49
	EMReadScreen case_residence, 		6, 8, 49
	EMReadScreen case_resource, 		6, 9, 49
	EMReadScreen case_retro_gross_inc, 	6, 10, 49
	EMReadScreen case_retro_net_inc, 	6, 11, 49
	EMReadScreen case_strike, 			6, 12, 49
	EMReadScreen case_xfer_resource_inc, 6, 13, 49
	EMReadScreen case_verif, 			6, 14, 49
	EMReadScreen case_voltry_quit, 		6, 15, 49
	EMReadScreen case_work_reg, 		6, 16, 49
	case_recert = trim(case_recert)
	case_residence = trim(case_residence)
	case_resource = trim(case_resource)
	case_retro_gross_inc = trim(case_retro_gross_inc)
	case_retro_net_inc = trim(case_retro_net_inc)
	case_strike = trim(case_strike)
	case_xfer_resource_inc = trim(case_xfer_resource_inc)
	case_verif = trim(case_verif)
	case_voltry_quit = trim(case_voltry_quit)
	case_work_reg = trim(case_work_reg)

	Call write_value_and_transmit("X", 14, 4)		''Fail to File Detail
	EMReadScreen fail_file_hrf, 6, 10, 32
	EMReadScreen fail_file_sr, 6, 11, 32
	transmit
	fail_file_hrf = trim(fail_file_hrf)
	fail_file_sr = trim(fail_file_sr)

	' Call write_value_and_transmit("X", 14, 4)		''Prosp Gross Income Detail
	' EMReadScreen case_gross_wages, 		10, 6, 30
	' EMReadScreen case_self_emp, 		10, 7, 30
	' EMReadScreen case_total_earned_inc, 10, 9, 30
	' EMReadScreen case_pa_grants_inc, 	10, 11, 30
	' EMReadScreen case_rsdi_inc, 		10, 12, 30
	' EMReadScreen case_ssi_inc, 			10, 13, 30
	' EMReadScreen case_va_inc, 			10, 14, 30
	' EMReadScreen case_us_wc_inc, 		10, 15, 30
	'
	' EMReadScreen case_cses_inc, 		10, 6, 67
	' EMReadScreen case_other_unea_inc, 	10, 7, 67
	' EMReadScreen case_total_unea_inc, 	10, 9, 67
	' EMReadScreen case_schl_inc, 		10, 11, 67
	' EMReadScreen case_farm_loss_ofset, 	10, 12, 67
	' EMReadScreen case_total_gross_inc, 	10, 13, 67
	' EMReadScreen case_max_gross_inc, 	10, 15, 67
	' transmit

	' Call write_value_and_transmit("X", 14, 4)		''prosp Net Income Detail
	' EMReadScreen case_gross_income, 		10, 5, 28
	' EMReadScreen case_deduct_standard, 		10, 8, 28
	' EMReadScreen case_deduct_earned_inc, 	10, 9, 28
	' EMReadScreen case_deduct_medical, 		10, 10, 28
	' EMReadScreen case_deduct_depndt_care, 	10, 11, 28
	' EMReadScreen case_deduct_cses, 			10, 12, 28
	' EMReadScreen case_total_deduct, 		10, 14, 28
	' EMReadScreen case_net_income, 			10, 16, 28
	' EMReadScreen case_shel_rent_mort, 		10, 19, 28
	' EMReadScreen case_shel_prop_tax, 		10, 20, 28
	' EMReadScreen case_shel_home_ins, 		10, 21, 28
	'
	' EMReadScreen case_hest_electricity, 10, 5, 65
	' EMReadScreen case_hest_heat_ac, 	10, 6, 65
	' EMReadScreen case_hest_water, 		10, 7, 65
	' EMReadScreen case_hest_garbage, 	10, 8, 65
	' EMReadScreen case_hest_phone, 		10, 9, 65
	' EMReadScreen case_hest_other, 		10, 10, 65
	' EMReadScreen case_total_shel_costs, 10, 12, 65
	' EMReadScreen case_50_perc_net_inc, 	10, 13, 65
	' EMReadScreen case_adj_shel_costs, 	10, 14, 65
	' EMReadScreen case_max_allow_shel, 	10, 15, 65
	' EMReadScreen case_shel_expense, 	10, 17, 65
	' EMReadScreen case_net_adj_inc, 		10, 19, 65
	' EMReadScreen case_max_net_adj_inc, 	10, 21, 65
	' transmit

	Call write_value_and_transmit("X", 14, 4)		''Resource Detail
	EMReadScreen resource_cash, 	10, 8, 47
	EMReadScreen resource_acct, 	10, 9, 47
	EMReadScreen resource_secu, 	10, 10, 47
	EMReadScreen resource_cars, 	10, 11, 47
	EMReadScreen resource_rest, 	10, 12, 47
	EMReadScreen resource_other, 	10, 13, 47
	EMReadScreen resource_burial, 	10, 14, 47
	EMReadScreen resource_spon, 	10, 15, 47
	EMReadScreen resource_total, 	10, 17, 47
	EMReadScreen resource_max, 		10, 18, 47
	transmit

	resource_cash = trim(resource_cash)
	resource_acct = trim(resource_acct)
	resource_secu = trim(resource_secu)
	resource_cars = trim(resource_cars)
	resource_rest = trim(resource_rest)
	resource_other = trim(resource_other)
	resource_burial = trim(resource_burial)
	resource_spon = trim(resource_spon)
	resource_total = trim(resource_total)
	resource_max = trim(resource_max)

	transmit 		'FSB1
	EMReadScreen fsb1_gross_wages, 		10, 5, 31
	EMReadScreen fsb1_self_emp, 		10, 6, 31
	EMReadScreen fsb1_total_earned_inc, 10, 8, 31

	fsb1_gross_wages = trim(fsb1_gross_wages)
	fsb1_self_emp = trim(fsb1_self_emp)
	fsb1_total_earned_inc = trim(fsb1_total_earned_inc)


	EMReadScreen fsb1_pa_grant_inc, 	10, 10, 31
	EMReadScreen fsb1_rsdi_inc, 		10, 11, 31
	EMReadScreen fsb1_ssi_inc, 			10, 12, 31
	EMReadScreen fsb1_va_inc, 			10, 13, 31
	EMReadScreen fsb1_uc_wc_inc, 		10, 14, 31
	EMReadScreen fsb1_cses_inc, 		10, 15, 31
	EMReadScreen fsb1_other_unea_inc, 	10, 16, 31
	EMReadScreen fsb1_total_unea_inc, 	10, 18, 31

	fsb1_pa_grant_inc = trim(fsb1_pa_grant_inc)
	fsb1_rsdi_inc = trim(fsb1_rsdi_inc)
	fsb1_ssi_inc = trim(fsb1_ssi_inc)
	fsb1_va_inc = trim(fsb1_va_inc)
	fsb1_uc_wc_inc = trim(fsb1_uc_wc_inc)
	fsb1_cses_inc = trim(fsb1_cses_inc)
	fsb1_other_unea_inc = trim(fsb1_other_unea_inc)
	fsb1_total_unea_inc = trim(fsb1_total_unea_inc)


	EMReadScreen fsb1_schl_inc, 			10, 5, 71
	EMReadScreen fsb1_farm_ofset, 			10, 6, 71
	EMReadScreen fsb1_total_gross_inc, 		10, 7, 71
	EMReadScreen fsb1_max_gross_inc, 		10, 8, 71

	EMReadScreen fsb1_deduct_standard, 		10, 10, 71
	EMReadScreen fsb1_deduct_earned, 		10, 11, 71
	EMReadScreen fsb1_deduct_medical, 		10, 12, 71
	EMReadScreen fsb1_deduct_depndt_care, 	10, 13, 71
	EMReadScreen fsb1_cses, 				10, 14, 71
	EMReadScreen fsb1_total_deduct, 		10, 16, 71

	EMReadScreen fsb1_net_inc, 				10, 18, 71

	fsb1_schl_inc = trim(fsb1_schl_inc)
	fsb1_farm_ofset = trim(fsb1_farm_ofset)
	fsb1_total_gross_inc = trim(fsb1_total_gross_inc)
	fsb1_max_gross_inc = trim(fsb1_max_gross_inc)
	fsb1_deduct_standard = trim(fsb1_deduct_standard)
	fsb1_deduct_earned = trim(fsb1_deduct_earned)
	fsb1_deduct_medical = trim(fsb1_deduct_medical)
	fsb1_deduct_depndt_care = trim(fsb1_deduct_depndt_care)
	fsb1_cses = trim(fsb1_cses)
	fsb1_total_deduct = trim(fsb1_total_deduct)
	fsb1_net_inc = trim(fsb1_net_inc)


	transmit 		'FSB2
	EMReadScreen fsb2_shel_rent_mort, 		10, 5, 27
	EMReadScreen fsb2_shel_prop_tax, 		10, 6, 27
	EMReadScreen fsb2_shel_home_ins, 		10, 7, 27
	EMReadScreen fsb2_shel_electricity, 	10, 8, 27
	EMReadScreen fsb2_shel_heat_ac, 		10, 9, 27
	EMReadScreen fsb2_shel_water_garbage, 	10, 10, 27
	EMReadScreen fsb2_shel_phone, 			10, 11, 27
	EMReadScreen fsb2_shel_other, 			10, 12, 27
	EMReadScreen fsb2_shel_total, 			10, 14, 27
	EMReadScreen fsb2_50_perc_net_inc, 		10, 15, 27
	EMReadScreen fsb2_adj_shel_costs, 		10, 17, 27

	fsb2_shel_rent_mort = trim(fsb2_shel_rent_mort)
	fsb2_shel_prop_tax = trim(fsb2_shel_prop_tax)
	fsb2_shel_home_ins = trim(fsb2_shel_home_ins)
	fsb2_shel_electricity = trim(fsb2_shel_electricity)
	fsb2_shel_heat_ac = trim(fsb2_shel_heat_ac)
	fsb2_shel_water_garbage = trim(fsb2_shel_water_garbage)
	fsb2_shel_phone = trim(fsb2_shel_phone)
	fsb2_shel_other = trim(fsb2_shel_other)
	fsb2_shel_total = trim(fsb2_shel_total)
	fsb2_50_perc_net_inc = trim(fsb2_50_perc_net_inc)
	fsb2_adj_shel_costs = trim(fsb2_adj_shel_costs)


	EMReadScreen fsb2_max_allow_shel, 		10, 5, 71
	EMReadScreen fsb2_shel_expenses, 		10, 6, 71
	EMReadScreen fsb2_net_adj_inc, 			10, 7, 71
	' EMReadScreen fsb2_max_net_adj_inc, 		10, 8, 71
	EMReadScreen fsb2_monthly_fs_allot, 	10, 10, 71
	EMReadScreen fsb2_drug_felon_sanc_amt, 	10, 12, 71
	EMReadScreen fsb2_recoup_amount, 		10, 14, 71
	EMReadScreen fsb2_benefit_amount, 		10, 16, 71
	EMReadScreen fsb2_state_food_amt, 		10, 17, 71
	EMReadScreen fsb2_fed_food_amt, 		10, 18, 71

	fsb2_max_allow_shel = trim(fsb2_max_allow_shel)
	fsb2_shel_expenses = trim(fsb2_shel_expenses)
	fsb2_net_adj_inc = trim(fsb2_net_adj_inc)
	' fsb2_max_net_adj_inc = trim(fsb2_max_net_adj_inc)
	fsb2_monthly_fs_allot = trim(fsb2_monthly_fs_allot)
	fsb2_drug_felon_sanc_amt = trim(fsb2_drug_felon_sanc_amt)
	fsb2_recoup_amount = trim(fsb2_recoup_amount)
	fsb2_benefit_amount = trim(fsb2_benefit_amount)
	fsb2_state_food_amt = trim(fsb2_state_food_amt)
	fsb2_fed_food_amt = trim(fsb2_fed_food_amt)


	Call write_value_and_transmit("X", 14, 4)		''Resource Detail
	row = 8
	Do
		EMReadScreen ref_numb, 2, row, 12

		For case_memb = 0 to UBound(MEMBER_ARRAY, 2)

			If MEMBER_ARRAY(ref_numb_const, case_memb) = ref_numb Then
				EMReadScreen memb_code, 1, row, 51
				EMReadScreen memb_drug_felon_test, 6, row, 64

				MEMBER_ARRAY(memb_drug_felon_test_const, case_memb) = trim(memb_drug_felon_test)
			End If
		Next

		row = row + 1
		EMReadScreen next_ref_numb, 2, row, 12
	Loop until next_ref_numb = "  "
	transmit

	Call write_value_and_transmit("X", 14, 4)		''Resource Detail
	EMReadScreen recoup_from_fed_fs, 10, 5, 51
	EMReadScreen recoup_from_state_fs, 10, 7, 51

	recoup_from_fed_fs = trim(recoup_from_fed_fs)
	recoup_from_state_fs = trim(recoup_from_state_fs)

	transmit

	transmit 		'FSSM
	EMReadScreen fssm_approved_date, 		8, 3, 14
	EMReadScreen fssm_date_last_approval, 	8, 5, 31
	EMReadScreen fssm_curr_prog_status, 	10, 6, 31
	EMReadScreen fssm_elig_result, 			10, 7, 31
	EMReadScreen fssm_reporting_status, 	12, 8, 31
	EMReadScreen fssm_info_source, 			4, 9, 31
	EMReadScreen fssm_benefit, 				12, 10, 31
	EMReadScreen fssm_elig_revw_date, 		8, 11, 31
	EMReadScreen fssm_budget_cycle, 		5, 12, 31
	EMReadScreen fssm_numb_in_assist_unit, 	2, 13, 31

	EMReadScreen fssm_total_resources, 		10, 5, 71
	EMReadScreen fssm_max_resources, 		10, 6, 71
	EMReadScreen fssm_net_adj_inc, 			10, 7, 71
	EMReadScreen fssm_monthly_fs_allotment, 10, 8, 71
	EMReadScreen fssm_prorated_amt, 		10, 9, 71
	EMReadScreen fssm_prorated_date,		8, 9, 58
	EMReadScreen fssm_benefit_amt, 			10, 13, 71

	fssm_approved_date = trim(fssm_approved_date)
	fssm_date_last_approval = trim(fssm_date_last_approval)
	fssm_curr_prog_status = trim(fssm_curr_prog_status)
	fssm_elig_result = trim(fssm_elig_result)
	fssm_reporting_status = trim(fssm_reporting_status)
	fssm_info_source = trim(fssm_info_source)
	fssm_benefit = trim(fssm_benefit)
	fssm_elig_revw_date = trim(fssm_elig_revw_date)
	fssm_budget_cycle = trim(fssm_budget_cycle)
	fssm_numb_in_assist_unit = trim(fssm_numb_in_assist_unit)
	fssm_total_resources = trim(fssm_total_resources)
	fssm_max_resources = trim(fssm_max_resources)
	fssm_net_adj_inc = trim(fssm_net_adj_inc)
	fssm_monthly_fs_allotment = trim(fssm_monthly_fs_allotment)
	fssm_prorated_amt = trim(fssm_prorated_amt)
	fssm_prorated_date = trim(fssm_prorated_date)
	fssm_benefit_amt = trim(fssm_benefit_amt)


	EMReadScreen fssm_expedited_info_exists, 16, 14, 44
	If fssm_expedited_info_exists = "EXPEDITED STATUS" Then
		Call write_value_and_transmit("X", 14, 72)		''Resource Detail
		EMReadScreen exp_status_issuance_on_or_before_15th, 1, 3, 5
		EMReadScreen exp_status_issuance_after_15th, 1, 5, 5
		EMReadScreen exp_status_issuance_app_month_fs_denial, 1, 9, 5

		EMReadScreen exp_criteria_migrant_destitute, 1, 15, 5
		EMReadScreen exp_criteria_resource_100_income_150, 1, 16, 5
		EMReadScreen exp_criteria_resource_income_less_shelter, 1, 19, 5

		EMReadScreen exp_verif_status_postponed, 1, 15, 52
		EMReadScreen exp_verif_status_out_of_state, 1, 17, 52
		EMReadScreen exp_verif_status_all_provided, 1, 19, 52
		transmit

		If exp_status_issuance_on_or_before_15th = "X" Then exp_package_includes_month_one = True
		If exp_status_issuance_after_15th = "X" Then
			exp_package_includes_month_one = True
			exp_package_includes_month_two = True
		End If
		If exp_status_issuance_app_month_fs_denial = "X" Then exp_package_includes_month_two = True

		If exp_criteria_migrant_destitute = "X" Then exp_criteria_migrant_destitute = True
		If exp_criteria_migrant_destitute = "_" Then exp_criteria_migrant_destitute = False
		If exp_criteria_resource_100_income_150 = "X" Then exp_criteria_resource_100_income_150 = True
		If exp_criteria_resource_100_income_150 = "_" Then exp_criteria_resource_100_income_150 = False
		If exp_criteria_resource_income_less_shelter = "X" Then exp_criteria_resource_income_less_shelter = True
		If exp_criteria_resource_income_less_shelter = "_" Then exp_criteria_resource_income_less_shelter = False

		If exp_verif_status_postponed = "X" Then exp_verif_status_postponed = True
		If exp_verif_status_postponed = "_" Then exp_verif_status_postponed = False
		If exp_verif_status_out_of_state = "X" Then exp_verif_status_out_of_state = True
		If exp_verif_status_out_of_state = "_" Then exp_verif_status_out_of_state = False
		If exp_verif_status_all_provided = "X" Then exp_verif_status_all_provided = True
		If exp_verif_status_all_provided = "_" Then exp_verif_status_all_provided = False


	End If

	EMReadScreen fssm_worker_message_one, 80, 17, 1
	EMReadScreen fssm_worker_message_two, 80, 18, 1

	fssm_worker_message_one = trim(fssm_worker_message_one)
	fssm_worker_message_two = trim(fssm_worker_message_two)

	MsgBox "fssm_benefit_amt - " & fssm_benefit_amt

	Call Back_to_SELF
end Function


function read_MFIP_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "MFIP")
	Call find_last_approved_ELIG_version(20, 79, version_number, version_date, elig_result)




	row = 7
	Do
		EMReadScreen ref_numb, 2, row, 6
		If ref_numb <> "  " Then
			EMReadScreen request_yn, 1, row, 32
			EMReadScreen memb_code, 1, row, 36
			EMReadScreen memb_count, 11, row, 41
			EMReadScreen memb_elig, 10, row, 53
			EMReadScreen memb_begin_date, 8, row, 67
			EMReadScreen memb_budg_cycle, 1, row, 78

			Call write_value_and_transmit("X", row, 3)
			EMReadScreen memb_absence, 			6, 7, 17
			EMReadScreen memb_child_age, 		6, 8, 17
			EMReadScreen memb_citizenship, 		6, 9, 17
			EMReadScreen memb_citizenship_ver, 	6, 10, 17
			EMReadScreen memb_dupl_assist, 		6, 11, 17
			EMReadScreen memb_fost_care, 		6, 12, 17
			EMReadScreen memb_fraud, 			6, 13, 17
			EMReadScreen memb_disq, 			6, 17, 17

			EMReadScreen memb_minor_living, 	6, 7, 52
			EMReadScreen memb_post_60, 			6, 8, 52
			EMReadScreen memb_ssi, 				6, 9, 52
			EMReadScreen memb_ssn_coop, 		6, 10, 52
			EMReadScreen memb_unit_memb, 		6, 11, 52
			EMReadScreen memb_unlawful_conduct, 6, 12, 52
			EMReadScreen memb_fs_recvd, 		6, 17, 52
			transmit

			Call write_value_and_transmit("X", row, 64)
			EMReadScreen memb_es_status_code, 2, 9, 22
			EMReadScreen memb_es_status_info, 30, 9, 25
			transmit

		End If

		row = row + 1
	Loop until ref_numb = "  "


	Call Back_to_SELF
end Function


function read_DWP_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "DWP ")



	Call Back_to_SELF
end Function


function read_GA_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "GA  ")



	Call Back_to_SELF
end Function

function read_MSA_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "MSA ")



	Call Back_to_SELF
end Function

function read_GRH_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "GRH ")



	Call Back_to_SELF
end Function

function read_MA_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "HC  ")



	Call Back_to_SELF
end Function

function read_MSP_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "HC  ")



	Call Back_to_SELF
end Function

function read_EMER_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "EMER")



	Call Back_to_SELF
end Function

function read_CASH_elig(footer_month, footer_year)
	call navigate_to_MAXIS_screen("ELIG", "DENY")



	Call Back_to_SELF
end Function
















'DECLARATIONS===============================================================================================================

class snap_eligibility_detail

	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public snap_expedited
	public snap_exp_package_includes_month_one
	public snap_exp_package_includes_month_two
	public snap_prorated
	public snap_earned_income_budgeted
	public snap_unearned_income_budgeted
	public snap_shel_costs_budgeted
	public snap_hest_costs_budgeted
	public snap_categorical_eligibility
	public snap_case_appl_withdrawn_test
	public snap_case_applct_elig_test
	public snap_case_comdty_test
	public snap_case_disq_test
	public snap_case_dupl_assist_test
	public snap_case_eligible_person_test
	public snap_case_fail_coop_test
	public snap_case_fail_file_test
	public snap_case_prosp_gross_inc_test
	public snap_case_prosp_net_inc_test
	public snap_case_recert_test
	public snap_case_residence_test
	public snap_case_resource_test
	public snap_case_retro_gross_inc_test
	public snap_case_retro_net_inc_test
	public snap_case_strike_test
	public snap_case_xfer_resource_inc_test
	public snap_case_verif_test
	public snap_case_voltry_quit_test
	public snap_case_work_reg_test
	public snap_fail_file_hrf
	public snap_fail_file_sr
	public snap_resource_cash
	public snap_resource_acct
	public snap_resource_secu
	public snap_resource_cars
	public snap_resource_rest
	public snap_resource_other
	public snap_resource_burial
	public snap_resource_spon
	public snap_resource_total
	public snap_resource_max
	public snap_budg_gross_wages
	public snap_budg_self_emp
	public snap_budg_total_earned_inc
	public snap_budg_pa_grant_inc
	public snap_budg_rsdi_inc
	public snap_budg_ssi_inc
	public snap_budg_va_inc
	public snap_budg_uc_wc_inc
	public snap_budg_cses_inc
	public snap_budg_other_unea_inc
	public snap_budg_total_unea_inc
	public snap_budg_schl_inc
	public snap_budg_farm_ofset
	public snap_budg_total_gross_inc
	public snap_budg_max_gross_inc
	public snap_budg_deduct_standard
	public snap_budg_deduct_earned
	public snap_budg_deduct_medical
	public snap_budg_deduct_depndt_care
	public snap_budg_cses
	public snap_budg_total_deduct
	public snap_budg_net_inc
	public snap_budg_shel_rent_mort
	public snap_budg_shel_prop_tax
	public snap_budg_shel_home_ins
	public snap_budg_shel_electricity
	public snap_budg_shel_heat_ac
	public snap_budg_shel_water_garbage
	public snap_budg_shel_phone
	public snap_budg_shel_other
	public snap_budg_shel_total
	public snap_budg_50_perc_net_inc
	public snap_budg_adj_shel_costs
	public snap_budg_max_allow_shel
	public snap_budg_shel_expenses
	' public snap_budg_net_adj_inc
	public snap_budg_max_net_adj_inc
	public snap_benefit_monthly_fs_allot
	public snap_benefit_drug_felon_sanc_amt
	public snap_benefit_recoup_amount
	public snap_benefit_benefit_amount
	public snap_benefit_state_food_amt
	public snap_benefit_fed_food_amt
	public snap_benefit_recoup_from_fed_fs
	public snap_benefit_recoup_from_state_fs
	public snap_approved_date
	public snap_date_last_approval
	public snap_curr_prog_status
	public snap_elig_result
	public snap_reporting_status
	public snap_info_source
	public snap_benefit
	public snap_elig_revw_date
	public snap_budget_cycle
	public snap_budg_numb_in_assist_unit
	public snap_budg_total_resources
	public snap_budg_max_resources
	public snap_budg_net_adj_inc
	public snap_benefit_monthly_fs_allotment
	public snap_benefit_prorated_amt
	public snap_benefit_prorated_date
	public snap_benefit_amt
	public snap_exp_criteria_migrant_destitute
	public snap_exp_criteria_resource_100_income_150
	public snap_exp_criteria_resource_income_less_shelter
	public snap_exp_verif_status_postponed
	public snap_exp_verif_status_out_of_state
	public snap_exp_verif_status_all_provided
	public snap_elig_worker_message_one
	public snap_elig_worker_message_two


end class

'Constants
const ref_numb_const				= 0

const access_denied					= 1
const full_name_const				= 2
const last_name_const				= 3
const first_name_const				= 4
const mid_initial					= 5
const other_names					= 6
const age							= 7
const date_of_birth					= 8
const ssn							= 9
const ssn_verif						= 10
const birthdate_verif				= 11


const fs_request_yn_const			= 12
const fs_memb_code_const			= 13
const fs_memb_status_info_const		= 14
const fs_memb_counted_const			= 15
const fs_memb_state_food_const		= 16
const fs_memb_elig_status_const		= 17
const fs_memb_begin_date_const		= 18
const fs_memb_budg_cycle_const		= 19
const fs_memb_abawd_const			= 20
const fs_memb_absence_const			= 21
const fs_memb_roomer_const			= 22
const fs_memb_boarder_const			= 23
const fs_memb_citizenship_const		= 24
const fs_memb_citizenship_coop_const = 25
const fs_memb_cmdty_const			= 26
const fs_memb_disq_const			= 27
const fs_memb_dupl_assist_const		= 28
const fs_memb_fraud_const			= 29
const fs_memb_eligible_student_const = 30
const fs_memb_institution_const		= 31
const fs_memb_mfip_elig_const		= 32
const fs_memb_non_applcnt_const		= 33
const fs_memb_residence_const		= 34
const fs_memb_ssn_coop_const		= 35
const fs_memb_unit_memb_const		= 36
const fs_memb_work_reg_const		= 37
const fs_memb_drug_felon_test_const	= 38

const last_const = 50

'Arrays
Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(last_const, 0)

Dim SNAP_ELIG_APPROVALS()
ReDim SNAP_ELIG_APPROVALS(0)

'===========================================================================================================================
EMConnect ""
Call check_for_MAXIS(True)

Call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog Dialog1, 0, 0, 366, 85, "Eligibility Summary Case Number Dialog"
  EditBox 65, 10, 60, 15, MAXIS_case_number
  EditBox 90, 30, 15, 15, first_footer_month
  EditBox 110, 30, 15, 15, first_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 65, 50, 15
    CancelButton 80, 65, 50, 15
  Text 15, 15, 50, 10, "Case Number"
  Text 20, 35, 65, 10, "First month of APP"
  Text 95, 45, 35, 10, "MM    YY"
  Text 160, 10, 140, 20, "This script will detail information about all APP actions for a this case taken today."
  Text 165, 30, 185, 10, "- Script will handle for approvals, denials, and closures."
  Text 165, 40, 155, 10, "- Script will handle for any program in MAXIS."
  Text 165, 50, 180, 10, "- To be handled by the script ELIG resulsts must be:"
  Text 185, 60, 75, 10, "CREATED Today"
  Text 185, 70, 75, 10, "APPROVED Today"
  ButtonGroup ButtonPressed
    PushButton 255, 65, 105, 15, "Script Instructions", intructions_btn
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1

		cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(first_footer_month, first_footer_year, err_msg, "*")

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

Call date_array_generator(first_footer_month, first_footer_year, MONTHS_ARRAY)

snap_elig_months_count = 0

For each footer_month in MONTHS_ARRAY
	' MsgBox footer_month
	Call convert_date_into_MAXIS_footer_month(footer_month, MAXIS_footer_month, MAXIS_footer_year)

	Call Navigate_to_MAXIS_screen("ELIG", "SUMM")

	EMReadScreen numb_SNAP_versions, 1, 17, 40
	' MsgBox "numb_SNAP_versions - " & numb_SNAP_versions
	'TODO MAKE THIS READ THE DATE AND COMPARE TO TODAY
	If numb_SNAP_versions <> " " Then
		ReDim Preserve SNAP_ELIG_APPROVALS(snap_elig_months_count)
		Set SNAP_ELIG_APPROVALS(snap_elig_months_count) = new snap_eligibility_detail

		SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month = MAXIS_footer_month
		SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_year = MAXIS_footer_year

		' MsgBox "SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month - " & SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month

		snap_elig_months_count = snap_elig_months_count + 1
	End If

	Call back_to_SELF
Next


For snap_approval = 0 to UBound(SNAP_ELIG_APPROVALS)
	' MsgBox "snap_approval - " & snap_approval
	MsgBox SNAP_ELIG_APPROVALS(snap_approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(snap_approval).elig_footer_year

	footer_month = SNAP_ELIG_APPROVALS(snap_approval).elig_footer_month
	footer_year = SNAP_ELIG_APPROVALS(snap_approval).elig_footer_year

	Call read_SNAP_elig(footer_month, footer_year, var_02_test, var_03_test, var_04_test, var_05_test, var_06_test, var_07_test, var_08_test, var_09_test, var_10_test, var_11_test, var_12_test, var_13_test, var_14_test, var_15_test, var_16_test, var_17_test, var_18_test, var_19_test, var_20_test, var_21_test, var_22_test, var_23_test, var_24_test, var_25_test, var_26_test, var_27_test, var_28_test, var_29_test, var_30_test, var_31_test, var_32_test, var_33_test, var_34_test, var_35_test, var_36_test, var_37_test, var_38_test, var_39_test, var_40_test, var_41_test, var_42_test, var_43_test, var_44_test, var_45_test, var_46_test, var_47_test, var_48_test, var_49_test, var_50_test, var_51_test, var_52_test, var_53_test, var_54_test, var_55_test, var_56_test, var_57_test, var_58_test, var_59_test, var_60_test, var_61_test, var_62_test, var_63_test, var_64_test, var_65_test, var_66_test, var_67_test, var_68_test, var_69_test, var_70_test, var_71_test, var_72_test, var_73_test, var_74_test, var_75_test, var_76_test, var_77_test, var_78_test, var_79_test, var_80_test, var_82_test, var_83_test, var_84_test, var_85_test, var_86_test, var_87_test, var_88_test, var_89_test, var_90_test, var_91_test, var_92_test, var_93_test, var_94_test, var_95_test, var_96_test, var_97_test, var_98_test, var_99_test, var_100_test, var_101_test, var_102_test, var_103_test, var_104_test, var_105_test, var_106_test, var_107_test, var_108_test, var_109_test, var_110_test, var_111_test, var_112_test, var_113_test, var_114_test, var_115_test, HH_MEMB_ARRAY, ref_numb_const, request_yn_const, memb_code_const, memb_status_info_const, memb_counted_const, memb_state_food_const, memb_elig_status_const, memb_begin_date_const, memb_budg_cycle_const, memb_abawd_const, memb_absence_const, memb_roomer_const, memb_boarder_const, memb_citizenship_const, memb_citizenship_coop_const, memb_cmdty_const, memb_disq_const, memb_dupl_assist_const, memb_fraud_const, memb_eligible_student_const, memb_institution_const, memb_mfip_elig_const, memb_non_applcnt_const, memb_residence_const, memb_ssn_coop_const, memb_unit_memb_const, memb_work_reg_const, memb_drug_felon_test_const)


	SNAP_ELIG_APPROVALS(snap_approval).elig_version_number = var_02_test
	SNAP_ELIG_APPROVALS(snap_approval).elig_version_date = var_03_test
	SNAP_ELIG_APPROVALS(snap_approval).elig_version_result = var_04_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_expedited = var_05_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_package_includes_month_one = var_06_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_package_includes_month_two = var_07_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_prorated = var_08_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_earned_income_budgeted = var_09_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_unearned_income_budgeted = var_10_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_shel_costs_budgeted = var_11_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_hest_costs_budgeted = var_12_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_categorical_eligibility = var_13_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_appl_withdrawn_test = var_14_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_applct_elig_test = var_15_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_comdty_test = var_16_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_disq_test = var_17_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_dupl_assist_test = var_18_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_eligible_person_test = var_19_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_fail_coop_test = var_20_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_fail_file_test = var_21_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_prosp_gross_inc_test = var_22_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_prosp_net_inc_test = var_23_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_recert_test = var_24_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_residence_test = var_25_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_resource_test = var_26_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_retro_gross_inc_test = var_27_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_retro_net_inc_test = var_28_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_strike_test = var_29_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_xfer_resource_inc_test = var_30_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_verif_test = var_31_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_voltry_quit_test = var_32_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_case_work_reg_test = var_33_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_fail_file_hrf = var_34_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_fail_file_sr = var_35_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_cash = var_36_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_acct = var_37_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_secu = var_38_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_cars = var_39_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_rest = var_40_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_other = var_41_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_burial = var_42_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_spon = var_43_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_total = var_44_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_resource_max = var_45_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_gross_wages = var_46_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_self_emp = var_47_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_total_earned_inc = var_48_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_pa_grant_inc = var_49_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_rsdi_inc = var_50_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_ssi_inc = var_51_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_va_inc = var_52_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_uc_wc_inc = var_53_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_cses_inc = var_54_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_other_unea_inc = var_55_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_total_unea_inc = var_56_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_schl_inc = var_57_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_farm_ofset = var_58_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_total_gross_inc = var_59_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_max_gross_inc = var_60_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_deduct_standard = var_61_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_deduct_earned = var_62_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_deduct_medical = var_63_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_deduct_depndt_care = var_64_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_cses = var_65_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_total_deduct = var_66_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_net_inc = var_67_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_rent_mort = var_68_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_prop_tax = var_69_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_home_ins = var_70_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_electricity = var_71_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_heat_ac = var_72_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_water_garbage = var_73_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_phone = var_74_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_other = var_75_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_total = var_76_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_50_perc_net_inc = var_77_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_adj_shel_costs = var_78_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_max_allow_shel = var_79_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_shel_expenses = var_80_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_max_net_adj_inc = 	var_82_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_monthly_fs_allot = var_83_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_drug_felon_sanc_amt = var_84_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_recoup_amount = var_85_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_benefit_amount = var_86_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_state_food_amt = var_87_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_fed_food_amt = var_88_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_recoup_from_fed_fs = var_89_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_recoup_from_state_fs = var_90_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_approved_date = var_91_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_date_last_approval = var_92_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_curr_prog_status = var_93_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_elig_result = var_94_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_reporting_status = var_95_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_info_source = var_96_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit = var_97_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_elig_revw_date = var_98_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budget_cycle = var_99_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_numb_in_assist_unit = var_100_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_total_resources = var_101_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_max_resources = var_102_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_budg_net_adj_inc = var_103_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_monthly_fs_allotment = var_104_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_prorated_amt = var_105_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_prorated_date = var_106_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_amt = var_107_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_criteria_migrant_destitute = var_108_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_criteria_resource_100_income_150 = var_109_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_criteria_resource_income_less_shelter = var_110_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_verif_status_postponed = var_111_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_verif_status_out_of_state = var_112_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_exp_verif_status_all_provided = var_113_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_elig_worker_message_one = var_114_test
	SNAP_ELIG_APPROVALS(snap_approval).snap_elig_worker_message_two = var_115_test

	MsgBox "SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_amt - " & SNAP_ELIG_APPROVALS(snap_approval).snap_benefit_amt


	' footer_month = SNAP_ELIG_APPROVALS(snap_approval).elig_footer_month
	' footer_year = SNAP_ELIG_APPROVALS(snap_approval).elig_footer_year
	' Call read_SNAP_elig(footer_month, footer_year, test_version_number, test_version_date, test_elig_result, test_fs_expedited, test_exp_package_includes_month_one, test_exp_package_includes_month_two, test_fs_prorated, test_earned_income_budgeted, test_unearned_income_budgeted, test_shel_costs_budgeted, test_hest_costs_budgeted, test_categorical_eligibility, test_case_appl_withdrawn, test_case_applct_elig, test_case_comdty, test_case_disq, test_case_dupl_assist, test_case_eligible_person)

	' Call read_SNAP_elig(footer_month, footer_year, test_version_number, test_version_date, test_elig_result, test_fs_expedited, test_exp_package_includes_month_one, test_exp_package_includes_month_two, test_fs_prorated, test_earned_income_budgeted, test_unearned_income_budgeted, test_shel_costs_budgeted, test_hest_costs_budgeted, test_categorical_eligibility, test_case_appl_withdrawn, test_case_applct_elig, test_case_comdty, test_case_disq, test_case_dupl_assist, test_case_eligible_person)
	' Call read_SNAP_elig(test_case_fail_coop, test_case_fail_file, test_case_prosp_gross_inc, test_case_prosp_net_inc, test_case_recert, test_case_residence, test_case_resource, test_case_retro_gross_inc, test_case_retro_net_inc, test_case_strike, test_case_xfer_resource_inc, test_case_verif, test_case_voltry_quit, test_case_work_reg, test_fail_file_hrf, test_fail_file_sr, test_resource_cash, test_resource_acct, test_resource_secu, test_resource_cars, test_resource_rest, test_resource_other)
	' Call read_SNAP_elig(test_resource_burial, test_resource_spon, test_resource_total, test_resource_max, test_fsb1_gross_wages, test_fsb1_self_emp, test_fsb1_total_earned_inc, test_fsb1_pa_grant_inc, test_fsb1_rsdi_inc, test_fsb1_ssi_inc, test_fsb1_va_inc, test_fsb1_uc_wc_inc, test_fsb1_cses_inc, test_fsb1_other_unea_inc, test_fsb1_total_unea_inc, test_fsb1_schl_inc, test_fsb1_farm_ofset, test_fsb1_total_gross_inc, test_fsb1_max_gross_inc, test_fsb1_deduct_standard, test_fsb1_deduct_earned)
	' Call read_SNAP_elig(test_fsb1_deduct_medical, test_fsb1_deduct_depndt_care, test_fsb1_cses, test_fsb1_total_deduct, test_fsb1_net_inc, test_fsb2_shel_rent_mort, test_fsb2_shel_prop_tax, test_fsb2_shel_home_ins, test_fsb2_shel_electricity, test_fsb2_shel_heat_ac, test_fsb2_shel_water_garbage, test_fsb2_shel_phone, test_fsb2_shel_other, test_fsb2_shel_total, test_fsb2_50_perc_net_inc, test_fsb2_adj_shel_costs, test_fsb2_max_allow_shel, test_fsb2_shel_expenses, test_fsb2_net_adj_inc)
	' Call read_SNAP_elig(test_fsb2_max_net_adj_inc, test_fsb2_monthly_fs_allot, test_fsb2_drug_felon_sanc_amt, test_fsb2_recoup_amount, test_fsb2_benefit_amount, test_fsb2_state_food_amt, test_fsb2_fed_food_amt, test_recoup_from_fed_fs, test_recoup_from_state_fs, test_fssm_approved_date, test_fssm_date_last_approval, test_fssm_curr_prog_status, test_fssm_elig_result, test_fssm_reporting_status, test_fssm_info_source, test_fssm_benefit, test_fssm_elig_revw_date, test_fssm_budget_cycle)
	' Call read_SNAP_elig(test_fssm_numb_in_assist_unit, test_fssm_total_resources, test_fssm_max_resources, test_fssm_net_adj_inc, test_fssm_monthly_fs_allotment, test_fssm_prorated_amt, test_fssm_prorated_date, test_fssm_benefit_amt, test_exp_criteria_migrant_destitute, test_exp_criteria_resource_100_income_150, test_exp_criteria_resource_income_less_shelter, test_exp_verif_status_postponed, test_exp_verif_status_out_of_state, test_exp_verif_status_all_provided, test_fssm_worker_message_one)
	' Call read_SNAP_elig(test_fssm_worker_message_two, HH_MEMB_ARRAY, ref_numb_const, fs_request_yn_const, fs_memb_code_const, fs_memb_status_info_const, fs_memb_counted_const, fs_memb_state_food_const, fs_memb_elig_status_const, fs_memb_begin_date_const, fs_memb_budg_cycle_const, fs_memb_abawd_const, fs_memb_absence_const, fs_memb_roomer_const, fs_memb_boarder_const, fs_memb_citizenship_const, fs_memb_citizenship_coop_const, fs_memb_cmdty_const, fs_memb_disq_const, fs_memb_dupl_assist_const)
	' Call read_SNAP_elig(fs_memb_fraud_const, fs_memb_eligible_student_const, fs_memb_institution_const, fs_memb_mfip_elig_const, fs_memb_non_applcnt_const, fs_memb_residence_const, fs_memb_ssn_coop_const, fs_memb_unit_memb_const, fs_memb_work_reg_const, fs_memb_drug_felon_test_const)



	' Call read_SNAP_elig(SNAP_ELIG_APPROVALS(snap_approval).elig_footer_month, SNAP_ELIG_APPROVALS(snap_approval).elig_footer_year, SNAP_ELIG_APPROVALS(snap_approval).version_number, SNAP_ELIG_APPROVALS(snap_approval).version_date, SNAP_ELIG_APPROVALS(snap_approval).elig_result, SNAP_ELIG_APPROVALS(snap_approval).fs_expedited, SNAP_ELIG_APPROVALS(snap_approval).exp_package_includes_month_one, SNAP_ELIG_APPROVALS(snap_approval).exp_package_includes_month_two,
	' SNAP_ELIG_APPROVALS(snap_approval).fs_prorated, SNAP_ELIG_APPROVALS(snap_approval).earned_income_budgeted, SNAP_ELIG_APPROVALS(snap_approval).unearned_income_budgeted, SNAP_ELIG_APPROVALS(snap_approval).shel_costs_budgeted, SNAP_ELIG_APPROVALS(snap_approval).hest_costs_budgeted, SNAP_ELIG_APPROVALS(snap_approval).categorical_eligibility, SNAP_ELIG_APPROVALS(snap_approval).case_appl_withdrawn, SNAP_ELIG_APPROVALS(snap_approval).case_applct_elig,
	' SNAP_ELIG_APPROVALS(snap_approval).case_comdty, SNAP_ELIG_APPROVALS(snap_approval).case_disq, SNAP_ELIG_APPROVALS(snap_approval).case_dupl_assist, SNAP_ELIG_APPROVALS(snap_approval).case_eligible_person, SNAP_ELIG_APPROVALS(snap_approval).case_fail_coop, SNAP_ELIG_APPROVALS(snap_approval).case_fail_file, SNAP_ELIG_APPROVALS(snap_approval).case_prosp_gross_inc, SNAP_ELIG_APPROVALS(snap_approval).case_prosp_net_inc, SNAP_ELIG_APPROVALS(snap_approval).case_recert,
	' SNAP_ELIG_APPROVALS(snap_approval).case_residence, SNAP_ELIG_APPROVALS(snap_approval).case_resource, SNAP_ELIG_APPROVALS(snap_approval).case_retro_gross_inc, SNAP_ELIG_APPROVALS(snap_approval).case_retro_net_inc, SNAP_ELIG_APPROVALS(snap_approval).case_strike, SNAP_ELIG_APPROVALS(snap_approval).case_xfer_resource_inc, SNAP_ELIG_APPROVALS(snap_approval).case_verif, SNAP_ELIG_APPROVALS(snap_approval).case_voltry_quit, SNAP_ELIG_APPROVALS(snap_approval).case_work_reg,
	' SNAP_ELIG_APPROVALS(snap_approval).fail_file_hrf, SNAP_ELIG_APPROVALS(snap_approval).fail_file_sr, SNAP_ELIG_APPROVALS(snap_approval).resource_cash, SNAP_ELIG_APPROVALS(snap_approval).resource_acct, SNAP_ELIG_APPROVALS(snap_approval).resource_secu, SNAP_ELIG_APPROVALS(snap_approval).resource_cars, SNAP_ELIG_APPROVALS(snap_approval).resource_rest, SNAP_ELIG_APPROVALS(snap_approval).resource_other, SNAP_ELIG_APPROVALS(snap_approval).resource_burial,
	' SNAP_ELIG_APPROVALS(snap_approval).resource_spon, SNAP_ELIG_APPROVALS(snap_approval).resource_total, SNAP_ELIG_APPROVALS(snap_approval).resource_max, SNAP_ELIG_APPROVALS(snap_approval).fsb1_gross_wages, SNAP_ELIG_APPROVALS(snap_approval).fsb1_self_emp, SNAP_ELIG_APPROVALS(snap_approval).fsb1_total_earned_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_pa_grant_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_rsdi_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_ssi_inc,
	' SNAP_ELIG_APPROVALS(snap_approval).fsb1_va_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_uc_wc_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_cses_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_other_unea_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_total_unea_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_schl_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_farm_ofset, SNAP_ELIG_APPROVALS(snap_approval).fsb1_total_gross_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb1_max_gross_inc,
	' SNAP_ELIG_APPROVALS(snap_approval).fsb1_deduct_standard, SNAP_ELIG_APPROVALS(snap_approval).fsb1_deduct_earned, SNAP_ELIG_APPROVALS(snap_approval).fsb1_deduct_medical, SNAP_ELIG_APPROVALS(snap_approval).fsb1_deduct_depndt_care, SNAP_ELIG_APPROVALS(snap_approval).fsb1_cses, SNAP_ELIG_APPROVALS(snap_approval).fsb1_total_deduct, SNAP_ELIG_APPROVALS(snap_approval).fsb1_net_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_rent_mort, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_prop_tax,
	' SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_home_ins, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_electricity, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_heat_ac, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_water_garbage, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_phone, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_other, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_total, SNAP_ELIG_APPROVALS(snap_approval).fsb2_50_perc_net_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb2_adj_shel_costs,
	' SNAP_ELIG_APPROVALS(snap_approval).fsb2_max_allow_shel, SNAP_ELIG_APPROVALS(snap_approval).fsb2_shel_expenses, SNAP_ELIG_APPROVALS(snap_approval).fsb2_net_adj_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb2_max_net_adj_inc, SNAP_ELIG_APPROVALS(snap_approval).fsb2_monthly_fs_allot, SNAP_ELIG_APPROVALS(snap_approval).fsb2_drug_felon_sanc_amt, SNAP_ELIG_APPROVALS(snap_approval).fsb2_recoup_amount, SNAP_ELIG_APPROVALS(snap_approval).fsb2_benefit_amount,
	' SNAP_ELIG_APPROVALS(snap_approval).fsb2_state_food_amt, SNAP_ELIG_APPROVALS(snap_approval).fsb2_fed_food_amt, SNAP_ELIG_APPROVALS(snap_approval).recoup_from_fed_fs, SNAP_ELIG_APPROVALS(snap_approval).recoup_from_state_fs, SNAP_ELIG_APPROVALS(snap_approval).fssm_approved_date, SNAP_ELIG_APPROVALS(snap_approval).fssm_date_last_approval, SNAP_ELIG_APPROVALS(snap_approval).fssm_curr_prog_status, SNAP_ELIG_APPROVALS(snap_approval).fssm_elig_result,
	' SNAP_ELIG_APPROVALS(snap_approval).fssm_reporting_status, SNAP_ELIG_APPROVALS(snap_approval).fssm_info_source, SNAP_ELIG_APPROVALS(snap_approval).fssm_benefit, SNAP_ELIG_APPROVALS(snap_approval).fssm_elig_revw_date, SNAP_ELIG_APPROVALS(snap_approval).fssm_budget_cycle, SNAP_ELIG_APPROVALS(snap_approval).fssm_numb_in_assist_unit, SNAP_ELIG_APPROVALS(snap_approval).fssm_total_resources, SNAP_ELIG_APPROVALS(snap_approval).fssm_max_resources, SNAP_ELIG_APPROVALS(snap_approval).fssm_net_adj_inc,
	' SNAP_ELIG_APPROVALS(snap_approval).fssm_monthly_fs_allotment, SNAP_ELIG_APPROVALS(snap_approval).fssm_prorated_amt, SNAP_ELIG_APPROVALS(snap_approval).fssm_prorated_date, SNAP_ELIG_APPROVALS(snap_approval).fssm_benefit_amt, SNAP_ELIG_APPROVALS(snap_approval).exp_criteria_migrant_destitute, SNAP_ELIG_APPROVALS(snap_approval).exp_criteria_resource_100_income_150, SNAP_ELIG_APPROVALS(snap_approval).exp_criteria_resource_income_less_shelter,
	' SNAP_ELIG_APPROVALS(snap_approval).exp_verif_status_postponed, SNAP_ELIG_APPROVALS(snap_approval).exp_verif_status_out_of_state, SNAP_ELIG_APPROVALS(snap_approval).exp_verif_status_all_provided, SNAP_ELIG_APPROVALS(snap_approval).fssm_worker_message_one, SNAP_ELIG_APPROVALS(snap_approval).fssm_worker_message_two, HH_MEMB_ARRAY, ref_numb_const, fs_request_yn_const, fs_memb_code_const, fs_memb_status_info_const, fs_memb_counted_const, fs_memb_state_food_const, fs_memb_elig_status_const,
	' fs_memb_begin_date_const, fs_memb_budg_cycle_const, fs_memb_abawd_const, fs_memb_absence_const, fs_memb_roomer_const, fs_memb_boarder_const, fs_memb_citizenship_const, fs_memb_citizenship_coop_const, fs_memb_cmdty_const, fs_memb_disq_const, fs_memb_dupl_assist_const, fs_memb_fraud_const, fs_memb_eligible_student_const, fs_memb_institution_const, fs_memb_mfip_elig_const, fs_memb_non_applcnt_const, fs_memb_residence_const, fs_memb_ssn_coop_const, fs_memb_unit_memb_const,
	' fs_memb_work_reg_const, fs_memb_drug_felon_test_const)




Next

MsgBox "PAUSE"

HH_MEMB_ARRAY
ref_numb_const
fs_request_yn_const
fs_memb_code_const
fs_memb_status_info_const
fs_memb_counted_const
fs_memb_state_food_const
fs_memb_elig_status_const
fs_memb_begin_date_const
fs_memb_budg_cycle_const
fs_memb_abawd_const
fs_memb_absence_const
fs_memb_roomer_const
fs_memb_boarder_const
fs_memb_citizenship_const
fs_memb_citizenship_coop_const
fs_memb_cmdty_const
fs_memb_disq_const
fs_memb_dupl_assist_const
fs_memb_fraud_const
fs_memb_eligible_student_const
fs_memb_institution_const
fs_memb_mfip_elig_const
fs_memb_non_applcnt_const
fs_memb_residence_const
fs_memb_ssn_coop_const
fs_memb_unit_memb_const
fs_memb_work_reg_const
fs_memb_drug_felon_test_const



MEMBER_ARRAY
ref_numb_const
request_yn_const
memb_code_const
memb_status_info_const
memb_counted_const
memb_state_food_const
memb_elig_status_const
memb_begin_date_const
memb_budg_cycle_const
memb_abawd_const
memb_absence_const
memb_roomer_const
memb_boarder_const
memb_citizenship_const
memb_citizenship_coop_const
memb_cmdty_const
memb_disq_const
memb_dupl_assist_const
memb_fraud_const
memb_eligible_student_const
memb_institution_const
memb_mfip_elig_const
memb_non_applcnt_const
memb_residence_const
memb_ssn_coop_const
memb_unit_memb_const
memb_work_reg_const
memb_drug_felon_test_const



















'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
