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
		EMReadScreen next_ref_numb, 2, row, 6
	Loop until next_ref_numb = "  "


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

	public snap_elig_ref_numbs()
	public snap_elig_membs_request_yn()
	public snap_elig_membs_code()
	public snap_elig_membs_status_info()
	public snap_elig_membs_counted()
	public snap_elig_membs_state_food()
	public snap_elig_membs_eligibility()
	public snap_elig_membs_begin_date()
	public snap_elig_membs_budget_cycle()

	public snap_elig_membs_abawd()
	public snap_elig_membs_absence()
	public snap_elig_membs_roomer()
	public snap_elig_membs_boarder()
	public snap_elig_membs_citizenship()
	public snap_elig_membs_citizenship_code()
	public snap_elig_membs_cmdty()
	public snap_elig_membs_disq()
	public snap_elig_membs_dupl_assist()
	public snap_elig_membs_fraud()
	public snap_elig_membs_eligible_student()
	public snap_elig_membs_institution()
	public snap_elig_membs_mfip_elig()
	public snap_elig_membs_non_applcnt()
	public snap_elig_membs_residence()
	public snap_elig_membs_ssn_coop()
	public snap_elig_membs_unit_memb()
	public snap_elig_membs_work_reg()
	public snap_elig_membs_drug_felon_test()
	' public snap_elig_membs
	' public snap_elig_membs
	' public snap_elig_membs
	' public snap_elig_membs
	' public snap_elig_membs

	public snap_expedited
	public snap_uhfs
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
	public snap_budg_deduct_cses
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


	public sub read_elig()
		snap_expedited = False
		snap_uhfs = False
		snap_exp_package_includes_month_one = False
		snap_exp_package_includes_month_two = False
		snap_prorated = False
		snap_earned_income_budgeted = False
		snap_unearned_income_budgeted = False
		snap_shel_costs_budgeted = False
		snap_hest_costs_budgeted = False
		snap_categorical_eligibility = ""

		ReDim snap_elig_ref_numbs(0)
		ReDim snap_elig_membs_request_yn(0)
		ReDim snap_elig_membs_code(0)
		ReDim snap_elig_membs_status_info(0)
		ReDim snap_elig_membs_counted(0)
		ReDim snap_elig_membs_state_food(0)
		ReDim snap_elig_membs_eligibility(0)
		ReDim snap_elig_membs_begin_date(0)
		ReDim snap_elig_membs_budget_cycle(0)
		ReDim snap_elig_membs_abawd(0)
		ReDim snap_elig_membs_absence(0)
		ReDim snap_elig_membs_roomer(0)
		ReDim snap_elig_membs_boarder(0)
		ReDim snap_elig_membs_citizenship(0)
		ReDim snap_elig_membs_citizenship_code(0)
		ReDim snap_elig_membs_cmdty(0)
		ReDim snap_elig_membs_disq(0)
		ReDim snap_elig_membs_dupl_assist(0)
		ReDim snap_elig_membs_fraud(0)
		ReDim snap_elig_membs_eligible_student(0)
		ReDim snap_elig_membs_institution(0)
		ReDim snap_elig_membs_mfip_elig(0)
		ReDim snap_elig_membs_non_applcnt(0)
		ReDim snap_elig_membs_residence(0)
		ReDim snap_elig_membs_ssn_coop(0)
		ReDim snap_elig_membs_unit_memb(0)
		ReDim snap_elig_membs_work_reg(0)
		ReDim snap_elig_membs_drug_felon_test(0)

		call navigate_to_MAXIS_screen("ELIG", "FS  ")
		EMWriteScreen elig_footer_month, 19, 54
		EMWriteScreen elig_footer_year, 19, 57
		Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result)

		row = 7
		elig_memb_count = 0
		Do
			EMReadScreen ref_numb, 2, row, 10

			ReDim preserve snap_elig_ref_numbs(elig_memb_count)
			ReDim preserve snap_elig_membs_request_yn(elig_memb_count)
			ReDim preserve snap_elig_membs_code(elig_memb_count)
			ReDim preserve snap_elig_membs_status_info(elig_memb_count)
			ReDim preserve snap_elig_membs_counted(elig_memb_count)
			ReDim preserve snap_elig_membs_state_food(elig_memb_count)
			ReDim preserve snap_elig_membs_eligibility(elig_memb_count)
			ReDim preserve snap_elig_membs_begin_date(elig_memb_count)
			ReDim preserve snap_elig_membs_budget_cycle(elig_memb_count)

			ReDim preserve snap_elig_membs_abawd(elig_memb_count)
			ReDim preserve snap_elig_membs_absence(elig_memb_count)
			ReDim preserve snap_elig_membs_roomer(elig_memb_count)
			ReDim preserve snap_elig_membs_boarder(elig_memb_count)
			ReDim preserve snap_elig_membs_citizenship(elig_memb_count)
			ReDim preserve snap_elig_membs_citizenship_code(elig_memb_count)
			ReDim preserve snap_elig_membs_cmdty(elig_memb_count)
			ReDim preserve snap_elig_membs_disq(elig_memb_count)
			ReDim preserve snap_elig_membs_dupl_assist(elig_memb_count)
			ReDim preserve snap_elig_membs_fraud(elig_memb_count)
			ReDim preserve snap_elig_membs_eligible_student(elig_memb_count)
			ReDim preserve snap_elig_membs_institution(elig_memb_count)
			ReDim preserve snap_elig_membs_mfip_elig(elig_memb_count)
			ReDim preserve snap_elig_membs_non_applcnt(elig_memb_count)
			ReDim preserve snap_elig_membs_residence(elig_memb_count)
			ReDim preserve snap_elig_membs_ssn_coop(elig_memb_count)
			ReDim preserve snap_elig_membs_unit_memb(elig_memb_count)
			ReDim preserve snap_elig_membs_work_reg(elig_memb_count)
			ReDim preserve snap_elig_membs_drug_felon_test(elig_memb_count)

			snap_elig_ref_numbs(elig_memb_count) = ref_numb
			EMReadScreen snap_elig_membs_request_yn(elig_memb_count), 1, row, 32
			EMReadScreen snap_elig_membs_code(elig_memb_count), 1, row, 35
			EMReadScreen memb_count, 11, row, 39
			EMReadScreen memb_state_food, 1, row, 50
			EMReadScreen memb_elig, 10, row, 57
			EMReadScreen snap_elig_membs_begin_date(elig_memb_count), 8, row, 68
			EMReadScreen memb_budg_cycle, 1, row, 78

			If snap_elig_membs_code(elig_memb_count) = "A" Then snap_elig_membs_status_info(elig_memb_count) = "Eligible"
			If snap_elig_membs_code(elig_memb_count) = "C" Then snap_elig_membs_status_info(elig_memb_count) = "Citizenship"
			If snap_elig_membs_code(elig_memb_count) = "F" Then snap_elig_membs_status_info(elig_memb_count) = "Fraud, DISQ, Work Reg"
			If snap_elig_membs_code(elig_memb_count) = "D" Then snap_elig_membs_status_info(elig_memb_count) = "Duplicate Assistance"
			If snap_elig_membs_code(elig_memb_count) = "I" Then snap_elig_membs_status_info(elig_memb_count) = "Ineligible"
			If snap_elig_membs_code(elig_memb_count) = "N" Then snap_elig_membs_status_info(elig_memb_count) = "Unit Member"
			If snap_elig_membs_code(elig_memb_count) = "S" Then snap_elig_membs_status_info(elig_memb_count) = "Ineligible Student"
			If snap_elig_membs_code(elig_memb_count) = "U" Then snap_elig_membs_status_info(elig_memb_count) = "Unknown"
			snap_elig_membs_counted(elig_memb_count) = trim(memb_count)
			If memb_state_food = "Y" Then snap_elig_membs_state_food(elig_memb_count) = True
			If memb_state_food = "N" Then snap_elig_membs_state_food(elig_memb_count) = False
			snap_elig_membs_eligibility(elig_memb_count) = trim(memb_elig)
			If memb_budg_cycle = "P" Then snap_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
			If memb_budg_cycle = "R" Then snap_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

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

			snap_elig_membs_abawd(elig_memb_count) = trim(memb_abawd)
			snap_elig_membs_absence(elig_memb_count) = trim(memb_absence)
			snap_elig_membs_roomer(elig_memb_count) = trim(memb_roomer)
			snap_elig_membs_boarder(elig_memb_count) = trim(memb_boarder)
			snap_elig_membs_citizenship(elig_memb_count) = trim(memb_citizenship)
			snap_elig_membs_citizenship_code(elig_memb_count) = trim(memb_citizenship_coop)
			snap_elig_membs_cmdty(elig_memb_count) = trim(memb_cmdty)
			snap_elig_membs_disq(elig_memb_count) = trim(memb_disq)
			snap_elig_membs_dupl_assist(elig_memb_count) = trim(memb_dupl_assist)

			EMReadScreen memb_fraud, 			6, 6, 54
			EMReadScreen memb_eligible_student, 6, 7, 54
			EMReadScreen memb_institution, 		6, 8, 54
			EMReadScreen memb_mfip_elig, 		6, 9, 54
			EMReadScreen memb_non_applcnt, 		6, 10, 54
			EMReadScreen memb_residence, 		6, 11, 54
			EMReadScreen memb_ssn_coop, 		6, 12, 54
			EMReadScreen memb_unit_memb, 		6, 13, 54
			EMReadScreen memb_work_reg, 		6, 14, 54

			snap_elig_membs_fraud(elig_memb_count) = trim(memb_fraud)
			snap_elig_membs_eligible_student(elig_memb_count) = trim(memb_eligible_student)
			snap_elig_membs_institution(elig_memb_count) = trim(memb_institution)
			snap_elig_membs_mfip_elig(elig_memb_count) = trim(memb_mfip_elig)
			snap_elig_membs_non_applcnt(elig_memb_count) = trim(memb_non_applcnt)
			snap_elig_membs_residence(elig_memb_count) = trim(memb_residence)
			snap_elig_membs_ssn_coop(elig_memb_count) = trim(memb_ssn_coop)
			snap_elig_membs_unit_memb(elig_memb_count) = trim(memb_unit_memb)
			snap_elig_membs_work_reg(elig_memb_count) = trim(memb_work_reg)
			transmit


			elig_memb_count = elig_memb_count + 1
			row = row + 1
			EMReadScreen next_ref_numb, 2, row, 10
		Loop until next_ref_numb = "  "

		transmit 		'FSCR
		EmReadScreen case_expedited_indicator, 9, 4, 3
		If case_expedited_indicator = "EXPEDITED" Then snap_expedited = True
		EMReadScreen case_uhfs_indicator, 11, 5, 4
		If case_uhfs_indicator = "UNCLE HARRY" Then snap_uhfs = True

		EMReadScreen snap_case_appl_withdrawn_test, 	6, 7, 9
		EMReadScreen snap_case_applct_elig_test, 		6, 8, 9
		EMReadScreen snap_case_comdty_test, 			6, 9, 9
		EMReadScreen snap_case_disq_test, 				6, 10, 9
		EMReadScreen snap_case_dupl_assist_test, 		6, 11, 9
		EMReadScreen snap_case_eligible_person_test, 	6, 12, 9
		EMReadScreen snap_case_fail_coop_test, 			6, 13, 9
		EMReadScreen snap_case_fail_file_test, 			6, 14, 9
		EMReadScreen snap_case_prosp_gross_inc_test, 	6, 15, 9
		EMReadScreen snap_case_prosp_net_inc_test, 		6, 16, 9
		snap_case_appl_withdrawn_test = trim(snap_case_appl_withdrawn_test)
		snap_case_applct_elig_test = trim(snap_case_applct_elig_test)
		snap_case_comdty_test = trim(snap_case_comdty_test)
		snap_case_disq_test = trim(snap_case_disq_test)
		snap_case_dupl_assist_test = trim(snap_case_dupl_assist_test)
		snap_case_eligible_person_test = trim(snap_case_eligible_person_test)
		snap_case_fail_coop_test = trim(snap_case_fail_coop_test)
		snap_case_fail_file_test = trim(snap_case_fail_file_test)
		snap_case_prosp_gross_inc_test = trim(snap_case_prosp_gross_inc_test)
		snap_case_prosp_net_inc_test = trim(snap_case_prosp_net_inc_test)

		EMReadScreen snap_case_recert_test, 			6, 7, 49
		EMReadScreen snap_case_residence_test, 			6, 8, 49
		EMReadScreen snap_case_resource_test, 			6, 9, 49
		EMReadScreen snap_case_retro_gross_inc_test, 	6, 10, 49
		EMReadScreen snap_case_retro_net_inc_test, 		6, 11, 49
		EMReadScreen snap_case_strike_test, 			6, 12, 49
		EMReadScreen snap_case_xfer_resource_inc_test, 	6, 13, 49
		EMReadScreen snap_case_verif_test, 				6, 14, 49
		EMReadScreen snap_case_voltry_quit_test, 		6, 15, 49
		EMReadScreen snap_case_work_reg_test, 			6, 16, 49
		snap_case_recert_test = trim(snap_case_recert_test)
		snap_case_residence_test = trim(snap_case_residence_test)
		snap_case_resource_test = trim(snap_case_resource_test)
		snap_case_retro_gross_inc_test = trim(snap_case_retro_gross_inc_test)
		snap_case_retro_net_inc_test = trim(snap_case_retro_net_inc_test)
		snap_case_strike_test = trim(snap_case_strike_test)
		snap_case_xfer_resource_inc_test = trim(snap_case_xfer_resource_inc_test)
		snap_case_verif_test = trim(snap_case_verif_test)
		snap_case_voltry_quit_test = trim(snap_case_voltry_quit_test)
		snap_case_work_reg_test = trim(snap_case_work_reg_test)

		Call write_value_and_transmit("X", 14, 4)		''Fail to File Detail
		EMReadScreen snap_fail_file_hrf, 6, 10, 32
		EMReadScreen snap_fail_file_sr, 6, 11, 32
		transmit
		snap_fail_file_hrf = trim(snap_fail_file_hrf)
		snap_fail_file_sr = trim(snap_fail_file_sr)

		Call write_value_and_transmit("X", 14, 4)		''Resource Detail
		EMReadScreen snap_resource_cash, 	10, 8, 47
		EMReadScreen snap_resource_acct, 	10, 9, 47
		EMReadScreen snap_resource_secu, 	10, 10, 47
		EMReadScreen snap_resource_cars, 	10, 11, 47
		EMReadScreen snap_resource_rest, 	10, 12, 47
		EMReadScreen snap_resource_other, 	10, 13, 47
		EMReadScreen snap_resource_burial, 	10, 14, 47
		EMReadScreen snap_resource_spon, 	10, 15, 47
		EMReadScreen snap_resource_total, 	10, 17, 47
		EMReadScreen snap_resource_max, 	10, 18, 47
		transmit

		snap_resource_cash = trim(snap_resource_cash)
		snap_resource_acct = trim(snap_resource_acct)
		snap_resource_secu = trim(snap_resource_secu)
		snap_resource_cars = trim(snap_resource_cars)
		snap_resource_rest = trim(snap_resource_rest)
		snap_resource_other = trim(snap_resource_other)
		snap_resource_burial = trim(snap_resource_burial)
		snap_resource_spon = trim(snap_resource_spon)
		snap_resource_total = trim(snap_resource_total)
		snap_resource_max = trim(snap_resource_max)

		transmit 		'FSB1
		EMReadScreen snap_budg_gross_wages, 		10, 5, 31
		EMReadScreen snap_budg_self_emp, 			10, 6, 31
		EMReadScreen snap_budg_total_earned_inc, 	10, 8, 31

		snap_budg_gross_wages = trim(snap_budg_gross_wages)
		snap_budg_self_emp = trim(snap_budg_self_emp)
		snap_budg_total_earned_inc = trim(snap_budg_total_earned_inc)


		EMReadScreen snap_budg_pa_grant_inc, 	10, 10, 31
		EMReadScreen snap_budg_rsdi_inc, 		10, 11, 31
		EMReadScreen snap_budg_ssi_inc, 		10, 12, 31
		EMReadScreen snap_budg_va_inc, 			10, 13, 31
		EMReadScreen snap_budg_uc_wc_inc, 		10, 14, 31
		EMReadScreen snap_budg_cses_inc, 		10, 15, 31
		EMReadScreen snap_budg_other_unea_inc, 	10, 16, 31
		EMReadScreen snap_budg_total_unea_inc, 	10, 18, 31

		snap_budg_pa_grant_inc = trim(snap_budg_pa_grant_inc)
		snap_budg_rsdi_inc = trim(snap_budg_rsdi_inc)
		snap_budg_ssi_inc = trim(snap_budg_ssi_inc)
		snap_budg_va_inc = trim(snap_budg_va_inc)
		snap_budg_uc_wc_inc = trim(snap_budg_uc_wc_inc)
		snap_budg_cses_inc = trim(snap_budg_cses_inc)
		snap_budg_other_unea_inc = trim(snap_budg_other_unea_inc)
		snap_budg_total_unea_inc = trim(snap_budg_total_unea_inc)

		EMReadScreen snap_budg_schl_inc, 			10, 5, 71
		EMReadScreen snap_budg_farm_ofset, 			10, 6, 71
		EMReadScreen snap_budg_total_gross_inc, 	10, 7, 71
		EMReadScreen snap_budg_max_gross_inc, 		10, 8, 71

		EMReadScreen snap_budg_deduct_standard, 	10, 10, 71
		EMReadScreen snap_budg_deduct_earned, 		10, 11, 71
		EMReadScreen snap_budg_deduct_medical, 		10, 12, 71
		EMReadScreen snap_budg_deduct_depndt_care, 	10, 13, 71
		EMReadScreen snap_budg_deduct_cses, 		10, 14, 71
		EMReadScreen snap_budg_total_deduct, 		10, 16, 71

		EMReadScreen snap_budg_net_inc, 			10, 18, 71

		snap_budg_schl_inc = trim(snap_budg_schl_inc)
		snap_budg_farm_ofset = trim(snap_budg_farm_ofset)
		snap_budg_total_gross_inc = trim(snap_budg_total_gross_inc)
		snap_budg_max_gross_inc = trim(snap_budg_max_gross_inc)
		snap_budg_deduct_standard = trim(snap_budg_deduct_standard)
		snap_budg_deduct_earned = trim(snap_budg_deduct_earned)
		snap_budg_deduct_medical = trim(snap_budg_deduct_medical)
		snap_budg_deduct_depndt_care = trim(snap_budg_deduct_depndt_care)
		snap_budg_deduct_cses = trim(snap_budg_deduct_cses)
		snap_budg_total_deduct = trim(snap_budg_total_deduct)
		snap_budg_net_inc = trim(snap_budg_net_inc)

		transmit 		'FSB2
		EMReadScreen snap_budg_shel_rent_mort, 		10, 5, 27
		EMReadScreen snap_budg_shel_prop_tax, 		10, 6, 27
		EMReadScreen snap_budg_shel_home_ins, 		10, 7, 27
		EMReadScreen snap_budg_shel_electricity, 	10, 8, 27
		EMReadScreen snap_budg_shel_heat_ac, 		10, 9, 27
		EMReadScreen snap_budg_shel_water_garbage, 	10, 10, 27
		EMReadScreen snap_budg_shel_phone, 			10, 11, 27
		EMReadScreen snap_budg_shel_other, 			10, 12, 27
		EMReadScreen snap_budg_shel_total, 			10, 14, 27
		EMReadScreen snap_budg_50_perc_net_inc, 	10, 15, 27
		EMReadScreen snap_budg_adj_shel_costs, 		10, 17, 27

		snap_budg_shel_rent_mort = trim(snap_budg_shel_rent_mort)
		snap_budg_shel_prop_tax = trim(snap_budg_shel_prop_tax)
		snap_budg_shel_home_ins = trim(snap_budg_shel_home_ins)
		snap_budg_shel_electricity = trim(snap_budg_shel_electricity)
		snap_budg_shel_heat_ac = trim(snap_budg_shel_heat_ac)
		snap_budg_shel_water_garbage = trim(snap_budg_shel_water_garbage)
		snap_budg_shel_phone = trim(snap_budg_shel_phone)
		snap_budg_shel_other = trim(snap_budg_shel_other)
		snap_budg_shel_total = trim(snap_budg_shel_total)
		snap_budg_50_perc_net_inc = trim(snap_budg_50_perc_net_inc)
		snap_budg_adj_shel_costs = trim(snap_budg_adj_shel_costs)


		EMReadScreen snap_budg_max_allow_shel, 			10, 5, 71
		EMReadScreen snap_budg_shel_expenses, 			10, 6, 71
		' EMReadScreen fsb2_net_adj_inc, 				10, 7, 71
		EMReadScreen snap_budg_max_net_adj_inc, 		10, 8, 71
		EMReadScreen snap_benefit_monthly_fs_allot, 	10, 10, 71
		EMReadScreen snap_benefit_drug_felon_sanc_amt, 	10, 12, 71
		EMReadScreen snap_benefit_recoup_amount, 		10, 14, 71
		EMReadScreen snap_benefit_benefit_amount, 		10, 16, 71
		EMReadScreen snap_benefit_state_food_amt, 		10, 17, 71
		EMReadScreen snap_benefit_fed_food_amt, 		10, 18, 71

		snap_budg_max_allow_shel = trim(snap_budg_max_allow_shel)
		snap_budg_shel_expenses = trim(snap_budg_shel_expenses)
		' fsb2_net_adj_inc = trim(fsb2_net_adj_inc)
		snap_budg_max_net_adj_inc = trim(snap_budg_max_net_adj_inc)
		snap_benefit_monthly_fs_allot = trim(snap_benefit_monthly_fs_allot)
		snap_benefit_drug_felon_sanc_amt = trim(snap_benefit_drug_felon_sanc_amt)
		snap_benefit_recoup_amount = trim(snap_benefit_recoup_amount)
		snap_benefit_benefit_amount = trim(snap_benefit_benefit_amount)
		snap_benefit_state_food_amt = trim(snap_benefit_state_food_amt)
		snap_benefit_fed_food_amt = trim(snap_benefit_fed_food_amt)


		Call write_value_and_transmit("X", 14, 4)		''Resource Detail
		row = 8
		Do
			EMReadScreen ref_numb, 2, row, 12

			For case_memb = 0 to UBound(snap_elig_ref_numbs)
				If ref_numb = snap_elig_ref_numbs(case_memb) Then
					EMReadScreen memb_drug_felon_test, 6, row, 64
					snap_elig_membs_drug_felon_test(case_memb) = trim(memb_drug_felon_test)
				End If
			Next

			row = row + 1
			EMReadScreen next_ref_numb, 2, row, 12
		Loop until next_ref_numb = "  "
		transmit

		Call write_value_and_transmit("X", 14, 4)		''Resource Detail
		EMReadScreen snap_benefit_recoup_from_fed_fs, 10, 5, 51
		EMReadScreen snap_benefit_recoup_from_state_fs, 10, 7, 51

		snap_benefit_recoup_from_fed_fs = trim(snap_benefit_recoup_from_fed_fs)
		snap_benefit_recoup_from_state_fs = trim(snap_benefit_recoup_from_state_fs)

		transmit

		transmit 		'FSSM
		EMReadScreen snap_approved_date, 			8, 3, 14
		EMReadScreen snap_date_last_approval, 		8, 5, 31
		EMReadScreen snap_curr_prog_status, 		10, 6, 31
		EMReadScreen snap_elig_result, 				10, 7, 31
		EMReadScreen snap_reporting_status, 		12, 8, 31
		EMReadScreen snap_info_source, 				4, 9, 31
		EMReadScreen snap_benefit, 					12, 10, 31
		EMReadScreen snap_elig_revw_date, 			8, 11, 31
		EMReadScreen snap_budget_cycle, 			5, 12, 31
		EMReadScreen snap_budg_numb_in_assist_unit, 2, 13, 31

		EMReadScreen snap_budg_total_resources, 		10, 5, 71
		EMReadScreen snap_budg_max_resources, 			10, 6, 71
		EMReadScreen snap_budg_net_adj_inc, 			10, 7, 71
		EMReadScreen snap_benefit_monthly_fs_allotment, 10, 8, 71
		EMReadScreen snap_benefit_prorated_amt, 		10, 9, 71
		EMReadScreen snap_benefit_prorated_date,		8, 9, 58
		EMReadScreen snap_benefit_amt, 					10, 13, 71

		snap_approved_date = trim(snap_approved_date)
		snap_date_last_approval = trim(snap_date_last_approval)
		snap_curr_prog_status = trim(snap_curr_prog_status)
		snap_elig_result = trim(snap_elig_result)
		snap_reporting_status = trim(snap_reporting_status)
		snap_info_source = trim(snap_info_source)
		snap_benefit = trim(snap_benefit)
		snap_elig_revw_date = trim(snap_elig_revw_date)
		snap_budget_cycle = trim(snap_budget_cycle)
		snap_budg_numb_in_assist_unit = trim(snap_budg_numb_in_assist_unit)
		snap_budg_total_resources = trim(snap_budg_total_resources)
		snap_budg_max_resources = trim(snap_budg_max_resources)
		snap_budg_net_adj_inc = trim(snap_budg_net_adj_inc)
		snap_benefit_monthly_fs_allotment = trim(snap_benefit_monthly_fs_allotment)
		snap_benefit_prorated_amt = trim(snap_benefit_prorated_amt)
		snap_benefit_prorated_date = trim(snap_benefit_prorated_date)
		snap_benefit_amt = trim(snap_benefit_amt)


		EMReadScreen fssm_expedited_info_exists, 16, 14, 44
		If fssm_expedited_info_exists = "EXPEDITED STATUS" Then
			Call write_value_and_transmit("X", 14, 72)		''Resource Detail
			EMReadScreen exp_status_issuance_on_or_before_15th, 1, 3, 5
			EMReadScreen exp_status_issuance_after_15th, 1, 5, 5
			EMReadScreen exp_status_issuance_app_month_fs_denial, 1, 9, 5

			EMReadScreen snap_exp_criteria_migrant_destitute, 1, 15, 5
			EMReadScreen snap_exp_criteria_resource_100_income_150, 1, 16, 5
			EMReadScreen snap_exp_criteria_resource_income_less_shelter, 1, 19, 5

			EMReadScreen snap_exp_verif_status_postponed, 1, 15, 52
			EMReadScreen snap_exp_verif_status_out_of_state, 1, 17, 52
			EMReadScreen snap_exp_verif_status_all_provided, 1, 19, 52
			transmit

			If exp_status_issuance_on_or_before_15th = "X" Then snap_exp_package_includes_month_one = True
			If exp_status_issuance_after_15th = "X" Then
				snap_exp_package_includes_month_one = True
				snap_exp_package_includes_month_two = True
			End If
			If exp_status_issuance_app_month_fs_denial = "X" Then snap_exp_package_includes_month_two = True

			If snap_exp_criteria_migrant_destitute = "X" Then snap_exp_criteria_migrant_destitute = True
			If snap_exp_criteria_migrant_destitute = "_" Then snap_exp_criteria_migrant_destitute = False
			If snap_exp_criteria_resource_100_income_150 = "X" Then snap_exp_criteria_resource_100_income_150 = True
			If snap_exp_criteria_resource_100_income_150 = "_" Then snap_exp_criteria_resource_100_income_150 = False
			If snap_exp_criteria_resource_income_less_shelter = "X" Then snap_exp_criteria_resource_income_less_shelter = True
			If snap_exp_criteria_resource_income_less_shelter = "_" Then snap_exp_criteria_resource_income_less_shelter = False

			If snap_exp_verif_status_postponed = "X" Then snap_exp_verif_status_postponed = True
			If snap_exp_verif_status_postponed = "_" Then snap_exp_verif_status_postponed = False
			If snap_exp_verif_status_out_of_state = "X" Then snap_exp_verif_status_out_of_state = True
			If snap_exp_verif_status_out_of_state = "_" Then snap_exp_verif_status_out_of_state = False
			If snap_exp_verif_status_all_provided = "X" Then snap_exp_verif_status_all_provided = True
			If snap_exp_verif_status_all_provided = "_" Then snap_exp_verif_status_all_provided = False


		End If

		EMReadScreen snap_elig_worker_message_one, 80, 17, 1
		EMReadScreen snap_elig_worker_message_two, 80, 18, 1

		snap_elig_worker_message_one = trim(snap_elig_worker_message_one)
		snap_elig_worker_message_two = trim(snap_elig_worker_message_two)

		If snap_budg_total_earned_inc <> "" Then snap_earned_income_budgeted = True
		If snap_budg_total_unea_inc <> "" Then snap_unearned_income_budgeted = True
		If snap_budg_shel_rent_mort <> "" or snap_budg_shel_prop_tax <> "" or snap_budg_shel_home_ins <> "" or snap_budg_shel_other <> ""Then snap_shel_costs_budgeted = True
		If snap_budg_shel_electricity <> "" or snap_budg_shel_heat_ac <> "" or snap_budg_shel_water_garbage <> "" or snap_budg_shel_phone <> ""Then snap_hest_costs_budgeted = True
		' categorical_eligibility = ""

		' MsgBox "snap_benefit_amt - " & snap_benefit_amt

		Call Back_to_SELF
	End sub

end class


class mfip_eligibility_detial
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result

	public mfip_elig_ref_numbs()
	public mfip_elig_membs_request_yn()
	public mfip_elig_membs_code()
	public mfip_elig_membs_status_info()
	public mfip_elig_membs_deemed()
	public mfip_elig_membs_counted()
	public mfip_elig_membs_eligibility()
	public mfip_elig_membs_begin_date()
	public mfip_elig_membs_budget_cycle()
	public mfip_elig_membs_absence()
	public mfip_elig_membs_child_age()
	public mfip_elig_membs_citizenship()
	public mfip_elig_membs_citizenship_verif()
	public mfip_elig_membs_dupl_assist()
	public mfip_elig_membs_foster_care()
	public mfip_elig_membs_fraud()
	public mfip_elig_membs_fs_disq()
	public mfip_elig_membs_minor_living_arngmt()
	public mfip_elig_membs_post_60_removal()
	public mfip_elig_membs_ssi()
	public mfip_elig_membs_ssn_code()
	public mfip_elig_membs_unit_memb()
	public mfip_elig_membs_unlawful_conduct()
	public mfip_elig_membs_fs_recvd()
	public mfip_elig_membs_es_status_code()
	public mfip_elig_membs_es_status_info()

	public sub read_elig()
		call navigate_to_MAXIS_screen("ELIG", "MFIP")
		EMWriteScreen elig_footer_month, 20, 56
		EMWriteScreen elig_footer_year, 20, 59
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)

		ReDim mfip_elig_ref_numbs(0)
		ReDim mfip_elig_membs_request_yn(0)
		ReDim mfip_elig_membs_code(0)
		ReDim mfip_elig_membs_status_info(0)
		ReDim mfip_elig_membs_deemed(0)
		ReDim mfip_elig_membs_counted(0)
		ReDim mfip_elig_membs_eligibility(0)
		ReDim mfip_elig_membs_begin_date(0)
		ReDim mfip_elig_membs_budget_cycle(0)
		ReDim mfip_elig_membs_absence(0)
		ReDim mfip_elig_membs_child_age(0)
		ReDim mfip_elig_membs_citizenship(0)
		ReDim mfip_elig_membs_citizenship_verif(0)
		ReDim mfip_elig_membs_dupl_assist(0)
		ReDim mfip_elig_membs_foster_care(0)
		ReDim mfip_elig_membs_fraud(0)
		ReDim mfip_elig_membs_fs_disq(0)
		ReDim mfip_elig_membs_minor_living_arngmt(0)
		ReDim mfip_elig_membs_post_60_removal(0)
		ReDim mfip_elig_membs_ssi(0)
		ReDim mfip_elig_membs_ssn_code(0)
		ReDim mfip_elig_membs_unit_memb(0)
		ReDim mfip_elig_membs_unlawful_conduct(0)
		ReDim mfip_elig_membs_fs_recvd(0)
		ReDim mfip_elig_membs_es_status_code(0)
		ReDim mfip_elig_membs_es_status_info(0)

		row = 7
		elig_memb_count = 0
		Do
			EMReadScreen ref_numb, 2, row, 6

			ReDim preserve mfip_elig_ref_numbs(elig_memb_count)
			ReDim preserve mfip_elig_membs_request_yn(elig_memb_count)
			ReDim preserve mfip_elig_membs_code(elig_memb_count)
			ReDim preserve mfip_elig_membs_status_info(elig_memb_count)
			ReDim preserve mfip_elig_membs_deemed(elig_memb_count)
			ReDim preserve mfip_elig_membs_counted(elig_memb_count)
			ReDim preserve mfip_elig_membs_eligibility(elig_memb_count)
			ReDim preserve mfip_elig_membs_begin_date(elig_memb_count)
			ReDim preserve mfip_elig_membs_budget_cycle(elig_memb_count)
			ReDim preserve mfip_elig_membs_absence(elig_memb_count)
			ReDim preserve mfip_elig_membs_child_age(elig_memb_count)
			ReDim preserve mfip_elig_membs_citizenship(elig_memb_count)
			ReDim preserve mfip_elig_membs_citizenship_verif(elig_memb_count)
			ReDim preserve mfip_elig_membs_dupl_assist(elig_memb_count)
			ReDim preserve mfip_elig_membs_foster_care(elig_memb_count)
			ReDim preserve mfip_elig_membs_fraud(elig_memb_count)
			ReDim preserve mfip_elig_membs_fs_disq(elig_memb_count)
			ReDim preserve mfip_elig_membs_minor_living_arngmt(elig_memb_count)
			ReDim preserve mfip_elig_membs_post_60_removal(elig_memb_count)
			ReDim preserve mfip_elig_membs_ssi(elig_memb_count)
			ReDim preserve mfip_elig_membs_ssn_code(elig_memb_count)
			ReDim preserve mfip_elig_membs_unit_memb(elig_memb_count)
			ReDim preserve mfip_elig_membs_unlawful_conduct(elig_memb_count)
			ReDim preserve mfip_elig_membs_fs_recvd(elig_memb_count)
			ReDim preserve mfip_elig_membs_es_status_code(elig_memb_count)
			ReDim preserve mfip_elig_membs_es_status_info(elig_memb_count)

			mfip_elig_ref_numbs(elig_memb_count) = ref_numb
			EMReadScreen mfip_elig_membs_request_yn(elig_memb_count), 1, row, 32
			EMReadScreen mfip_elig_membs_code(elig_memb_count), 1, row, 36
			EMReadScreen mfip_elig_membs_counted(elig_memb_count), 11, row, 41
			EMReadScreen mfip_elig_membs_eligibility(elig_memb_count), 10, row, 53
			EMReadScreen mfip_elig_membs_begin_date(elig_memb_count), 8, row, 67
			EMReadScreen mfip_elig_membs_budget_cycle(elig_memb_count), 1, row, 78

			If mfip_elig_membs_code(elig_memb_count) = "A" Then mfip_elig_membs_status_info(elig_memb_count) = "Eligible"
			If mfip_elig_membs_code(elig_memb_count) = "D" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed SSI, IV-E ADOPTION ASSISTANCE"
			If mfip_elig_membs_code(elig_memb_count) = "F" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed FRAUD, SSN COOP, UNLAWFUL CONDUCT"
			If mfip_elig_membs_code(elig_memb_count) = "G" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Parent of a minor caregiver"
			If mfip_elig_membs_code(elig_memb_count) = "H" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed CITIZENSHIP, CITIZENSHIP VERIFICATION"
			If mfip_elig_membs_code(elig_memb_count) = "I" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed ABSENCE, DUPLICATE ASSISTANCE, CHILD AGE"
			If mfip_elig_membs_code(elig_memb_count) = "J" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed MFIP PERSON POST 60 REMOVAL"
			If mfip_elig_membs_code(elig_memb_count) = "N" Then mfip_elig_membs_status_info(elig_memb_count) = "Not a Unit Member"
			If mfip_elig_membs_code(elig_memb_count) = "A" Then mfip_elig_membs_deemed(elig_memb_count) = "Unit Member"
			If mfip_elig_membs_code(elig_memb_count) = "F" or mfip_elig_membs_code(elig_memb_count) = "G" or mfip_elig_membs_code(elig_memb_count) = "H" or mfip_elig_membs_code(elig_memb_count) = "J" Then mfip_elig_membs_deemed(elig_memb_count) = "Deemed"
			If mfip_elig_membs_code(elig_memb_count) = "D" or mfip_elig_membs_code(elig_memb_count) = "I" or mfip_elig_membs_code(elig_memb_count) = "N" Then mfip_elig_membs_deemed(elig_memb_count) = "Not Deemed"
			mfip_elig_membs_counted(elig_memb_count) = trim(mfip_elig_membs_counted(elig_memb_count))
			mfip_elig_membs_eligibility(elig_memb_count) = trim(mfip_elig_membs_eligibility(elig_memb_count))
			If mfip_elig_membs_budget_cycle(elig_memb_count) = "P" Then mfip_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
			If mfip_elig_membs_budget_cycle(elig_memb_count) = "R" Then mfip_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

			Call write_value_and_transmit("X", row, 3)
			EMReadScreen mfip_elig_membs_absence(elig_memb_count), 			6, 7, 17
			EMReadScreen mfip_elig_membs_child_age(elig_memb_count), 		6, 8, 17
			EMReadScreen mfip_elig_membs_citizenship(elig_memb_count), 		6, 9, 17
			EMReadScreen mfip_elig_membs_citizenship_verif(elig_memb_count),6, 10, 17
			EMReadScreen mfip_elig_membs_dupl_assist(elig_memb_count), 		6, 11, 17
			EMReadScreen mfip_elig_membs_foster_care(elig_memb_count), 		6, 12, 17
			EMReadScreen mfip_elig_membs_fraud(elig_memb_count), 			6, 13, 17
			EMReadScreen mfip_elig_membs_fs_disq(elig_memb_count), 			6, 17, 17

			mfip_elig_membs_absence(elig_memb_count) = trim(mfip_elig_membs_absence(elig_memb_count))
			mfip_elig_membs_child_age(elig_memb_count) = trim(mfip_elig_membs_child_age(elig_memb_count))
			mfip_elig_membs_citizenship(elig_memb_count) = trim(mfip_elig_membs_citizenship(elig_memb_count))
			mfip_elig_membs_citizenship_verif(elig_memb_count) = trim(mfip_elig_membs_citizenship_verif(elig_memb_count))
			mfip_elig_membs_dupl_assist(elig_memb_count) = trim(mfip_elig_membs_dupl_assist(elig_memb_count))
			mfip_elig_membs_foster_care(elig_memb_count) = trim(mfip_elig_membs_foster_care(elig_memb_count))
			mfip_elig_membs_fraud(elig_memb_count) = trim(mfip_elig_membs_fraud(elig_memb_count))
			mfip_elig_membs_fs_disq(elig_memb_count) = trim(mfip_elig_membs_fs_disq(elig_memb_count))


			EMReadScreen mfip_elig_membs_minor_living_arngmt(elig_memb_count), 	6, 7, 52
			EMReadScreen mfip_elig_membs_post_60_removal(elig_memb_count), 		6, 8, 52
			EMReadScreen mfip_elig_membs_ssi(elig_memb_count), 					6, 9, 52
			EMReadScreen mfip_elig_membs_ssn_code(elig_memb_count), 			6, 10, 52
			EMReadScreen mfip_elig_membs_unit_memb(elig_memb_count), 			6, 11, 52
			EMReadScreen mfip_elig_membs_unlawful_conduct(elig_memb_count), 	6, 12, 52
			EMReadScreen mfip_elig_membs_fs_recvd(elig_memb_count), 			6, 17, 52

			mfip_elig_membs_minor_living_arngmt(elig_memb_count) = trim(mfip_elig_membs_minor_living_arngmt(elig_memb_count))
			mfip_elig_membs_post_60_removal(elig_memb_count) = trim(mfip_elig_membs_post_60_removal(elig_memb_count))
			mfip_elig_membs_ssi(elig_memb_count) = trim(mfip_elig_membs_ssi(elig_memb_count))
			mfip_elig_membs_ssn_code(elig_memb_count) = trim(mfip_elig_membs_ssn_code(elig_memb_count))
			mfip_elig_membs_unit_memb(elig_memb_count) = trim(mfip_elig_membs_unit_memb(elig_memb_count))
			mfip_elig_membs_unlawful_conduct(elig_memb_count) = trim(mfip_elig_membs_unlawful_conduct(elig_memb_count))
			mfip_elig_membs_fs_recvd(elig_memb_count) = trim(mfip_elig_membs_fs_recvd(elig_memb_count))

			transmit

			Call write_value_and_transmit("X", row, 64)
			EMReadScreen mfip_elig_membs_es_status_code(elig_memb_count), 2, 9, 22
			EMReadScreen mfip_elig_membs_es_status_info(elig_memb_count), 30, 9, 25

			mfip_elig_membs_es_status_code(elig_memb_count) = trim(mfip_elig_membs_es_status_code(elig_memb_count))
			mfip_elig_membs_es_status_info(elig_memb_count) = trim(mfip_elig_membs_es_status_info(elig_memb_count))
			transmit

			row = row + 1
			EMReadScreen next_ref_numb, 2, row, 6
		Loop until next_ref_numb = "  "
	end sub

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

		Call SNAP_ELIG_APPROVALS(snap_elig_months_count).read_elig

		' MsgBox "SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month - " & SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month

		snap_elig_months_count = snap_elig_months_count + 1
	End If

	Call back_to_SELF
Next

For approval_month = 0 to UBound(SNAP_ELIG_APPROVALS)
	For snap_memb = 0 to UBound(SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs)
		MsgBox SNAP_ELIG_APPROVALS(approval_month).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval_month).elig_footer_year & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs(snap_memb) & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_membs_eligibility(snap_memb)
	Next
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
