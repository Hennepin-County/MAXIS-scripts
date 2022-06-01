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

	public mfip_memb_cash_portion_code()
	public mfip_memb_food_portion_code()
	public mfip_memb_state_food_code()
	public mfip_memb_sanction_yn()
	public mfip_memb_sanction_child_support_test()
	public mfip_memb_sanction_drug_felon_test()
	public mfip_memb_sanction_emp_services_test()
	public mfip_memb_sanction_fin_orient_test()
	public mfip_memb_sanction_occurence()
	public mfip_memb_sanction_begin_date()
	public mfip_memb_sanction_last_sanc_month()

	public mfip_cash_opt_out
	public mfip_HG_opt_out
	public mfip_child_only

	public mfip_case_test_appl_withdraw
	public mfip_case_test_asset
	public mfip_case_test_death_applicant
	public mfip_case_test_dupl_assist
	public mfip_case_test_elig_child
	public mfip_case_test_fail_coop
	public mfip_case_test_fail_file
	public mfip_case_test_initial_income
	public mfip_case_test_minor_liv_arrange
	public mfip_case_test_monthly_income
	public mfip_case_test_post_60_disq
	public mfip_case_test_residence
	public mfip_case_test_sanction_limit
	public mfip_case_test_strike
	public mfip_case_test_TANF_time_limit
	public mfip_case_test_transfer_asset
	public mfip_case_test_verif
	public mfip_case_test_275_new_spouse_income
	public mfip_fs_case_test_fail_coop_snap_qc
	public mfip_fs_case_test_opt_out_cash
	public mfip_fs_case_test_opt_out_housing_grant

	public mfip_counted_asset_CASH
	public mfip_counted_asset_ACCT
	public mfip_counted_asset_SECU
	public mfip_counted_asset_CARS
	public mfip_counted_asset_SPON
	public mfip_counted_asset_total
	public mfip_counted_asset_max

	public mfip_initial_income_earned
	public mfip_initial_income_deoendant_care
	public mfip_initial_income_unearned
	public mfip_initial_income_deemed
	public mfip_initial_income_cses_exclusion
	public mfip_initial_income_total
	public mfip_initial_income_family_wage_level

	public mfip_12_month_start_date
	public mfip_designated_spouse_ref_numb
	public mfip_new_spouse_inc_earned
	public mfip_new_spouse_inc_unearned
	public mfip_new_spouse_inc_deemed_earned
	public mfip_new_spouse_inc_deemed_unearned
	public mfip_new_spouse_inc_total
	public mfip_275_fpg_amt
	public mfip_hh_size_fornew_spouse_calc

	public mfip_case_sanction_percent
	public mfip_case_sanction_vendor_yn
	public mfip_case_sanction_last_vendor_month

	public mfip_case_budg_family_wage_level
	public mfip_case_budg_monthly_earned_income
	public mfip_case_budg_wage_level_earned_inc_difference
	public mfip_case_budg_transitional_standard
	public mfip_case_budg_monthly_need
	public mfip_case_budg_unearned_income
	public mfip_case_budg_deemed_income
	public mfip_case_budg_cses_exclusion
	public mfip_case_budg_unmet_need
	public mfip_case_budg_unmet_need_food_potion
	public mfip_case_budg_tribal_counted_income
	public mfip_case_budg_unmet_need_cash_portion
	public mfip_case_budg_deduction_subsidy_tribal_cses

	public mfip_case_budg_net_food_portion
	public mfip_case_budg_net_cash_portion
	public mfip_case_budg_net_unmet_need
	public mfip_case_budg_deduction_sanction_vendor
	public mfip_case_budg_unmet_neet_subtotal
	public mfip_case_budg_subtotal_food_portion
	public mfip_case_budg_food_portion_deduction
	public mfip_case_budg_entitlement_food_portion
	public mfip_case_budg_entitlement_housing_grant

	public mfip_budg_cses_excln_cses_income
	public mfip_budg_cses_excln_child_count
	public mfip_budg_cses_excln_total

	public mfip_case_budg_10_perc_sanc
	public mfip_case_budg_unmet_need_after_pre_vndr_sanc
	public mfip_case_budg_sanc_calc_food_portion
	public mfip_case_budg_sanc_calc_cash_portion
	public mfip_case_budg_pot_mand_vndr_pymts
	public mfip_case_budg_30_perc_sanc

	public mfip_case_budg_non_citzn_fs_inelig_pers_count
	public mfip_case_budg_non_citzn_fs_inelig_amt
	public mfip_case_budg_other_fs_inelig_pers_count
	public mfip_case_budg_other_fs_inelig_amt

	public mfip_case_budg_prorate_date
	public mfip_case_budg_fed_food_benefit
	public mfip_case_budg_food_prorated_amt
	public mfip_case_budg_entitlement_cash_portion
	public mfip_case_budg_mand_sanc_vendor
	public mfip_case_budg_net_cash_portion
	public mfip_case_budg_cash_prorated_amt
	public mfip_case_budg_state_food_benefit
	public mfip_case_budg_state_food_prorated_amt
	public mfip_case_budg_grant_amount
	public mfip_case_budg_amt_already_issued
	public mfip_case_budg_supplement_due
	public mfip_case_budg_overpayment
	public mfip_case_budg_adjusted_grant_amt
	public mfip_case_budg_recoupment
	public mfip_case_budg_total_food_issuance
	public mfip_case_budg_total_cash_issuance
	public mfip_case_budg_total_housing_grant_issuance

	public mfip_case_budg_food_supplement
	public mfip_case_budg_state_food_supplement
	public mfip_case_budg_cash_supplement
	public mfip_case_budg_housing_grant_supplement

	public mfip_case_budg_cash_recoupment
	public mfip_case_budg_state_food_recoupment
	public mfip_case_budg_food_recoupment

	public mfip_case_budg_fed_food_memb_count
	public mfip_case_budg_fed_food_benefit_amt
	public mfip_case_budg_state_food_memb_count
	public mfip_case_budg_state_food_benefit_amt

	public mfip_case_budg_tanf_cash_memb_count
	public mfip_case_budg_tanf_cash_benefit_amt
	public mfip_case_budg_state_cash_memb_count
	public mfip_case_budg_state_cash_benefit_amt

	public mfip_approved_date
	public mfip_process_date
	public mfip_prev_approval
	public mfip_case_last_approval_date
	public mfip_case_current_prog_status
	public mfip_case_eligibility_result
	public mfip_case_hrf_reporting
	public mfip_case_source_of_info
	public mfip_case_benefit_impact
	public mfip_case_review_date
	public mfip_case_budget_cycle
	public mfip_case_vendor_reason_code
	public mfip_case_vendor_reason_info
	public mfip_case_responsible_county
	public mfip_case_service_county
	public mfip_case_asst_unit_caregivers
	public mfip_case_asst_unit_children
	public mfip_case_total_assets
	public mfip_case_maximum_assets
	public mfip_case_summary_grant_amount
	public mfip_case_summary_net_grant_amount
	public mfip_case_summary_cash_portion
	public mfip_case_summary_food_portion
	public mfip_case_summary_housing_grant

	public sub read_elig()
		mfip_cash_opt_out = False
		mfip_HG_opt_out = False
		mfip_child_only = False

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
		ReDim mfip_memb_cash_portion_code(0)
		ReDim mfip_memb_food_portion_code(0)
		ReDim mfip_memb_state_food_code(0)
		ReDim mfip_memb_sanction_yn(0)
		ReDim mfip_memb_sanction_child_support_test(0)
		ReDim mfip_memb_sanction_drug_felon_test(0)
		ReDim mfip_memb_sanction_emp_services_test(0)
		ReDim mfip_memb_sanction_fin_orient_test(0)
		ReDim mfip_memb_sanction_occurence(0)
		ReDim mfip_memb_sanction_begin_date(0)
		ReDim mfip_memb_sanction_last_sanc_month(0)

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
			ReDim preserve mfip_memb_cash_portion_code(elig_memb_count)
			ReDim preserve mfip_memb_food_portion_code(elig_memb_count)
			ReDim preserve mfip_memb_state_food_code(elig_memb_count)
			ReDim preserve mfip_memb_sanction_yn(elig_memb_count)
			ReDim preserve mfip_memb_sanction_child_support_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_drug_felon_test()
			ReDim preserve mfip_memb_sanction_emp_services_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_fin_orient_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_occurence(elig_memb_count)
			ReDim preserve mfip_memb_sanction_begin_date(elig_memb_count)
			ReDim preserve mfip_memb_sanction_last_sanc_month(elig_memb_count)

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

		transmit			'MFCR

		EMReadScreen mfip_case_test_appl_withdraw, 		6, 6, 7
		EMReadScreen mfip_case_test_asset, 				6, 7, 7
		EMReadScreen mfip_case_test_death_applicant, 	6, 8, 7
		EMReadScreen mfip_case_test_dupl_assist, 		6, 9, 7
		EMReadScreen mfip_case_test_elig_child, 		6, 10, 7
		EMReadScreen mfip_case_test_fail_coop, 			6, 11, 7
		EMReadScreen mfip_case_test_fail_file, 			6, 12, 7
		EMReadScreen mfip_case_test_initial_income, 	6, 13, 7
		EMReadScreen mfip_case_test_minor_liv_arrange, 	6, 14, 7

		EMReadScreen mfip_case_test_monthly_income, 		6, 6, 46
		EMReadScreen mfip_case_test_post_60_disq, 			6, 7, 46
		EMReadScreen mfip_case_test_residence, 				6, 8, 46
		EMReadScreen mfip_case_test_sanction_limit, 		6, 9, 46
		EMReadScreen mfip_case_test_strike, 				6, 10, 46
		EMReadScreen mfip_case_test_TANF_time_limit, 		6, 11, 46
		EMReadScreen mfip_case_test_transfer_asset, 		6, 12, 46
		EMReadScreen mfip_case_test_verif, 					6, 13, 46
		EMReadScreen mfip_case_test_275_new_spouse_income, 	6, 14, 46

		EMReadScreen mfip_fs_case_test_fail_coop_snap_qc, 		6, 17, 7
		EMReadScreen mfip_fs_case_test_opt_out_cash, 			6, 17, 46
		EMReadScreen mfip_fs_case_test_opt_out_housing_grant, 	6, 18, 46

		If mfip_fs_case_test_opt_out_cash = "FAILED" Then mfip_cash_opt_out = True
		If mfip_fs_case_test_opt_out_housing_grant = "FAILED" Then mfip_HG_opt_out = True

		mfip_case_test_appl_withdraw = trim(mfip_case_test_appl_withdraw)
		mfip_case_test_asset = trim(mfip_case_test_asset)
		mfip_case_test_death_applicant = trim(mfip_case_test_death_applicant)
		mfip_case_test_dupl_assist = trim(mfip_case_test_dupl_assist)
		mfip_case_test_elig_child = trim(mfip_case_test_elig_child)
		mfip_case_test_fail_coop = trim(mfip_case_test_fail_coop)
		mfip_case_test_fail_file = trim(mfip_case_test_fail_file)
		mfip_case_test_initial_income = trim(mfip_case_test_initial_income)
		mfip_case_test_minor_liv_arrange = trim(mfip_case_test_minor_liv_arrange)
		mfip_case_test_monthly_income = trim(mfip_case_test_monthly_income)
		mfip_case_test_post_60_disq = trim(mfip_case_test_post_60_disq)
		mfip_case_test_residence = trim(mfip_case_test_residence)
		mfip_case_test_sanction_limit = trim(mfip_case_test_sanction_limit)
		mfip_case_test_strike = trim(mfip_case_test_strike)
		mfip_case_test_TANF_time_limit = trim(mfip_case_test_TANF_time_limit)
		mfip_case_test_transfer_asset = trim(mfip_case_test_transfer_asset)
		mfip_case_test_verif = trim(mfip_case_test_verif)
		mfip_case_test_275_new_spouse_income = trim(mfip_case_test_275_new_spouse_income)
		mfip_fs_case_test_fail_coop_snap_qc = trim(mfip_fs_case_test_fail_coop_snap_qc)
		mfip_fs_case_test_opt_out_cash = trim(mfip_fs_case_test_opt_out_cash)
		mfip_fs_case_test_opt_out_housing_grant = trim(mfip_fs_case_test_opt_out_housing_grant)

		Call write_value_and_transmit("X", 7, 5)						'ASSETS
		EMReadScreen mfip_counted_asset_CASH, 	10, 6, 47
		EMReadScreen mfip_counted_asset_ACCT, 	10, 7, 47
		EMReadScreen mfip_counted_asset_SECU, 	10, 8, 47
		EMReadScreen mfip_counted_asset_CARS, 	10, 9, 47
		EMReadScreen mfip_counted_asset_SPON, 	10, 10, 47
		EMReadScreen mfip_counted_asset_total, 	10, 12, 47
		EMReadScreen mfip_counted_asset_max, 	10, 13, 47

		mfip_counted_asset_CASH = trim(mfip_counted_asset_CASH)
		mfip_counted_asset_ACCT = trim(mfip_counted_asset_ACCT)
		mfip_counted_asset_SECU = trim(mfip_counted_asset_SECU)
		mfip_counted_asset_CARS = trim(mfip_counted_asset_CARS)
		mfip_counted_asset_SPON = trim(mfip_counted_asset_SPON)
		mfip_counted_asset_total = trim(mfip_counted_asset_total)
		mfip_counted_asset_max = trim(mfip_counted_asset_max)

		transmit

		Call write_value_and_transmit("X", 13, 5)						'INITIAL INCOME
		EMReadScreen mfip_initial_income_earned, 			10, 8, 51
		EMReadScreen mfip_initial_income_deoendant_care, 	10, 9, 51
		EMReadScreen mfip_initial_income_unearned, 			10, 10, 51
		EMReadScreen mfip_initial_income_deemed, 			10, 11, 51
		EMReadScreen mfip_initial_income_cses_exclusion, 	10, 12, 51
		EMReadScreen mfip_initial_income_total, 			10, 13, 51
		EMReadScreen mfip_initial_income_family_wage_level, 10, 15, 51

		mfip_initial_income_earned = trim(mfip_initial_income_earned)
		mfip_initial_income_deoendant_care = trim(mfip_initial_income_deoendant_care)
		mfip_initial_income_unearned = trim(mfip_initial_income_unearned)
		mfip_initial_income_deemed = trim(mfip_initial_income_deemed)
		mfip_initial_income_cses_exclusion = trim(mfip_initial_income_cses_exclusion)
		mfip_initial_income_total = trim(mfip_initial_income_total)
		mfip_initial_income_family_wage_level = trim(mfip_initial_income_family_wage_level)

		'TODO - Read each person's information in the pop-ups

		PF3

		Call write_value_and_transmit("X", 14, 44)						'NEW SPOUSE 275% INCOME

		EMReadScreen mfip_12_month_start_date, 				8, 6, 46
		EMReadScreen mfip_designated_spouse_ref_numb, 		2, 7, 46
		EMReadScreen mfip_new_spouse_inc_earned, 			10, 11, 57
		EMReadScreen mfip_new_spouse_inc_unearned, 			10, 12, 57
		EMReadScreen mfip_new_spouse_inc_deemed_earned, 	10, 13, 57
		EMReadScreen mfip_new_spouse_inc_deemed_unearned, 	10, 14, 57
		EMReadScreen mfip_new_spouse_inc_total, 			10, 16, 57
		EMReadScreen mfip_275_fpg_amt, 						10, 18, 57
		EMReadScreen mfip_hh_size_fornew_spouse_calc, 		2, 18, 51

		mfip_12_month_start_date = trim(mfip_12_month_start_date)
		mfip_designated_spouse_ref_numb = trim(mfip_designated_spouse_ref_numb)
		mfip_new_spouse_inc_earned = trim(mfip_new_spouse_inc_earned)
		mfip_new_spouse_inc_unearned = trim(mfip_new_spouse_inc_unearned)
		mfip_new_spouse_inc_deemed_earned = trim(mfip_new_spouse_inc_deemed_earned)
		mfip_new_spouse_inc_deemed_unearned = trim(mfip_new_spouse_inc_deemed_unearned)
		mfip_new_spouse_inc_total = trim(mfip_new_spouse_inc_total)
		mfip_275_fpg_amt = trim(mfip_275_fpg_amt)
		mfip_hh_size_fornew_spouse_calc = trim(mfip_hh_size_fornew_spouse_calc)

		'TODO - Read each person's information in the pop-ups

		PF3

		transmit			'MFBF
		mfbf_row = 7
		Do
			EMReadScreen ref_numb, 2, mfbf_row, 3

			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If ref_numb = mfip_elig_ref_numbs(case_memb) Then
					EMReadScreen mfip_memb_cash_portion_code(case_memb), 	1, mfbf_row, 37
					EMReadScreen mfip_memb_food_portion_code(case_memb), 	1, mfbf_row, 45
					EMReadScreen mfip_memb_state_food_code(case_memb), 		1, mfbf_row, 54
					EMReadScreen mfip_memb_sanction_yn(case_memb), 			1, mfbf_row, 68

					Call write_value_and_transmit("X", mfbf_row, 62)
					EMReadScreen mfip_memb_sanction_child_support_test(case_memb),	6, 7, 12
					EMReadScreen mfip_memb_sanction_drug_felon_test(case_memb), 	6, 7, 12
					EMReadScreen mfip_memb_sanction_emp_services_test(case_memb), 	6, 7, 12
					EMReadScreen mfip_memb_sanction_fin_orient_test(case_memb), 	6, 7, 12

					EMReadScreen mfip_memb_sanction_occurence(case_memb), 1, 12, 21
					EMReadScreen mfip_memb_sanction_begin_date(case_memb), 7, 12, 40
					EMReadScreen mfip_memb_sanction_last_sanc_month(case_memb), 55, 12, 62
					transmit
				End If
			Next

			mfbf_row = mfbf_row + 1
			EMReadScreen next_ref_numb, 2, mfbf_row, 3
		Loop until next_ref_numb = "  "

		EMReadScreen mfip_case_sanction_percent, 3, 18, 28
		EMReadScreen mfip_case_sanction_vendor_yn, 1, 18, 48
		EMReadScreen mfip_case_sanction_last_vendor_month, 7, 18, 68

		mfip_case_sanction_percent = trim(mfip_case_sanction_percent)
		mfip_case_sanction_vendor_yn = trim(mfip_case_sanction_vendor_yn)
		mfip_case_sanction_last_vendor_month = trim(mfip_case_sanction_last_vendor_month)

		transmit			'MFB1

		EMReadScreen mfip_case_budg_family_wage_level, 				10, 5, 32
		EMReadScreen mfip_case_budg_monthly_earned_income, 			10, 5, 32
		EMReadScreen mfip_case_budg_wage_level_earned_inc_difference, 10, 5, 32
		EMReadScreen mfip_case_budg_transitional_standard, 			10, 5, 32
		EMReadScreen mfip_case_budg_monthly_need, 					10, 5, 32
		EMReadScreen mfip_case_budg_unearned_income, 				10, 5, 32
		EMReadScreen mfip_case_budg_deemed_income, 					10, 5, 32
		EMReadScreen mfip_case_budg_cses_exclusion, 				10, 5, 32
		EMReadScreen mfip_case_budg_unmet_need, 					10, 5, 32
		EMReadScreen mfip_case_budg_unmet_need_food_potion, 		10, 5, 32
		EMReadScreen mfip_case_budg_tribal_counted_income, 			10, 5, 32
		EMReadScreen mfip_case_budg_unmet_need_cash_portion, 		10, 5, 32
		EMReadScreen mfip_case_budg_deduction_subsidy_tribal_cses, 	10, 5, 32


		EMReadScreen mfip_case_budg_net_food_portion, 			10, 5, 71
		EMReadScreen mfip_case_budg_net_cash_portion, 			10, 5, 71
		EMReadScreen mfip_case_budg_net_unmet_need, 			10, 5, 71
		EMReadScreen mfip_case_budg_deduction_sanction_vendor, 	10, 5, 71
		EMReadScreen mfip_case_budg_unmet_neet_subtotal, 		10, 5, 71
		EMReadScreen mfip_case_budg_subtotal_food_portion, 		10, 5, 71
		EMReadScreen mfip_case_budg_food_portion_deduction, 	10, 5, 71
		EMReadScreen mfip_case_budg_entitlement_food_portion, 	10, 5, 71
		EMReadScreen mfip_case_budg_entitlement_housing_grant, 	10, 5, 71

		mfip_case_budg_family_wage_level = trim(mfip_case_budg_family_wage_level)
		mfip_case_budg_monthly_earned_income = trim(mfip_case_budg_monthly_earned_income)
		mfip_case_budg_wage_level_earned_inc_difference = trim(mfip_case_budg_wage_level_earned_inc_difference)
		mfip_case_budg_transitional_standard = trim(mfip_case_budg_transitional_standard)
		mfip_case_budg_monthly_need = trim(mfip_case_budg_monthly_need)
		mfip_case_budg_unearned_income = trim(mfip_case_budg_unearned_income)
		mfip_case_budg_deemed_income = trim(mfip_case_budg_deemed_income)
		mfip_case_budg_cses_exclusion = trim(mfip_case_budg_cses_exclusion)
		mfip_case_budg_unmet_need = trim(mfip_case_budg_unmet_need)
		mfip_case_budg_unmet_need_food_potion = trim(mfip_case_budg_unmet_need_food_potion)
		mfip_case_budg_tribal_counted_income = trim(mfip_case_budg_tribal_counted_income)
		mfip_case_budg_unmet_need_cash_portion = trim(mfip_case_budg_unmet_need_cash_portion)
		mfip_case_budg_deduction_subsidy_tribal_cses = trim(mfip_case_budg_deduction_subsidy_tribal_cses)

		mfip_case_budg_net_food_portion = trim(mfip_case_budg_net_food_portion)
		mfip_case_budg_net_cash_portion = trim(mfip_case_budg_net_cash_portion)
		mfip_case_budg_net_unmet_need = trim(mfip_case_budg_net_unmet_need)
		mfip_case_budg_deduction_sanction_vendor = trim(mfip_case_budg_deduction_sanction_vendor)
		mfip_case_budg_unmet_neet_subtotal = trim(mfip_case_budg_unmet_neet_subtotal)
		mfip_case_budg_subtotal_food_portion = trim(mfip_case_budg_subtotal_food_portion)
		mfip_case_budg_food_portion_deduction = trim(mfip_case_budg_food_portion_deduction)
		mfip_case_budg_entitlement_food_portion = trim(mfip_case_budg_entitlement_food_portion)
		mfip_case_budg_entitlement_housing_grant = trim(mfip_case_budg_entitlement_housing_grant)

		' Call write_value_and_transmit("X", 6, 3)		'TODO member specific EARNED INCOME
		' Call write_value_and_transmit("X", 11, 3)		'TODO member specific UNEARNED INCOME
		' Call write_value_and_transmit("X", 12, 3)		'TODO member specific DEEMED INCOME

		Call write_value_and_transmit("X", 13, 3)		'Child Support Exclusion'
		EMReadScreen mfip_budg_cses_excln_cses_income, 10, 9, 52
		EMReadScreen mfip_budg_cses_excln_child_count, 2, 11, 37
		EMReadScreen mfip_budg_cses_excln_total, 10, 13, 52

		mfip_budg_cses_excln_cses_income = trim(mfip_budg_cses_excln_cses_income)
		mfip_budg_cses_excln_child_count = trim(mfip_budg_cses_excln_child_count)
		mfip_budg_cses_excln_total = trim(mfip_budg_cses_excln_total)

		transmit

		' Call write_value_and_transmit("X", 16, 5)		'TODO member specific TRIBAL INCOME
		' Call write_value_and_transmit("X", 18, 5)		'TODO member specific SUBSIDY

		Call write_value_and_transmit("X", 8, 44)		'Sanction and Vendor
		EMReadScreen mfip_case_budg_10_perc_sanc, 					10, 7, 55
		EMReadScreen mfip_case_budg_unmet_need_after_pre_vndr_sanc, 10, 7, 55
		EMReadScreen mfip_case_budg_sanc_calc_food_portion, 		10, 7, 55
		EMReadScreen mfip_case_budg_sanc_calc_cash_portion, 		10, 7, 55
		EMReadScreen mfip_case_budg_pot_mand_vndr_pymts, 			10, 7, 55
		EMReadScreen mfip_case_budg_30_perc_sanc, 					10, 7, 55

		mfip_case_budg_10_perc_sanc = trim(mfip_case_budg_10_perc_sanc)
		mfip_case_budg_unmet_need_after_pre_vndr_sanc = trim(mfip_case_budg_unmet_need_after_pre_vndr_sanc)
		mfip_case_budg_sanc_calc_food_portion = trim(mfip_case_budg_sanc_calc_food_portion)
		mfip_case_budg_sanc_calc_cash_portion = trim(mfip_case_budg_sanc_calc_cash_portion)
		mfip_case_budg_pot_mand_vndr_pymts = trim(mfip_case_budg_pot_mand_vndr_pymts)
		mfip_case_budg_30_perc_sanc = trim(mfip_case_budg_30_perc_sanc)
		transmit

		Call write_value_and_transmit("X", 12, 44)		'Food portion Deduction
		EMReadScreen mfip_case_budg_non_citzn_fs_inelig_pers_count, 1, 10, 17
		EMReadScreen mfip_case_budg_non_citzn_fs_inelig_amt, 		10, 10, 45
		EMReadScreen mfip_case_budg_other_fs_inelig_pers_count, 	1, 12, 17
		EMReadScreen mfip_case_budg_other_fs_inelig_amt, 			10, 12, 45

		mfip_case_budg_non_citzn_fs_inelig_pers_count = trim(mfip_case_budg_non_citzn_fs_inelig_pers_count)
		mfip_case_budg_non_citzn_fs_inelig_amt = trim(mfip_case_budg_non_citzn_fs_inelig_amt)
		mfip_case_budg_other_fs_inelig_pers_count = trim(mfip_case_budg_other_fs_inelig_pers_count)
		mfip_case_budg_other_fs_inelig_amt = trim(mfip_case_budg_other_fs_inelig_amt)
		transmit

		transmit			'MFB2
		EmReadScreen mfip_case_budg_prorate_date, 8, 5, 19

		EMReadScreen mfip_case_budg_fed_food_benefit, 10, 8, 32
		EMReadScreen mfip_case_budg_food_prorated_amt, 10, 8, 32
		EMReadScreen mfip_case_budg_entitlement_cash_portion, 10, 10, 32
		EMReadScreen mfip_case_budg_mand_sanc_vendor, 10, 10, 32
		EMReadScreen mfip_case_budg_net_cash_portion, 10, 10, 32
		EMReadScreen mfip_case_budg_cash_prorated_amt, 10, 10, 32
		EMReadScreen mfip_case_budg_state_food_benefit, 10, 10, 32
		EMReadScreen mfip_case_budg_state_food_prorated_amt, 10, 10, 32
		' EMReadScreen mfip_case_budg_entitlement_cash_portion, 10, 10, 32

		EMReadScreen mfip_case_budg_grant_amount, 10, 5, 71
		EMReadScreen mfip_case_budg_amt_already_issued, 10, 88, 71
		EMReadScreen mfip_case_budg_supplement_due, 10, 88, 71
		EMReadScreen mfip_case_budg_overpayment, 10, 88, 71
		EMReadScreen mfip_case_budg_adjusted_grant_amt, 10, 88, 71
		EMReadScreen mfip_case_budg_recoupment, 10, 88, 71
		EMReadScreen mfip_case_budg_total_food_issuance, 10, 88, 71
		EMReadScreen mfip_case_budg_total_cash_issuance, 10, 88, 71
		EMReadScreen mfip_case_budg_total_housing_grant_issuance, 10, 88, 71

		mfip_case_budg_prorate_date = trim(mfip_case_budg_prorate_date)
		mfip_case_budg_fed_food_benefit = trim(mfip_case_budg_fed_food_benefit)
		mfip_case_budg_food_prorated_amt = trim(mfip_case_budg_food_prorated_amt)
		mfip_case_budg_entitlement_cash_portion = trim(mfip_case_budg_entitlement_cash_portion)
		mfip_case_budg_mand_sanc_vendor = trim(mfip_case_budg_mand_sanc_vendor)
		mfip_case_budg_net_cash_portion = trim(mfip_case_budg_net_cash_portion)
		mfip_case_budg_cash_prorated_amt = trim(mfip_case_budg_cash_prorated_amt)
		mfip_case_budg_state_food_benefit = trim(mfip_case_budg_state_food_benefit)
		mfip_case_budg_state_food_prorated_amt = trim(mfip_case_budg_state_food_prorated_amt)
		mfip_case_budg_grant_amount = trim(mfip_case_budg_grant_amount)
		mfip_case_budg_amt_already_issued = trim(mfip_case_budg_amt_already_issued)
		mfip_case_budg_supplement_due = trim(mfip_case_budg_supplement_due)
		mfip_case_budg_overpayment = trim(mfip_case_budg_overpayment)
		mfip_case_budg_adjusted_grant_amt = trim(mfip_case_budg_adjusted_grant_amt)
		mfip_case_budg_recoupment = trim(mfip_case_budg_recoupment)
		mfip_case_budg_total_food_issuance = trim(mfip_case_budg_total_food_issuance)
		mfip_case_budg_total_cash_issuance = trim(mfip_case_budg_total_cash_issuance)
		mfip_case_budg_total_housing_grant_issuance = trim(mfip_case_budg_total_housing_grant_issuance)

		' Call write_value_and_transmit("X", 15, 3)			'State food benefit pop-up - I think this is duplicate
		Call write_value_and_transmit("X", 9, 44)			'Supplement pop-up
		EMReadScreen mfip_case_budg_food_supplement, 		10, 11, 32
		EMReadScreen mfip_case_budg_state_food_supplement, 	10, 16, 32
		EMReadScreen mfip_case_budg_cash_supplement, 		10, 11, 68
		EMReadScreen mfip_case_budg_housing_grant_supplement, 10, 16, 68

		mfip_case_budg_food_supplement = trim(mfip_case_budg_food_supplement)
		mfip_case_budg_state_food_supplement = trim(mfip_case_budg_state_food_supplement)
		mfip_case_budg_cash_supplement = trim()
		mfip_case_budg_housing_grant_supplement = trim(mfip_case_budg_housing_grant_supplement)
		transmit

		' Call write_value_and_transmit("X", 10, 44)			'Overpayment pop-up - MAYBE WE DON"T NEED THIS?
		Call write_value_and_transmit("X", 12, 44)			'Recoupment pop-up
		EMReadScreen mfip_case_budg_cash_recoupment, 10, 7, 51
		EMReadScreen mfip_case_budg_state_food_recoupment, 10, 7, 51
		EMReadScreen mfip_case_budg_food_recoupment, 10, 7, 51

		mfip_case_budg_cash_recoupment = trim(mfip_case_budg_cash_recoupment)
		mfip_case_budg_state_food_recoupment = trim(mfip_case_budg_state_food_recoupment)
		mfip_case_budg_food_recoupment = trim(mfip_case_budg_food_recoupment)
		transmit

		Call write_value_and_transmit("X", 14, 44)			'Total Food issuance pop-up
		EMReadScreen mfip_case_budg_fed_food_memb_count, 1, 7, 17
		EMReadScreen mfip_case_budg_fed_food_benefit_amt, 10, 7, 45
		EMReadScreen mfip_case_budg_state_food_memb_count, 1, 9, 17
		EMReadScreen mfip_case_budg_state_food_benefit_amt, 10, 9, 45

		mfip_case_budg_fed_food_memb_count = trim(mfip_case_budg_fed_food_memb_count)
		mfip_case_budg_fed_food_benefit_amt = trim(mfip_case_budg_fed_food_benefit_amt)
		mfip_case_budg_state_food_memb_count = trim(mfip_case_budg_state_food_memb_count)
		mfip_case_budg_state_food_benefit_amt = trim(mfip_case_budg_state_food_benefit_amt)
		transmit

		Call write_value_and_transmit("X", 1, 44)			'Total Cash Issuance pop-up
		EMReadScreen mfip_case_budg_tanf_cash_memb_count, 1, 8, 17
		EMReadScreen mfip_case_budg_tanf_cash_benefit_amt, 10, 8, 45
		EMReadScreen mfip_case_budg_state_cash_memb_count, 1, 10, 17
		EMReadScreen mfip_case_budg_state_cash_benefit_amt, 10, 10, 45

		mfip_case_budg_tanf_cash_memb_count = trim(mfip_case_budg_tanf_cash_memb_count)
		mfip_case_budg_tanf_cash_benefit_amt = trim(mfip_case_budg_tanf_cash_benefit_amt)
		mfip_case_budg_state_cash_memb_count = trim(mfip_case_budg_state_cash_memb_count)
		mfip_case_budg_state_cash_benefit_amt = trim(mfip_case_budg_state_cash_benefit_amt)
		transmit

		' Call write_value_and_transmit("X", 16, 44)			'MFIP Housing Grant Issuance pop-up - there is not federal housing grant
		transmit			'MFSM

		EMReadScreen mfip_approved_date, 8, 3, 14
		EMReadScreen mfip_process_date, 8, 2, 73
		EMReadScreen mfip_prev_approval, 4, 3, 73

		EMReadScreen mfip_case_last_approval_date, 8, 5, 31
		EMReadScreen mfip_case_current_prog_status, 12, 6, 31
		EMReadScreen mfip_case_eligibility_result, 12,  7, 31
		EMReadScreen mfip_case_hrf_reporting, 12, 8, 31
		EMReadScreen mfip_case_source_of_info, 4, 9, 31
		EMReadScreen mfip_case_benefit_impact, 12, 10, 31
		EMReadScreen mfip_case_review_date, 8, 11, 31
		EMReadScreen mfip_case_budget_cycle, 12, 12, 31
		EMReadScreen mfip_case_vendor_reason_code, 2, 13, 31

		EMReadScreen mfip_case_responsible_county, 2, 5, 73
		EMReadScreen mfip_case_service_county, 2, 6, 73
		EMReadScreen mfip_case_asst_unit_caregivers, 1, 7, 73
		EMReadScreen mfip_case_asst_unit_children, 2, 8, 73
		EMReadScreen mfip_case_total_assets, 10, 9, 71
		EMReadScreen mfip_case_maximum_assets, 10, 10, 71
		EMReadScreen mfip_case_summary_grant_amount, 10, 11, 71
		EMReadScreen mfip_case_summary_net_grant_amount, 10, 13, 71
		EMReadScreen mfip_case_summary_cash_portion, 10, 14, 71
		EMReadScreen mfip_case_summary_food_portion, 10, 15, 71
		EMReadScreen mfip_case_summary_housing_grant, 10, 16, 71

		If mfip_case_vendor_reason_code = "01" Then mfip_case_vendor_reason_info = "Client Request"
		If mfip_case_vendor_reason_code = "05" Then mfip_case_vendor_reason_info = "Money Mismanagement"
		If mfip_case_vendor_reason_code = "06" Then mfip_case_vendor_reason_info = "Social Service Non-Coop"
		If mfip_case_vendor_reason_code = "07" Then mfip_case_vendor_reason_info = "Residing in a Facility"
		If mfip_case_vendor_reason_code = "21" Then mfip_case_vendor_reason_info = "MFIP Sanction Related Vendor"
		If mfip_case_vendor_reason_code = "22" Then mfip_case_vendor_reason_info = "Convicted Drug Felon in Household"

		mfip_prev_approval = trim(mfip_prev_approval)
		mfip_case_last_approval_date = trim(mfip_case_last_approval_date)

		mfip_case_current_prog_status = trim(mfip_case_current_prog_status)
		mfip_case_eligibility_result = trim(mfip_case_eligibility_result)
		mfip_case_hrf_reporting = trim(mfip_case_hrf_reporting)
		mfip_case_source_of_info = trim(mfip_case_source_of_info)
		mfip_case_benefit_impact = trim(mfip_case_benefit_impact)

		mfip_case_budget_cycle = trim(mfip_case_budget_cycle)
		mfip_case_vendor_reason_code = trim(mfip_case_vendor_reason_code)

		mfip_case_asst_unit_caregivers = trim(mfip_case_asst_unit_caregivers)
		mfip_case_asst_unit_children = trim(mfip_case_asst_unit_children)
		mfip_case_total_assets = trim(mfip_case_total_assets)
		mfip_case_maximum_assets = trim(mfip_case_maximum_assets)
		mfip_case_summary_grant_amount = trim(mfip_case_summary_grant_amount)
		mfip_case_summary_net_grant_amount = trim(mfip_case_summary_net_grant_amount)
		mfip_case_summary_cash_portion = trim(mfip_case_summary_cash_portion)
		mfip_case_summary_food_portion = trim(mfip_case_summary_food_portion)
		mfip_case_summary_housing_grant = trim(mfip_case_summary_housing_grant)
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

const snap_elig_indicator			= 12
const mfip_elig_indicator			= 13


' const fs_request_yn_const			= 12
' const fs_memb_code_const			= 13
' const fs_memb_status_info_const		= 14
' const fs_memb_counted_const			= 15
' const fs_memb_state_food_const		= 16
' const fs_memb_elig_status_const		= 17
' const fs_memb_begin_date_const		= 18
' const fs_memb_budg_cycle_const		= 19
' const fs_memb_abawd_const			= 20
' const fs_memb_absence_const			= 21
' const fs_memb_roomer_const			= 22
' const fs_memb_boarder_const			= 23
' const fs_memb_citizenship_const		= 24
' const fs_memb_citizenship_coop_const = 25
' const fs_memb_cmdty_const			= 26
' const fs_memb_disq_const			= 27
' const fs_memb_dupl_assist_const		= 28
' const fs_memb_fraud_const			= 29
' const fs_memb_eligible_student_const = 30
' const fs_memb_institution_const		= 31
' const fs_memb_mfip_elig_const		= 32
' const fs_memb_non_applcnt_const		= 33
' const fs_memb_residence_const		= 34
' const fs_memb_ssn_coop_const		= 35
' const fs_memb_unit_memb_const		= 36
' const fs_memb_work_reg_const		= 37
' const fs_memb_drug_felon_test_const	= 38

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
