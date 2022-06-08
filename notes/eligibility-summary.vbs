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
		EMReadScreen elig_date, 8, row, 26
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




'DECLARATIONS===============================================================================================================
class dwp_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result

	public dwp_elig_ref_numbs()
	public dwp_elig_membs_full_name()
	public dwp_elig_membs_request_yn()
	public dwp_elig_membs_member_code()
	public dwp_elig_membs_member_info()
	public dwp_elig_membs_funding_source_code()
	public dwp_elig_membs_funding_source_info()
	public dwp_elig_membs_elig_status()
	public dwp_elig_membs_begin_date()
	public dwp_elig_membs_adult_or_child()
	public dwp_elig_membs_test_absence()
	public dwp_elig_membs_test_child_age()
	public dwp_elig_membs_test_citizenship()
	public dwp_elig_membs_test_citizenship_verif()
	public dwp_elig_membs_test_dupl_assistance()
	public dwp_elig_membs_test_foster_care()
	public dwp_elig_membs_test_fraud()
	public dwp_elig_membs_test_minor_living_arrangement()
	public dwp_elig_membs_test_post_60_removal()
	public dwp_elig_membs_test_ssi()
	public dwp_elig_membs_test_ssn_coop()
	public dwp_elig_membs_test_unit_member()
	public dwp_elig_membs_test_unlawful_conduct()
	public dwp_elig_membs_es_status_code()
	public dwp_elig_membs_es_status_info()

	public dwp_elig_case_test_application_withdrawn
	public dwp_elig_case_test_assets
	public dwp_elig_case_test_CS_disqualification
	public dwp_elig_case_test_death_of_applicant
	public dwp_elig_case_test_dupl_assistance
	public dwp_elig_case_test_eligible_child
	public dwp_elig_case_test_ES_disqualification
	public dwp_elig_case_test_fail_coop
	public dwp_elig_case_test_four_month_limit
	public dwp_elig_case_test_initial_income
	public dwp_elig_case_test_MFIP_conversion
	public dwp_elig_case_test_residence
	public dwp_elig_case_test_strike
	public dwp_elig_case_test_TANF_time_limit
	public dwp_elig_case_test_transfer_of_assets
	public dwp_elig_case_test_verif
	public dwp_elig_case_test_new_spouse_income
	public dwp_elig_asset_CASH
	public dwp_elig_asset_ACCT
	public dwp_elig_asset_SECU
	public dwp_elig_asset_CARS
	public dwp_elig_asset_SPON
	public dwp_elig_asset_total
	public dwp_elig_asset_maximum
	public dwp_elig_test_fail_coop_applied_other_benefits
	public dwp_elig_test_fail_coop_provide_requested_info
	public dwp_elig_test_fail_coop_IEVS
	public dwp_elig_test_fail_coop_vendor_info
	public dwp_elig_initial_counted_earned_income
	public dwp_elig_initial_dependent_care_expense
	public dwp_elig_initial_counted_unearned_incom
	public dwp_elig_initial_counted_deemed_income
	public dwp_elig_initial_child_support_exclusion
	public dwp_elig_initial_total_counted_income
	public dwp_elig_initial_family_wage_level
	public dwp_elig_test_verif_ACCT
	public dwp_elig_test_verif_BUSI
	public dwp_elig_test_verif_CARS
	public dwp_elig_test_verif_JOBS
	public dwp_elig_test_verif_MEMB_dob
	public dwp_elig_test_verif_MEMB_id
	public dwp_elig_test_verif_PARE
	public dwp_elig_test_verif_PREG
	public dwp_elig_test_verif_RBIC
	public dwp_elig_test_verif_ADDR
	public dwp_elig_test_verif_SCHL
	public dwp_elig_test_verif_SECU
	public dwp_elig_test_verif_SPON
	public dwp_elig_test_verif_UNEA

	public dwp_elig_budg_shel_rent_mortgage
	public dwp_elig_budg_shel_property_tax
	public dwp_elig_budg_shel_house_insurance
	public dwp_elig_budg_hest_electricity
	public dwp_elig_budg_hest_heat_air
	public dwp_elig_budg_hest_water_sewer_garbage
	public dwp_elig_budg_hest_phone
	public dwp_elig_budg_shel_other
	public dwp_elig_budg_total_shelter_costs
	public dwp_elig_budg_personal_needs
	public dwp_elig_budg_total_DWP_need
	public dwp_elig_budg_earned_income
	public dwp_elig_budg_unearned_income
	public dwp_elig_budg_deemed_income
	public dwp_elig_budg_child_support_exclusion
	public dwp_elig_budg_budget_month_total
	public dwp_elig_budg_prior_low
	public dwp_elig_budg_DWP_countable_income
	public dwp_elig_budg_unmet_need
	public dwp_elig_budg_DWP_max_grant
	public dwp_elig_budg_DWP_grant
	public dwp_elig_cses_income
	public dwp_elig_child_count

	public dwp_elig_prorated_date
	public dwp_elig_prorated_amount
	public dwp_elig_amount_already_issued
	public dwp_elig_supplement_due
	public dwp_elig_overpayment
	public dwp_elig_adjusted_grant_amount
	public dwp_elig_recoupment_amount
	public dwp_elig_shelter_benefit_grant
	public dwp_elig_personal_needs_grant
	public dwp_elig_overpayment_fed_hh_count
	public dwp_elig_overpayment_fed_amount
	public dwp_elig_overpayment_state_hh_count
	public dwp_elig_overpayment_state_amount
	public dwp_elig_adjusted_grant_fed_hh_count
	public dwp_elig_adjusted_grant_fed_amount
	public dwp_elig_adjusted_grant_state_hh_count
	public dwp_elig_adjusted_grant_state_amount

	public dwp_approved_date
	public dwp_process_date
	public dwp_prev_approval
	public dwp_case_last_approval_date
	public dwp_case_current_prog_status
	public dwp_case_eligibility_result
	public dwp_case_source_of_info
	public dwp_case_benefit_impact
	public dwp_case_4th_month_of_elig
	public dwp_case_caregivers_have_es_plan
	public dwp_case_responsible_county
	public dwp_case_service_county
	public dwp_case_asst_unit_caregivers
	public dwp_case_asst_unit_children
	public dwp_case_total_assets
	public dwp_case_maximum_assets
	public dwp_case_summary_grant_amount
	public dwp_case_summary_net_grant_amount
	public dwp_case_summary_shelter_benefit_portion
	public dwp_case_summary_personal_needs_portion


	public sub read_elig()
		call navigate_to_MAXIS_screen("ELIG", "DWP ")
		EMWriteScreen elig_footer_month, 20, 56
		EMWriteScreen elig_footer_year, 20, 59
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)

		ReDim dwp_elig_ref_numbs(0)
		ReDim dwp_elig_membs_full_name(0)
		ReDim dwp_elig_membs_request_yn(0)
		ReDim dwp_elig_membs_member_code(0)
		ReDim dwp_elig_membs_member_info(0)
		ReDim dwp_elig_membs_funding_source_code(0)
		ReDim dwp_elig_membs_funding_source_info(0)
		ReDim dwp_elig_membs_elig_status(0)
		ReDim dwp_elig_membs_begin_date(0)
		ReDim dwp_elig_membs_adult_or_child(0)
		ReDim dwp_elig_membs_test_absence(0)
		ReDim dwp_elig_membs_test_child_age(0)
		ReDim dwp_elig_membs_test_citizenship(0)
		ReDim dwp_elig_membs_test_citizenship_verif(0)
		ReDim dwp_elig_membs_test_dupl_assistance(0)
		ReDim dwp_elig_membs_test_foster_care(0)
		ReDim dwp_elig_membs_test_fraud(0)
		ReDim dwp_elig_membs_test_minor_living_arrangement(0)
		ReDim dwp_elig_membs_test_post_60_removal(0)
		ReDim dwp_elig_membs_test_ssi(0)
		ReDim dwp_elig_membs_test_ssn_coop(0)
		ReDim dwp_elig_membs_test_unit_member(0)
		ReDim dwp_elig_membs_test_unlawful_conduct(0)
		ReDim dwp_elig_membs_es_status_code(0)
		ReDim dwp_elig_membs_es_status_info(0)

		row = 7
		elig_memb_count = 0
		Do
			EMReadScreen ref_numb, 2, row, 5

			ReDim preserve dwp_elig_ref_numbs(elig_memb_count)
			ReDim preserve dwp_elig_membs_full_name(elig_memb_count)
			ReDim preserve dwp_elig_membs_request_yn(elig_memb_count)
			ReDim preserve dwp_elig_membs_member_code(elig_memb_count)
			ReDim preserve dwp_elig_membs_member_info(elig_memb_count)
			ReDim preserve dwp_elig_membs_funding_source_code(elig_memb_count)
			ReDim preserve dwp_elig_membs_funding_source_info(elig_memb_count)
			ReDim preserve dwp_elig_membs_elig_status(elig_memb_count)
			ReDim preserve dwp_elig_membs_begin_date(elig_memb_count)
			ReDim preserve dwp_elig_membs_adult_or_child(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_absence(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_child_age(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_citizenship(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_citizenship_verif(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_dupl_assistance(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_foster_care(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_fraud(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_minor_living_arrangement(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_post_60_removal(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_ssi(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_ssn_coop(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_unit_member(elig_memb_count)
			ReDim preserve dwp_elig_membs_test_unlawful_conduct(elig_memb_count)
			ReDim preserve dwp_elig_membs_es_status_code(elig_memb_count)
			ReDim preserve dwp_elig_membs_es_status_info(elig_memb_count)

			dwp_elig_ref_numbs(elig_memb_count) = ref_numb
			EMReadScreen full_name_information, 20, row, 9
			full_name_information = trim(full_name_information)
			name_array = split(full_name_information, " ")
			For each name_parts in name_array
				If len(name_parts) <> 1 Then dwp_elig_membs_full_name(elig_memb_count) = dwp_elig_membs_full_name(elig_memb_count) & " " & name_parts
			Next
			dwp_elig_membs_full_name(elig_memb_count) = trim((dwp_elig_membs_full_name(elig_memb_count)))

			EMReadScreen dwp_elig_membs_request_yn(elig_memb_count), 1, row, 31
			EMReadScreen dwp_elig_membs_member_code(elig_memb_count), 1, row, 35
			EMReadScreen dwp_elig_membs_funding_source_code(elig_memb_count), 1, row, 53
			EMReadScreen dwp_elig_membs_elig_status(elig_memb_count), 12, row, 57
			EMReadScreen dwp_elig_membs_begin_date(elig_memb_count), 8, row, 73

			dwp_elig_membs_elig_status(elig_memb_count) = trim(dwp_elig_membs_elig_status(elig_memb_count))

			If dwp_elig_membs_member_code(elig_memb_count) = "A" Then dwp_elig_membs_member_info(elig_memb_count) = "Eligible"
			If dwp_elig_membs_member_code(elig_memb_count) = "D" Then dwp_elig_membs_member_info(elig_memb_count) = "SSI/IVE/Adoption Assistance Recipient"
			If dwp_elig_membs_member_code(elig_memb_count) = "F" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Deemer"
			If dwp_elig_membs_member_code(elig_memb_count) = "G" Then dwp_elig_membs_member_info(elig_memb_count) = "Parent of Minor Caregiver, Deemer"
			If dwp_elig_membs_member_code(elig_memb_count) = "H" Then dwp_elig_membs_member_info(elig_memb_count) = "Other Deemer"
			If dwp_elig_membs_member_code(elig_memb_count) = "I" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Pare of Unit"
			If dwp_elig_membs_member_code(elig_memb_count) = "J" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Deemer"
			If dwp_elig_membs_member_code(elig_memb_count) = "N" Then dwp_elig_membs_member_info(elig_memb_count) = "Not Counted"

			If dwp_elig_membs_funding_source_code(elig_memb_count) = "F" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Federal Funds (TANF Cash)"
			If dwp_elig_membs_funding_source_code(elig_memb_count) = "S" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "State Funds (Cash)"
			If dwp_elig_membs_funding_source_code(elig_memb_count) = "I" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Ineligible for DWP"
			If dwp_elig_membs_funding_source_code(elig_memb_count) = "N" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Not Applicable"

			Call write_value_and_transmit("X", row, 3)		'open member test information
			EMReadScreen dwp_elig_membs_adult_or_child(elig_memb_count), 1, 7, 51

			If dwp_elig_membs_adult_or_child(elig_memb_count) = "A" Then dwp_elig_membs_adult_or_child(elig_memb_count) = "Adult"
			If dwp_elig_membs_adult_or_child(elig_memb_count) = "C" Then dwp_elig_membs_adult_or_child(elig_memb_count) = "Child"

			EMReadScreen dwp_elig_membs_test_absence(elig_memb_count), 			6, 10, 7
			EMReadScreen dwp_elig_membs_test_child_age(elig_memb_count), 		6, 11, 7
			EMReadScreen dwp_elig_membs_test_citizenship(elig_memb_count), 		6, 12, 7
			EMReadScreen dwp_elig_membs_test_citizenship_verif(elig_memb_count), 6, 13, 7
			EMReadScreen dwp_elig_membs_test_dupl_assistance(elig_memb_count), 	6, 14, 7
			EMReadScreen dwp_elig_membs_test_foster_care(elig_memb_count), 		6, 15, 7
			EMReadScreen dwp_elig_membs_test_fraud(elig_memb_count), 			6, 16, 7

			EMReadScreen dwp_elig_membs_test_minor_living_arrangement(elig_memb_count), 6, 10, 43
			EMReadScreen dwp_elig_membs_test_post_60_removal(elig_memb_count), 			6, 11, 43
			EMReadScreen dwp_elig_membs_test_ssi(elig_memb_count), 						6, 12, 43
			EMReadScreen dwp_elig_membs_test_ssn_coop(elig_memb_count), 				6, 13, 43
			EMReadScreen dwp_elig_membs_test_unit_member(elig_memb_count), 				6, 14, 43
			EMReadScreen dwp_elig_membs_test_unlawful_conduct(elig_memb_count), 		6, 15, 43

			dwp_elig_membs_test_absence(elig_memb_count) = trim(dwp_elig_membs_test_absence(elig_memb_count))
			dwp_elig_membs_test_child_age(elig_memb_count) = trim(dwp_elig_membs_test_child_age(elig_memb_count))
			dwp_elig_membs_test_citizenship(elig_memb_count) = trim(dwp_elig_membs_test_citizenship(elig_memb_count))
			dwp_elig_membs_test_citizenship_verif(elig_memb_count) = trim(dwp_elig_membs_test_citizenship_verif(elig_memb_count))
			dwp_elig_membs_test_dupl_assistance(elig_memb_count) = trim(dwp_elig_membs_test_dupl_assistance(elig_memb_count))
			dwp_elig_membs_test_foster_care(elig_memb_count) = trim(dwp_elig_membs_test_foster_care(elig_memb_count))
			dwp_elig_membs_test_fraud(elig_memb_count) = trim(dwp_elig_membs_test_fraud(elig_memb_count))

			dwp_elig_membs_test_minor_living_arrangement(elig_memb_count) = trim(dwp_elig_membs_test_minor_living_arrangement(elig_memb_count))
			dwp_elig_membs_test_post_60_removal(elig_memb_count) = trim(dwp_elig_membs_test_post_60_removal(elig_memb_count))
			dwp_elig_membs_test_ssi(elig_memb_count) = trim(dwp_elig_membs_test_ssi(elig_memb_count))
			dwp_elig_membs_test_ssn_coop(elig_memb_count) = trim(dwp_elig_membs_test_ssn_coop(elig_memb_count))
			dwp_elig_membs_test_unit_member(elig_memb_count) = trim(dwp_elig_membs_test_unit_member(elig_memb_count))
			dwp_elig_membs_test_unlawful_conduct(elig_memb_count) = trim(dwp_elig_membs_test_unlawful_conduct(elig_memb_count))

			transmit

			Call write_value_and_transmit("X", row, 69)		'open member EMPS information
			EMReadScreen emps_exists_for_memb, 19, 24, 2
			If emps_exists_for_memb = "EMPS DOES NOT EXIST" Then
				EMWriteScreen " ", row, 69
			Else
				EMReadScreen dwp_elig_membs_es_status_code(elig_memb_count), 2, 9, 22
				EMReadScreen dwp_elig_membs_es_status_info(elig_memb_count), 30, 9, 25

				dwp_elig_membs_es_status_code(elig_memb_count) = trim(dwp_elig_membs_es_status_code(elig_memb_count))
				dwp_elig_membs_es_status_info(elig_memb_count) = trim(dwp_elig_membs_es_status_info(elig_memb_count))
				transmit
			End If

			row = row + 1
			elig_memb_count = elig_memb_count + 1
			EMReadScreen next_ref_numb, 2, row, 6
		Loop until next_ref_numb = "  "

		transmit 		'going to the next panel - DWCR
		MsgBox "AT DWCR"

		EMReadScreen dwp_elig_case_test_application_withdrawn, 6, 6, 7
		EMReadScreen dwp_elig_case_test_assets, 6, 6, 7
		EMReadScreen dwp_elig_case_test_CS_disqualification, 6, 6, 7
		EMReadScreen dwp_elig_case_test_death_of_applicant, 6, 6, 7
		EMReadScreen dwp_elig_case_test_dupl_assistance, 6, 6, 7
		EMReadScreen dwp_elig_case_test_eligible_child, 6, 6, 7
		EMReadScreen dwp_elig_case_test_ES_disqualification, 6, 6, 7
		EMReadScreen dwp_elig_case_test_fail_coop, 6, 6, 7
		EMReadScreen dwp_elig_case_test_four_month_limit, 6, 6, 7

		EMReadScreen dwp_elig_case_test_initial_income, 6, 6, 45
		EMReadScreen dwp_elig_case_test_MFIP_conversion, 6, 6, 45
		EMReadScreen dwp_elig_case_test_residence, 6, 6, 45
		EMReadScreen dwp_elig_case_test_strike, 6, 6, 45
		EMReadScreen dwp_elig_case_test_TANF_time_limit, 6, 6, 45
		EMReadScreen dwp_elig_case_test_transfer_of_assets, 6, 6, 45
		EMReadScreen dwp_elig_case_test_verif, 6, 6, 45

		EMReadScreen dwp_elig_case_test_new_spouse_income, 6, 6, 45

		dwp_elig_case_test_application_withdrawn = trim(dwp_elig_case_test_application_withdrawn)
		dwp_elig_case_test_assets = trim(dwp_elig_case_test_assets)
		dwp_elig_case_test_CS_disqualification = trim(dwp_elig_case_test_CS_disqualification)
		dwp_elig_case_test_death_of_applicant = trim(dwp_elig_case_test_death_of_applicant)
		dwp_elig_case_test_dupl_assistance = trim(dwp_elig_case_test_dupl_assistance)
		dwp_elig_case_test_eligible_child = trim(dwp_elig_case_test_eligible_child)
		dwp_elig_case_test_ES_disqualification = trim(dwp_elig_case_test_ES_disqualification)
		dwp_elig_case_test_fail_coop = trim(dwp_elig_case_test_fail_coop)
		dwp_elig_case_test_four_month_limit = trim(dwp_elig_case_test_four_month_limit)

		dwp_elig_case_test_initial_income = trim(dwp_elig_case_test_initial_income)
		dwp_elig_case_test_MFIP_conversion = trim(dwp_elig_case_test_MFIP_conversion)
		dwp_elig_case_test_residence = trim(dwp_elig_case_test_residence)
		dwp_elig_case_test_strike = trim(dwp_elig_case_test_strike)
		dwp_elig_case_test_TANF_time_limit = trim(dwp_elig_case_test_TANF_time_limit)
		dwp_elig_case_test_transfer_of_assets = trim(dwp_elig_case_test_transfer_of_assets)
		dwp_elig_case_test_verif = trim(dwp_elig_case_test_verif)

		dwp_elig_case_test_new_spouse_income = trim(dwp_elig_case_test_new_spouse_income)

		If dwp_elig_case_test_assets <> "NA" Then
			Call write_value_and_transmit("X", 7, 5)

			EMReadScreen dwp_elig_asset_CASH, 9, 8, 54
			EMReadScreen dwp_elig_asset_ACCT, 9, 9, 54
			EMReadScreen dwp_elig_asset_SECU, 9, 10, 54
			EMReadScreen dwp_elig_asset_CARS, 9, 11, 54
			EMReadScreen dwp_elig_asset_SPON, 9, 12, 54

			EMReadScreen dwp_elig_asset_total, 9, 17, 54
			EMReadScreen dwp_elig_asset_maximum, 9, 18, 54

			dwp_elig_asset_CASH = trim(dwp_elig_asset_CASH)
			dwp_elig_asset_ACCT = trim(dwp_elig_asset_ACCT)
			dwp_elig_asset_SECU = trim(dwp_elig_asset_SECU)
			dwp_elig_asset_CARS = trim(dwp_elig_asset_CARS)
			dwp_elig_asset_SPON = trim(dwp_elig_asset_SPON)
			dwp_elig_asset_total = trim(dwp_elig_asset_total)
			dwp_elig_asset_maximum = trim(dwp_elig_asset_maximum)

			transmit
		End If

		If dwp_elig_case_test_fail_coop <> "NA" Then
			Call write_value_and_transmit("X", 13, 5)

			EMReadScreen dwp_elig_test_fail_coop_applied_other_benefits, 6, 10, 30
			EMReadScreen dwp_elig_test_fail_coop_provide_requested_info, 6, 11, 30
			EMReadScreen dwp_elig_test_fail_coop_IEVS, 6, 12, 30
			EMReadScreen dwp_elig_test_fail_coop_vendor_info, 6, 13, 30

			dwp_elig_test_fail_coop_applied_other_benefits = trim(dwp_elig_test_fail_coop_applied_other_benefits)
			dwp_elig_test_fail_coop_provide_requested_info = trim(dwp_elig_test_fail_coop_provide_requested_info)
			dwp_elig_test_fail_coop_IEVS = trim(dwp_elig_test_fail_coop_IEVS)
			dwp_elig_test_fail_coop_vendor_info = trim(dwp_elig_test_fail_coop_vendor_info)

			transmit
		End If

		If dwp_elig_case_test_initial_income <> "NA" Then
			Call write_value_and_transmit("X", 6, 43)

			EMReadScreen dwp_elig_initial_counted_earned_income, 	9, 8, 42
			EMReadScreen dwp_elig_initial_dependent_care_expense, 	9, 9, 42
			EMReadScreen dwp_elig_initial_counted_unearned_incom, 	9, 10, 42
			EMReadScreen dwp_elig_initial_counted_deemed_income, 	9, 11, 42
			EMReadScreen dwp_elig_initial_child_support_exclusion, 	9, 12, 42
			EMReadScreen dwp_elig_initial_total_counted_income, 	9, 13, 42
			EMReadScreen dwp_elig_initial_family_wage_level, 		9, 15, 42

			dwp_elig_initial_counted_earned_income = trim(dwp_elig_initial_counted_earned_income)
			dwp_elig_initial_dependent_care_expense = trim(dwp_elig_initial_dependent_care_expense)
			dwp_elig_initial_counted_unearned_incom = trim(dwp_elig_initial_counted_unearned_incom)
			dwp_elig_initial_counted_deemed_income = trim(dwp_elig_initial_counted_deemed_income)
			dwp_elig_initial_child_support_exclusion = trim(dwp_elig_initial_child_support_exclusion)
			dwp_elig_initial_total_counted_income = trim(dwp_elig_initial_total_counted_income)
			dwp_elig_initial_family_wage_level = trim(dwp_elig_initial_family_wage_level)

			'TODO - read member specific detail'

			transmit
		End If

		If dwp_elig_case_test_verif <> "NA" Then
			Call write_value_and_transmit("X", 12, 43)

			EMReadScreen dwp_elig_test_verif_ACCT, 		6, 5, 32
			EMReadScreen dwp_elig_test_verif_BUSI, 		6, 6, 32
			EMReadScreen dwp_elig_test_verif_CARS, 		6, 7, 32
			EMReadScreen dwp_elig_test_verif_JOBS, 		6, 8, 32
			EMReadScreen dwp_elig_test_verif_MEMB_dob, 	6, 9, 32
			EMReadScreen dwp_elig_test_verif_MEMB_id, 	6, 10, 32
			EMReadScreen dwp_elig_test_verif_PARE, 		6, 11, 32
			EMReadScreen dwp_elig_test_verif_PREG, 		6, 12, 32
			EMReadScreen dwp_elig_test_verif_RBIC, 		6, 13, 32
			EMReadScreen dwp_elig_test_verif_ADDR, 		6, 14, 32
			EMReadScreen dwp_elig_test_verif_SCHL, 		6, 15, 32
			EMReadScreen dwp_elig_test_verif_SECU, 		6, 16, 32
			EMReadScreen dwp_elig_test_verif_SPON, 		6, 17, 32
			EMReadScreen dwp_elig_test_verif_UNEA, 		6, 18, 32

			dwp_elig_test_verif_ACCT = trim(dwp_elig_test_verif_ACCT)
			dwp_elig_test_verif_BUSI = trim(dwp_elig_test_verif_BUSI)
			dwp_elig_test_verif_CARS = trim(dwp_elig_test_verif_CARS)
			dwp_elig_test_verif_JOBS = trim(dwp_elig_test_verif_JOBS)
			dwp_elig_test_verif_MEMB_dob = trim(dwp_elig_test_verif_MEMB_dob)
			dwp_elig_test_verif_MEMB_id = trim(dwp_elig_test_verif_MEMB_id)
			dwp_elig_test_verif_PARE = trim(dwp_elig_test_verif_PARE)
			dwp_elig_test_verif_PREG = trim(dwp_elig_test_verif_PREG)
			dwp_elig_test_verif_RBIC = trim(dwp_elig_test_verif_RBIC)
			dwp_elig_test_verif_ADDR = trim(dwp_elig_test_verif_ADDR)
			dwp_elig_test_verif_SCHL = trim(dwp_elig_test_verif_SCHL)
			dwp_elig_test_verif_SECU = trim(dwp_elig_test_verif_SECU)
			dwp_elig_test_verif_SPON = trim(dwp_elig_test_verif_SPON)
			dwp_elig_test_verif_UNEA = trim(dwp_elig_test_verif_UNEA)

			transmit
		End If

		If dwp_elig_case_test_new_spouse_income <> "NA" Then
			Call write_value_and_transmit("X", 17, 5)

			'TODO - Read New Spouse Income Information

			transmit
		End If

		transmit 		'going to the next panel - DWCB1
		MsgBox "AT DWB1"


		EMReadScreen dwp_elig_budg_shel_rent_mortgage, 		9, 5, 29
		EMReadScreen dwp_elig_budg_shel_property_tax, 		9, 6, 29
		EMReadScreen dwp_elig_budg_shel_house_insurance, 	9, 7, 29
		EMReadScreen dwp_elig_budg_hest_electricity, 		9, 8, 29
		EMReadScreen dwp_elig_budg_hest_heat_air, 			9, 9, 29
		EMReadScreen dwp_elig_budg_hest_water_sewer_garbage, 9, 10, 29
		EMReadScreen dwp_elig_budg_hest_phone, 				9, 11, 29
		EMReadScreen dwp_elig_budg_shel_other, 				9, 12, 29

		EMReadScreen dwp_elig_budg_total_shelter_costs, 	9, 14, 29
		EMReadScreen dwp_elig_budg_personal_needs, 			9, 15, 29

		EMReadScreen dwp_elig_budg_total_DWP_need, 			9, 17, 29

		EMReadScreen dwp_elig_budg_earned_income, 			9, 7, 71
		EMReadScreen dwp_elig_budg_unearned_income, 		9, 8, 71
		EMReadScreen dwp_elig_budg_deemed_income, 			9, 9, 71
		EMReadScreen dwp_elig_budg_child_support_exclusion, 9, 10, 71
		EMReadScreen dwp_elig_budg_budget_month_total, 		9, 11, 71
		EMReadScreen dwp_elig_budg_prior_low, 				9, 12, 71
		EMReadScreen dwp_elig_budg_DWP_countable_income, 	9, 13, 71

		EMReadScreen dwp_elig_budg_unmet_need, 				9, 15, 71
		EMReadScreen dwp_elig_budg_DWP_max_grant, 			9, 16, 71
		EMReadScreen dwp_elig_budg_DWP_grant, 				9, 17, 71

		dwp_elig_budg_shel_rent_mortgage = trim(dwp_elig_budg_shel_rent_mortgage)
		dwp_elig_budg_shel_property_tax = trim(dwp_elig_budg_shel_property_tax)
		dwp_elig_budg_shel_house_insurance = trim(dwp_elig_budg_shel_house_insurance)
		dwp_elig_budg_hest_electricity = trim(dwp_elig_budg_hest_electricity)
		dwp_elig_budg_hest_heat_air = trim(dwp_elig_budg_hest_heat_air)
		dwp_elig_budg_hest_water_sewer_garbage = trim(dwp_elig_budg_hest_water_sewer_garbage)
		dwp_elig_budg_hest_phone = trim(dwp_elig_budg_hest_phone)
		dwp_elig_budg_shel_other = trim(dwp_elig_budg_shel_other)
		dwp_elig_budg_total_shelter_costs = trim(dwp_elig_budg_total_shelter_costs)
		dwp_elig_budg_personal_needs = trim(dwp_elig_budg_personal_needs)
		dwp_elig_budg_total_DWP_need = trim(dwp_elig_budg_total_DWP_need)
		dwp_elig_budg_earned_income = trim(dwp_elig_budg_earned_income)
		dwp_elig_budg_unearned_income = trim(dwp_elig_budg_unearned_income)
		dwp_elig_budg_deemed_income = trim(dwp_elig_budg_deemed_income)
		dwp_elig_budg_child_support_exclusion = trim(dwp_elig_budg_child_support_exclusion)
		dwp_elig_budg_budget_month_total = trim(dwp_elig_budg_budget_month_total)
		dwp_elig_budg_prior_low = trim(dwp_elig_budg_prior_low)
		dwp_elig_budg_DWP_countable_income = trim(dwp_elig_budg_DWP_countable_income)
		dwp_elig_budg_unmet_need = trim(dwp_elig_budg_unmet_need)
		dwp_elig_budg_DWP_max_grant = trim(dwp_elig_budg_DWP_max_grant)
		dwp_elig_budg_DWP_grant = trim(dwp_elig_budg_DWP_grant)

		Call write_value_and_transmit("X", 7, 41)
		EmReadScreen pop_up_menu_title, 13, 3, 46
		If pop_up_menu_title = "Earned Income" Then
			'TODO - read member specific unearned income
			transmit
		End If

		Call write_value_and_transmit("X", 8, 41)
		EmReadScreen pop_up_menu_title, 15, 5, 32
		If pop_up_menu_title = "Unearned Income" Then
			'TODO - read member specific unearned income
			transmit
		End If

		Call write_value_and_transmit("X", 9, 41)
		EmReadScreen pop_up_menu_title, 13, 3, 36
		If pop_up_menu_title = "Deemed Income" Then
			'TODO - read member specific unearned income
			' EMReadScreen dwp_elig_membs_budg_deemed_self_emp(member_sel), 				9, 8, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_wages(member_sel), 					9, 9, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_counted_earned(member_sel), 		9, 10, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_standard_EI_disregard(member_sel), 	9, 11, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_earned_subtotal(member_sel), 		9, 12, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_earned_disregard(member_sel), 		9, 13, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_unearned_income(member_sel), 		9, 14, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_subtotal_counted_income(member_sel), 9, 15, 56
			'
			' EMReadScreen dwp_elig_membs_budg_deemed_deemer_unmet_need(member_sel), 		9, 18, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_allocation(member_sel), 			9, 19, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_child_support(member_sel), 			9, 20, 56
			' EMReadScreen dwp_elig_membs_budg_deemed_counted_income(member_sel), 		9, 21, 56
			transmit
		End If

		Call write_value_and_transmit("X", 10, 41)
		EMReadScreen dwp_elig_cses_income, 9, 10, 54
		EMReadScreen dwp_elig_child_count, 2, 12, 36
		'TODO - read member specific unearned income

		dwp_elig_cses_income = trim(dwp_elig_cses_income)
		dwp_elig_child_count = trim(dwp_elig_child_count)

		transmit


		transmit 		'going to the next panel - DWB2
		MsgBox "AT DWB2"

		EMReadScreen dwp_elig_prorated_date, 8, 6, 18
		If dwp_elig_prorated_date = "__ __ __" then dwp_elig_prorated_date = ""
		dwp_elig_prorated_date = replace(dwp_elig_prorated_date, " ", "/")

		EMReadScreen dwp_elig_prorated_amount, 9, 6, 35

		EMReadScreen dwp_elig_amount_already_issued, 	9, 9, 35
		EMReadScreen dwp_elig_supplement_due, 			9, 10, 35
		EMReadScreen dwp_elig_overpayment, 				9, 11, 35
		EMReadScreen dwp_elig_adjusted_grant_amount, 	9, 12, 35
		EMReadScreen dwp_elig_recoupment_amount, 		9, 13, 35

		EMReadScreen dwp_elig_shelter_benefit_grant, 	9, 15, 35
		EMReadScreen dwp_elig_personal_needs_grant, 	9, 16, 35

		Call write_value_and_transmit("X", 11, 3)
		EMReadScreen dwp_elig_overpayment_fed_hh_count, 	2, 10, 31
		EMReadScreen dwp_elig_overpayment_fed_amount, 		9, 10, 50
		EMReadScreen dwp_elig_overpayment_state_hh_count, 	2, 12, 31
		EMReadScreen dwp_elig_overpayment_state_amount, 	9, 12, 50
		transmit

		Call write_value_and_transmit("X", 12, 3)
		EMReadScreen dwp_elig_adjusted_grant_fed_hh_count, 		2, 10, 25
		EMReadScreen dwp_elig_adjusted_grant_fed_amount, 		9, 10, 45
		EMReadScreen dwp_elig_adjusted_grant_state_hh_count, 	2, 12, 25
		EMReadScreen dwp_elig_adjusted_grant_state_amount, 		9, 12, 45
		transmit

		dwp_elig_prorated_amount = trim(dwp_elig_prorated_amount)
		dwp_elig_amount_already_issued = trim(dwp_elig_amount_already_issued)
		dwp_elig_supplement_due = trim(dwp_elig_supplement_due)
		dwp_elig_overpayment = trim(dwp_elig_overpayment)
		dwp_elig_adjusted_grant_amount = trim(dwp_elig_adjusted_grant_amount)
		dwp_elig_recoupment_amount = trim(dwp_elig_recoupment_amount)
		dwp_elig_shelter_benefit_grant = trim(dwp_elig_shelter_benefit_grant)
		dwp_elig_personal_needs_grant = trim(dwp_elig_personal_needs_grant)
		dwp_elig_overpayment_fed_amount = trim(dwp_elig_overpayment_fed_amount)
		dwp_elig_overpayment_state_amount = trim(dwp_elig_overpayment_state_amount)
		dwp_elig_adjusted_grant_fed_amount = trim(dwp_elig_adjusted_grant_fed_amount)
		dwp_elig_adjusted_grant_state_amount = trim(dwp_elig_adjusted_grant_state_amount)

		transmit 		'going to the next panel - DWSM
		MsgBox "AT DWSM"

		EMReadScreen dwp_approved_date, 8, 3, 14
		EMReadScreen dwp_process_date, 8, 2, 73
		EMReadScreen dwp_prev_approval, 4, 3, 73

		EMReadScreen dwp_case_last_approval_date, 8, 5, 31
		EMReadScreen dwp_case_current_prog_status, 12, 6, 31
		EMReadScreen dwp_case_eligibility_result, 12,  7, 31
		EMReadScreen dwp_case_source_of_info, 4, 9, 31
		EMReadScreen dwp_case_benefit_impact, 12, 10, 31
		EMReadScreen dwp_case_4th_month_of_elig, 5, 11, 31
		EMReadScreen dwp_case_caregivers_have_es_plan, 1, 12, 31
		EMReadScreen dwp_case_responsible_county, 2, 13, 31
		EMReadScreen dwp_case_service_county, 2, 14, 31

		EMReadScreen dwp_case_asst_unit_caregivers, 3, 5, 72
		EMReadScreen dwp_case_asst_unit_children, 3, 6, 72
		EMReadScreen dwp_case_total_assets, 10, 7, 71
		EMReadScreen dwp_case_maximum_assets, 10, 8, 71
		EMReadScreen dwp_case_summary_grant_amount, 10, 10, 71
		EMReadScreen dwp_case_summary_net_grant_amount, 10, 12, 71
		EMReadScreen dwp_case_summary_shelter_benefit_portion, 10, 13, 71
		EMReadScreen dwp_case_summary_personal_needs_portion, 10, 14, 71

		dwp_prev_approval = trim(dwp_prev_approval)
		dwp_case_last_approval_date = trim(dwp_case_last_approval_date)

		dwp_case_current_prog_status = trim(dwp_case_current_prog_status)
		dwp_case_eligibility_result = trim(dwp_case_eligibility_result)
		dwp_case_source_of_info = trim(dwp_case_source_of_info)
		dwp_case_benefit_impact = trim(dwp_case_benefit_impact)

		dwp_case_asst_unit_caregivers = trim(dwp_case_asst_unit_caregivers)
		dwp_case_asst_unit_children = trim(dwp_case_asst_unit_children)
		dwp_case_total_assets = trim(dwp_case_total_assets)
		dwp_case_maximum_assets = trim(dwp_case_maximum_assets)
		dwp_case_summary_grant_amount = trim(dwp_case_summary_grant_amount)
		dwp_case_summary_net_grant_amount = trim(dwp_case_summary_net_grant_amount)
		dwp_case_summary_shelter_benefit_portion = trim(dwp_case_summary_shelter_benefit_portion)
		dwp_case_summary_personal_needs_portion = trim(dwp_case_summary_personal_needs_portion)

		Call back_to_SELF
	end sub
end class

class mfip_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result

	public mfip_elig_ref_numbs()
	public mfip_elig_membs_full_name()
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

	public mfip_elig_membs_initial_BUSI_inc_total()
	public mfip_elig_membs_initial_JOBS_inc_total()
	public mfip_elig_membs_initial_earned_inc_total()
	public mfip_elig_membs_initial_stndrd_ei_disregard()
	public mfip_elig_membs_initial_earned_inc_subtotal()
	public mfip_elig_membs_initial_earned_inc_disregard()
	public mfip_elig_membs_initial_avail_earned_inc()
	public mfip_elig_membs_initial_allocation()
	public mfip_elig_membs_initial_child_support()
	public mfip_elig_membs_initial_counted_earned_inc_total()
	public mfip_elig_membs_initial_UNEA_inc_total()
	public mfip_elig_membs_initial_allocation_balance()
	public mfip_elig_membs_initial_child_support_balance()
	public mfip_elig_membs_initial_counted_UNEA_inc_total()
	public mfip_elig_membs_initial_income_cses_retro_income()
	public mfip_elig_membs_initial_income_cses_prosp_income()
	public mfip_elig_membs_new_spouse_earned_income()
	public mfip_elig_membs_new_spouse_unearned_income()
	public mfip_elig_membs_new_spouse_total_income()

	public mfip_elig_membs_self_emp_income()
	public mfip_elig_membs_wages_income()
	public mfip_elig_membs_total_earned_income()
	public mfip_elig_membs_standard_EI_disregard()
	public mfip_elig_membs_earned_income_subtotal()
	public mfip_elig_membs_earned_income_50_perc_disregard()
	public mfip_elig_membs_available_earned_income()
	public mfip_elig_membs_allocation_deduction()
	public mfip_elig_membs_child_support_deduction()
	public mfip_elig_membs_counted_earned_income()

	public mfip_elig_membs_total_unearned_income()
	public mfip_elig_membs_allocation_balance()
	public mfip_elig_membs_child_support_balance()
	public mfip_elig_membs_counted_unearned_income()

	public mfip_elig_membs_county_88_cses_income()
	public mfip_elig_membs_county_88_gaming_income()
	public mfip_elig_membs_county_88_200_perc_fpg()
	public mfip_elig_membs_county_88_deemers_unmet_need()
	public mfip_elig_membs_county_88_allocation()
	public mfip_elig_membs_county_88_child_support()
	public mfip_elig_membs_county_88_counted_gaming_income()

	public mfip_elig_membs_retro_subsidy_amount()
	public mfip_elig_membs_prosp_subsidy_amount()

	public mfip_cash_opt_out
	public mfip_HG_opt_out
	public mfip_child_only
	public mfip_case_in_sancttion

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
	public mfip_initial_income_cses_income
	public mfip_initial_income_cses_child_count
	public mfip_initial_income_net_cses_income
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
	public mfip_budg_total_county_88_child_support_income
	public mfip_budg_total_county_88_gaming_income
	public mfip_budg_total_tribal_income_fs_portion_deduction
	public mfip_budg_total_housing_subsidy_amount
	public mfip_budg_total_tribal_child_support
	public mfip_budg_total_subsidy_tribal_cash_portion_deduction
	public mfip_elig_budg_total_countable_housing_subsidy
	public mfip_elig_budg_housing_subsidy_exempt

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
	public mfip_case_budg_net_cash_after_sanc_portion
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
		mfip_case_in_sancttion = False

		call navigate_to_MAXIS_screen("ELIG", "MFIP")
		EMWriteScreen elig_footer_month, 20, 55
		EMWriteScreen elig_footer_year, 20, 58
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)

		ReDim mfip_elig_ref_numbs(0)
		ReDim mfip_elig_membs_full_name(0)
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
		ReDim mfip_elig_membs_initial_BUSI_inc_total(0)
		ReDim mfip_elig_membs_initial_JOBS_inc_total(0)
		ReDim mfip_elig_membs_initial_earned_inc_total(0)
		ReDim mfip_elig_membs_initial_stndrd_ei_disregard(0)
		ReDim mfip_elig_membs_initial_earned_inc_subtotal(0)
		ReDim mfip_elig_membs_initial_earned_inc_disregard(0)
		ReDim mfip_elig_membs_initial_avail_earned_inc(0)
		ReDim mfip_elig_membs_initial_allocation(0)
		ReDim mfip_elig_membs_initial_child_support(0)
		ReDim mfip_elig_membs_initial_counted_earned_inc_total(0)
		ReDim mfip_elig_membs_initial_UNEA_inc_total(0)
		ReDim mfip_elig_membs_initial_allocation_balance(0)
		ReDim mfip_elig_membs_initial_child_support_balance(0)
		ReDim mfip_elig_membs_initial_counted_UNEA_inc_total(0)
		ReDim mfip_elig_membs_initial_income_cses_retro_income(0)
		ReDim mfip_elig_membs_initial_income_cses_prosp_income(0)
		ReDim mfip_elig_membs_new_spouse_earned_income(0)
		ReDim mfip_elig_membs_new_spouse_unearned_income(0)
		ReDim mfip_elig_membs_new_spouse_total_income(0)
		ReDim mfip_elig_membs_self_emp_income(0)
		ReDim mfip_elig_membs_wages_income(0)
		ReDim mfip_elig_membs_total_earned_income(0)
		ReDim mfip_elig_membs_standard_EI_disregard(0)
		ReDim mfip_elig_membs_earned_income_subtotal(0)
		ReDim mfip_elig_membs_earned_income_50_perc_disregard(0)
		ReDim mfip_elig_membs_available_earned_income(0)
		ReDim mfip_elig_membs_allocation_deduction(0)
		ReDim mfip_elig_membs_child_support_deduction(0)
		ReDim mfip_elig_membs_counted_earned_income(0)
		ReDim mfip_elig_membs_total_unearned_income(0)
		ReDim mfip_elig_membs_allocation_balance(0)
		ReDim mfip_elig_membs_child_support_balance(0)
		ReDim mfip_elig_membs_counted_unearned_income(0)
		ReDim mfip_elig_membs_county_88_cses_income(0)
		ReDim mfip_elig_membs_county_88_gaming_income(0)
		ReDim mfip_elig_membs_county_88_200_perc_fpg(0)
		ReDim mfip_elig_membs_county_88_deemers_unmet_need(0)
		ReDim mfip_elig_membs_county_88_allocation(0)
		ReDim mfip_elig_membs_county_88_child_support(0)
		ReDim mfip_elig_membs_county_88_counted_gaming_income(0)
		ReDim mfip_elig_membs_retro_subsidy_amount(0)
		ReDim mfip_elig_membs_prosp_subsidy_amount(0)

		row = 7
		elig_memb_count = 0
		Do
			EMReadScreen ref_numb, 2, row, 6

			ReDim preserve mfip_elig_ref_numbs(elig_memb_count)
			ReDim preserve mfip_elig_membs_full_name(elig_memb_count)
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
			ReDim preserve mfip_memb_sanction_drug_felon_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_emp_services_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_fin_orient_test(elig_memb_count)
			ReDim preserve mfip_memb_sanction_occurence(elig_memb_count)
			ReDim preserve mfip_memb_sanction_begin_date(elig_memb_count)
			ReDim preserve mfip_memb_sanction_last_sanc_month(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_BUSI_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_JOBS_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_earned_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_stndrd_ei_disregard(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_earned_inc_subtotal(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_earned_inc_disregard(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_avail_earned_inc(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_allocation(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_child_support(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_counted_earned_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_UNEA_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_allocation_balance(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_child_support_balance(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_counted_UNEA_inc_total(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_income_cses_retro_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_initial_income_cses_prosp_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_new_spouse_earned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_new_spouse_unearned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_new_spouse_total_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_self_emp_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_wages_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_total_earned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_standard_EI_disregard(elig_memb_count)
			ReDim preserve mfip_elig_membs_earned_income_subtotal(elig_memb_count)
			ReDim preserve mfip_elig_membs_earned_income_50_perc_disregard(elig_memb_count)
			ReDim preserve mfip_elig_membs_available_earned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_allocation_deduction(elig_memb_count)
			ReDim preserve mfip_elig_membs_child_support_deduction(elig_memb_count)
			ReDim preserve mfip_elig_membs_counted_earned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_total_unearned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_allocation_balance(elig_memb_count)
			ReDim preserve mfip_elig_membs_child_support_balance(elig_memb_count)
			ReDim preserve mfip_elig_membs_counted_unearned_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_cses_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_gaming_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_200_perc_fpg(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_deemers_unmet_need(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_allocation(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_child_support(elig_memb_count)
			ReDim preserve mfip_elig_membs_county_88_counted_gaming_income(elig_memb_count)
			ReDim preserve mfip_elig_membs_retro_subsidy_amount(elig_memb_count)
			ReDim preserve mfip_elig_membs_prosp_subsidy_amount(elig_memb_count)

			mfip_elig_ref_numbs(elig_memb_count) = ref_numb
			EMReadScreen full_name_information, 20, row, 10
			full_name_information = trim(full_name_information)
			name_array = split(full_name_information, " ")
			For each name_parts in name_array
				If len(name_parts) <> 1 Then mfip_elig_membs_full_name(elig_memb_count) = mfip_elig_membs_full_name(elig_memb_count) & " " & name_parts
			Next
			mfip_elig_membs_full_name(elig_memb_count) = trim((mfip_elig_membs_full_name(elig_memb_count)))
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
			EMReadScreen emps_exists_for_memb, 19, 24, 2
			If emps_exists_for_memb = "EMPS DOES NOT EXIST" Then
				EMWriteScreen " ", row, 64
			Else
				EMReadScreen mfip_elig_membs_es_status_code(elig_memb_count), 2, 9, 22
				EMReadScreen mfip_elig_membs_es_status_info(elig_memb_count), 30, 9, 25
				mfip_elig_membs_es_status_code(elig_memb_count) = trim(mfip_elig_membs_es_status_code(elig_memb_count))
				mfip_elig_membs_es_status_info(elig_memb_count) = trim(mfip_elig_membs_es_status_info(elig_memb_count))
				transmit
			End If


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
		Call write_value_and_transmit("X", 8, 20)		'Member Initial Earned Income
		Do
			EMReadScreen pop_up_name, 40, 8, 28
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then
					EMReadScreen mfip_elig_membs_initial_BUSI_inc_total(case_memb), 		10, 11, 54
					EMReadScreen mfip_elig_membs_initial_JOBS_inc_total(case_memb), 		10, 12, 54
					EMReadScreen mfip_elig_membs_initial_earned_inc_total(case_memb), 		10, 13, 54
					EMReadScreen mfip_elig_membs_initial_stndrd_ei_disregard(case_memb), 	10, 14, 54
					EMReadScreen mfip_elig_membs_initial_earned_inc_subtotal(case_memb), 	10, 15, 54
					EMReadScreen mfip_elig_membs_initial_earned_inc_disregard(case_memb), 	10, 16, 54
					EMReadScreen mfip_elig_membs_initial_avail_earned_inc(case_memb), 		10, 17, 54
					EMReadScreen mfip_elig_membs_initial_allocation(case_memb), 			10, 18, 54
					EMReadScreen mfip_elig_membs_initial_child_support(case_memb), 			10, 19, 54
					EMReadScreen mfip_elig_membs_initial_counted_earned_inc_total(case_memb), 10, 20, 54

					mfip_elig_membs_initial_BUSI_inc_total(case_memb) = trim(mfip_elig_membs_initial_BUSI_inc_total(case_memb))
					mfip_elig_membs_initial_JOBS_inc_total(case_memb) = trim(mfip_elig_membs_initial_JOBS_inc_total(case_memb))
					mfip_elig_membs_initial_earned_inc_total(case_memb) = trim(mfip_elig_membs_initial_earned_inc_total(case_memb))
					mfip_elig_membs_initial_stndrd_ei_disregard(case_memb) = trim(mfip_elig_membs_initial_stndrd_ei_disregard(case_memb))
					mfip_elig_membs_initial_earned_inc_subtotal(case_memb) = trim(mfip_elig_membs_initial_earned_inc_subtotal(case_memb))
					mfip_elig_membs_initial_earned_inc_disregard(case_memb) = trim(mfip_elig_membs_initial_earned_inc_disregard(case_memb))
					mfip_elig_membs_initial_avail_earned_inc(case_memb) = trim(mfip_elig_membs_initial_avail_earned_inc(case_memb))
					mfip_elig_membs_initial_allocation(case_memb) = trim(mfip_elig_membs_initial_allocation(case_memb))
					mfip_elig_membs_initial_child_support(case_memb) = trim(mfip_elig_membs_initial_child_support(case_memb))
					mfip_elig_membs_initial_counted_earned_inc_total(case_memb) = trim(mfip_elig_membs_initial_counted_earned_inc_total(case_memb))

					' If mfip_elig_membs_initial_BUSI_inc_total(case_memb) <> "0.00" Then 			'this will likely not be used - opening these pop ups do not provide details on different jobs
					' 	Call write_value_and_transmit("X", 11, 20)
					' End If
					' If mfip_elig_membs_initial_JOBS_inc_total(case_memb) <> "0.00" Then
					' 	Call write_value_and_transmit("X", 12, 20)
					' End If
				End If
			Next
			transmit

			EMReadScreen back_to_menu, 14, 6, 29
		Loop until back_to_menu = "Initial Income"

		If mfip_initial_income_deoendant_care <> "0.00" Then 			''Depended Care Initial Income calculation pop-up
			Call write_value_and_transmit("X", 9, 20)
		End If

		Call write_value_and_transmit("X", 10, 20)		'Member Initial Unearned Income
		Do
			EMReadScreen pop_up_name, 40, 8, 28
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then
					EMReadScreen mfip_elig_membs_initial_UNEA_inc_total(case_memb), 		10, 11, 49
					EMReadScreen mfip_elig_membs_initial_allocation_balance(case_memb), 	10, 12, 49
					EMReadScreen mfip_elig_membs_initial_child_support_balance(case_memb), 	10, 13, 49
					EMReadScreen mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb), 10, 14, 49

					mfip_elig_membs_initial_UNEA_inc_total(case_memb) = trim(mfip_elig_membs_initial_UNEA_inc_total(case_memb))
					mfip_elig_membs_initial_allocation_balance(case_memb) = trim(mfip_elig_membs_initial_allocation_balance(case_memb))
					mfip_elig_membs_initial_child_support_balance(case_memb) = trim(mfip_elig_membs_initial_child_support_balance(case_memb))
					 mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb) = trim(mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb))
				End If
			Next
			transmit

			EMReadScreen back_to_menu, 14, 6, 29
		Loop until back_to_menu = "Initial Income"

		If mfip_initial_income_deemed <> "0.00" Then 			'Deemed Initial Income calculation pop-up
			Call write_value_and_transmit("X", 11, 20)
		End If

		Call write_value_and_transmit("X", 12, 20)				'CSES Exclusion Initiall Income calculation pop-up
		EMReadScreen mfip_initial_income_cses_income, 10, 9, 52
		EMReadScreen mfip_initial_income_cses_child_count, 2, 11, 37

		mfip_initial_income_cses_income = trim(mfip_initial_income_cses_income)
		mfip_initial_income_cses_child_count = trim(mfip_initial_income_cses_child_count)

		Call write_value_and_transmit("X", 9, 20)				'open cses initial income pop-up'

		EMReadScreen mfip_initial_income_net_cses_income, 10, 19, 44
		mfip_initial_income_net_cses_income = trim(mfip_initial_income_net_cses_income)
		mfcr_row = 7
		Do
			EMReadScreen ref_numb, 2, mfcr_row, 7

			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If ref_numb = mfip_elig_ref_numbs(case_memb) Then
					EMReadScreen mfip_elig_membs_initial_income_cses_retro_income(case_memb), 10, mfcr_row, 41
					EMReadScreen mfip_elig_membs_initial_income_cses_prosp_income(case_memb), 10, mfcr_row, 54

					mfip_elig_membs_initial_income_cses_retro_income(case_memb) = trim(mfip_elig_membs_initial_income_cses_retro_income(case_memb))
					mfip_elig_membs_initial_income_cses_prosp_income(case_memb) = trim(mfip_elig_membs_initial_income_cses_prosp_income(case_memb))
				End If
			Next

			mfcr_row = mfcr_row + 1
			EMReadScreen next_ref_numb, 2, mfcr_row, 3
		Loop until next_ref_numb = "  "

		PF3			'back to CSES Exclusion caclulaiton
		PF3			'back to initial income calculation
		PF3			'back to main mf elig panel'
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

		Call write_value_and_transmit("X", 11, 20)		'Member earned and unearned for New Spouse calculation
		Do
			EMReadScreen pop_up_name, 35, 7, 25
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

					EMReadScreen mfip_elig_membs_new_spouse_earned_income(case_memb), 	10, 9, 48
					EMReadScreen mfip_elig_membs_new_spouse_unearned_income(case_memb), 10, 10, 48
					EMReadScreen mfip_elig_membs_new_spouse_total_income(case_memb), 	10, 11, 48

					mfip_elig_membs_new_spouse_earned_income(case_memb) = trim(mfip_elig_membs_new_spouse_earned_income(case_memb))
					mfip_elig_membs_new_spouse_unearned_income(case_memb) = trim(mfip_elig_membs_new_spouse_unearned_income(case_memb))
					mfip_elig_membs_new_spouse_total_income(case_memb) = trim(mfip_elig_membs_new_spouse_total_income(case_memb))
				End If
			Next
			transmit

			EMReadScreen back_to_menu, 17, 7, 22
		Loop until back_to_menu = "Designated Spouse"

		'TODO - Read the deemed pop-ups
		If mfip_new_spouse_inc_deemed_earned <> "0.00" Then
			' Call write_value_and_transmit("X", 13, 20)		'Member deemed earned for New Spouse calculation
		End If
		If mfip_new_spouse_inc_deemed_unearned <> "0.00" Then
			' Call write_value_and_transmit("X", 14, 20)		'Member deemed unearned for New Spouse calculation
		End If

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
					If mfip_memb_sanction_yn(case_memb) = "Y" Then mfip_case_in_sancttion = True

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
		EMReadScreen mfip_case_budg_monthly_earned_income, 			10, 6, 32
		EMReadScreen mfip_case_budg_wage_level_earned_inc_difference, 10, 7, 32
		EMReadScreen mfip_case_budg_transitional_standard, 			10, 9, 32
		EMReadScreen mfip_case_budg_monthly_need, 					10, 10, 32
		EMReadScreen mfip_case_budg_unearned_income, 				10, 11, 32
		EMReadScreen mfip_case_budg_deemed_income, 					10, 12, 32
		EMReadScreen mfip_case_budg_cses_exclusion, 				10, 13, 32
		EMReadScreen mfip_case_budg_unmet_need, 					10, 14, 32
		EMReadScreen mfip_case_budg_unmet_need_food_potion, 		10, 15, 32
		EMReadScreen mfip_case_budg_tribal_counted_income, 			10, 16, 32
		EMReadScreen mfip_case_budg_unmet_need_cash_portion, 		10, 17, 32
		EMReadScreen mfip_case_budg_deduction_subsidy_tribal_cses, 	10, 18, 32


		EMReadScreen mfip_case_budg_net_food_portion, 			10, 5, 71
		EMReadScreen mfip_case_budg_net_cash_portion, 			10, 6, 71
		EMReadScreen mfip_case_budg_net_unmet_need, 			10, 7, 71
		EMReadScreen mfip_case_budg_deduction_sanction_vendor, 	10, 8, 71
		EMReadScreen mfip_case_budg_unmet_neet_subtotal, 		10, 9, 71
		EMReadScreen mfip_case_budg_subtotal_food_portion, 		10, 11, 71
		EMReadScreen mfip_case_budg_food_portion_deduction, 	10, 12, 71
		EMReadScreen mfip_case_budg_entitlement_food_portion, 	10, 13, 71
		EMReadScreen mfip_case_budg_entitlement_housing_grant, 	10, 15, 71

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

		Call write_value_and_transmit("X", 6, 3)		' member specific EARNED INCOME
		Do
			EMReadScreen pop_up_name, 40, 8, 28
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

					EMReadScreen mfip_elig_membs_self_emp_income(case_memb), 				10, 11, 54
					EMReadScreen mfip_elig_membs_wages_income(case_memb), 					10, 12, 54
					EMReadScreen mfip_elig_membs_total_earned_income(case_memb), 			10, 13, 54
					EMReadScreen mfip_elig_membs_standard_EI_disregard(case_memb), 			10, 14, 54
					EMReadScreen mfip_elig_membs_earned_income_subtotal(case_memb), 		10, 15, 54
					EMReadScreen mfip_elig_membs_earned_income_50_perc_disregard(case_memb), 10, 16, 54
					EMReadScreen mfip_elig_membs_available_earned_income(case_memb), 		10, 17, 54
					EMReadScreen mfip_elig_membs_allocation_deduction(case_memb), 			10, 18, 54
					EMReadScreen mfip_elig_membs_child_support_deduction(case_memb), 		10, 19, 54
					EMReadScreen mfip_elig_membs_counted_earned_income(case_memb), 			10, 20, 54

					mfip_elig_membs_self_emp_income(case_memb) = trim(mfip_elig_membs_self_emp_income(case_memb))
					mfip_elig_membs_wages_income(case_memb) = trim(mfip_elig_membs_wages_income(case_memb))
					mfip_elig_membs_total_earned_income(case_memb) = trim(mfip_elig_membs_total_earned_income(case_memb))
					mfip_elig_membs_standard_EI_disregard(case_memb) = trim(mfip_elig_membs_standard_EI_disregard(case_memb))
					mfip_elig_membs_earned_income_subtotal(case_memb) = trim(mfip_elig_membs_earned_income_subtotal(case_memb))
					mfip_elig_membs_earned_income_50_perc_disregard(case_memb) = trim(mfip_elig_membs_earned_income_50_perc_disregard(case_memb))
					mfip_elig_membs_available_earned_income(case_memb) = trim(mfip_elig_membs_available_earned_income(case_memb))
					mfip_elig_membs_allocation_deduction(case_memb) = trim(mfip_elig_membs_allocation_deduction(case_memb))
					mfip_elig_membs_child_support_deduction(case_memb) = trim(mfip_elig_membs_child_support_deduction(case_memb))
					mfip_elig_membs_counted_earned_income(case_memb) = trim(mfip_elig_membs_counted_earned_income(case_memb))

				End If
			Next
			transmit
			EMReadScreen still_in_menu, 12, 5, 32
		Loop until still_in_menu <> "Maxis Person"

		Call write_value_and_transmit("X", 11, 3)		' member specific UNEARNED INCOME
		Do
			EMReadScreen pop_up_name, 25, 8, 34
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

					EMReadScreen mfip_elig_membs_total_unearned_income(case_memb), 	10, 11, 54
					EMReadScreen mfip_elig_membs_allocation_balance(case_memb), 	10, 12, 54
					EMReadScreen mfip_elig_membs_child_support_balance(case_memb), 	10, 13, 54
					EMReadScreen mfip_elig_membs_counted_unearned_income(case_memb), 10, 14, 54

					mfip_elig_membs_total_unearned_income(case_memb) = trim(mfip_elig_membs_total_unearned_income(case_memb))
					mfip_elig_membs_allocation_balance(case_memb) = trim(mfip_elig_membs_allocation_balance(case_memb))
					mfip_elig_membs_child_support_balance(case_memb) = trim(mfip_elig_membs_child_support_balance(case_memb))
					mfip_elig_membs_counted_unearned_income(case_memb) = trim(mfip_elig_membs_counted_unearned_income(case_memb))

				End If
			Next
			transmit
			EMReadScreen still_in_menu, 15, 6, 34
		Loop until still_in_menu <> "Unearned Income"

		' Call write_value_and_transmit("X", 12, 3)		'TODO member specific DEEMED INCOME

		Call write_value_and_transmit("X", 13, 3)		'Child Support Exclusion'
		EMReadScreen mfip_budg_cses_excln_cses_income, 10, 9, 52
		EMReadScreen mfip_budg_cses_excln_child_count, 2, 11, 37
		EMReadScreen mfip_budg_cses_excln_total, 10, 13, 52

		mfip_budg_cses_excln_cses_income = trim(mfip_budg_cses_excln_cses_income)
		mfip_budg_cses_excln_child_count = trim(mfip_budg_cses_excln_child_count)
		mfip_budg_cses_excln_total = trim(mfip_budg_cses_excln_total)

		transmit

		Call write_value_and_transmit("X", 16, 5)		' member specific TRIBAL INCOME
		EMReadScreen mfip_budg_total_county_88_child_support_income, 	10, 6, 55
		EMReadScreen mfip_budg_total_county_88_gaming_income, 			10, 7, 55
		EMReadScreen mfip_budg_total_tribal_income_fs_portion_deduction, 10, 8, 55
		mfip_budg_total_county_88_child_support_income = trim(mfip_budg_total_county_88_child_support_income)
		mfip_budg_total_county_88_gaming__income = trim(mfip_budg_total_county_88_gaming__income)
		mfip_budg_total_tribal_income_fs_portion_deduction = trim(mfip_budg_total_tribal_income_fs_portion_deduction)

		Call write_value_and_transmit("X", 6, 12)		' member specific Tribal Child Support Income
		Do
			EMReadScreen pop_up_name, 25, 8, 34
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

					EMReadScreen mfip_elig_membs_county_88_cses_income(case_memb), 10, 11, 54

					mfip_elig_membs_county_88_cses_income(case_memb) = trim(mfip_elig_membs_county_88_cses_income(case_memb))
				End If
			Next
			transmit
			EMReadScreen back_to_menu, 21, 4, 31
		Loop until back_to_menu = "Tribal Counted Income"

		Call write_value_and_transmit("X", 7, 12)		' member specific Tribal Gaming Income
		Do
			EMReadScreen pop_up_name, 30, 7, 37
			pop_up_name = trim(pop_up_name)
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

					EMReadScreen mfip_elig_membs_county_88_gaming_income(case_memb), 	10, 10, 61
					EMReadScreen mfip_elig_membs_county_88_200_perc_fpg(case_memb), 	10, 11, 61
					EMReadScreen mfip_elig_membs_county_88_deemers_unmet_need(case_memb), 10, 12, 61
					EMReadScreen mfip_elig_membs_county_88_allocation(case_memb), 		10, 13, 61
					EMReadScreen mfip_elig_membs_county_88_child_support(case_memb), 	10, 14, 61
					EMReadScreen mfip_elig_membs_county_88_counted_gaming_income(case_memb), 10, 15, 61

					mfip_elig_membs_county_88_gaming_income(case_memb) = trim(mfip_elig_membs_county_88_gaming_income(case_memb))
					mfip_elig_membs_county_88_200_perc_fpg(case_memb) = trim(mfip_elig_membs_county_88_200_perc_fpg(case_memb))
					mfip_elig_membs_county_88_deemers_unmet_need(case_memb) = trim(mfip_elig_membs_county_88_deemers_unmet_need(case_memb))
					mfip_elig_membs_county_88_allocation(case_memb) = trim(mfip_elig_membs_county_88_allocation(case_memb))
					mfip_elig_membs_county_88_child_support(case_memb) = trim(mfip_elig_membs_county_88_child_support(case_memb))
					mfip_elig_membs_county_88_counted_gaming_income(case_memb) = trim(mfip_elig_membs_county_88_counted_gaming_income(case_memb))
				End If
			Next
			transmit
			EMReadScreen back_to_menu, 21, 4, 31
		Loop until back_to_menu = "Tribal Counted Income"
		transmit                  ''back to MFB1

		Call write_value_and_transmit("X", 18, 5)		' member specific SUBSIDY
		EMReadScreen mfip_budg_total_housing_subsidy_amount, 10, 8, 51
		EMReadScreen mfip_budg_total_tribal_child_support, 10, 9, 51
		EMReadScreen mfip_budg_total_subsidy_tribal_cash_portion_deduction, 10, 10, 51
		mfip_budg_total_housing_subsidy_amount = trim(mfip_budg_total_housing_subsidy_amount)
		mfip_budg_total_tribal_child_support = trim(mfip_budg_total_tribal_child_support)
		mfip_budg_total_subsidy_tribal_cash_portion_deduction = trim(mfip_budg_total_subsidy_tribal_cash_portion_deduction)

		Call write_value_and_transmit("X", 8, 13)		' member specific subsidy Income
		EMReadScreen mfip_elig_budg_total_countable_housing_subsidy, 10, 19, 48
		EMReadScreen mfip_elig_budg_housing_subsidy_exempt, 1, 21, 47

		mfip_elig_budg_total_countable_housing_subsidy = trim(mfip_elig_budg_total_countable_housing_subsidy)
		mfip_elig_budg_housing_subsidy_exempt = trim(mfip_elig_budg_housing_subsidy_exempt)

		Do
			row = 8
			EMReadScreen memb_ref_numb, 2, row, 6
			For case_memb = 0 to UBound(mfip_elig_ref_numbs)
				If memb_ref_numb = mfip_elig_ref_numbs(case_memb) Then

					EMReadScreen mfip_elig_membs_retro_subsidy_amount(case_memb), 10, row, 38
					EMReadScreen mfip_elig_membs_prosp_subsidy_amount(case_memb), 10, row, 49

					mfip_elig_membs_retro_subsidy_amount(case_memb) = trim(mfip_elig_membs_retro_subsidy_amount(case_memb))
					mfip_elig_membs_prosp_subsidy_amount(case_memb) = trim(mfip_elig_membs_prosp_subsidy_amount(case_memb))
				End If
			Next
			row = row + 1
			EMReadScreen next_memb_ref_numb, 2, row, 6
		Loop until next_memb_ref_numb = "  "
		transmit 					'back to pop-up

		transmit                 	'back to MFB1

		Call write_value_and_transmit("X", 8, 44)		'Sanction and Vendor
		EMReadScreen mfip_case_budg_10_perc_sanc, 					10, 7, 55
		EMReadScreen mfip_case_budg_unmet_need_after_pre_vndr_sanc, 10, 8, 55
		EMReadScreen mfip_case_budg_sanc_calc_food_portion, 		10, 9, 55
		EMReadScreen mfip_case_budg_sanc_calc_cash_portion, 		10, 10, 55
		EMReadScreen mfip_case_budg_pot_mand_vndr_pymts, 			10, 11, 55
		EMReadScreen mfip_case_budg_30_perc_sanc, 					10, 12, 55

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
		EMReadScreen mfip_case_budg_prorate_date, 8, 5, 19

		EMReadScreen mfip_case_budg_fed_food_benefit, 			10, 7, 32
		EMReadScreen mfip_case_budg_food_prorated_amt, 			10, 8, 32
		EMReadScreen mfip_case_budg_entitlement_cash_portion, 	10, 10, 32
		EMReadScreen mfip_case_budg_mand_sanc_vendor, 			10, 11, 32
		EMReadScreen mfip_case_budg_net_cash_after_sanc_portion, 10, 12, 32
		EMReadScreen mfip_case_budg_cash_prorated_amt, 			10, 13, 32
		EMReadScreen mfip_case_budg_state_food_benefit, 		10, 15, 32
		EMReadScreen mfip_case_budg_state_food_prorated_amt, 	10, 16, 32
		' EMReadScreen mfip_case_budg_entitlement_cash_portion, 10, 10, 32

		EMReadScreen mfip_case_budg_grant_amount, 				10, 5, 71
		EMReadScreen mfip_case_budg_amt_already_issued, 		10, 8, 71
		EMReadScreen mfip_case_budg_supplement_due, 			10, 9, 71
		EMReadScreen mfip_case_budg_overpayment, 				10, 10, 71
		EMReadScreen mfip_case_budg_adjusted_grant_amt, 		10, 11, 71
		EMReadScreen mfip_case_budg_recoupment, 				10, 12, 71
		EMReadScreen mfip_case_budg_total_food_issuance, 		10, 14, 71
		EMReadScreen mfip_case_budg_total_cash_issuance, 		10, 15, 71
		EMReadScreen mfip_case_budg_total_housing_grant_issuance, 10, 16, 71

		mfip_case_budg_prorate_date = trim(mfip_case_budg_prorate_date)
		mfip_case_budg_fed_food_benefit = trim(mfip_case_budg_fed_food_benefit)
		mfip_case_budg_food_prorated_amt = trim(mfip_case_budg_food_prorated_amt)
		mfip_case_budg_entitlement_cash_portion = trim(mfip_case_budg_entitlement_cash_portion)
		mfip_case_budg_mand_sanc_vendor = trim(mfip_case_budg_mand_sanc_vendor)
		mfip_case_budg_net_cash_after_sanc_portion = trim(mfip_case_budg_net_cash_after_sanc_portion)
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
		mfip_case_budg_cash_supplement = trim(mfip_case_budg_cash_supplement)
		mfip_case_budg_housing_grant_supplement = trim(mfip_case_budg_housing_grant_supplement)
		transmit

		' Call write_value_and_transmit("X", 10, 44)			'Overpayment pop-up - MAYBE WE DON"T NEED THIS?
		Call write_value_and_transmit("X", 12, 44)			'Recoupment pop-up
		EMReadScreen mfip_case_budg_cash_recoupment, 10, 7, 51
		EMReadScreen mfip_case_budg_state_food_recoupment, 10, 8, 51
		EMReadScreen mfip_case_budg_food_recoupment, 10, 9, 51

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

		Call write_value_and_transmit("X", 15, 44)			'Total Cash Issuance pop-up
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
		' Msgbox mfip_case_summary_net_grant_amount

		If mfip_case_asst_unit_caregivers = "0" Then mfip_child_only = True

		Call Back_to_SELF
	end sub

end class

class msa_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result

	public msa_elig_ref_numbs()
	public msa_elig_membs_full_name()
	public msa_elig_membs_request_yn()
	public msa_elig_membs_member_code()
	public msa_elig_membs_member_info()
	public msa_elig_membs_elig_status()
	public msa_elig_membs_elig_basis_code()
	public msa_elig_membs_elig_basis_info()
	public msa_elig_membs_begin_date()
	public msa_elig_membs_budget_cycle()
	public msa_elig_membs_test_absence()
	public msa_elig_membs_test_age()
	public msa_elig_membs_test_basis_of_eligibility()
	public msa_elig_membs_test_citizenship()
	public msa_elig_membs_test_dupl_assistance()
	public msa_elig_membs_test_fail_coop()
	public msa_elig_membs_test_fraud()
	public msa_elig_membs_test_ive_eligible()
	public msa_elig_membs_test_living_arrangement()
	public msa_elig_membs_test_ssi_basis()
	public msa_elig_membs_test_ssn_coop()
	public msa_elig_membs_test_unit_member()
	public msa_elig_membs_test_verif()
	public msa_elig_membs_test_absence_absent()
	public msa_elig_membs_test_absence_death()
	public msa_elig_membs_test_fail_coop_sign_iaas()
	public msa_elig_membs_test_fail_coop_applied_other_benefits()
	public msa_elig_membs_test_unit_member_faci()
	public msa_elig_membs_test_unit_member_relationship()
	public msa_elig_membs_test_verif_date_of_birth()
	public msa_elig_membs_test_verif_disability()
	public msa_elig_membs_test_verif_ssi()

	public msa_elig_budg_memb_gross_earned_income()
	public msa_elig_budg_memb_blind_disa_student()
	public msa_elig_budg_memb_standard_disregard()
	public msa_elig_budg_memb_earned_income()
	public msa_elig_budg_memb_standard_EI_disregard()
	public msa_elig_budg_memb_work_expense_disa()
	public msa_elig_budg_memb_earned_inc_subtotal()
	public msa_elig_budg_memb_earned_inc_disregard()
	public msa_elig_budg_memb_work_expense_blind()
	public msa_elig_budg_memb_net_earned_income()
	public msa_elig_budg_memb_special_needs_total()

	public msa_elig_case_test_applicant_eligible
	public msa_elig_case_test_application_withdrawn
	public msa_elig_case_test_eligible_member
	public msa_elig_case_test_fail_file
	public msa_elig_case_test_prosp_gross_income
	public msa_elig_case_test_prosp_net_income
	public msa_elig_case_test_residence
	public msa_elig_case_test_assets
	public msa_elig_case_test_retro_net_income
	public msa_elig_case_test_verif
	public msa_elig_case_shared_hh_yn

	public msa_elig_case_test_fail_file_revw
	public msa_elig_case_test_fail_file_hrf
	public msa_elig_case_test_prosp_gross_earned_income
	public msa_elig_case_test_prosp_gross_unearned_income
	public msa_elig_case_test_prosp_gross_deemed_income
	public msa_elig_case_test_prosp_total_gross_income
	public msa_elig_case_test_prosp_gross_ssi_need_standard
	public msa_elig_case_test_prosp_gross_ssi_standard_multiplier
	public msa_elig_case_test_prosp_gross_income_limit
	public msa_elig_case_test_total_countable_assets
	public msa_elig_case_test_maximum_assets
	public msa_elig_case_test_verif_acct
	public msa_elig_case_test_verif_addr
	public msa_elig_case_test_verif_busi
	public msa_elig_case_test_verif_cars
	public msa_elig_case_test_verif_jobs
	public msa_elig_case_test_verif_lump
	public msa_elig_case_test_verif_pact
	public msa_elig_case_test_verif_rbic
	public msa_elig_case_test_verif_secu
	public msa_elig_case_test_verif_spon
	public msa_elig_case_test_verif_stin
	public msa_elig_case_test_verif_unea

	public msa_elig_case_budg_type

	public msa_elig_budg_ssi_standard_fbr
	public msa_elig_budg_standard_disregard
	public msa_elig_budg_unearned_income
	public msa_elig_budg_deemed_income
	public msa_elig_budg_net_unearned_income
	public msa_elig_budg_net_earned_income

	public msa_elig_budg_spec_standard_ref_numb()
	public msa_elig_budg_spec_standard_type_code()
	public msa_elig_budg_spec_standard_type_info()
	public msa_elig_budg_spec_standard_amount()

	public msa_elig_budg_need_standard
	public msa_elig_budg_net_income
	public msa_elig_budg_msa_grant
	public msa_elig_budg_amount_already_issued
	public msa_elig_budg_supplement_due
	public msa_elig_budg_overpayment
	public msa_elig_budg_adjusted_grant_amount
	public msa_elig_budg_recoupment
	public msa_elig_budg_current_payment

	public msa_elig_budg_basic_needs_assistance_standard
	public msa_elig_budg_special_needs
	public msa_elig_budg_household_total_needs

	public msa_elig_summ_approved_date
	public msa_elig_summ_process_date
	public msa_elig_summ_date_last_approval
	public msa_elig_summ_curr_prog_status
	public msa_elig_summ_eligibility_result
	public msa_elig_summ_reporting_status
	public msa_elig_summ_source_of_info
	public msa_elig_summ_benefit
	public msa_elig_summ_recertification_date
	public msa_elig_summ_budget_cycle
	public msa_elig_summ_eligible_houshold_members
	public msa_elig_summ_shared_houshold
	public msa_elig_summ_vendor_reason_code
	public msa_elig_summ_vendor_reason_info
	public msa_elig_summ_responsible_county
	public msa_elig_summ_servicing_county
	public msa_elig_summ_total_assets
	public msa_elig_summ_maximum_assets
	public msa_elig_summ_grant
	public msa_elig_summ_current_payment
	public msa_elig_summ_worker_message

	public sub read_elig()

		call navigate_to_MAXIS_screen("ELIG", "MSA ")
		EMWriteScreen elig_footer_month, 20, 56
		EMWriteScreen elig_footer_year, 20, 59
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)

		ReDim msa_elig_ref_numbs(0)
		ReDim msa_elig_membs_full_name(0)
		ReDim msa_elig_membs_request_yn(0)
		ReDim msa_elig_membs_member_code(0)
		ReDim msa_elig_membs_member_info(0)
		ReDim msa_elig_membs_elig_status(0)
		ReDim msa_elig_membs_elig_basis_code(0)
		ReDim msa_elig_membs_elig_basis_info(0)
		ReDim msa_elig_membs_begin_date(0)
		ReDim msa_elig_membs_budget_cycle(0)
		ReDim msa_elig_membs_test_absence(0)
		ReDim msa_elig_membs_test_age(0)
		ReDim msa_elig_membs_test_basis_of_eligibility(0)
		ReDim msa_elig_membs_test_citizenship(0)
		ReDim msa_elig_membs_test_dupl_assistance(0)
		ReDim msa_elig_membs_test_fail_coop(0)
		ReDim msa_elig_membs_test_fraud(0)
		ReDim msa_elig_membs_test_ive_eligible(0)
		ReDim msa_elig_membs_test_living_arrangement(0)
		ReDim msa_elig_membs_test_ssi_basis(0)
		ReDim msa_elig_membs_test_ssn_coop(0)
		ReDim msa_elig_membs_test_unit_member(0)
		ReDim msa_elig_membs_test_verif(0)
		ReDim msa_elig_membs_test_absence_absent(0)
		ReDim msa_elig_membs_test_absence_death(0)
		ReDim msa_elig_membs_test_fail_coop_sign_iaas(0)
		ReDim msa_elig_membs_test_fail_coop_applied_other_benefits(0)
		ReDim msa_elig_membs_test_unit_member_faci(0)
		ReDim msa_elig_membs_test_unit_member_relationship(0)
		ReDim msa_elig_membs_test_verif_date_of_birth(0)
		ReDim msa_elig_membs_test_verif_disability(0)
		ReDim msa_elig_membs_test_verif_ssi(0)
		ReDim msa_elig_budg_memb_special_needs_total(0)


		ReDim msa_elig_budg_spec_standard_ref_numb(0)
		ReDim msa_elig_budg_spec_standard_type_code(0)
		ReDim msa_elig_budg_spec_standard_type_info(0)
		ReDim msa_elig_budg_spec_standard_amount(0)

		elig_memb_count = 0
		msa_row = 7
		Do
			EMReadScreen ref_numb, 2, msa_row, 5

			ReDim preserve msa_elig_ref_numbs(elig_memb_count)
			ReDim preserve msa_elig_membs_full_name(elig_memb_count)
			ReDim preserve msa_elig_membs_request_yn(elig_memb_count)
			ReDim preserve msa_elig_membs_member_code(elig_memb_count)
			ReDim preserve msa_elig_membs_member_info(elig_memb_count)
			ReDim preserve msa_elig_membs_elig_status(elig_memb_count)
			ReDim preserve msa_elig_membs_elig_basis_code(elig_memb_count)
			ReDim preserve msa_elig_membs_elig_basis_info(elig_memb_count)
			ReDim preserve msa_elig_membs_begin_date(elig_memb_count)
			ReDim preserve msa_elig_membs_budget_cycle(elig_memb_count)
			ReDim preserve msa_elig_membs_test_absence(elig_memb_count)
			ReDim preserve msa_elig_membs_test_age(elig_memb_count)
			ReDim preserve msa_elig_membs_test_basis_of_eligibility(elig_memb_count)
			ReDim preserve msa_elig_membs_test_citizenship(elig_memb_count)
			ReDim preserve msa_elig_membs_test_dupl_assistance(elig_memb_count)
			ReDim preserve msa_elig_membs_test_fail_coop(elig_memb_count)
			ReDim preserve msa_elig_membs_test_fraud(elig_memb_count)
			ReDim preserve msa_elig_membs_test_ive_eligible(elig_memb_count)
			ReDim preserve msa_elig_membs_test_living_arrangement(elig_memb_count)
			ReDim preserve msa_elig_membs_test_ssi_basis(elig_memb_count)
			ReDim preserve msa_elig_membs_test_ssn_coop(elig_memb_count)
			ReDim preserve msa_elig_membs_test_unit_member(elig_memb_count)
			ReDim preserve msa_elig_membs_test_verif(elig_memb_count)
			ReDim preserve msa_elig_membs_test_absence_absent(elig_memb_count)
			ReDim preserve msa_elig_membs_test_absence_death(elig_memb_count)
			ReDim preserve msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count)
			ReDim preserve msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count)
			ReDim preserve msa_elig_membs_test_unit_member_faci(elig_memb_count)
			ReDim preserve msa_elig_membs_test_unit_member_relationship(elig_memb_count)
			ReDim preserve msa_elig_membs_test_verif_date_of_birth(elig_memb_count)
			ReDim preserve msa_elig_membs_test_verif_disability(elig_memb_count)
			ReDim preserve msa_elig_membs_test_verif_ssi(elig_memb_count)
			ReDim preserve msa_elig_budg_memb_special_needs_total(elig_memb_count)

			msa_elig_ref_numbs(elig_memb_count) = ref_numb

			EMReadScreen msa_elig_membs_request_yn(elig_memb_count), 1, msa_row, 25

			EMReadScreen msa_elig_membs_member_code(elig_memb_count), 1, msa_row, 29
			If msa_elig_membs_member_code(elig_memb_count) = "A" Then msa_elig_membs_member_info(elig_memb_count) = "Eligible"
			If msa_elig_membs_member_code(elig_memb_count) = "1" Then msa_elig_membs_member_info(elig_memb_count) = "Non-MSA Spouse"
			If msa_elig_membs_member_code(elig_memb_count) = "2" Then msa_elig_membs_member_info(elig_memb_count) = "Non-MSA Parent - Deem Income/Resources"
			If msa_elig_membs_member_code(elig_memb_count) = "4" Then msa_elig_membs_member_info(elig_memb_count) = "Step Parent - Deem Resources"
			If msa_elig_membs_member_code(elig_memb_count) = "N" Then msa_elig_membs_member_info(elig_memb_count) = "Not Counted"
			If msa_elig_membs_member_code(elig_memb_count) = "I" Then msa_elig_membs_member_info(elig_memb_count) = "Ineligible"

			EMReadScreen msa_elig_membs_elig_status(elig_memb_count), 10, msa_row, 46
			msa_elig_membs_elig_status(elig_memb_count) = trim(msa_elig_membs_elig_status(elig_memb_count))

			EMReadScreen msa_elig_membs_elig_basis_code(elig_memb_count), 1, msa_row, 59
			If msa_elig_membs_elig_basis_code(elig_memb_count) = "A" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Aged"
			If msa_elig_membs_elig_basis_code(elig_memb_count) = "B" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Blind"
			If msa_elig_membs_elig_basis_code(elig_memb_count) = "D" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Disabled"
			If msa_elig_membs_elig_basis_code(elig_memb_count) = "S" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "SSI"
			If msa_elig_membs_elig_basis_code(elig_memb_count) = " " Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Blank"

			EMReadScreen msa_elig_membs_begin_date(elig_memb_count), 8, msa_row, 63
			msa_elig_membs_begin_date(elig_memb_count) = trim(msa_elig_membs_begin_date(elig_memb_count))
			If msa_elig_membs_begin_date(elig_memb_count) <> "" then msa_elig_membs_begin_date(elig_memb_count) = replace(msa_elig_membs_begin_date(elig_memb_count), " ", "/")

			EMReadScreen msa_elig_membs_budget_cycle(elig_memb_count), 1, msa_row, 76
			If msa_elig_membs_budget_cycle(elig_memb_count) = "P" Then msa_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
			If msa_elig_membs_budget_cycle(elig_memb_count) = "R" Then msa_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

			Call write_value_and_transmit("X", msa_row, 3)

			EMReadScreen full_name_information, 20, 7, 10
			full_name_information = trim(full_name_information)
			name_array = split(full_name_information, " ")
			For each name_parts in name_array
				If len(name_parts) <> 1 Then msa_elig_membs_full_name(elig_memb_count) = msa_elig_membs_full_name(elig_memb_count) & " " & name_parts
			Next
			msa_elig_membs_full_name(elig_memb_count) = trim((msa_elig_membs_full_name(elig_memb_count)))

			EMReadScreen msa_elig_membs_test_absence(elig_memb_count), 				6, 10, 8
			EMReadScreen msa_elig_membs_test_age(elig_memb_count), 					6, 11, 8
			EMReadScreen msa_elig_membs_test_basis_of_eligibility(elig_memb_count), 6, 12, 8
			EMReadScreen msa_elig_membs_test_citizenship(elig_memb_count), 			6, 13, 8
			EMReadScreen msa_elig_membs_test_dupl_assistance(elig_memb_count), 		6, 14, 8
			EMReadScreen msa_elig_membs_test_fail_coop(elig_memb_count), 			6, 15, 8
			EMReadScreen msa_elig_membs_test_fraud(elig_memb_count), 				6, 16, 8

			EMReadScreen msa_elig_membs_test_ive_eligible(elig_memb_count), 		6, 10, 47
			EMReadScreen msa_elig_membs_test_living_arrangement(elig_memb_count), 	6, 11, 47
			EMReadScreen msa_elig_membs_test_ssi_basis(elig_memb_count), 			6, 12, 47
			EMReadScreen msa_elig_membs_test_ssn_coop(elig_memb_count), 			6, 13, 47
			EMReadScreen msa_elig_membs_test_unit_member(elig_memb_count), 			6, 14, 47
			EMReadScreen msa_elig_membs_test_verif(elig_memb_count), 				6, 15, 47

			msa_elig_membs_test_absence(elig_memb_count) = trim(msa_elig_membs_test_absence(elig_memb_count))
			msa_elig_membs_test_age(elig_memb_count) = trim(msa_elig_membs_test_age(elig_memb_count))
			msa_elig_membs_test_basis_of_eligibility(elig_memb_count) = trim(msa_elig_membs_test_basis_of_eligibility(elig_memb_count))
			msa_elig_membs_test_citizenship(elig_memb_count) = trim(msa_elig_membs_test_citizenship(elig_memb_count))
			msa_elig_membs_test_dupl_assistance(elig_memb_count) = trim(msa_elig_membs_test_dupl_assistance(elig_memb_count))
			msa_elig_membs_test_fail_coop(elig_memb_count) = trim(msa_elig_membs_test_fail_coop(elig_memb_count))
			msa_elig_membs_test_fraud(elig_memb_count) = trim(msa_elig_membs_test_fraud(elig_memb_count))

			msa_elig_membs_test_ive_eligible(elig_memb_count) = trim(msa_elig_membs_test_ive_eligible(elig_memb_count))
			msa_elig_membs_test_living_arrangement(elig_memb_count) = trim(msa_elig_membs_test_living_arrangement(elig_memb_count))
			msa_elig_membs_test_ssi_basis(elig_memb_count) = trim(msa_elig_membs_test_ssi_basis(elig_memb_count))
			msa_elig_membs_test_ssn_coop(elig_memb_count) = trim(msa_elig_membs_test_ssn_coop(elig_memb_count))
			msa_elig_membs_test_unit_member(elig_memb_count) = trim(msa_elig_membs_test_unit_member(elig_memb_count))
			msa_elig_membs_test_verif(elig_memb_count) = trim(msa_elig_membs_test_verif(elig_memb_count))

			Call write_value_and_transmit("X", 10, 6)
			EMReadScreen msa_elig_membs_test_absence_absent(elig_memb_count), 	6, 12, 40
			EMReadScreen msa_elig_membs_test_absence_death(elig_memb_count), 	6, 13, 40

			msa_elig_membs_test_absence_absent(elig_memb_count) = trim(msa_elig_membs_test_absence_absent(elig_memb_count))
			msa_elig_membs_test_absence_death(elig_memb_count) = trim(msa_elig_membs_test_absence_death(elig_memb_count))
			transmit

			Call write_value_and_transmit("X", 15, 6)
			EMReadScreen msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count), 				6, 12, 24
			EMReadScreen msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count), 6, 13, 24

			msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count) = trim(msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count))
			msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count) = trim(msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count))
			transmit

			Call write_value_and_transmit("X", 14, 45)
			EMReadScreen msa_elig_membs_test_unit_member_faci(elig_memb_count), 		6, 12, 24
			EMReadScreen msa_elig_membs_test_unit_member_relationship(elig_memb_count), 6, 13, 24

			msa_elig_membs_test_unit_member_faci(elig_memb_count) = trim(msa_elig_membs_test_unit_member_faci(elig_memb_count))
			msa_elig_membs_test_unit_member_relationship(elig_memb_count) = trim(msa_elig_membs_test_unit_member_relationship(elig_memb_count))
			transmit

			Call write_value_and_transmit("X", 15, 45)
			EMReadScreen msa_elig_membs_test_verif_date_of_birth(elig_memb_count), 	6, 12, 30
			EMReadScreen msa_elig_membs_test_verif_disability(elig_memb_count), 	6, 13, 30
			EMReadScreen msa_elig_membs_test_verif_ssi(elig_memb_count), 			6, 14, 30

			msa_elig_membs_test_verif_date_of_birth(elig_memb_count) = trim(msa_elig_membs_test_verif_date_of_birth(elig_memb_count))
			msa_elig_membs_test_verif_disability(elig_memb_count) = trim(msa_elig_membs_test_verif_disability(elig_memb_count))
			msa_elig_membs_test_verif_ssi(elig_memb_count) = trim(msa_elig_membs_test_verif_ssi(elig_memb_count))
			transmit

			transmit

			msa_row = msa_row + 1
			elig_memb_count = elig_memb_count + 1
			EMReadScreen next_ref_numb, 2, msa_row, 5
		Loop until next_ref_numb = "  "
		transmit 		'going to the next panel - MSCR

		EMReadScreen msa_elig_case_test_applicant_eligible, 	6, 6, 7
		EMReadScreen msa_elig_case_test_application_withdrawn, 	6, 7, 7
		EMReadScreen msa_elig_case_test_eligible_member, 		6, 8, 7
		EMReadScreen msa_elig_case_test_fail_file, 				6, 9, 7
		EMReadScreen msa_elig_case_test_prosp_gross_income, 	6, 10, 7

		EMReadScreen msa_elig_case_test_prosp_net_income, 	6, 6, 45
		EMReadScreen msa_elig_case_test_residence, 			6, 7, 45
		EMReadScreen msa_elig_case_test_assets, 			6, 8, 45
		EMReadScreen msa_elig_case_test_retro_net_income, 	6, 9, 45
		EMReadScreen msa_elig_case_test_verif, 				6, 10, 45

		EMReadScreen msa_elig_case_shared_hh_yn, 1, 13, 61

		msa_elig_case_test_applicant_eligible = trim(msa_elig_case_test_applicant_eligible)
		msa_elig_case_test_application_withdrawn = trim(msa_elig_case_test_application_withdrawn)
		msa_elig_case_test_eligible_member = trim(msa_elig_case_test_eligible_member)
		msa_elig_case_test_fail_file = trim(msa_elig_case_test_fail_file)
		msa_elig_case_test_prosp_gross_income = trim(msa_elig_case_test_prosp_gross_income)

		msa_elig_case_test_prosp_net_income = trim(msa_elig_case_test_prosp_net_income)
		msa_elig_case_test_residence = trim(msa_elig_case_test_residence)
		msa_elig_case_test_assets = trim(msa_elig_case_test_assets)
		msa_elig_case_test_retro_net_income = trim(msa_elig_case_test_retro_net_income)
		msa_elig_case_test_verif = trim(msa_elig_case_test_verif)

		If msa_elig_case_test_fail_file <> "NA" Then
			Call write_value_and_transmit("X", 9, 5)

			EMReadScreen msa_elig_case_test_fail_file_revw, 6, 8, 28
			EMReadScreen msa_elig_case_test_fail_file_hrf, 6, 9, 28

			msa_elig_case_test_fail_file_revw = trim(msa_elig_case_test_fail_file_revw)
			msa_elig_case_test_fail_file_hrf = trim(msa_elig_case_test_fail_file_hrf)
			transmit
		End If

		If msa_elig_case_test_prosp_gross_income <> "NA" Then
			Call write_value_and_transmit("X", 10, 5)

			EMReadScreen msa_elig_case_test_prosp_gross_earned_income, 		9, 9, 55
			EMReadScreen msa_elig_case_test_prosp_gross_unearned_income, 	9, 10, 55
			EMReadScreen msa_elig_case_test_prosp_gross_deemed_income, 		9, 11, 55

			EMReadScreen msa_elig_case_test_prosp_total_gross_income, 			9, 13, 55
			EMReadScreen msa_elig_case_test_prosp_gross_ssi_need_standard, 		9, 14, 55
			EMReadScreen msa_elig_case_test_prosp_gross_ssi_standard_multiplier, 1, 15, 63
			EMReadScreen msa_elig_case_test_prosp_gross_income_limit, 			9, 16, 55


			msa_elig_case_test_prosp_gross_earned_income = trim(msa_elig_case_test_prosp_gross_earned_income)
			msa_elig_case_test_prosp_gross_unearned_income = trim(msa_elig_case_test_prosp_gross_unearned_income)
			msa_elig_case_test_prosp_gross_deemed_income = trim(msa_elig_case_test_prosp_gross_deemed_income)

			msa_elig_case_test_prosp_total_gross_income = trim(msa_elig_case_test_prosp_total_gross_income)
			msa_elig_case_test_prosp_gross_ssi_need_standard = trim(msa_elig_case_test_prosp_gross_ssi_need_standard)
			msa_elig_case_test_prosp_gross_income_limit = trim(msa_elig_case_test_prosp_gross_income_limit)
			transmit
		End If

		If msa_elig_case_test_assets <> "NA" Then
			Call write_value_and_transmit("X", 8, 43)

			EMReadScreen msa_elig_case_test_total_countable_assets, 10, 8, 48
			EMReadScreen msa_elig_case_test_maximum_assets, 		10, 9, 48

			msa_elig_case_test_total_countable_assets = replace(msa_elig_case_test_total_countable_assets, "_", "")
			msa_elig_case_test_maximum_assets = replace(msa_elig_case_test_maximum_assets, "_", "")
			transmit
		End If

		If msa_elig_case_test_verif <> "NA" Then
			Call write_value_and_transmit("X", 10, 43)

			EMReadScreen msa_elig_case_test_verif_acct, 6, 6, 32
			EMReadScreen msa_elig_case_test_verif_addr, 6, 7, 32
			EMReadScreen msa_elig_case_test_verif_busi, 6, 8, 32
			EMReadScreen msa_elig_case_test_verif_cars, 6, 9, 32
			EMReadScreen msa_elig_case_test_verif_jobs, 6, 10, 32
			EMReadScreen msa_elig_case_test_verif_lump, 6, 11, 32
			EMReadScreen msa_elig_case_test_verif_pact, 6, 12, 32
			EMReadScreen msa_elig_case_test_verif_rbic, 6, 13, 32
			EMReadScreen msa_elig_case_test_verif_secu, 6, 14, 32
			EMReadScreen msa_elig_case_test_verif_spon, 6, 15, 32
			EMReadScreen msa_elig_case_test_verif_stin, 6, 16, 32
			EMReadScreen msa_elig_case_test_verif_unea, 6, 17, 32

			msa_elig_case_test_verif_acct = trim(msa_elig_case_test_verif_acct)
			msa_elig_case_test_verif_addr = trim(msa_elig_case_test_verif_addr)
			msa_elig_case_test_verif_busi = trim(msa_elig_case_test_verif_busi)
			msa_elig_case_test_verif_cars = trim(msa_elig_case_test_verif_cars)
			msa_elig_case_test_verif_jobs = trim(msa_elig_case_test_verif_jobs)
			msa_elig_case_test_verif_lump = trim(msa_elig_case_test_verif_lump)
			msa_elig_case_test_verif_pact = trim(msa_elig_case_test_verif_pact)
			msa_elig_case_test_verif_rbic = trim(msa_elig_case_test_verif_rbic)
			msa_elig_case_test_verif_secu = trim(msa_elig_case_test_verif_secu)
			msa_elig_case_test_verif_spon = trim(msa_elig_case_test_verif_spon)
			msa_elig_case_test_verif_stin = trim(msa_elig_case_test_verif_stin)
			msa_elig_case_test_verif_unea = trim(msa_elig_case_test_verif_unea)
			transmit
		End If

		transmit 		'going to the next panel - MSCB

		EmReadScreen msa_elig_case_budg_type, 12, 3, 25
		msa_elig_case_budg_type = trim(msa_elig_case_budg_type)

		If msa_elig_case_budg_type = "SSI TYPE" Then
			EMReadScreen msa_elig_budg_ssi_standard_fbr, 	9, 6, 32
			EMReadScreen msa_elig_budg_standard_disregard, 	9, 7, 32

			msa_elig_budg_ssi_standard_fbr = trim(msa_elig_budg_ssi_standard_fbr)
			msa_elig_budg_standard_disregard = trim(msa_elig_budg_standard_disregard)
		End If

		If msa_elig_case_budg_type = "Non-SSI TYPE" Then
			EMReadScreen msa_elig_budg_unearned_income, 	9, 6, 32
			EMReadScreen msa_elig_budg_deemed_income, 		9, 7, 32
			EMReadScreen msa_elig_budg_standard_disregard, 	9, 8, 32
			EMReadScreen msa_elig_budg_net_unearned_income, 9, 9, 32
			EMReadScreen msa_elig_budg_net_earned_income, 	9, 10, 32

			msa_elig_budg_unearned_income = trim(msa_elig_budg_unearned_income)
			msa_elig_budg_deemed_income = trim(msa_elig_budg_deemed_income)
			msa_elig_budg_standard_disregard = trim(msa_elig_budg_standard_disregard)
			msa_elig_budg_net_unearned_income = trim(msa_elig_budg_net_unearned_income)
			msa_elig_budg_net_earned_income = trim(msa_elig_budg_net_earned_income)

			Call write_value_and_transmit("X", 10, 3)

			EMReadScreen msa_elig_budg_gross_earned_income, 	9, 9, 42
			EMReadScreen msa_elig_budg_blind_disa_student, 		9, 10, 42
			EMReadScreen msa_elig_budg_earned_standard_disregard, 9, 11, 42
			EMReadScreen msa_elig_budg_earned_income, 			9, 12, 42
			EMReadScreen msa_elig_budg_standard_EI_disregard, 	9, 13, 42
			EMReadScreen msa_elig_budg_work_expense_disa, 		9, 14, 42
			EMReadScreen msa_elig_budg_earned_inc_subtotal, 	9, 15, 42
			EMReadScreen msa_elig_budg_earned_inc_disregard, 	9, 16, 42
			EMReadScreen msa_elig_budg_work_expense_blind, 		9, 17, 42

			EMReadScreen ref_numb_one, 2, 7, 62
			If ref_numb_one <> "  " Then
				For memn_count = 0 to UBound(msa_elig_ref_numbs)
					If ref_numb_one = msa_elig_ref_numbs(memn_count) Then
						EMReadScreen msa_elig_budg_memb_gross_earned_income(memn_count), 	9, 9, 54
						EMReadScreen msa_elig_budg_memb_blind_disa_student(memn_count), 	9, 10, 54
						EMReadScreen msa_elig_budg_memb_standard_disregard(memn_count), 	9, 11, 54
						EMReadScreen msa_elig_budg_memb_earned_income(memn_count), 			9, 12, 54
						EMReadScreen msa_elig_budg_memb_standard_EI_disregard(memn_count), 	9, 13, 54
						EMReadScreen msa_elig_budg_memb_work_expense_disa(memn_count), 		9, 14, 54
						EMReadScreen msa_elig_budg_memb_earned_inc_subtotal(memn_count), 	9, 15, 54
						EMReadScreen msa_elig_budg_memb_earned_inc_disregard(memn_count), 	9, 16, 54
						EMReadScreen msa_elig_budg_memb_work_expense_blind(memn_count), 	9, 17, 54
						EMReadScreen msa_elig_budg_memb_net_earned_income(memn_count), 		9, 18, 54

						msa_elig_budg_memb_gross_earned_income(memn_count) = trim(msa_elig_budg_memb_gross_earned_income(memn_count))
						msa_elig_budg_memb_blind_disa_student(memn_count) = trim(msa_elig_budg_memb_blind_disa_student(memn_count))
						msa_elig_budg_memb_standard_disregard(memn_count) = trim(msa_elig_budg_memb_standard_disregard(memn_count))
						msa_elig_budg_memb_earned_income(memn_count) = trim(msa_elig_budg_memb_earned_income(memn_count))
						msa_elig_budg_memb_standard_EI_disregard(memn_count) = trim(msa_elig_budg_memb_standard_EI_disregard(memn_count))
						msa_elig_budg_memb_work_expense_disa(memn_count) = trim(msa_elig_budg_memb_work_expense_disa(memn_count))
						msa_elig_budg_memb_earned_inc_subtotal(memn_count) = trim(msa_elig_budg_memb_earned_inc_subtotal(memn_count))
						msa_elig_budg_memb_earned_inc_disregard(memn_count) = trim(msa_elig_budg_memb_earned_inc_disregard(memn_count))
						msa_elig_budg_memb_work_expense_blind(memn_count) = trim(msa_elig_budg_memb_work_expense_blind(memn_count))
						msa_elig_budg_memb_net_earned_income(memn_count) = trim(msa_elig_budg_memb_net_earned_income(memn_count))
					End If
				Next
			End if

			EMReadScreen ref_numb_two, 2, 7, 75
			If ref_numb_two <> "  " Then
				For memn_count = 0 to UBound(msa_elig_ref_numbs)
					If ref_numb_two = msa_elig_ref_numbs(memn_count) Then
						EMReadScreen msa_elig_budg_memb_gross_earned_income(memn_count), 	9, 9, 67
						EMReadScreen msa_elig_budg_memb_blind_disa_student(memn_count), 	9, 10, 67
						EMReadScreen msa_elig_budg_memb_standard_disregard(memn_count), 	9, 11, 67
						EMReadScreen msa_elig_budg_memb_earned_income(memn_count), 			9, 12, 67
						EMReadScreen msa_elig_budg_memb_standard_EI_disregard(memn_count), 	9, 13, 67
						EMReadScreen msa_elig_budg_memb_work_expense_disa(memn_count), 		9, 14, 67
						EMReadScreen msa_elig_budg_memb_earned_inc_subtotal(memn_count), 	9, 15, 67
						EMReadScreen msa_elig_budg_memb_earned_inc_disregard(memn_count), 	9, 16, 67
						EMReadScreen msa_elig_budg_memb_work_expense_blind(memn_count), 	9, 17, 67
						EMReadScreen msa_elig_budg_memb_net_earned_income(memn_count), 		9, 18, 67

						msa_elig_budg_memb_gross_earned_income(memn_count) = trim(msa_elig_budg_memb_gross_earned_income(memn_count))
						msa_elig_budg_memb_blind_disa_student(memn_count) = trim(msa_elig_budg_memb_blind_disa_student(memn_count))
						msa_elig_budg_memb_standard_disregard(memn_count) = trim(msa_elig_budg_memb_standard_disregard(memn_count))
						msa_elig_budg_memb_earned_income(memn_count) = trim(msa_elig_budg_memb_earned_income(memn_count))
						msa_elig_budg_memb_standard_EI_disregard(memn_count) = trim(msa_elig_budg_memb_standard_EI_disregard(memn_count))
						msa_elig_budg_memb_work_expense_disa(memn_count) = trim(msa_elig_budg_memb_work_expense_disa(memn_count))
						msa_elig_budg_memb_earned_inc_subtotal(memn_count) = trim(msa_elig_budg_memb_earned_inc_subtotal(memn_count))
						msa_elig_budg_memb_earned_inc_disregard(memn_count) = trim(msa_elig_budg_memb_earned_inc_disregard(memn_count))
						msa_elig_budg_memb_work_expense_blind(memn_count) = trim(msa_elig_budg_memb_work_expense_blind(memn_count))
						msa_elig_budg_memb_net_earned_income(memn_count) = trim(msa_elig_budg_memb_net_earned_income(memn_count))
					End If
				Next
			End if
			transmit
		End If

		EMReadScreen msa_elig_budg_need_standard, 			9, 6, 72
		EMReadScreen msa_elig_budg_net_income, 				9, 7, 72
		EMReadScreen msa_elig_budg_msa_grant, 				9, 8, 72

		EMReadScreen msa_elig_budg_amount_already_issued, 	9, 11, 72
		EMReadScreen msa_elig_budg_supplement_due, 			9, 12, 72
		EMReadScreen msa_elig_budg_overpayment, 			9, 13, 72

		EMReadScreen msa_elig_budg_adjusted_grant_amount, 	9, 15, 72
		EMReadScreen msa_elig_budg_recoupment, 				9, 16, 72
		EMReadScreen msa_elig_budg_current_payment, 		9, 17, 72

		msa_elig_budg_need_standard = trim(msa_elig_budg_need_standard)
		msa_elig_budg_net_income = trim(msa_elig_budg_net_income)
		msa_elig_budg_msa_grant = trim(msa_elig_budg_msa_grant)

		msa_elig_budg_amount_already_issued = trim(msa_elig_budg_amount_already_issued)
		msa_elig_budg_supplement_due = trim(msa_elig_budg_supplement_due)
		msa_elig_budg_overpayment = trim(msa_elig_budg_overpayment)

		msa_elig_budg_adjusted_grant_amount = trim(msa_elig_budg_adjusted_grant_amount)
		msa_elig_budg_recoupment = trim(msa_elig_budg_recoupment)
		msa_elig_budg_current_payment = trim(msa_elig_budg_current_payment)


		Call write_value_and_transmit("X", 6, 43)
		EMReadScreen msa_elig_budg_basic_needs_assistance_standard, 10, 16, 59
		EMReadScreen msa_elig_budg_special_needs, 					10, 17, 59
		EMReadScreen msa_elig_budg_household_total_needs, 			10, 18, 59

		msa_elig_budg_basic_needs_assistance_standard = trim(msa_elig_budg_basic_needs_assistance_standard)
		msa_elig_budg_special_needs = trim(msa_elig_budg_special_needs)
		msa_elig_budg_household_total_needs = trim(msa_elig_budg_household_total_needs)

		msa_col = 6
		spec_needs_count = 0
		For msa_col = 6 to 42 step 36
			EMReadScreen ref_numb, 2, 5, msa_col+9
			If ref_numb <> "  " Then
				For msa_membs = 0 to UBound(msa_elig_ref_numbs)
					If msa_elig_ref_numbs(msa_membs) = ref_numb Then
						EMReadScreen amount_total, 8, 15, msa_col+26
						msa_elig_budg_memb_special_needs_total(msa_membs) = amount_total
					End If
				Next

				EMReadScreen info_code, 2, 8, msa_col
				Do while info_code <> "__"
					ReDim preserve msa_elig_budg_spec_standard_ref_numb(spec_needs_count)
					ReDim preserve msa_elig_budg_spec_standard_type_code(spec_needs_count)
					ReDim preserve msa_elig_budg_spec_standard_type_info(spec_needs_count)
					ReDim preserve msa_elig_budg_spec_standard_amount(spec_needs_count)

					msa_elig_budg_spec_standard_ref_numb(spec_needs_count) = ref_numb
					msa_elig_budg_spec_standard_type_code(spec_needs_count) = info_code
					If info_code = "" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = ""
					If info_code = "01" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - High Protein > 79 Gr/Day"
					If info_code = "02" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Control Protein 40-60 GR/DAY"
					If info_code = "03" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Control Protein < 40 GR/DAY"
					If info_code = "04" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Low Cholesterol"
					If info_code = "05" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - High Residue"
					If info_code = "06" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Pregnancy and Lactation"
					If info_code = "07" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Gluten Free"
					If info_code = "08" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Lactose Free"
					If info_code = "09" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Anti Dumping"
					If info_code = "10" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Hypoglycemic"
					If info_code = "11" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Ketogenic"
					If info_code = "RP" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Representative Payee"
					If info_code = "GF" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Guardianship Fee Max"
					If info_code = "SN" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Shelter Need"
					If info_code = "RM" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Restaurant Meals"
					If info_code = "EN" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Excess Need"
					If info_code = "OT" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Other Need"

					EMReadScreen msa_elig_budg_spec_standard_amount(spec_needs_count), 8, msa_row, msa_col+26
					msa_elig_budg_spec_standard_amount(spec_needs_count) = trim(msa_elig_budg_spec_standard_amount(spec_needs_count))

					msa_row = msa_row + 1
					If msa_row = 14 Then MsgBox "MORE THAN SIX?"
					spec_needs_count = spec_needs_count + 1
					EMReadScreen info_code, 2, msa_row, msa_col
				Loop
			End If
		Next
		transmit

		transmit 		'going to the next panel - MSSM

		EMReadScreen msa_elig_summ_approved_date, 8, 3, 14
		EMReadScreen msa_elig_summ_process_date, 8, 2, 72
		EMReadScreen msa_elig_summ_date_last_approval, 8, 5, 29
		EMReadScreen msa_elig_summ_curr_prog_status, 12, 6, 29
		EMReadScreen msa_elig_summ_eligibility_result, 12, 7, 29
		EMReadScreen msa_elig_summ_reporting_status, 12, 8, 29
		EMReadScreen msa_elig_summ_source_of_info, 4, 10, 29
		EMReadScreen msa_elig_summ_benefit, 12, 11, 29
		EMReadScreen msa_elig_summ_recertification_date, 8, 12, 29
		EMReadScreen msa_elig_summ_budget_cycle, 5, 13, 29
		EMReadScreen msa_elig_summ_eligible_houshold_members, 1, 14, 29
		EMReadScreen msa_elig_summ_shared_houshold, 3, 15, 29
		EMReadScreen msa_elig_summ_vendor_reason_code, 2, 16, 18

		EMReadScreen msa_elig_summ_responsible_county, 2, 5, 73
		EMReadScreen msa_elig_summ_servicing_county, 2, 6, 73
		EMReadScreen msa_elig_summ_total_assets, 9, 7, 72
		EMReadScreen msa_elig_summ_maximum_assets, 9, 8, 72
		EMReadScreen msa_elig_summ_grant, 9, 11, 72
		EMReadScreen msa_elig_summ_current_payment, 9, 17, 72

		EMReadScreen msa_elig_summ_worker_message, 80, 18, 1

		msa_elig_summ_curr_prog_status = trim(msa_elig_summ_curr_prog_status)
		msa_elig_summ_eligibility_result = trim(msa_elig_summ_eligibility_result)
		msa_elig_summ_reporting_status = trim(msa_elig_summ_reporting_status)
		msa_elig_summ_benefit = trim(msa_elig_summ_benefit)
		msa_elig_summ_shared_houshold = trim(msa_elig_summ_shared_houshold)

		If msa_elig_summ_vendor_reason_code = "01" Then msa_elig_summ_vendor_reason_info = "Client Request"
		If msa_elig_summ_vendor_reason_code = "05" Then msa_elig_summ_vendor_reason_info = "Money Mismanagement"
		If msa_elig_summ_vendor_reason_code = "09" Then msa_elig_summ_vendor_reason_info = "Emergency"
		If msa_elig_summ_vendor_reason_code = "10" Then msa_elig_summ_vendor_reason_info = "Chemical Dependency"
		If msa_elig_summ_vendor_reason_code = "11" Then msa_elig_summ_vendor_reason_info = "No Residence"
		If msa_elig_summ_vendor_reason_code = "20" Then msa_elig_summ_vendor_reason_info = "Grant Diversion"

		msa_elig_summ_total_assets = trim(msa_elig_summ_total_assets)
		msa_elig_summ_maximum_assets = trim(msa_elig_summ_maximum_assets)
		msa_elig_summ_grant = trim(msa_elig_summ_grant)
		msa_elig_summ_current_payment = trim(msa_elig_summ_current_payment)

		msa_elig_summ_worker_message = trim(msa_elig_summ_worker_message)

		Call back_to_SELF
	end sub
end class

class ga_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result

	public ga_elig_case_status
	public ga_elig_file_unit_type_code
	public ga_elig_faci_file_unit_type_code
	public ga_elig_file_unit_type_info
	public ga_elig_faci_file_unit_type_info

	public ga_elig_ref_numbs()
	public ga_elig_membs_full_name()
	public ga_elig_membs_relationship_code()
	public ga_elig_membs_relationship_info()
	public ga_elig_membs_code()
	public ga_elig_membs_info()
	public ga_elig_membs_elig_basis_code()
	public ga_elig_membs_elig_basis_info()
	public ga_elig_membs_elig_status()
	public ga_elig_membs_budget_cycle()
	public ga_elig_membs_elig_begin_date()
	public ga_elig_membs_test_absence()
	public ga_elig_membs_test_dupl_assistance()
	public ga_elig_membs_test_ga_coop()
	public ga_elig_membs_test_ive()
	public ga_elig_membs_test_ssi()
	public ga_elig_membs_test_lump_sum_payment()
	public ga_elig_membs_test_unit_member()
	public ga_elig_membs_test_imig_status_verif()
	public ga_elig_membs_test_imig_status()
	public ga_elig_membs_test_basis_of_elig()
	public ga_elig_membs_test_elig_other_prgm()
	public ga_elig_membs_test_ssn_coop()

	public ga_elig_case_test_appl_withdrawn
	public ga_elig_case_test_dupl_assistance
	public ga_elig_case_test_fail_coop
	public ga_elig_case_test_fail_file
	public ga_elig_case_test_eligible_member
	public ga_elig_case_test_prosp_net_income
	public ga_elig_case_test_retro_net_income
	public ga_elig_case_test_residence
	public ga_elig_case_test_assets
	public ga_elig_case_test_eligible_other_prgm
	public ga_elig_case_test_verif
	public ga_elig_case_test_lump_sum_payment

	public ga_elig_case_budg_gross_wages
	public ga_elig_case_budg_gross_self_emp
	public ga_elig_case_budg_total_gross_income
	public ga_elig_case_budg_standard_EI_disregard
	public ga_elig_case_budg_earned_income_subtotal
	public ga_elig_case_budg_earned_income_disregard_percent
	public ga_elig_case_budg_earned_income_disregard_amount
	public ga_elig_case_budg_total_deductions
	public ga_elig_case_budg_net_earned_income
	public ga_elig_case_budg_unearned_income
	public ga_elig_case_budg_counted_school_income
	public ga_elig_case_budg_total_deemed_income
	public ga_elig_case_budg_total_countable_income

	public ga_elig_case_budg_payment_standard
	public ga_elig_case_budg_payment_subtotal
	public ga_elig_case_budg_prorated_from
	public ga_elig_case_budg_prorated_to
	public ga_elig_case_budg_grant_subtotal
	public ga_elig_case_budg_total_assets
	public ga_elig_case_budg_ga_exclusion
	public ga_elig_case_budg_countable_assets
	public ga_elig_case_budg_maximum_assets
	public ga_elig_case_budg_reason_ga_exclusion
	public ga_elig_case_budg_pers_needs_payment_standard
	public ga_elig_case_budg_pers_needs_payment_subtotal
	public ga_elig_case_budg_pers_needs_prorated_from
	public ga_elig_case_budg_pers_needs_prorated_to
	public ga_elig_case_budg_pers_needs_grant_subtotal
	public ga_elig_case_budg_total_ga_grant_amount

	public ga_elig_summ_approved_date
	public ga_elig_summ_process_date
	public ga_elig_summ_date_last_approval
	public ga_elig_summ_curr_prog_status
	public ga_elig_summ_eligibility_result
	public ga_elig_summ_hrf_reporting
	public ga_elig_summ_source_of_info
	public ga_elig_summ_eligibility_begin_date
	public ga_elig_summ_eligiblity_review_date
	public ga_elig_summ_budget_cycle
	public ga_elig_summ_filing_unit_type_code
	public ga_elig_summ_filing_unit_type_info
	public ga_elig_summ_faci_unit_type_code
	public ga_elig_summ_faci_unit_type_info
	public ga_elig_summ_responsible_county
	public ga_elig_summ_vendor_reason_code
	public ga_elig_summ_vendor_reason_info
	public ga_elig_summ_total_assets
	public ga_elig_summ_client_faci_obligation
	public ga_elig_summ_standards
	public ga_elig_summ_counted_income
	public ga_elig_summ_monthly_grant
	public ga_elig_summ_amount_to_be_paid
	public ga_elig_summ_action_code
	public ga_elig_summ_action_info
	public ga_elig_summ_reason_code
	public ga_elig_summ_reason_info
	public ga_elig_summ_worker_message

	public sub read_elig()

		call navigate_to_MAXIS_screen("ELIG", "GA  ")
		EMWriteScreen elig_footer_month, 20, 54
		EMWriteScreen elig_footer_year, 20, 57
		Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result)

 		EMReadScreen ga_elig_case_status, 12, 18, 23
		EMReadScreen ga_elig_file_unit_type_code, 1, 18, 52
		EMReadScreen ga_elig_faci_file_unit_type_code, 1, 18, 77

		ga_elig_case_status = trim(ga_elig_case_status)

		If ga_elig_file_unit_type_code = "1" Then ga_elig_file_unit_type_info = "Single Adult"
		If ga_elig_file_unit_type_code = "2" Then ga_elig_file_unit_type_info = "Single Adult living with Parents"
		If ga_elig_file_unit_type_code = "3" Then ga_elig_file_unit_type_info = "Minor Child Outside the Home"
		If ga_elig_file_unit_type_code = "6" Then ga_elig_file_unit_type_info = "Married Couple"
		If ga_elig_file_unit_type_code = "9" Then ga_elig_file_unit_type_info = "Family State Food Program"

		If ga_elig_faci_file_unit_type_code = "5" Then ga_elig_faci_file_unit_type_info = "Personal Needs"


		ReDim ga_elig_ref_numbs(0)
		ReDim ga_elig_membs_full_name(0)
		ReDim ga_elig_membs_relationship_code(0)
		ReDim ga_elig_membs_relationship_info(0)
		ReDim ga_elig_membs_code(0)
		ReDim ga_elig_membs_info(0)
		ReDim ga_elig_membs_elig_basis_code(0)
		ReDim ga_elig_membs_elig_basis_info(0)
		ReDim ga_elig_membs_elig_status(0)
		ReDim ga_elig_membs_budget_cycle(0)
		ReDim ga_elig_membs_elig_begin_date(0)
		ReDim ga_elig_membs_test_absence(0)
		ReDim ga_elig_membs_test_dupl_assistance(0)
		ReDim ga_elig_membs_test_ga_coop(0)
		ReDim ga_elig_membs_test_ive(0)
		ReDim ga_elig_membs_test_ssi(0)
		ReDim ga_elig_membs_test_lump_sum_payment(0)
		ReDim ga_elig_membs_test_unit_member(0)
		ReDim ga_elig_membs_test_imig_status_verif(0)
		ReDim ga_elig_membs_test_imig_status(0)
		ReDim ga_elig_membs_test_basis_of_elig(0)
		ReDim ga_elig_membs_test_elig_other_prgm(0)
		ReDim ga_elig_membs_test_ssn_coop(0)

		elig_memb_count = 0
		ga_row = 8
		Do
			EMReadScreen ref_numb, 2, ga_row, 9

			ReDim preserve ga_elig_ref_numbs(elig_memb_count)
			ReDim preserve ga_elig_membs_full_name(elig_memb_count)
			ReDim preserve ga_elig_membs_relationship_code(elig_memb_count)
			ReDim preserve ga_elig_membs_relationship_info(elig_memb_count)
			ReDim preserve ga_elig_membs_code(elig_memb_count)
			ReDim preserve ga_elig_membs_info(elig_memb_count)
			ReDim preserve ga_elig_membs_elig_basis_code(elig_memb_count)
			ReDim preserve ga_elig_membs_elig_basis_info(elig_memb_count)
			ReDim preserve ga_elig_membs_elig_status(elig_memb_count)
			ReDim preserve ga_elig_membs_budget_cycle(elig_memb_count)
			ReDim preserve ga_elig_membs_elig_begin_date(elig_memb_count)
			ReDim preserve ga_elig_membs_test_absence(elig_memb_count)
			ReDim preserve ga_elig_membs_test_dupl_assistance(elig_memb_count)
			ReDim preserve ga_elig_membs_test_ga_coop(elig_memb_count)
			ReDim preserve ga_elig_membs_test_ive(elig_memb_count)
			ReDim preserve ga_elig_membs_test_ssi(elig_memb_count)
			ReDim preserve ga_elig_membs_test_lump_sum_payment(elig_memb_count)
			ReDim preserve ga_elig_membs_test_unit_member(elig_memb_count)
			ReDim preserve ga_elig_membs_test_imig_status_verif(elig_memb_count)
			ReDim preserve ga_elig_membs_test_imig_status(elig_memb_count)
			ReDim preserve ga_elig_membs_test_basis_of_elig(elig_memb_count)
			ReDim preserve ga_elig_membs_test_elig_other_prgm(elig_memb_count)
			ReDim preserve ga_elig_membs_test_ssn_coop(elig_memb_count)

			ga_elig_ref_numbs(elig_memb_count) = ref_numb
			EMReadScreen full_name_information, 20, ga_row, 12
			full_name_information = trim(full_name_information)
			name_array = split(full_name_information, " ")
			For each name_parts in name_array
				If len(name_parts) <> 1 Then ga_elig_membs_full_name(elig_memb_count) = ga_elig_membs_full_name(elig_memb_count) & " " & name_parts
			Next
			ga_elig_membs_full_name(elig_memb_count) = trim((ga_elig_membs_full_name(elig_memb_count)))
			EMReadScreen ga_elig_membs_relationship_code(elig_memb_count), 2, ga_row, 33


			If ga_elig_membs_relationship_code(elig_memb_count) = "01" Then ga_elig_membs_relationship_info(elig_memb_count) = "Applicant"
			If ga_elig_membs_relationship_code(elig_memb_count) = "02" Then ga_elig_membs_relationship_info(elig_memb_count) = "Spouse"
			If ga_elig_membs_relationship_code(elig_memb_count) = "03" Then ga_elig_membs_relationship_info(elig_memb_count) = "Child"
			If ga_elig_membs_relationship_code(elig_memb_count) = "04" Then ga_elig_membs_relationship_info(elig_memb_count) = "Parent"
			If ga_elig_membs_relationship_code(elig_memb_count) = "05" Then ga_elig_membs_relationship_info(elig_memb_count) = "Sibling"
			If ga_elig_membs_relationship_code(elig_memb_count) = "06" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Sibling"
			If ga_elig_membs_relationship_code(elig_memb_count) = "08" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Child"
			If ga_elig_membs_relationship_code(elig_memb_count) = "09" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Parent"
			If ga_elig_membs_relationship_code(elig_memb_count) = "10" Then ga_elig_membs_relationship_info(elig_memb_count) = "Aunt"
			If ga_elig_membs_relationship_code(elig_memb_count) = "11" Then ga_elig_membs_relationship_info(elig_memb_count) = "Uncle"
			If ga_elig_membs_relationship_code(elig_memb_count) = "12" Then ga_elig_membs_relationship_info(elig_memb_count) = "Niece"
			If ga_elig_membs_relationship_code(elig_memb_count) = "13" Then ga_elig_membs_relationship_info(elig_memb_count) = "Nephew"
			If ga_elig_membs_relationship_code(elig_memb_count) = "14" Then ga_elig_membs_relationship_info(elig_memb_count) = "Cousin"
			If ga_elig_membs_relationship_code(elig_memb_count) = "15" Then ga_elig_membs_relationship_info(elig_memb_count) = "Grandparent"
			If ga_elig_membs_relationship_code(elig_memb_count) = "16" Then ga_elig_membs_relationship_info(elig_memb_count) = "Grandchild"
			If ga_elig_membs_relationship_code(elig_memb_count) = "17" Then ga_elig_membs_relationship_info(elig_memb_count) = "Other Relative"
			If ga_elig_membs_relationship_code(elig_memb_count) = "18" Then ga_elig_membs_relationship_info(elig_memb_count) = "Legal Guardian"
			If ga_elig_membs_relationship_code(elig_memb_count) = "24" Then ga_elig_membs_relationship_info(elig_memb_count) = "Not Related"
			If ga_elig_membs_relationship_code(elig_memb_count) = "25" Then ga_elig_membs_relationship_info(elig_memb_count) = "Live-In Attendant"
			If ga_elig_membs_relationship_code(elig_memb_count) = "27" Then ga_elig_membs_relationship_info(elig_memb_count) = "Unknown"

			EMReadScreen ga_elig_membs_code(elig_memb_count), 1, ga_row, 48

			If ga_elig_membs_code(elig_memb_count) = "A" Then ga_elig_membs_info(elig_memb_count) = "Assistance Unit Member"
			If ga_elig_membs_code(elig_memb_count) = "C" Then ga_elig_membs_info(elig_memb_count) = "Deemer"
			If ga_elig_membs_code(elig_memb_count) = "F" Then ga_elig_membs_info(elig_memb_count) = "Ineligible - Counted without Deductions"
			If ga_elig_membs_code(elig_memb_count) = "S" Then ga_elig_membs_info(elig_memb_count) = "Ineligible - Counted with Deduction"
			If ga_elig_membs_code(elig_memb_count) = "G" Then ga_elig_membs_info(elig_memb_count) = "Ineligible Affects Grant"
			If ga_elig_membs_code(elig_memb_count) = "I" Then ga_elig_membs_info(elig_memb_count) = "Ineligible Par of Unit"
			If ga_elig_membs_code(elig_memb_count) = "L" Then ga_elig_membs_info(elig_memb_count) = "Other Adult Applicant"
			If ga_elig_membs_code(elig_memb_count) = "M" Then ga_elig_membs_info(elig_memb_count) = "Allocation Only"
			If ga_elig_membs_code(elig_memb_count) = "N" Then ga_elig_membs_info(elig_memb_count) = "Not Counted"
			If ga_elig_membs_code(elig_memb_count) = "U" Then ga_elig_membs_info(elig_memb_count) = "Unknown"

			EMReadScreen ga_elig_membs_elig_basis_code(elig_memb_count), 2, row, 52

			If ga_elig_membs_elig_basis_code(elig_memb_count) = "04" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Permanent Ill Or Incap"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "05" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Temporary Ill Or Incap"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "06" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Care Of Ill Or Incap Mbr"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "07" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Requires Services In Residence"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "09" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Mntl Ill Or Dev Disabled"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "10" then ga_elig_membs_elig_basis_info(elig_memb_count) = "SSI/RSDI Pend"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "11" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Appealing SSI/RSDI Denial"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "12" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Advanced Age"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "13" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Learning Disability"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "17" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Protect/Court Ordered"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "20" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Age 16 Or 17 SS Approval"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "25" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Emancipated Minor"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "28" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Unemployable"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "29" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Displaced Hmkr(Ft Student)"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "30" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Minor W/ Adult Unrelated"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "32" then ga_elig_membs_elig_basis_info(elig_memb_count) = "ESL, Adult/HS At Least Half Time, Adult"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "35" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Drug/Alcohol Addiction(DAA)"
			If ga_elig_membs_elig_basis_code(elig_memb_count) = "99" then ga_elig_membs_elig_basis_info(elig_memb_count) = "No Elig Basis"

			EMReadScreen ga_elig_membs_elig_status(elig_memb_count), 4, row, 57

			If ga_elig_membs_elig_status(elig_memb_count) = "ELIG" then ga_elig_membs_elig_status(elig_memb_count) = "ELIGIBLE"
			If ga_elig_membs_elig_status(elig_memb_count) = "INEL" then ga_elig_membs_elig_status(elig_memb_count) = "INELIGIBLE"

			EMReadScreen ga_elig_membs_budget_cycle(elig_memb_count), 1, row, 63

			If ga_elig_membs_budget_cycle(elig_memb_count) = "P" then ga_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
			If ga_elig_membs_budget_cycle(elig_memb_count) = "R" then ga_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

			EMReadScreen ga_elig_membs_elig_begin_date(elig_memb_count), 8, row, 66

			Call write_value_and_transmit("X", row, 6)

			EMReadScreen ga_elig_membs_test_absence(elig_memb_count), 			6, 11, 12
			EMReadScreen ga_elig_membs_test_dupl_assistance(elig_memb_count), 	6, 12, 12
			EMReadScreen ga_elig_membs_test_ga_coop(elig_memb_count), 			6, 13, 12
			EMReadScreen ga_elig_membs_test_ive(elig_memb_count), 				6, 14, 12
			EMReadScreen ga_elig_membs_test_ssi(elig_memb_count), 				6, 15, 12
			EMReadScreen ga_elig_membs_test_lump_sum_payment(elig_memb_count), 	6, 16, 12


			EMReadScreen ga_elig_membs_test_unit_member(elig_memb_count), 		6, 11, 42
			EMReadScreen ga_elig_membs_test_imig_status_verif(elig_memb_count), 6, 12, 42
			EMReadScreen ga_elig_membs_test_imig_status(elig_memb_count), 		6, 13, 42
			EMReadScreen ga_elig_membs_test_basis_of_elig(elig_memb_count), 	6, 14, 42
			EMReadScreen ga_elig_membs_test_elig_other_prgm(elig_memb_count), 	6, 15, 42
			EMReadScreen ga_elig_membs_test_ssn_coop(elig_memb_count), 			6, 16, 42

			ga_elig_membs_test_absence(elig_memb_count) = trim(ga_elig_membs_test_absence(elig_memb_count))
			ga_elig_membs_test_dupl_assistance(elig_memb_count) = trim(ga_elig_membs_test_dupl_assistance(elig_memb_count))
			ga_elig_membs_test_ga_coop(elig_memb_count) = trim(ga_elig_membs_test_ga_coop(elig_memb_count))
			ga_elig_membs_test_ive(elig_memb_count) = trim(ga_elig_membs_test_ive(elig_memb_count))
			ga_elig_membs_test_ssi(elig_memb_count) = trim(ga_elig_membs_test_ssi(elig_memb_count))
			ga_elig_membs_test_lump_sum_payment(elig_memb_count) = trim(ga_elig_membs_test_lump_sum_payment(elig_memb_count))

			ga_elig_membs_test_unit_member(elig_memb_count) = trim(ga_elig_membs_test_unit_member(elig_memb_count))
			ga_elig_membs_test_imig_status_verif(elig_memb_count) = trim(ga_elig_membs_test_imig_status_verif(elig_memb_count))
			ga_elig_membs_test_imig_status(elig_memb_count) = trim(ga_elig_membs_test_imig_status(elig_memb_count))
			ga_elig_membs_test_basis_of_elig(elig_memb_count) = trim(ga_elig_membs_test_basis_of_elig(elig_memb_count))
			ga_elig_membs_test_elig_other_prgm(elig_memb_count) = trim(ga_elig_membs_test_elig_other_prgm(elig_memb_count))
			ga_elig_membs_test_ssn_coop(elig_memb_count) = trim(ga_elig_membs_test_ssn_coop(elig_memb_count))

			transmit

			ga_row = ga_row + 1
			elig_memb_count = elig_memb_count + 1
			EMReadScreen next_ref_numb, 2, ga_row, 9
		Loop until next_ref_numb = "  "

		transmit 		'going to the next panel - GACR

		EMReadScreen ga_elig_case_test_appl_withdrawn, 		6, 8, 10
		EMReadScreen ga_elig_case_test_dupl_assistance, 	6, 9, 10
		EMReadScreen ga_elig_case_test_fail_coop, 			6, 10, 10
		EMReadScreen ga_elig_case_test_fail_file, 			6, 11, 10
		EMReadScreen ga_elig_case_test_eligible_member, 	6, 12, 10
		EMReadScreen ga_elig_case_test_prosp_net_income, 	6, 13, 10

		EMReadScreen ga_elig_case_test_retro_net_income, 	6, 8, 46
		EMReadScreen ga_elig_case_test_residence, 			6, 9, 46
		EMReadScreen ga_elig_case_test_assets, 				6, 10, 46
		EMReadScreen ga_elig_case_test_eligible_other_prgm, 6, 11, 46
		EMReadScreen ga_elig_case_test_verif, 				6, 12, 46
		EMReadScreen ga_elig_case_test_lump_sum_payment, 	6, 13, 46

		ga_elig_case_test_appl_withdrawn = trim(ga_elig_case_test_appl_withdrawn)
		ga_elig_case_test_dupl_assistance = trim(ga_elig_case_test_dupl_assistance)
		ga_elig_case_test_fail_coop = trim(ga_elig_case_test_fail_coop)
		ga_elig_case_test_fail_file = trim(ga_elig_case_test_fail_file)
		ga_elig_case_test_eligible_member = trim(ga_elig_case_test_eligible_member)
		ga_elig_case_test_prosp_net_income = trim(ga_elig_case_test_prosp_net_income)

		ga_elig_case_test_retro_net_income = trim(ga_elig_case_test_retro_net_income)
		ga_elig_case_test_residence = trim(ga_elig_case_test_residence)
		ga_elig_case_test_assets = trim(ga_elig_case_test_assets)
		ga_elig_case_test_eligible_other_prgm = trim(ga_elig_case_test_eligible_other_prgm)
		ga_elig_case_test_verif = trim(ga_elig_case_test_verif)
		ga_elig_case_test_lump_sum_payment = trim(ga_elig_case_test_lump_sum_payment)

		' Call write_value_and_transmit("X", 13, 4)		'This is the Prosp Net Income Pop-Up - this appears to match the information on GAb1 - so we are not reading it'

		transmit 		'going to the next panel - GAB1

		EMReadScreen ga_elig_case_budg_gross_wages, 					10, 6, 29
		EMReadScreen ga_elig_case_budg_gross_self_emp, 					10, 7, 29
		EMReadScreen ga_elig_case_budg_total_gross_income, 				10, 9, 29
		EMReadScreen ga_elig_case_budg_standard_EI_disregard, 			10, 13, 29
		EMReadScreen ga_elig_case_budg_earned_income_subtotal, 			10, 14, 29
		EMReadScreen ga_elig_case_budg_earned_income_disregard_percent, 2, 15, 23
		EMReadScreen ga_elig_case_budg_earned_income_disregard_amount, 	10, 15, 29
		EMReadScreen ga_elig_case_budg_total_deductions, 				10, 17, 29

		EMReadScreen ga_elig_case_budg_net_earned_income, 				10, 6, 71
		EMReadScreen ga_elig_case_budg_unearned_income, 				10, 8, 71
		EMReadScreen ga_elig_case_budg_counted_school_income, 			10, 10, 71
		EMReadScreen ga_elig_case_budg_total_deemed_income, 			10, 14, 71
		EMReadScreen ga_elig_case_budg_total_countable_income, 			10, 17, 71

		ga_elig_case_budg_gross_wages = trim(ga_elig_case_budg_gross_wages)
		ga_elig_case_budg_gross_self_emp = trim(ga_elig_case_budg_gross_self_emp)
		ga_elig_case_budg_total_gross_income = trim(ga_elig_case_budg_total_gross_income)
		ga_elig_case_budg_standard_EI_disregard = trim(ga_elig_case_budg_standard_EI_disregard)
		ga_elig_case_budg_earned_income_subtotal = trim(ga_elig_case_budg_earned_income_subtotal)
		ga_elig_case_budg_earned_income_disregard_percent = trim(ga_elig_case_budg_earned_income_disregard_percent)
		ga_elig_case_budg_earned_income_disregard_amount = trim(ga_elig_case_budg_earned_income_disregard_amount)
		ga_elig_case_budg_total_deductions = trim(ga_elig_case_budg_total_deductions)

		ga_elig_case_budg_net_earned_income = trim(ga_elig_case_budg_net_earned_income)
		ga_elig_case_budg_unearned_income = trim(ga_elig_case_budg_unearned_income)
		ga_elig_case_budg_counted_school_income = trim(ga_elig_case_budg_counted_school_income)
		ga_elig_case_budg_total_deemed_income = trim(ga_elig_case_budg_total_deemed_income)
		ga_elig_case_budg_total_countable_income = trim(ga_elig_case_budg_total_countable_income)

		transmit 		'going to the next panel - GAB2

		EMReadScreen ga_elig_case_budg_payment_standard, 	10, 6, 34
		' EMReadScreen ga_elig_case_budg_total_countable_income, 10, 7, 34
		EMReadScreen ga_elig_case_budg_payment_subtotal, 	10, 8, 34
		EMReadScreen ga_elig_case_budg_prorated_from, 		5, 10, 15
		EMReadScreen ga_elig_case_budg_prorated_to, 		5, 10, 25
		EMReadScreen ga_elig_case_budg_grant_subtotal, 		10, 11, 34
		EMReadScreen ga_elig_case_budg_total_assets, 		10, 14, 34
		EMReadScreen ga_elig_case_budg_ga_exclusion, 		10, 15, 34
		EMReadScreen ga_elig_case_budg_countable_assets, 	10, 16, 34
		EMReadScreen ga_elig_case_budg_maximum_assets, 		10, 17, 34
		EMReadScreen ga_elig_case_budg_reason_ga_exclusion, 10, 18, 34

		EMReadScreen ga_elig_case_budg_pers_needs_payment_standard, 10, 6, 72
		' EMReadScreen ga_elig_case_budg_total_countable_income, 10, 7, 72
		EMReadScreen ga_elig_case_budg_pers_needs_payment_subtotal, 10, 8, 72
		EMReadScreen ga_elig_case_budg_pers_needs_prorated_from, 	5, 10, 58
		EMReadScreen ga_elig_case_budg_pers_needs_prorated_to, 		5, 10, 68
		EMReadScreen ga_elig_case_budg_pers_needs_grant_subtotal, 	10, 11, 72
		EMReadScreen ga_elig_case_budg_total_ga_grant_amount, 		10, 13, 72

		ga_elig_case_budg_payment_standard = trim(ga_elig_case_budg_payment_standard)
		ga_elig_case_budg_payment_subtotal = trim(ga_elig_case_budg_payment_subtotal)
		ga_elig_case_budg_prorated_from = trim(ga_elig_case_budg_prorated_from)
		ga_elig_case_budg_prorated_to = trim(ga_elig_case_budg_prorated_to)
		ga_elig_case_budg_grant_subtotal = trim(ga_elig_case_budg_grant_subtotal)
		ga_elig_case_budg_total_assets = trim(ga_elig_case_budg_total_assets)
		ga_elig_case_budg_ga_exclusion = trim(ga_elig_case_budg_ga_exclusion)
		ga_elig_case_budg_countable_assets = trim(ga_elig_case_budg_countable_assets)
		ga_elig_case_budg_maximum_assets = trim(ga_elig_case_budg_maximum_assets)
		ga_elig_case_budg_reason_ga_exclusion = trim(ga_elig_case_budg_reason_ga_exclusion)

		ga_elig_case_budg_pers_needs_payment_standard = trim(ga_elig_case_budg_pers_needs_payment_standard)
		ga_elig_case_budg_pers_needs_payment_subtotal = trim(ga_elig_case_budg_pers_needs_payment_subtotal)
		ga_elig_case_budg_pers_needs_prorated_from = trim(ga_elig_case_budg_pers_needs_prorated_from)
		ga_elig_case_budg_pers_needs_prorated_to = trim(ga_elig_case_budg_pers_needs_prorated_to)
		ga_elig_case_budg_pers_needs_grant_subtotal = trim(ga_elig_case_budg_pers_needs_grant_subtotal)
		ga_elig_case_budg_total_ga_grant_amount = trim(ga_elig_case_budg_total_ga_grant_amount)

		If ga_elig_case_budg_prorated_from <> "" Then
			ga_elig_case_budg_prorated_from = replace(ga_elig_case_budg_prorated_from, " ", "/")
			ga_elig_case_budg_prorated_from = ga_elig_case_budg_prorated_from & "/" & elig_footer_year
		End If
		If ga_elig_case_budg_prorated_to <> "" Then
			ga_elig_case_budg_prorated_to = replace(ga_elig_case_budg_prorated_to, " ", "/")
			ga_elig_case_budg_prorated_to = ga_elig_case_budg_prorated_to & "/" & elig_footer_year
		End If
		If ga_elig_case_budg_pers_needs_prorated_from <> "" Then
			ga_elig_case_budg_pers_needs_prorated_from = replace(ga_elig_case_budg_pers_needs_prorated_from, " ", "/")
			ga_elig_case_budg_pers_needs_prorated_from = ga_elig_case_budg_pers_needs_prorated_from & "/" & elig_footer_year
		End If
		If ga_elig_case_budg_pers_needs_prorated_to <> "" Then
			ga_elig_case_budg_pers_needs_prorated_to = replace(ga_elig_case_budg_pers_needs_prorated_to, " ", "/")
			ga_elig_case_budg_pers_needs_prorated_to = ga_elig_case_budg_pers_needs_prorated_to & "/" & elig_footer_year
		End If

		transmit 		'going to the next panel - GASM

		EMReadScreen ga_elig_summ_approved_date, 8, 3, 15
		EMReadScreen ga_elig_summ_process_date, 8, 2, 73
		EMReadScreen ga_elig_summ_date_last_approval, 8, 5, 32
		EMReadScreen ga_elig_summ_curr_prog_status, 12, 6, 32
		EMReadScreen ga_elig_summ_eligibility_result, 12, 7, 32
		EMReadScreen ga_elig_summ_hrf_reporting, 12, 8, 32
		EMReadScreen ga_elig_summ_source_of_info, 4, 9, 32
		EMReadScreen ga_elig_summ_eligibility_begin_date, 8, 10, 32
		EMReadScreen ga_elig_summ_eligiblity_review_date, 8, 11, 32
		EMReadScreen ga_elig_summ_budget_cycle, 5, 12, 32
		EMReadScreen ga_elig_summ_filing_unit_type_code, 1, 13, 32
		EMReadScreen ga_elig_summ_faci_unit_type_code, 1, 14, 32
		EMReadScreen ga_elig_summ_responsible_county, 2, 15, 32
		EMReadScreen ga_elig_summ_vendor_reason_code, 2, 16, 32

		EMReadScreen ga_elig_summ_total_assets, 10, 5, 71
		EMReadScreen ga_elig_summ_client_faci_obligation, 10, 6, 71
		EMReadScreen ga_elig_summ_standards, 10, 7, 71
		EMReadScreen ga_elig_summ_counted_income, 10, 8, 71
		EMReadScreen ga_elig_summ_monthly_grant, 10, 9, 71
		EMReadScreen ga_elig_summ_amount_to_be_paid, 10, 14, 71
		EMReadScreen ga_elig_summ_action_code, 1, 15, 53
		EMReadScreen ga_elig_summ_reason_code, 2, 16, 53

		EMReadScreen ga_elig_summ_worker_message, 80, 19, 1

		ga_elig_summ_curr_prog_status = trim(ga_elig_summ_curr_prog_status)
		ga_elig_summ_eligibility_result = trim(ga_elig_summ_eligibility_result)
		ga_elig_summ_hrf_reporting = trim(ga_elig_summ_hrf_reporting)

		If ga_elig_summ_filing_unit_type_code = "1" Then ga_elig_summ_filing_unit_type_info = "Single Adult"
		If ga_elig_summ_filing_unit_type_code = "2" Then ga_elig_summ_filing_unit_type_info = "Single Adult Lv W/ Parents"
		If ga_elig_summ_filing_unit_type_code = "3" Then ga_elig_summ_filing_unit_type_info = "Minor Child Outside Home"
		If ga_elig_summ_filing_unit_type_code = "6" Then ga_elig_summ_filing_unit_type_info = "Married Couple"
		If ga_elig_summ_filing_unit_type_code = "9" Then ga_elig_summ_filing_unit_type_info = "Family State Food Program"

		If ga_elig_summ_faci_unit_type_code = "5" Then ga_elig_summ_faci_unit_type_info = "Personal Needs"

		If ga_elig_summ_vendor_reason_code = "01" Then ga_elig_summ_vendor_reason_info = "Client Request"
		If ga_elig_summ_vendor_reason_code = "05" Then ga_elig_summ_vendor_reason_info = "Money Mismanagement"
		If ga_elig_summ_vendor_reason_code = "09" Then ga_elig_summ_vendor_reason_info = "Emergency"
		If ga_elig_summ_vendor_reason_code = "10" Then ga_elig_summ_vendor_reason_info = "Chemical Dependency"
		If ga_elig_summ_vendor_reason_code = "11" Then ga_elig_summ_vendor_reason_info = "No Residence"
		If ga_elig_summ_vendor_reason_code = "20" Then ga_elig_summ_vendor_reason_info = "Grant Diversion"


		ga_elig_summ_total_assets = trim(ga_elig_summ_total_assets)
		ga_elig_summ_client_faci_obligation = trim(ga_elig_summ_client_faci_obligation)
		ga_elig_summ_standards = trim(ga_elig_summ_standards)
		ga_elig_summ_counted_income = trim(ga_elig_summ_counted_income)
		ga_elig_summ_monthly_grant = trim(ga_elig_summ_monthly_grant)
		ga_elig_summ_amount_to_be_paid = trim(ga_elig_summ_amount_to_be_paid)

		If ga_elig_summ_action_code = "1" Then ga_elig_summ_action_info = "Open"
		If ga_elig_summ_action_code = "2" Then ga_elig_summ_action_info = "Suspend"
		If ga_elig_summ_action_code = "3" Then ga_elig_summ_action_info = "Unsuspend"
		If ga_elig_summ_action_code = "4" Then ga_elig_summ_action_info = "Review - Grant Change"
		If ga_elig_summ_action_code = "5" Then ga_elig_summ_action_info = "Close"
		If ga_elig_summ_action_code = "7" Then ga_elig_summ_action_info = "Grant Change - Chng Reported"
		If ga_elig_summ_action_code = "8" Then ga_elig_summ_action_info = "Review - No Grant Chng"
		If ga_elig_summ_action_code = "9" Then ga_elig_summ_action_info = "No Grant Chng - Chng Reported"
		If ga_elig_summ_action_code = "0" Then ga_elig_summ_action_info = "STAT Change - No Notice Rqrd"
		If ga_elig_summ_action_code = "C" Then ga_elig_summ_action_info = "Reinstate Closed Case"

		If ga_elig_summ_reason_code = "01" Then ga_elig_summ_reason_info = "Earned Income Increased"
		If ga_elig_summ_reason_code = "02" Then ga_elig_summ_reason_info = "Earned Income Decreased"
		If ga_elig_summ_reason_code = "03" Then ga_elig_summ_reason_info = "Unearned Income Increased"
		If ga_elig_summ_reason_code = "04" Then ga_elig_summ_reason_info = "Unearned Income Decreased"
		If ga_elig_summ_reason_code = "05" Then ga_elig_summ_reason_info = "Expenses/Deductions Increased"
		If ga_elig_summ_reason_code = "06" Then ga_elig_summ_reason_info = "Expenses/Deductions Decr"
		If ga_elig_summ_reason_code = "08" Then ga_elig_summ_reason_info = "No Proof Given"
		If ga_elig_summ_reason_code = "09" Then ga_elig_summ_reason_info = "Did Not Return Review Form"
		If ga_elig_summ_reason_code = "10" Then ga_elig_summ_reason_info = "Non Coop With GA Rules"
		If ga_elig_summ_reason_code = "12" Then ga_elig_summ_reason_info = "Must Apply For Other Benefit"
		If ga_elig_summ_reason_code = "14" Then ga_elig_summ_reason_info = "Not At Given Address"
		If ga_elig_summ_reason_code = "16" Then ga_elig_summ_reason_info = "Request Close"
		If ga_elig_summ_reason_code = "17" Then ga_elig_summ_reason_info = "Eligibility For Other Cash Program"
		If ga_elig_summ_reason_code = "18" Then ga_elig_summ_reason_info = "Non State Resident"
		If ga_elig_summ_reason_code = "19" Then ga_elig_summ_reason_info = "Client Died"
		If ga_elig_summ_reason_code = "20" Then ga_elig_summ_reason_info = "Household Member Died"
		If ga_elig_summ_reason_code = "22" Then ga_elig_summ_reason_info = "Excess Income"
		If ga_elig_summ_reason_code = "23" Then ga_elig_summ_reason_info = "Assets over the GA Limit"
		If ga_elig_summ_reason_code = "24" Then ga_elig_summ_reason_info = "Tranfer of Assets - No GA Eligiblity"
		If ga_elig_summ_reason_code = "27" Then ga_elig_summ_reason_info = "Fail To Sign Interim Assistance Agreemnt"
		If ga_elig_summ_reason_code = "28" Then ga_elig_summ_reason_info = "Program Reqquirements Have Been Met"
		If ga_elig_summ_reason_code = "30" Then ga_elig_summ_reason_info = "Household Size Change"
		If ga_elig_summ_reason_code = "31" Then ga_elig_summ_reason_info = "Review - No Change"
		If ga_elig_summ_reason_code = "32" Then ga_elig_summ_reason_info = "Begin Recoupment"
		If ga_elig_summ_reason_code = "33" Then ga_elig_summ_reason_info = "Change Recoupment"
		If ga_elig_summ_reason_code = "34" Then ga_elig_summ_reason_info = "End Recoupment"
		If ga_elig_summ_reason_code = "35" Then ga_elig_summ_reason_info = "New GA Basis Of Eligiblity"
		If ga_elig_summ_reason_code = "36" Then ga_elig_summ_reason_info = "Add/Change/Delete Vendor"
		If ga_elig_summ_reason_code = "39" Then ga_elig_summ_reason_info = "Person In/Out Facility"
		If ga_elig_summ_reason_code = "49" Then ga_elig_summ_reason_info = "No HRF"
		If ga_elig_summ_reason_code = "51" Then ga_elig_summ_reason_info = "Under Control Of Penal System"
		If ga_elig_summ_reason_code = "52" Then ga_elig_summ_reason_info = "Court Order Mitchell et al"
		If ga_elig_summ_reason_code = "54" Then ga_elig_summ_reason_info = "Not a GRH Facility"
		If ga_elig_summ_reason_code = "57" Then ga_elig_summ_reason_info = "Undocumented/Inelig Imig"
		If ga_elig_summ_reason_code = "59" Then ga_elig_summ_reason_info = "Imig-status not ver"
		If ga_elig_summ_reason_code = "61" Then ga_elig_summ_reason_info = "No GA Basis or Spouse w/none"
		If ga_elig_summ_reason_code = "62" Then ga_elig_summ_reason_info = "Lump Sum Payment"
		If ga_elig_summ_reason_code = "63" Then ga_elig_summ_reason_info = "Disqualified/Lump Sum"
		If ga_elig_summ_reason_code = "64" Then ga_elig_summ_reason_info = "Failed provide or apply SSN"
		If ga_elig_summ_reason_code = "66" Then ga_elig_summ_reason_info = "Eligible State wide MFIP"
		If ga_elig_summ_reason_code = "96" Then ga_elig_summ_reason_info = "April 2010 Legislation"
		If ga_elig_summ_reason_code = "97" Then ga_elig_summ_reason_info = "GRH Mass Change"
		If ga_elig_summ_reason_code = "98" Then ga_elig_summ_reason_info = "PNA Mass Change"

		ga_elig_summ_worker_message = trim(ga_elig_summ_worker_message)


		Call back_to_SELF
	end sub
end class



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
		EMReadScreen case_expedited_indicator, 9, 4, 3
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

		Call Back_to_SELF
	End sub

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

Dim DWP_ELIG_APPROVALS()
ReDim DWP_ELIG_APPROVALS(0)

Dim MFIP_ELIG_APPROVALS()
ReDim MFIP_ELIG_APPROVALS(0)

Dim MSA_ELIG_APPROVALS()
ReDim MSA_ELIG_APPROVALS(0)

Dim GA_ELIG_APPROVALS()
ReDim GA_ELIG_APPROVALS(0)

Dim CASH_DENIAL_APPROVALS()
ReDim CASH_DENIAL_APPROVALS(0)

Dim GRH_ELIG_APPROVALS()
ReDim GRH_ELIG_APPROVALS(0)

Dim IVE_ELIG_APPROVALS()
ReDim IVE_ELIG_APPROVALS(0)

Dim EMER_ELIG_APPROVALS()
ReDim EMER_ELIG_APPROVALS(0)

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

dwp_elig_months_count = 0
mfip_elig_months_count = 0
msa_elig_months_count = 0
ga_elig_months_count = 0
cash_deny_months_count = 0
grh_elig_months_count = 0
ive_elig_months_count = 0
emer_elig_months_count = 0
snap_elig_months_count = 0

For each footer_month in MONTHS_ARRAY
	' MsgBox footer_month
	Call convert_date_into_MAXIS_footer_month(footer_month, MAXIS_footer_month, MAXIS_footer_year)

	Call Navigate_to_MAXIS_screen("ELIG", "SUMM")

	EMReadScreen numb_DWP_versions, 		1, 7, 40
	EMReadScreen numb_MFIP_versions, 		1, 8, 40
	EMReadScreen numb_MSA_versions, 		1, 11, 40
	EMReadScreen numb_GA_versions, 			1, 12, 40
	' EMReadScreen numb_CASH_denial_versions, 1, 13, 40
	' EMReadScreen numb_GRH_versions, 		1, 14, 40
	' EMReadScreen numb_IVE_versions, 		1, 15, 40
	' EMReadScreen numb_EMER_versions, 		1, 16, 40
	EMReadScreen numb_SNAP_versions, 		1, 17, 40

	' MsgBox "numb_SNAP_versions - " & numb_SNAP_versions
	'TODO MAKE THIS READ THE DATE AND COMPARE TO TODAY

	If numb_DWP_versions <> " " Then
		ReDim Preserve DWP_ELIG_APPROVALS(dwp_elig_months_count)
		Set DWP_ELIG_APPROVALS(dwp_elig_months_count) = new dwp_eligibility_detail

		DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month = MAXIS_footer_month
		DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call DWP_ELIG_APPROVALS(dwp_elig_months_count).read_elig

		MsgBox "DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month - " & DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month & vbCr & "DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year - " & DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year & vbCr &_
		"DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_approved_date: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_approved_date & vbCr & "DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_summary_grant_amount: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_summary_grant_amount & vbCr &_
		"DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_eligibility_result: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_eligibility_result

		dwp_elig_months_count = dwp_elig_months_count + 1
	End If

	If numb_MFIP_versions <> " " Then
		ReDim Preserve MFIP_ELIG_APPROVALS(mfip_elig_months_count)
		Set MFIP_ELIG_APPROVALS(mfip_elig_months_count) = new mfip_eligibility_detail

		MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month = MAXIS_footer_month
		MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call MFIP_ELIG_APPROVALS(mfip_elig_months_count).read_elig

		' MsgBox "MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month - " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month & vbCr & "MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year - " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_approved_date: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_approved_date & vbCr & "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_grant_amount: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_grant_amount & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_cash_portion: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_cash_portion & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_food_portion: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_food_portion & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_housing_grant: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_housing_grant & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_eligibility_result: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_eligibility_result

		mfip_elig_months_count = mfip_elig_months_count + 1
	End If

	If numb_MSA_versions <> " " Then
		ReDim Preserve MSA_ELIG_APPROVALS(msa_elig_months_count)
		Set MSA_ELIG_APPROVALS(msa_elig_months_count) = new msa_eligibility_detail

		MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month = MAXIS_footer_month
		MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call MSA_ELIG_APPROVALS(msa_elig_months_count).read_elig

		' MsgBox "MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month - " & MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month & vbCr & "MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year - " & MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year & vbCr &_
		' "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_approved_date: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_approved_date & vbCr & "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_grant: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_grant & vbCr &_
		' "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_eligibility_result: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_eligibility_result

		msa_elig_months_count = msa_elig_months_count + 1
	End If

	If numb_GA_versions <> " " Then
		ReDim Preserve GA_ELIG_APPROVALS(ga_elig_months_count)
		Set GA_ELIG_APPROVALS(ga_elig_months_count) = new ga_eligibility_detail

		GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month = MAXIS_footer_month
		GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call GA_ELIG_APPROVALS(ga_elig_months_count).read_elig

		' MsgBox "GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month - " & GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month & vbCr & "GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year - " & GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year & vbCr &_
		' "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_approved_date: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_approved_date & vbCr & "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_monthly_grant: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_monthly_grant & vbCr &_
		' "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_eligibility_result: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_eligibility_result

		ga_elig_months_count = ga_elig_months_count + 1
	End If

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

' For approval_month = 0 to UBound(SNAP_ELIG_APPROVALS)
' 	For snap_memb = 0 to UBound(SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs)
' 		MsgBox SNAP_ELIG_APPROVALS(approval_month).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval_month).elig_footer_year & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs(snap_memb) & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_membs_eligibility(snap_memb)
' 	Next
' Next


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
