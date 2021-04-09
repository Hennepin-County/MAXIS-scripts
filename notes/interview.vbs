'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
' run_locally = TRUE
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
call changelog_update("04/00/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================

Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(0)



'===========================================================================================================================

'FUNCTIONS =================================================================================================================


class mx_hh_member

	public access_denied
	public selected
	'stuff about the members
	public first_name
	public last_name
	public mid_initial
	public other_names
	public date_of_birth
	public age
	public ref_number
	public ssn
	public ssn_verif
	public birthdate_verif
	public gender
	public id_verif
	public rel_to_applcnt
	public race
	public race_a_checkbox
	public race_b_checkbox
	public race_n_checkbox
	public race_p_checkbox
	public race_w_checkbox
	public snap_minor
	public cash_minor
	public written_lang
	public spoken_lang
	public interpreter
	public alias_yn
	public ethnicity_yn

	public marital_status
	public spouse_ref
	public spouse_name
	public last_grade_completed
	public citizen
	public other_st_FS_end_date
	public in_mn_12_mo
	public residence_verif
	public mn_entry_date
	public former_state
	public intend_to_reside_in_mn

	public parents_in_home
	public parents_in_home_notes
	public parent_one_name
	public parent_one_type
	public parent_one_verif
	public parent_one_in_home
	public parent_two_name
	public parent_two_type
	public parent_two_verif
	public parent_two_in_home

	public pare_exists
	public pare_child_ref_nbr
	public pare_child_name
	public pare_child_member_index
	public pare_relationship_type
	public pare_verification

	public remo_exists
	public left_hh_date
	public left_hh_reason
	public left_hh_expected_return_date
	public left_hh_expected_return_verif
	public left_hh_actual_return_date
	public left_hh_HC_temp_out_of_state
	public left_hh_date_reported
	public left_hh_12_months_or_more

	public adme_exists
	public adme_arrival_date
	public adme_cash_date
	public adme_emer_date
	public adme_snap_date
	public adme_within_12_months

	public imig_exists
	public imig_status
	public us_entry_date
	public imig_status_date
	public imig_status_verif
	public lpr_adj_from
	public nationality
	public nationality_detail
	public alien_id_nbr
	public imig_active_doc
	public imig_recvd_doc
	public imig_q_2_required
	public imig_q_4_required
	public imig_q_5_required
	public imig_clt_current_doc
	public imig_doc_on_file
	public imig_save_completed
	public imig_prev_status

	public new_imig_status
	public new_us_entry_date
	public new_imig_status_date
	public new_imig_status_verif
	public new_lpr_adj_from
	public new_nationality
	public new_nationality_detail
	public new_imig_active_doc
	public new_imig_recvd_doc
	public new_imig_clt_current_doc
	public new_imig_doc_on_file
	public new_imig_save_completed
	public new_imig_prev_status
	public new_spon_name
	public new_spon_street
	public new_spon_city
	public new_spon_state
	public new_spon_zip
	public new_spon_phone
	public new_spon_gross_income
	public new_spon_income_freq
	public new_spon_spouse_name
	public new_spon_spouse_income
	public new_spon_spouse_income_freq

	' public ans_us_entry_date
	' public ans_nationality
	' public ans_nationality_detail
	' public ans_imig_status
	' public ans_imig_prev_status
	' public ans_imig_status_date
	' public ans_imig_clt_current_doc
	' public ans_imig_doc_on_file
	' public ans_imig_save_completed
	' public ans_clt_has_sponsor
	' public ans_spon_name
	' public ans_live_with_spon
	' public ans_spon_street
	' public ans_spon_city
	' public ans_spon_state
	' public ans_spon_zip
	' public ans_spon_phone
	' public ans_spon_gross_income
	' public ans_spon_income_freq
	' public ans_spon_married_yn
	' public ans_spon_children_yn
	' public ans_spon_spouse_name
	' public ans_spon_spouse_income
	' public ans_spon_spouse_income_freq
	' public ans_spon_numb_children
	' public ans_spon_hh_notes

	public spon_exists
	public clt_has_sponsor
	' public ask_about_spon
	public spon_type
	public spon_verif
	public spon_name
	public spon_street
	public spon_city
	public spon_state
	public spon_zip
	public spon_phone
	public spon_cash_retro_income
	public spon_cash_prosp_income
	public spon_cash_assets
	public spon_snap_retro_income
	public spon_snap_prosp_income
	public spon_snap_assets
	public spon_spouse
	public spon_hh_size
	public spon_numb_children
	public spon_possible_indigent_exemption
	public spon_gross_income
	public spon_spouse_income
	public live_with_spon
	public spon_income_freq
	public spon_spouse_income_freq
	public spon_married_yn
	public spon_children_yn
	public spon_hh_notes
	public spon_spouse_name

	public disa_exists
	public disa_begin_date
	public disa_end_date
	public disa_cert_begin_date
	public disa_cert_end_date
	public cash_disa_status
	public cash_disa_verif
	public fs_disa_status
	public fs_disa_verif
	public hc_disa_status
	public hc_disa_verif
	public disa_waiver
	public disa_1619
	public disa_detail
	public mof_file
	public mof_detail
	public mof_end_date
	public iaa_file
	public iaa_received_date
	public iaa_complete
	public disa_review

	public fs_pwe
	public wreg_exists

	public schl_exists
	public school_status
	public school_grade
	public school_name
	public school_verif
	public school_type
	public school_district
	public kinder_start_date
	public grad_date
	public grad_date_verif
	public school_funding
	public school_elig_status
	public higher_ed

	public stin_exists
	public total_stin
	public stin_type_array
	public stin_amount_array
	public stin_avail_date_array
	public stin_months_cov_array
	public stin_verif_array

	public stec_exists
	public total_stec
	public stec_type_array
	public stec_amount_array
	public stec_months_cov_array
	public stec_verif_array
	public stec_earmarked_amount_array
	public stec_earmarked_months_cov_array

	public shel_exists
	public shel_summary
	public shel_hud_subsidy_yn
	public shel_shared_yn
	public shel_paid_to
	public shel_retro_rent_amount
	public shel_retro_rent_verif
	public shel_retro_lot_rent_amount
	public shel_retro_lot_rent_verif
	public shel_retro_mortgage_amount
	public shel_retro_mortgage_verif
	public shel_retro_insurance_amount
	public shel_retro_insurance_verif
	public shel_retro_taxes_amount
	public shel_retro_taxes_verif
	public shel_retro_room_amount
	public shel_retro_room_verif
	public shel_retro_garage_amount
	public shel_retro_garage_verif
	public shel_retro_subsidy_amount
	public shel_retro_subsidy_verif

	public shel_prosp_rent_amount
	public shel_prosp_rent_verif
	public shel_prosp_lot_rent_amount
	public shel_prosp_lot_rent_verif
	public shel_prosp_mortgage_amount
	public shel_prosp_mortgage_verif
	public shel_prosp_insurance_amount
	public shel_prosp_insurance_verif
	public shel_prosp_taxes_amount
	public shel_prosp_taxes_verif
	public shel_prosp_room_amount
	public shel_prosp_room_verif
	public shel_prosp_garage_amount
	public shel_prosp_garage_verif
	public shel_prosp_subsidy_amount
	public shel_prosp_subsidy_verif

	public coex_exists
	public coex_support_verif
	public coex_support_retro_amount
	public coex_support_prosp_amount
	public coex_support_hc_est_amount
	public coex_alimony_verif
	public coex_alimony_retro_amount
	public coex_alimony_prosp_amount
	public coex_alimony_hc_est_amount
	public coex_tax_dep_verif
	public coex_tax_dep_retro_amount
	public coex_tax_dep_prosp_amount
	public coex_tax_dep_hc_est_amount
	public coex_other_verif
	public coex_other_retro_amount
	public coex_other_prosp_amount
	public coex_other_hc_est_amount
	public coex_total_retro_amount
	public coex_total_prosp_amount
	public coex_total_hc_est_amount
	public coex_change_in_financial_circumstances

	public stwk_exists
	public stwk_employer
	public stwk_work_stop_date
	public stwk_income_stop_date
	public stwk_verification
	public stwk_refused_employment
	public stwk_vol_quit
	public stwk_refused_employment_date
	public stwk_cash_good_cause_yn
	public stwk_grh_good_cause_yn
	public stwk_snap_good_cause_yn
	public stwk_snap_pwe
	public stwk_ma_epd_extension
	public stwk_summary

	public fmed_exists
	public fmed_miles
	public fmed_rate
	public fmed_milage_expense
	public fmed_page()
	public fmed_row()
	public fmed_type()
	public fmed_verif()
	public fmed_ref()
	public fmed_catgry()
	public fmed_begin()
	public fmed_end()
	public fmed_expense()
	public fmed_notes()

	public pded_exists
	public pded_guardian_fee
	public pded_rep_payee_fee
	public pded_shel_spec_need

	public diet_exists
	public diet_mf_type_one
	public diet_mf_verif_one
	public diet_mf_type_two
	public diet_mf_verif_two
	public diet_msa_type_one
	public diet_msa_verif_one
	public diet_msa_type_two
	public diet_msa_verif_two
	public diet_msa_type_three
	public diet_msa_verif_three
	public diet_msa_type_four
	public diet_msa_verif_four
	public diet_msa_type_five
	public diet_msa_verif_five
	public diet_msa_type_six
	public diet_msa_verif_six
	public diet_msa_type_seven
	public diet_msa_verif_seven
	public diet_msa_type_eight
	public diet_msa_verif_eight

	public checkbox_one
	public checkbox_two
	public checkbox_three
	public checkbox_four

	public detail_one
	public detail_two
	public detail_three
	public detail_four

	public button_one
	public button_two
	public button_three
	public button_four

	public clt_has_cs_income
	public clt_cs_counted
	public cs_paid_to
	public clt_has_ss_income
	public clt_has_BUSI
	public clt_has_JOBS

	public snap_req_checkbox
	public cash_req_checkbox
	public emer_req_checkbox
	public grh_req_checkbox
	public hc_req_checkbox
	public none_req_checkbox
	public client_verification
	public client_verification_details
	public client_notes

	public property get full_name
		full_name = first_name & " " & last_name
	end property

	Public sub define_the_member()

		pare_child_ref_nbr = array("")
		pare_child_name = array("")
		pare_child_member_index = array("")
		pare_relationship_type = array("")
		pare_verification = array("")

		Call navigate_to_MAXIS_screen("STAT", "MEMB")		'===============================================================================================
		EMWriteScreen ref_number, 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			last_name = "UNABLE TO FIND"
			first_name = "Access Denied"
			mid_initial = ""
			access_denied = TRUE
		Else
			access_denied = FALSE
			EMReadscreen last_name, 25, 6, 30
			EMReadscreen first_name, 12, 6, 63
			EMReadscreen mid_initial, 1, 6, 79
			EMReadScreen age, 3, 8, 76

			EMReadScreen date_of_birth, 10, 8, 42
			EMReadScreen ssn, 11, 7, 42
			EMReadScreen ssn_verif, 1, 7, 68
			EMReadScreen birthdate_verif, 2, 8, 68
			EMReadScreen gender, 1, 9, 42
			EMReadScreen race, 30, 17, 42
			EMReadScreen spoken_lang, 20, 12, 42
			EMReadScreen written_lang, 29, 13, 42
			EMReadScreen interpreter, 1, 14, 68
			EMReadScreen alias_yn, 1, 15, 42
			EMReadScreen ethnicity_yn, 1, 16, 68

			age = trim(age)
			If age = "" Then age = 0
			age = age * 1
			last_name = trim(replace(last_name, "_", ""))
			first_name = trim(replace(first_name, "_", ""))
			mid_initial = replace(mid_initial, "_", "")
			EMReadScreen id_verif, 2, 9, 68

			EMReadScreen rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
			If rel_to_applcnt = "01" Then rel_to_applcnt = "Self"
			If rel_to_applcnt = "02" Then rel_to_applcnt = "Spouse"
			If rel_to_applcnt = "03" Then rel_to_applcnt = "Child"
			If rel_to_applcnt = "04" Then rel_to_applcnt = "Parent"
			If rel_to_applcnt = "05" Then rel_to_applcnt = "Sibling"
			If rel_to_applcnt = "06" Then rel_to_applcnt = "Step Sibling"
			If rel_to_applcnt = "08" Then rel_to_applcnt = "Step Child"
			If rel_to_applcnt = "09" Then rel_to_applcnt = "Step Parent"
			If rel_to_applcnt = "10" Then rel_to_applcnt = "Aunt"
			If rel_to_applcnt = "11" Then rel_to_applcnt = "Uncle"
			If rel_to_applcnt = "12" Then rel_to_applcnt = "Niece"
			If rel_to_applcnt = "13" Then rel_to_applcnt = "Nephew"
			If rel_to_applcnt = "14" Then rel_to_applcnt = "Cousin"
			If rel_to_applcnt = "15" Then rel_to_applcnt = "Grandparent"
			If rel_to_applcnt = "16" Then rel_to_applcnt = "Grandchild"
			If rel_to_applcnt = "17" Then rel_to_applcnt = "Other Relative"
			If rel_to_applcnt = "18" Then rel_to_applcnt = "Legal Guardian"
			If rel_to_applcnt = "24" Then rel_to_applcnt = "Not Related"
			If rel_to_applcnt = "25" Then rel_to_applcnt = "Live-in Attendant"
			If rel_to_applcnt = "27" Then rel_to_applcnt = "Unknown"

			If id_verif = "BC" Then id_verif = "BC - Birth Certificate"
			If id_verif = "RE" Then id_verif = "RE - Religious Record"
			If id_verif = "DL" Then id_verif = "DL - Drivers License/ST ID"
			If id_verif = "DV" Then id_verif = "DV - Divorce Decree"
			If id_verif = "AL" Then id_verif = "AL - Alien Card"
			If id_verif = "AD" Then id_verif = "AD - Arrival//Depart"
			If id_verif = "DR" Then id_verif = "DR - Doctor Stmt"
			If id_verif = "PV" Then id_verif = "PV - Passport/Visa"
			If id_verif = "OT" Then id_verif = "OT - Other Document"
			If id_verif = "NO" Then id_verif = "NO - No Veer Prvd"

			If age > 18 then
				cash_minor = FALSE
			Else
				cash_minor = TRUE
			End If
			If age > 21 then
				snap_minor = FALSE
			Else
				snap_minor = TRUE
			End If

			date_of_birth = replace(date_of_birth, " ", "/")
			If birthdate_verif = "BC" Then birthdate_verif = "BC - Birth Certificate"
			If birthdate_verif = "RE" Then birthdate_verif = "RE - Religious Record"
			If birthdate_verif = "DL" Then birthdate_verif = "DL - Drivers License/State ID"
			If birthdate_verif = "DV" Then birthdate_verif = "DV - Divorce Decree"
			If birthdate_verif = "AL" Then birthdate_verif = "AL - Alien Card"
			If birthdate_verif = "DR" Then birthdate_verif = "DR - Doctor Statement"
			If birthdate_verif = "OT" Then birthdate_verif = "OT - Other Document"
			If birthdate_verif = "PV" Then birthdate_verif = "PV - Passport/Visa"
			If birthdate_verif = "NO" Then birthdate_verif = "NO - No Verif Provided"

			ssn = replace(ssn, " ", "-")
			if ssn = "___-__-____" Then ssn = ""
			If ssn_verif = "A" THen ssn_verif = "A - SSN Applied For"
			If ssn_verif = "P" THen ssn_verif = "P - SSN Provided, verif Pending"
			If ssn_verif = "N" THen ssn_verif = "N - SSN Not Provided"
			If ssn_verif = "V" THen ssn_verif = "V - SSN Verified via Interface"

			If gender = "M" Then gender = "Male"
			If gender = "F" Then gender = "Female"

			race = trim(race)

			spoken_lang = replace(replace(spoken_lang, "_", ""), "  ", " - ")
			written_lang = trim(replace(replace(replace(written_lang, "_", ""), "  ", " - "), "(HRF)", ""))

			clt_has_cs_income = FALSE
			clt_has_ss_income = FALSE
			clt_has_BUSI = FALSE
			clt_has_JOBS = FALSE
		End If

		If access_denied = FALSE Then
			Call navigate_to_MAXIS_screen("STAT", "MEMI")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMReadScreen marital_status, 1, 7, 40
			EMReadScreen spouse_ref, 2, 9, 49
			EMReadScreen spouse_name, 40, 9, 52
			EMReadScreen last_grade_completed, 2, 10, 49
			EMReadScreen citizen, 1, 11, 49
			EMReadScreen other_st_FS_end_date, 8, 13, 49
			EMReadScreen in_mn_12_mo, 1, 14, 49
			EMReadScreen residence_verif, 1, 14, 78
			EMReadScreen mn_entry_date, 8, 15, 49
			EMReadScreen former_state, 2, 15, 78

			If marital_status = "N" Then marital_status = "N - Never Married"
			If marital_status = "M" Then marital_status = "M - Married Living with Spouse"
			If marital_status = "S" Then marital_status = "S - Married Living Apart"
			If marital_status = "L" Then marital_status = "L - Legally Seperated"
			If marital_status = "D" Then marital_status = "D - Divorced"
			If marital_status = "W" Then marital_status = "W - Widowed"
			If spouse_ref = "__" Then spouse_ref = ""
			spouse_name = trim(spouse_name)

			If last_grade_completed = "00" Then last_grade_completed = "Not Attended or Pre-Grade 1 - 00"
			If last_grade_completed = "12" Then last_grade_completed = "High School Diploma or GED - 12"
			If last_grade_completed = "13" Then last_grade_completed = "Some Post Sec Education - 13"
			If last_grade_completed = "14" Then last_grade_completed = "High School Plus Certiificate - 14"
			If last_grade_completed = "15" Then last_grade_completed = "Four Year Degree - 15"
			If last_grade_completed = "16" Then last_grade_completed = "Grad Degree - 16"
			If len(last_grade_completed) = 2 Then last_grade_completed = "Grade " & last_grade_completed
			If citizen = "Y" Then citizen = "Yes"
			If citizen = "N" Then citizen = "No"

			other_st_FS_end_date = replace(other_st_FS_end_date, " ", "/")
			If other_st_FS_end_date = "__/__/__" Then other_st_FS_end_date = ""
			If in_mn_12_mo = "Y" Then in_mn_12_mo = "Yes"
			If in_mn_12_mo = "N" Then in_mn_12_mo = "No"
			If residence_verif = "1" Then residence_verif = "1 - Rent Receipt"
			If residence_verif = "2" Then residence_verif = "2 - Landlord's Statement"
			If residence_verif = "3" Then residence_verif = "3 - Utility Bill"
			If residence_verif = "4" Then residence_verif = "4 - Other"
			If residence_verif = "N" Then residence_verif = "N - Verif Not Provided"
			mn_entry_date = replace(mn_entry_date, " ", "/")
			If mn_entry_date = "__/__/__" Then mn_entry_date = ""
			If former_state = "__" Then former_state = ""


			Call navigate_to_MAXIS_screen("STAT", "IMIG")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen imig_version, 1, 2, 73
			If imig_version = "0" Then imig_exists = FALSE
			If imig_version = "1" Then imig_exists = TRUE

			If imig_exists = TRUE Then
				EMReadScreen imig_status_code, 2, 6, 45
				EMReadScreen imig_status_desc, 32, 6, 48
				EMReadScreen us_entry_date, 10, 7, 45
				EMReadScreen imig_status_date, 10, 7, 71
				EMReadScreen imig_status_verif, 2, 8, 45
				EMReadScreen lpr_adj_from, 40, 9, 45
				EMReadScreen nationality, 2, 10, 45
				EMReadScreen alien_id_nbr, 10, 10, 71

				imig_status_desc = trim(imig_status_desc)
				imig_status = imig_status_code & " - " & imig_status_desc
				us_entry_date = replace(us_entry_date, " ", "/")
				imig_status_date = replace(imig_status_date, " ", "/")

				If imig_status_verif = "S1" Then imig_status_verif = "S1 - SAVE Primary"
				If imig_status_verif = "S2" Then imig_status_verif = "S2 - SAVE Secondary"
				If imig_status_verif = "AL" Then imig_status_verif = "AL - Alien Card"
				If imig_status_verif = "PV" Then imig_status_verif = "PV - Passport/Visa"
				If imig_status_verif = "RE" Then imig_status_verif = "RE - Re-Entry Permit"
				If imig_status_verif = "IM" Then imig_status_verif = "IN - INS Correspondence"
				If imig_status_verif = "OT" Then imig_status_verif = "OT - Other Document"
				If imig_status_verif = "NO" Then imig_status_verif = "NO - No Verif Provided"

				lpr_adj_from = trim(lpr_adj_from)

				If nationality = "AA" Then nationality = "AA - Amerasian"
				If nationality = "EH" Then nationality = "EH - Ethnic Chinese"
				If nationality = "EL" Then nationality = "EL - Ethnic Lao"
				If nationality = "HG" Then nationality = "HG - Hmong"
				If nationality = "KD" Then nationality = "KD - Kurd"
				If nationality = "SJ" Then nationality = "SJ - Soviet Jew"
				If nationality = "TT" Then nationality = "TT - Tinh"
				If nationality = "AF" Then nationality = "AF - Afghanistan"
				If nationality = "BK" Then nationality = "BK - Bosnia"
				If nationality = "CB" Then nationality = "CB - Cambodia"
				If nationality = "CH" Then nationality = "CH - China, Mainland"
				If nationality = "CU" Then nationality = "CU - Cuba"
				If nationality = "ES" Then nationality = "ES - El Salvador"
				If nationality = "ER" Then nationality = "ER - Eritrea"
				If nationality = "ET" Then nationality = "ET - Ethiopia"
				If nationality = "GT" Then nationality = "GT - Guatemala"
				If nationality = "HA" Then nationality = "HA - Haiti"
				If nationality = "HO" Then nationality = "HO - Honduras"
				If nationality = "IR" Then nationality = "IR - Iran"
				If nationality = "IZ" Then nationality = "IZ - Iraq"
				If nationality = "LI" Then nationality = "LI - Liberia"
				If nationality = "MC" Then nationality = "MC - Micronesia"
				If nationality = "MI" Then nationality = "MI - Marshall Islands"
				If nationality = "MX" Then nationality = "MX - Mexico"
				If nationality = "WA" Then nationality = "WA - Namibia (SW Africa)"
				If nationality = "PK" Then nationality = "PK - Pakistan"
				If nationality = "RP" Then nationality = "RP - Philippines"
				If nationality = "PL" Then nationality = "PL - Poland"
				If nationality = "RO" Then nationality = "RO - Romania"
				If nationality = "RS" Then nationality = "RS - Russia"
				If nationality = "SO" Then nationality = "SO - Somalia"
				If nationality = "SF" Then nationality = "SF - South Africa"
				If nationality = "TH" Then nationality = "TH - Thailand"
				If nationality = "VM" Then nationality = "VM - Vietnam"
				If nationality = "OT" Then nationality = "OT - All Others"

				imig_q_2_required = TRUE
				imig_q_4_required = TRUE
				imig_q_5_required = TRUE

			End If

			Call navigate_to_MAXIS_screen("STAT", "SPON")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen spon_version, 1, 2, 73
			If spon_version = "0" Then spon_exists = FALSE
			If spon_version = "1" Then spon_exists = TRUE
			clt_has_sponsor = "No"

			If spon_exists = TRUE Then
				clt_has_sponsor = "Yes"
				' new_spon_name			=
				' new_spon_street			=
				' new_spon_city			=
				' new_spon_state			=
				' new_spon_zip			=
				' new_spon_phone			=
				' new_spon_gross_income	=
				' new_spon_income_freq	=
				' new_spon_spouse_name	=
				' new_spon_spouse_income	=
				' new_spon_spouse_income_freq =


			End If
			' public spon_exists

			Call navigate_to_MAXIS_screen("STAT", "REMO")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen remo_version, 1, 2, 73
			If remo_version = "0" Then remo_exists = FALSE
			If remo_version = "1" Then remo_exists = TRUE

			If remo_exists = TRUE Then
				EMReadScreen left_hh_date, 8, 8, 53
				EMReadScreen left_hh_reason, 2, 8, 71
				EMReadScreen left_hh_expected_return_date, 8, 13, 53
				EMReadScreen left_hh_expected_return_verif, 2, 13, 71
				EMReadScreen left_hh_actual_return_date, 8, 14, 53
				EMReadScreen left_hh_HC_temp_out_of_state, 1, 16, 64
				EMReadScreen left_hh_date_reported, 8, 17, 64

				left_hh_date = replace(left_hh_date, " ", "/")
				If left_hh_date = "__/__/__" Then left_hh_date = ""

				If left_hh_reason = "01" Then left_hh_reason = "01 - Death"
				If left_hh_reason = "02" Then left_hh_reason = "02 - Moved out of Household"
				If left_hh_reason = "03" Then left_hh_reason = "03 - Institutional Placement"
				If left_hh_reason = "04" Then left_hh_reason = "04 - IV-E Foster Care Placement"
				If left_hh_reason = "05" Then left_hh_reason = "05 - Non IV-E Foster Care Placement"
				If left_hh_reason = "06" Then left_hh_reason = "06 - Illness"
				If left_hh_reason = "07" Then left_hh_reason = "07 - Vacation or Visit"
				If left_hh_reason = "08" Then left_hh_reason = "08 - Runaway"
				If left_hh_reason = "09" Then left_hh_reason = "09 - Away for Education"
				If left_hh_reason = "10" Then left_hh_reason = "10 - Relative Ill/Deceased"
				If left_hh_reason = "11" Then left_hh_reason = "11 - Training of Employment Search"
				If left_hh_reason = "12" Then left_hh_reason = "12 - Incarceration"
				If left_hh_reason = "13" Then left_hh_reason = "13 - Other Allowed Return before 10th"
				If left_hh_reason = "14" Then left_hh_reason = "14 - Non-Allowed Absent Cash"
				If left_hh_reason = "15" Then left_hh_reason = "15 - Military Service"
				If left_hh_reason = "16" Then left_hh_reason = "16 - Other"
				If left_hh_reason = "__" Then left_hh_reason = ""

				left_hh_expected_return_date = replace(left_hh_expected_return_date, " ", "/")
				If left_hh_expected_return_date = "__/__/__" Then left_hh_expected_return_date = ""

				If left_hh_expected_return_verif = "01" Then left_hh_expected_return_verif = "01 - Social Worker Statement"
				If left_hh_expected_return_verif = "02" Then left_hh_expected_return_verif = "02 - Court Papers"
				If left_hh_expected_return_verif = "03" Then left_hh_expected_return_verif = "03 - Doctor Statement"
				If left_hh_expected_return_verif = "04" Then left_hh_expected_return_verif = "04 - Other Document"
				If left_hh_expected_return_verif = "__" Then left_hh_expected_return_verif = ""

				left_hh_actual_return_date = replace(left_hh_actual_return_date, " ", "/")
				If left_hh_actual_return_date = "__/__/__" Then left_hh_actual_return_date = ""

				If left_hh_HC_temp_out_of_state = "_" Then left_hh_HC_temp_out_of_state = ""

				left_hh_date_reported = replace(left_hh_date_reported, " ", "/")
				If left_hh_date_reported = "__/__/__" Then left_hh_date_reported = ""

				If IsDate(left_hh_date) = TRUE Then
					If DateDiff("m", left_hh_date, date) >= 12 Then
						left_hh_12_months_or_more = TRUE
					Else
						left_hh_12_months_or_more = FALSE
					End If
				End If
			End If

			Call navigate_to_MAXIS_screen("STAT", "ADME")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen adme_version, 1, 2, 73
			If adme_version = "0" Then adme_exists = FALSE
			If adme_version = "1" Then adme_exists = TRUE

			If adme_exists = TRUE Then
				EMReadScreen adme_arrival_date, 8, 7, 38
				EMReadScreen adme_cash_date, 8, 12, 38
				EMReadScreen adme_emer_date, 8, 14, 38
				EMReadScreen adme_snap_date, 8, 16, 38

				adme_arrival_date = trim(adme_arrival_date)
				If adme_arrival_date = "////////" Then adme_arrival_date = ""

				adme_cash_date = replace(adme_cash_date, " ", "/")
				If adme_cash_date = "__/__/__" Then adme_cash_date = ""

				adme_emer_date = replace(adme_emer_date, " ", "/")
				If adme_emer_date = "__/__/__" Then adme_emer_date = ""

				adme_snap_date = replace(adme_snap_date, " ", "/")
				If adme_snap_date = "__/__/__" Then adme_snap_date = ""

				adme_within_12_months = FALSE
				If IsDate(adme_arrival_date) = TRUE Then
					If DateDiff("m", adme_arrival_date, date) < 13 Then adme_within_12_months = TRUE
				End If
			End If


			Call navigate_to_MAXIS_screen("STAT", "COEX")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen coex_version, 1, 2, 73
			If coex_version = "0" Then coex_exists = FALSE
			If coex_version = "1" Then coex_exists = TRUE

			If coex_exists = TRUE Then
				EMReadScreen coex_support_verif, 1, 10, 36
				EMReadScreen coex_support_retro_amount, 8, 10, 45
				EMReadScreen coex_support_prosp_amount, 8, 10, 63

				EMReadScreen coex_alimony_verif, 1, 11, 36
				EMReadScreen coex_alimony_retro_amount, 8, 11, 45
				EMReadScreen coex_alimony_prosp_amount, 8, 11, 63

				EMReadScreen coex_tax_dep_verif, 1, 12, 36
				EMReadScreen coex_tax_dep_retro_amount, 8, 12, 45
				EMReadScreen coex_tax_dep_prosp_amount, 8, 12, 63

				EMReadScreen coex_other_verif, 1, 13, 36
				EMReadScreen coex_other_retro_amount, 8, 13, 45
				EMReadScreen coex_other_prosp_amount, 8, 13, 63

				EMReadScreen coex_total_retro_amount, 8, 15, 45
				EMReadScreen coex_total_prosp_amount, 8, 15, 63

				EMReadScreen coex_change_in_financial_circumstances, 1, 17, 61

				EMWriteScreen "X", 18, 44
				transmit

				EMReadScreen coex_support_hc_est_amount, 8, 6, 38
				EMReadScreen coex_alimony_hc_est_amount, 8, 7, 38
				EMReadScreen coex_tax_dep_hc_est_amount, 8, 8, 38
				EMReadScreen coex_other_hc_est_amount, 8, 9, 38
				EMReadScreen coex_total_hc_est_amount, 8, 11, 38

				PF3

			End If

			Call navigate_to_MAXIS_screen("STAT", "DISA")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen disa_version, 1, 2, 73
			If disa_version = "0" Then disa_exists = FALSE
			If disa_version = "1" Then disa_exists = TRUE

			If disa_exists = TRUE Then
				EMReadScreen disa_begin_date, 10, 6, 47
				EMReadScreen disa_end_date, 10, 6, 69
				EMReadScreen disa_cert_begin_date, 10, 7, 47
				EMReadScreen disa_cert_end_date, 10, 7, 69
				EMReadScreen cash_disa_status, 2, 11, 59
				EMReadScreen cash_disa_verif, 1, 11, 69
				EMReadScreen fs_disa_status, 2, 12, 59
				EMReadScreen fs_disa_verif, 1, 12, 69
				EMReadScreen hc_disa_status, 2, 13, 59
				EMReadScreen hc_disa_verif, 1, 13, 69
				EMReadScreen disa_waiver, 1, 14, 59
				EMReadScreen disa_1619, 1, 16, 59

				disa_begin_date = replace(disa_begin_date, " ", "/")
				If disa_begin_date = "__/__/____" Then disa_begin_date = ""
				disa_end_date = replace(disa_end_date, " ", "/")
				If disa_end_date = "__/__/____" Then disa_end_date = ""
				disa_cert_begin_date = replace(disa_cert_begin_date, " ", "/")
				If disa_cert_begin_date = "__/__/____" Then disa_cert_begin_date = ""
				disa_cert_end_date = replace(disa_cert_end_date, " ", "/")
				If disa_cert_end_date = "__/__/____" Then disa_cert_end_date = ""

				If hc_disa_verif = "1" OR fs_disa_verif = "1" OR cash_disa_status = "1" Then disa_detail = "DISA based on Doctor's Statement"
				If hc_disa_verif = "2" OR fs_disa_verif = "2" OR cash_disa_status = "2" Then disa_detail = "SMRT Certified Disability"
				If hc_disa_verif = "3" OR fs_disa_verif = "3" OR cash_disa_status = "3" Then disa_detail = "SSA Certified Disability"
				If cash_disa_status = "7" Then disa_detail = "Disability based on Professional Statement of Need"

				If cash_disa_status = "01" Then cash_disa_status = "01 - RSDI Only Disability"
				If cash_disa_status = "02" Then cash_disa_status = "02 - RSDI Only Blindness"
				If cash_disa_status = "03" Then cash_disa_status = "03 - SSI, SSI/RSDI Disability"
				If cash_disa_status = "04" Then cash_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If cash_disa_status = "06" Then cash_disa_status = "06 - SMRT/SSA Pend"
				If cash_disa_status = "08" Then cash_disa_status = "08 - SMRT Certified Blindness"
				If cash_disa_status = "09" Then cash_disa_status = "09 - Ill/Incapacity"
				If cash_disa_status = "10" Then cash_disa_status = "10 - SMRT Certified Disability"
				If cash_disa_status = "__" Then cash_disa_status = ""
				If cash_disa_verif = "1" Then cash_disa_verif = "1 - DHS 161/Dr Stmt"
				If cash_disa_verif = "2" Then cash_disa_verif = "2 - SMRT Certified"
				If cash_disa_verif = "3" Then cash_disa_verif = "3 - Certified for RSDI or SSI"
				If cash_disa_verif = "6" Then cash_disa_verif = "6 - Other Document"
				If cash_disa_verif = "7" Then cash_disa_verif = "7 - Professional Stmt of Need"
				If cash_disa_verif = "N" Then cash_disa_verif = "N - No Verif Provided"

				If fs_disa_status = "01" Then fs_disa_status = "01 - RSDI Only Disability"
				If fs_disa_status = "02" Then fs_disa_status = "02 - RSDI Only Blindness"
				If fs_disa_status = "03" Then fs_disa_status = "03 - SSI, SSI/RSDI Disability"
				If fs_disa_status = "04" Then fs_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If fs_disa_status = "08" Then fs_disa_status = "08 - SMRT Certified Blindness"
				If fs_disa_status = "09" Then fs_disa_status = "09 - Ill/Incapacity"
				If fs_disa_status = "10" Then fs_disa_status = "10 - SMRT Certified Disability"
				If fs_disa_status = "11" Then fs_disa_status = "11 - VA Determined Pd - 100% Disa"
				If fs_disa_status = "12" Then fs_disa_status = "12 - VA (Other Accept Disa)"
				If fs_disa_status = "13" Then fs_disa_status = "13 - Certified RR Retirement Disa"
				If fs_disa_status = "14" Then fs_disa_status = "14 - Other Govt Permanent Disa"
				If fs_disa_status = "15" Then fs_disa_status = "15 - Disability from MINE List"
				If fs_disa_status = "16" Then fs_disa_status = "16 - Unable to Prepare Purch Own Meal"
				If fs_disa_status = "__" Then fs_disa_status = ""
				If fs_disa_verif = "1" Then fs_disa_verif = "1 - DHS 161/Dr Stmt"
				If fs_disa_verif = "2" Then fs_disa_verif = "2 - SMRT Certified"
				If fs_disa_verif = "3" Then fs_disa_verif = "3 - Certified for RSDI or SSI"
				If fs_disa_verif = "4" Then fs_disa_verif = "4 - Receipt of HC for Disa/Blind"
				If fs_disa_verif = "5" Then fs_disa_verif = "5 - Work Judgement"
				If fs_disa_verif = "6" Then fs_disa_verif = "6 - Other Document"
				If fs_disa_verif = "7" Then fs_disa_verif = "7 - Out of State Verif Pending"
				If fs_disa_verif = "N" Then fs_disa_verif = "N - No Verif Provided"

				If hc_disa_status = "01" Then hc_disa_status = "01 - RSDI Only Disability"
				If hc_disa_status = "02" Then hc_disa_status = "02 - RSDI Only Blindness"
				If hc_disa_status = "03" Then hc_disa_status = "03 - SSI, SSI/RSDI Disability"
				If hc_disa_status = "04" Then hc_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If hc_disa_status = "06" Then hc_disa_status = "06 - SMRT Pend or SSA Pend"
				If hc_disa_status = "08" Then hc_disa_status = "08 - Certified Blind"
				If hc_disa_status = "10" Then hc_disa_status = "10 - Certified Disabled"
				If hc_disa_status = "11" Then hc_disa_status = "11 - Special Category - Disabled Child"
				If hc_disa_status = "20" Then hc_disa_status = "20 - TEFRA - Disabled"
				If hc_disa_status = "21" Then hc_disa_status = "21 - TEFRA - Blind"
				If hc_disa_status = "22" Then hc_disa_status = "22 - MA-EPD"
				If hc_disa_status = "23" Then hc_disa_status = "23 - MA/Waiver"
				If hc_disa_status = "24" Then hc_disa_status = "24 - SSA/SMRT Appeal Pending"
				If hc_disa_status = "26" Then hc_disa_status = "26 - SSA/SMRT Disa Deny"
				If hc_disa_status = "__" Then hc_disa_status = ""
				If hc_disa_verif = "1" Then hc_disa_verif = "1 - DHS 161/Dr Stmt"
				If hc_disa_verif = "2" Then hc_disa_verif = "2 - SMRT Certified"
				If hc_disa_verif = "3" Then hc_disa_verif = "3 - Certified for RSDI or SSI"
				If hc_disa_verif = "6" Then hc_disa_verif = "6 - Other Document"
				If hc_disa_verif = "7" Then hc_disa_verif = "7 - Case Manager Determination"
				If hc_disa_verif = "8" Then hc_disa_verif = "8 - LTC Consult Services"
				If hc_disa_verif = "N" Then hc_disa_verif = "N - No Verif Provided"

				If disa_waiver = "F" Then disa_waiver = "F - LTC CADI Conversion"
				If disa_waiver = "G" Then disa_waiver = "G - LTC CADI DIversion"
				If disa_waiver = "H" Then disa_waiver = "H - LTC CAC Conversion"
				If disa_waiver = "I" Then disa_waiver = "I - LTC CAC Diversion"
				If disa_waiver = "J" Then disa_waiver = "J - LTC EW Conversion"
				If disa_waiver = "K" Then disa_waiver = "K - LTC EW Diversion"
				If disa_waiver = "L" Then disa_waiver = "L - LTC TBI NF Conversion"
				If disa_waiver = "M" Then disa_waiver = "M - LTC TBI NF Diversion"
				If disa_waiver = "P" Then disa_waiver = "P - LTC TBI NB Conversion"
				If disa_waiver = "Q" Then disa_waiver = "Q - LTC TBI NB Diversion"
				If disa_waiver = "R" Then disa_waiver = "R - DD Conversion"
				If disa_waiver = "S" Then disa_waiver = "S - DD Conversion"
				If disa_waiver = "Y" Then disa_waiver = "Y - CSG Conversion"
				If disa_waiver = "_" Then disa_waiver = ""

				If disa_1619 = "A" Then disa_1619 = "A - 1619A Status"
				If disa_1619 = "B" Then disa_1619 = "B - 1619B Status"
				If disa_1619 = "N" Then disa_1619 = "N - No 1619 Status"
				If disa_1619 = "T" Then disa_1619 = "T - 1619 Status Terminated"
				If disa_1619 = "_" Then disa_1619 = ""
			End If

			Call navigate_to_MAXIS_screen("STAT", "WREG")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen wreg_version, 1, 2, 73
			If wreg_version = "0" Then wreg_exists = FALSE
			If wreg_version = "1" Then wreg_exists = TRUE

			If wreg_exists = TRUE Then
				EMReadScreen wreg_pwe, 1, 6, 68

				If wreg_pwe = "Y" Then fs_pwe = "Yes"
				If wreg_pwe = "N" OR wreg_pwe = "_" Then fs_pwe = "No"
			End If


			Call navigate_to_MAXIS_screen("STAT", "SCHL")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen schl_version, 1, 2, 73
			If schl_version = "0" Then schl_exists = FALSE
			If schl_version = "1" Then schl_exists = TRUE

			If schl_exists = TRUE Then
				EMReadScreen schl_status, 1, 6, 40
				EMReadScreen schl_verif, 2, 6, 63
				EMReadScreen schl_type, 2, 7, 40
				EMReadScreen school_district, 4, 8, 40
				EMReadScreen schl_start_date, 8, 10, 63
				EMReadScreen schl_grad_date, 5, 11, 63
				EMReadScreen schl_grad_verif, 2, 12, 63
				EMReadScreen schl_fund, 1, 14, 63
				EMReadScreen schl_elig, 2, 16, 63
				EMReadScreen schl_higher_ed_yn, 1, 18, 63

				If schl_status = "F" Then school_status = "Fulltime"
				If schl_status = "H" Then school_status = "Halftime"
				If schl_status = "L" Then school_status = "Less than Half "
				If schl_status = "N" Then school_status = "Not Attending"

				If schl_verif = "SC" Then school_verif = "SC - School Statement"
				If schl_verif = "OT" Then school_verif = "OT - Other Document"
				If schl_verif = "NO" Then school_verif = "NO - No Verif Provided"
				If schl_verif = "__" Then school_verif = "Blank"

				If schl_type = "01" Then school_type = "01 - Preschool - 6"
				If schl_type = "11" Then school_type = "11 - 7 - 8"
				If schl_type = "02" Then school_type = "02 - 9 - 12"
				If schl_type = "03" Then school_type = "03 - GED Or Equiv"
				If schl_type = "06" Then school_type = "06 - Child, Not In School"
				If schl_type = "07" Then school_type = "07 - Individual Ed Plan/IEP"
				If schl_type = "08" Then school_type = "08 - Post-Sec Not Grad Student"
				If schl_type = "09" Then school_type = "09 - Post-Sec Grad Student"
				If schl_type = "10" Then school_type = "10 - Post-Sec Tech Schl"
				If schl_type = "12" Then school_type = "11 - Adult Basic Ed (ABE)"
				If schl_type = "13" Then school_type = "13 - English As A 2nd Language"

				If school_district = "____" Then school_district = ""

				kinder_start_date = replace(schl_start_date, " ", "/")
				If kinder_start_date = "__/__/__" Then kinder_start_date = ""

				grad_date = replace(schl_grad_date, " ", "/")
				If grad_date = "__/__" Then grad_date = ""

				If schl_grad_verif = "SC" Then grad_date_verif = "SC - School Statement"
				If schl_grad_verif = "OT" Then grad_date_verif = "OT - Other Document"
				If schl_grad_verif = "NO" Then grad_date_verif = "NO - No Verif Provided"
				If schl_grad_verif = "__" Then grad_date_verif = "Blank"

				If schl_fund = "1" Then school_funding = "1 - Not Attending in MN"
				If schl_fund = "2" Then school_funding = "2 - Attending Pub School"
				If schl_fund = "3" Then school_funding = "3 - Attending private/Parochial"
				If schl_fund = "4" Then school_funding = "4 - Not in Pre-12"

				If schl_elig = "01" Then school_elig_status = "01 - Under 18 or Over 50"
				If schl_elig = "02" Then school_elig_status = "02 - Disabled"
				If schl_elig = "03" Then school_elig_status = "03 - Not Higher Ed or < Halftime"
				If schl_elig = "04" Then school_elig_status = "04 - Employed 20 hrs/wk"
				If schl_elig = "05" Then school_elig_status = "05 - Work Study Program"
				If schl_elig = "06" Then school_elig_status = "06 - Dependant under 6"
				If schl_elig = "07" Then school_elig_status = "07 - Dep 6-11 No Child Care"
				If schl_elig = "09" Then school_elig_status = "09 - WIA, TAA, TRA or FSET"
				If schl_elig = "10" Then school_elig_status = "10 - Single Parent w/ Child < 12"
				If schl_elig = "99" Then school_elig_status = "99 - Not Eligible"

				If schl_higher_ed_yn = "Y" Then higher_ed = "Yes"
				If schl_higher_ed_yn = "N" Then higher_ed = "No"
				If schl_higher_ed_yn = "_" Then higher_ed = "Blank"

			End If

			Call navigate_to_MAXIS_screen("STAT", "STIN")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen stin_version, 1, 2, 73
			If stin_version = "0" Then stin_exists = FALSE
			If stin_version = "1" Then stin_exists = TRUE

			If stin_exists = TRUE Then
				total_stin = 0

				stin_type_array = ARRAY("")
				stin_amount_array = ARRAY("")
				stin_avail_date_array = ARRAY("")
				stin_months_cov_array = ARRAY("")
				stin_verif_array = ARRAY("")

				stin_row = 8
				stin_counter = 0
				Do
					EMReadScreen stin_type, 2, stin_row, 27
					EMReadScreen stin_amount, 8, stin_row, 34
					EMReadScreen stin_date, 8, stin_row, 46
					EMReadScreen stin_month_one, 5, stin_row, 58
					EmReadscreen stin_month_two, 5, stin_row, 67
					EMReadScreen stin_verif, 1, stin_row, 76


					ReDim Preserve stin_type_array(stin_counter)
					ReDim Preserve stin_amount_array(stin_counter)
					ReDim Preserve stin_avail_date_array(stin_counter)
					ReDim Preserve stin_months_cov_array(stin_counter)
					ReDim Preserve stin_verif_array(stin_counter)

					If stin_type = "01" Then stin_type_array(stin_counter) = stin_type & " - Perkins Loan"
					If stin_type = "02" Then stin_type_array(stin_counter) = stin_type & " - Stafford Loan"
					If stin_type = "03" Then stin_type_array(stin_counter) = stin_type & " - Pell Grant"
					If stin_type = "04" Then stin_type_array(stin_counter) = stin_type & " - BIA Grant"
					If stin_type = "05" Then stin_type_array(stin_counter) = stin_type & " - SEOG"
					If stin_type = "06" Then stin_type_array(stin_counter) = stin_type & " - MN State Scholarship"
					If stin_type = "07" Then stin_type_array(stin_counter) = stin_type & " - Robert C Byrd Scholarship"
					If stin_type = "46" Then stin_type_array(stin_counter) = stin_type & " - Plus Loan (Deferred)"
					If stin_type = "16" Then stin_type_array(stin_counter) = stin_type & " - Plus Loan (Non-Deferred)"
					If stin_type = "47" Then stin_type_array(stin_counter) = stin_type & " - SLS (ALAS) Loan (Deferred)"
					If stin_type = "17" Then stin_type_array(stin_counter) = stin_type & " - SLS (ALAS) Loan (Non-Deferred)"
					If stin_type = "08" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Deferred Income"
					If stin_type = "09" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Grant"
					If stin_type = "10" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Scholarship"
					If stin_type = "11" Then stin_type_array(stin_counter) = stin_type & " - VA/GI Bill"
					If stin_type = "51" Then stin_type_array(stin_counter) = stin_type & " - VA/GI Bill (Earmarked)"
					If stin_type = "12" Then stin_type_array(stin_counter) = stin_type & " - Other Deferred Loan"
					If stin_type = "52" Then stin_type_array(stin_counter) = stin_type & " - Other Deferred Loan (Earmarked)"
					If stin_type = "13" Then stin_type_array(stin_counter) = stin_type & " - Other Grant"
					If stin_type = "53" Then stin_type_array(stin_counter) = stin_type & " - Other Grant (Earmarked)"
					If stin_type = "14" Then stin_type_array(stin_counter) = stin_type & " - Other Scholarship"
					If stin_type = "54" Then stin_type_array(stin_counter) = stin_type & " - Other Scholarship (Earmarked)"
					If stin_type = "15" Then stin_type_array(stin_counter) = stin_type & " - Other Aid"
					If stin_type = "55" Then stin_type_array(stin_counter) = stin_type & " - Other Aid (Earmarked)"
					If stin_type = "60" Then stin_type_array(stin_counter) = stin_type & " - MFIP Empl Svc (Earmarked)"
					If stin_type = "61" Then stin_type_array(stin_counter) = stin_type & " - WIOA, Unearned (Earmarked)"
					If stin_type = "18" Then stin_type_array(stin_counter) = stin_type & " - Other Exempt Loan"
					If stin_type = "62" Then stin_type_array(stin_counter) = stin_type & " - Tribal DSARLP"

					stin_amount_array(stin_counter) = trim(stin_amount)

					stin_avail_date_array(stin_counter) = replace(stin_date, " ", "/")

					stin_month_one = replace(stin_month_one, " ", "/")
					stin_month_two = replace(stin_month_two, " ", "/")
					stin_months_cov_array(stin_counter) = stin_month_one & " - " & stin_month_two

					If stin_verif = "1" Then stin_verif_array(stin_counter) = stin_verif & " - Award Letter"
					If stin_verif = "2" Then stin_verif_array(stin_counter) = stin_verif & " - DHS Financial Aid Form"
					If stin_verif = "3" Then stin_verif_array(stin_counter) = stin_verif & " - Student Profile Bulletin"
					If stin_verif = "4" Then stin_verif_array(stin_counter) = stin_verif & " - Pay Stubs"
					If stin_verif = "5" Then stin_verif_array(stin_counter) = stin_verif & " - Source Document"
					If stin_verif = "6" Then stin_verif_array(stin_counter) = stin_verif & " - Pend Out State Verif"
					If stin_verif = "7" Then stin_verif_array(stin_counter) = stin_verif & " - Other Document"
					If stin_verif = "N" Then stin_verif_array(stin_counter) = stin_verif & " - No Ver Prvd"

					stin_amount = stin_amount * 1
					total_stin = total_stin + stin_amount

					stin_row = stin_row + 1
					stin_counter = stin_counter + 1

					If stin_row = 18 Then
						PF20
						EMReadscreen last_page, 9, 24, 14
						If last_page = "LAST PAGE" Then Exit Do
						stin_row = 8
					End If
					EMReadScreen next_stin_type, 2, stin_row, 27
				Loop until next_stin_type = "__"

			End If

			Call navigate_to_MAXIS_screen("STAT", "STEC")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen stec_version, 1, 2, 73
			If stec_version = "0" Then stec_exists = FALSE
			If stec_version = "1" Then stec_exists = TRUE

			If stec_exists = TRUE Then
				total_stec = 0

				stec_type_array = ARRAY("")
				stec_amount_array = ARRAY("")
				stec_months_cov_array = ARRAY("")
				stec_verif_array = ARRAY("")
				stec_earmarked_amount_array = ARRAY("")
				stec_earmarked_months_cov_array = ARRAY("")

				stec_row = 8
				stec_counter = 0
				Do
					EMReadScreen stec_type, 2, stec_row, 25
					EMReadScreen stec_amount, 8, stec_row, 31
					EMReadScreen stec_month_one, 5, stec_row, 41
					EMReadScreen stec_month_two, 5, stec_row, 48
					EMReadScreen stec_verif, 1, stec_row, 55
					EMReadScreen stec_earmarked_amount, 8, stec_row, 59
					EMReadScreen stec_earmarked_month_one, 2, stec_row, 69
					EMReadScreen stec_earmarked_month_two, 2, stec_row, 76

					ReDim Preserve stec_type_array(stec_counter)
					ReDim Preserve stec_amount_array(stec_counter)
					ReDim Preserve stec_months_cov_array(stec_counter)
					ReDim Preserve stec_verif_array(stec_counter)
					ReDim Preserve stec_earmarked_amount_array(stec_counter)
					ReDim Preserve stec_earmarked_months_cov_array(stec_counter)

					If stec_type = "" Then stec_type_array(stec_counter) = stec_type & " - "

					stec_amount_array(stec_counter) = trim(stec_amount)

					stec_month_one = replace(stec_month_one, " ", "/")
					stec_month_two = replace(stec_month_two, " ", "/")
					stec_months_cov_array(stec_counter) = stec_month_one & " - " & stec_month_two

					If stec_verif = "" Then stec_verif_array(stec_counter) = stec_verif & " - "

					stec_earmarked_amount_array(stec_counter) = trim(stec_earmarked_amount)

					stec_earmarked_month_one = replace(stec_earmarked_month_one, " ", "/")
					stec_earmarked_month_two = replace(stec_earmarked_month_two, " ", "/")
					stec_earmarked_months_cov_array(stec_counter) = stec_earmarked_month_one & " - " & stec_earmarked_month_two

					stec_amount = stec_amount * 1
					total_stec = total_stec + stec_amount

					stec_row = stec_row + 1
					stec_counter = stec_counter + 1

					If stec_row = 17 Then
						PF20
						EMReadscreen last_page, 9, 24, 14
						If last_page = "LAST PAGE" Then Exit Do
						stec_row = 8
					End If
					EMReadScreen next_stec_type, 2, stec_row, 25
				Loop until next_stec_type = "__"
			End If

			Call navigate_to_MAXIS_screen("STAT", "SHEL")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen shel_version, 1, 2, 73
			If shel_version = "0" Then shel_exists = FALSE
			If shel_version = "1" Then shel_exists = TRUE

			If shel_exists = TRUE Then
				EMReadScreen shel_hud_subsidy_yn, 1, 6, 46
				EMReadScreen shel_shared_yn, 1, 6, 64

				EMReadScreen shel_paid_to, 25, 7, 50

				EMReadScreen shel_retro_rent_amount, 8, 11, 37
				EMReadScreen shel_retro_rent_verif, 2, 11, 48
				EMReadScreen shel_retro_lot_rent_amount, 8, 12, 37
				EMReadScreen shel_retro_lot_rent_verif, 2, 12, 48
				EMReadScreen shel_retro_mortgage_amount, 8, 13, 37
				EMReadScreen shel_retro_mortgage_verif, 2, 13, 48
				EMReadScreen shel_retro_insurance_amount, 8, 14, 37
				EMReadScreen shel_retro_insurance_verif, 2, 14, 48
				EMReadScreen shel_retro_taxes_amount, 8, 15, 37
				EMReadScreen shel_retro_taxes_verif, 2, 15, 48
				EMReadScreen shel_retro_room_amount, 8, 16, 37
				EMReadScreen shel_retro_room_verif, 2, 16, 48
				EMReadScreen shel_retro_garage_amount, 8, 17, 37
				EMReadScreen shel_retro_garage_verif, 2, 17, 48
				EMReadScreen shel_retro_subsidy_amount, 8, 18, 37
				EMReadScreen shel_retro_subsidy_verif, 2, 18, 48

				EMReadScreen shel_prosp_rent_amount, 8, 11, 56
				EMReadScreen shel_prosp_rent_verif, 2, 11, 67
				EMReadScreen shel_prosp_lot_rent_amount, 8, 12, 56
				EMReadScreen shel_prosp_lot_rent_verif, 2, 12, 67
				EMReadScreen shel_prosp_mortgage_amount, 8, 13, 56
				EMReadScreen shel_prosp_mortgage_verif, 2, 13, 67
				EMReadScreen shel_prosp_insurance_amount, 8, 14, 56
				EMReadScreen shel_prosp_insurance_verif, 2, 14, 67
				EMReadScreen shel_prosp_taxes_amount, 8, 15, 56
				EMReadScreen shel_prosp_taxes_verif, 2, 15, 67
				EMReadScreen shel_prosp_room_amount, 8, 16, 56
				EMReadScreen shel_prosp_room_verif, 2, 16, 67
				EMReadScreen shel_prosp_garage_amount, 8, 17, 56
				EMReadScreen shel_prosp_garage_verif, 2, 17, 67
				EMReadScreen shel_prosp_subsidy_amount, 8, 18, 56
				EMReadScreen shel_prosp_subsidy_verif, 2, 18, 67

				shel_paid_to = replace(shel_paid_to, "_", "")

				shel_retro_rent_amount = trim(replace(shel_retro_rent_amount, "_", ""))
				shel_retro_lot_rent_amount = trim(replace(shel_retro_lot_rent_amount, "_", ""))
				shel_retro_mortgage_amount = trim(replace(shel_retro_mortgage_amount, "_", ""))
				shel_retro_insurance_amount = trim(replace(shel_retro_insurance_amount, "_", ""))
				shel_retro_taxes_amount = trim(replace(shel_retro_taxes_amount, "_", ""))
				shel_retro_room_amount = trim(replace(shel_retro_room_amount, "_", ""))
				shel_retro_garage_amount = trim(replace(shel_retro_garage_amount, "_", ""))
				shel_retro_subsidy_amount = trim(replace(shel_retro_subsidy_amount, "_", ""))

				shel_prosp_rent_amount = trim(replace(shel_prosp_rent_amount, "_", ""))
				shel_prosp_lot_rent_amount = trim(replace(shel_prosp_lot_rent_amount, "_", ""))
				shel_prosp_mortgage_amount = trim(replace(shel_prosp_mortgage_amount, "_", ""))
				shel_prosp_insurance_amount = trim(replace(shel_prosp_insurance_amount, "_", ""))
				shel_prosp_taxes_amount = trim(replace(shel_prosp_taxes_amount, "_", ""))
				shel_prosp_room_amount = trim(replace(shel_prosp_room_amount, "_", ""))
				shel_prosp_garage_amount = trim(replace(shel_prosp_garage_amount, "_", ""))
				shel_prosp_subsidy_amount = trim(replace(shel_prosp_subsidy_amount, "_", ""))

				If shel_prosp_rent_amount <> "" Then shel_summary = shel_summary & " Rent: $" & shel_prosp_rent_amount & " - Verif: " & shel_prosp_rent_verif & " | "
				If shel_prosp_lot_rent_amount <> "" Then shel_summary = shel_summary & " Lot Rent: $" & shel_prosp_lot_rent_amount & " - Verif: " & shel_prosp_lot_rent_verif & " | "
				If shel_prosp_mortgage_amount <> "" Then shel_summary = shel_summary & " Mortgage: $" & shel_prosp_mortgage_amount & " - Verif: " & shel_prosp_mortgage_verif & " | "
				If shel_prosp_insurance_amount <> "" Then shel_summary = shel_summary & " Insurance: $" & shel_prosp_insurance_amount & " - Verif: " & shel_prosp_insurance_verif & " | "
				If shel_prosp_taxes_amount <> "" Then shel_summary = shel_summary & " Taxes: $" & shel_prosp_taxes_amount & " - Verif: " & shel_prosp_taxes_verif & " | "
				If shel_prosp_room_amount <> "" Then shel_summary = shel_summary & " Room: $" & shel_prosp_room_amount & " - Verif: " & shel_prosp_room_verif & " | "
				If shel_prosp_garage_amount <> "" Then shel_summary = shel_summary & " Garage: $" & shel_prosp_garage_amount & " - Verif: " & shel_prosp_garage_verif & " | "
				If shel_prosp_subsidy_amount <> "" Then shel_summary = shel_summary & " Subsidy: $" & shel_prosp_subsidy_amount & " - Verif: " & shel_prosp_subsidy_verif & " | "

				If shel_retro_rent_verif = "SF" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Shelter Form"
				If shel_retro_rent_verif = "LE" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Lease"
				If shel_retro_rent_verif = "RE" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Rent Receipts"
				If shel_retro_rent_verif = "OT" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Other Document"
				If shel_retro_rent_verif = "NC" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Not Verif, Neg Impact"
				If shel_retro_rent_verif = "PC" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Not Verif, Pos Impact"
				If shel_retro_rent_verif = "NO" Then shel_retro_rent_verif = shel_retro_rent_verif & " - No Verif Provided"

				If shel_retro_lot_rent_verif = "LE" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Lease"
				If shel_retro_lot_rent_verif = "RE" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Rent Receipts"
				If shel_retro_lot_rent_verif = "BI" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Billing Statement"
				If shel_retro_lot_rent_verif = "OT" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Other Document"
				If shel_retro_lot_rent_verif = "NC" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Not Verif, Neg Impact"
				If shel_retro_lot_rent_verif = "PC" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Not Verif, Pos Impact"
				If shel_retro_lot_rent_verif = "NO" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - No Verif Provided"

				If shel_retro_mortgage_verif = "MO" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Mortgage Payment"
				If shel_retro_mortgage_verif = "CD" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Contract for Deed"
				If shel_retro_mortgage_verif = "OT" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Other Document"
				If shel_retro_mortgage_verif = "NC" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Not Verif, Neg Impact"
				If shel_retro_mortgage_verif = "PC" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Not Verif, Pos Impact"
				If shel_retro_mortgage_verif = "NO" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - No Verif Provided"

				If shel_retro_insurance_verif = "BI" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Billing Statement"
				If shel_retro_insurance_verif = "OT" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Other Document"
				If shel_retro_insurance_verif = "NC" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Not Verif, Neg Impact"
				If shel_retro_insurance_verif = "PC" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Not Verif, Pos Impact"
				If shel_retro_insurance_verif = "NO" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - No Verif Provided"

				If shel_retro_taxes_verif = "TX" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Property Tax Statement"
				If shel_retro_taxes_verif = "OT" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Other Document"
				If shel_retro_taxes_verif = "NC" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Not Verif, Neg Impact"
				If shel_retro_taxes_verif = "PC" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Not Verif, Pos Impact"
				If shel_retro_taxes_verif = "NO" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - No Verif Provided"

				If shel_retro_room_verif = "SF" Then shel_retro_room_verif = shel_retro_room_verif & " - Shelter Form"
				If shel_retro_room_verif = "LE" Then shel_retro_room_verif = shel_retro_room_verif & " - Lease"
				If shel_retro_room_verif = "RE" Then shel_retro_room_verif = shel_retro_room_verif & " - Rent Receipts"
				If shel_retro_room_verif = "OT" Then shel_retro_room_verif = shel_retro_room_verif & " - Other Document"
				If shel_retro_room_verif = "NC" Then shel_retro_room_verif = shel_retro_room_verif & " - Not Verif, Neg Impact"
				If shel_retro_room_verif = "PC" Then shel_retro_room_verif = shel_retro_room_verif & " - Not Verif, Pos Impact"
				If shel_retro_room_verif = "NO" Then shel_retro_room_verif = shel_retro_room_verif & " - No Verif Provided"

				If shel_retro_garage_verif = "SF" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Shelter Form"
				If shel_retro_garage_verif = "LE" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Lease"
				If shel_retro_garage_verif = "RE" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Rent Receipts"
				If shel_retro_garage_verif = "OT" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Other Document"
				If shel_retro_garage_verif = "NC" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Not Verif, Neg Impact"
				If shel_retro_garage_verif = "PC" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Not Verif, Pos Impact"
				If shel_retro_garage_verif = "NO" Then shel_retro_garage_verif = shel_retro_garage_verif & " - No Verif Provided"

				If shel_retro_subsidy_verif = "SF" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Shelter Form"
				If shel_retro_subsidy_verif = "LE" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Lease"
				If shel_retro_subsidy_verif = "OT" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Other Document"
				If shel_retro_subsidy_verif = "NO" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - No Verif Provided"


				If shel_prosp_rent_verif = "SF" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Shelter Form"
				If shel_prosp_rent_verif = "LE" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Lease"
				If shel_prosp_rent_verif = "RE" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Rent Receipts"
				If shel_prosp_rent_verif = "OT" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Other Document"
				If shel_prosp_rent_verif = "NC" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Not Verif, Neg Impact"
				If shel_prosp_rent_verif = "PC" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Not Verif, Pos Impact"
				If shel_prosp_rent_verif = "NO" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - No Verif Provided"

				If shel_prosp_lot_rent_verif = "LE" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Lease"
				If shel_prosp_lot_rent_verif = "RE" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Rent Receipts"
				If shel_prosp_lot_rent_verif = "BI" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Billing Statement"
				If shel_prosp_lot_rent_verif = "OT" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Other Document"
				If shel_prosp_lot_rent_verif = "NC" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Not Verif, Neg Impact"
				If shel_prosp_lot_rent_verif = "PC" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Not Verif, Pos Impact"
				If shel_prosp_lot_rent_verif = "NO" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - No Verif Provided"

				If shel_prosp_mortgage_verif = "MO" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Mortgage Payment"
				If shel_prosp_mortgage_verif = "CD" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Contract for Deed"
				If shel_prosp_mortgage_verif = "OT" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Other Document"
				If shel_prosp_mortgage_verif = "NC" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Not Verif, Neg Impact"
				If shel_prosp_mortgage_verif = "PC" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Not Verif, Pos Impact"
				If shel_prosp_mortgage_verif = "NO" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - No Verif Provided"

				If shel_prosp_insurance_verif = "BI" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Billing Statement"
				If shel_prosp_insurance_verif = "OT" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Other Document"
				If shel_prosp_insurance_verif = "NC" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Not Verif, Neg Impact"
				If shel_prosp_insurance_verif = "PC" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Not Verif, Pos Impact"
				If shel_prosp_insurance_verif = "NO" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - No Verif Provided"

				If shel_prosp_taxes_verif = "TX" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Property Tax Statement"
				If shel_prosp_taxes_verif = "OT" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Other Document"
				If shel_prosp_taxes_verif = "NC" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Not Verif, Neg Impact"
				If shel_prosp_taxes_verif = "PC" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Not Verif, Pos Impact"
				If shel_prosp_taxes_verif = "NO" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - No Verif Provided"

				If shel_prosp_room_verif = "SF" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Shelter Form"
				If shel_prosp_room_verif = "LE" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Lease"
				If shel_prosp_room_verif = "RE" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Rent Receipts"
				If shel_prosp_room_verif = "OT" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Other Document"
				If shel_prosp_room_verif = "NC" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Not Verif, Neg Impact"
				If shel_prosp_room_verif = "PC" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Not Verif, Pos Impact"
				If shel_prosp_room_verif = "NO" Then shel_prosp_room_verif = shel_prosp_room_verif & " - No Verif Provided"

				If shel_prosp_garage_verif = "SF" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Shelter Form"
				If shel_prosp_garage_verif = "LE" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Lease"
				If shel_prosp_garage_verif = "RE" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Rent Receipts"
				If shel_prosp_garage_verif = "OT" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Other Document"
				If shel_prosp_garage_verif = "NC" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Not Verif, Neg Impact"
				If shel_prosp_garage_verif = "PC" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Not Verif, Pos Impact"
				If shel_prosp_garage_verif = "NO" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - No Verif Provided"

				If shel_prosp_subsidy_verif = "SF" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Shelter Form"
				If shel_prosp_subsidy_verif = "LE" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Lease"
				If shel_prosp_subsidy_verif = "OT" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Other Document"
				If shel_prosp_subsidy_verif = "NO" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - No Verif Provided"

			End If

			Call navigate_to_MAXIS_screen("STAT", "STWK")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen stwk_version, 1, 2, 73
			If stwk_version = "0" Then stwk_exists = FALSE
			If stwk_version = "1" Then stwk_exists = TRUE

			If stwk_exists = TRUE Then
				EMReadScreen stwk_employer, 30, 6, 46
				EMReadScreen stwk_work_stop_date, 8, 7, 46
				EMReadScreen stwk_income_stop_date, 8, 8, 46
				EMReadScreen stwk_verification, 1, 7, 63
				EMReadScreen stwk_refused_employment, 1, 8, 78
				EMReadScreen stwk_vol_quit, 1, 10, 46
				EMReadScreen stwk_refused_employment_date, 8, 10, 72
				EMReadScreen stwk_cash_good_cause_yn, 1, 12, 52
				EMReadScreen stwk_grh_good_cause_yn, 1, 12, 60
				EMReadScreen stwk_snap_good_cause_yn, 1, 12, 67
				EMReadScreen stwk_snap_pwe, 1, 14, 46
				EMReadScreen stwk_ma_epd_extension, 1, 16, 46

				stwk_employer = replace(stwk_employer, "_", "")
				stwk_work_stop_date = replace(stwk_work_stop_date, " ", "/")
				stwk_income_stop_date = replace(stwk_income_stop_date, " ", "/")
				If stwk_verification = "1" Then stwk_verification = "Employers Statement"
				If stwk_verification = "2" Then stwk_verification = "Seperation Notice"
				If stwk_verification = "3" Then stwk_verification = "Colateral Statement"
				If stwk_verification = "4" Then stwk_verification = "Other Document"
				If stwk_verification = "N" Then stwk_verification = "No Verif Provided"
				If stwk_verification = "_" Then stwk_verification = "Blank"
				If stwk_verification = "?" Then stwk_verification = "Postponed Verif"
				stwk_refused_employment_date = replace(stwk_refused_employment_date, " ", "/")
				stwk_summary = "Work ended at " & stwk_employer & " on " & stwk_work_stop_date

			End If

			Call navigate_to_MAXIS_screen("STAT", "FMED")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen fmed_version, 1, 2, 73
			If fmed_version = "0" Then fmed_exists = FALSE
			If fmed_version = "1" Then fmed_exists = TRUE

			If fmed_exists = TRUE Then
				EMReadScreen fmed_miles, 4, 17, 34
				EMReadScreen fmed_rate, 6, 17, 58
				EMReadScreen fmed_milage_expense, 8, 17, 70

				panel_row = 9
				fmed_count = 0
				scroll_page = 1
				Do
					EMReadScreen the_type, 2, panel_row, 25

					If the_type <> "__" Then
						' ReDim Preserve fmed_expense_array(fmed_count, fmed_notes)
						ReDim Preserve fmed_page(fmed_count)
						ReDim Preserve fmed_row(fmed_count)
						ReDim Preserve fmed_type(fmed_count)
						ReDim Preserve fmed_verif(fmed_count)
						ReDim Preserve fmed_ref(fmed_count)
						ReDim Preserve fmed_catgry(fmed_count)
						ReDim Preserve fmed_begin(fmed_count)
						ReDim Preserve fmed_end(fmed_count)
						ReDim Preserve fmed_expense(fmed_count)
						ReDim Preserve fmed_notes(fmed_count)

						EMReadScreen the_ver, 2, panel_row, 32
						EMReadScreen the_ref, 2, panel_row, 38
						EMReadScreen the_cat, 1, panel_row, 44
						EMReadScreen the_begin, 5, panel_row, 50
						EMReadScreen the_end, 5, panel_row, 60
						EMReadScreen the_amt, 8, panel_row, 70

						fmed_page(fmed_count) = scroll_page
						fmed_row(fmed_count) = panel_row

						If the_type = "01" Then fmed_type(fmed_count) = "01 Nursing Home"
						If the_type = "02" Then fmed_type(fmed_count) = "02 Hosp/Clinic"
						If the_type = "03" Then fmed_type(fmed_count) = "03 Physicians"
						If the_type = "04" Then fmed_type(fmed_count) = "04 Prescriptions"
						If the_type = "05" Then fmed_type(fmed_count) = "05 Ins Premiums"
						If the_type = "06" Then fmed_type(fmed_count) = "06 Dental"
						If the_type = "07" Then fmed_type(fmed_count) = "07 Medical Trans/Flat Amount"
						If the_type = "08" Then fmed_type(fmed_count) = "08 Vision Care"
						If the_type = "09" Then fmed_type(fmed_count) = "09 Medicare Prem"
						If the_type = "10" Then fmed_type(fmed_count) = "10 Mo Spdwn Amt/Waiver Oblig"
						If the_type = "11" Then fmed_type(fmed_count) = "11 Home Care"
						If the_type = "12" Then fmed_type(fmed_count) = "12 Medical Trans/Mileage Calc"
						If the_type = "15" Then fmed_type(fmed_count) = "15 Medi Part D Premium"

						If the_ver = "BI" Then fmed_verif(fmed_count) = "BI Billing Stmt"
						If the_ver = "EB" Then fmed_verif(fmed_count) = "EB Expl Of Bnft (Medicare/Ins)"
						If the_ver = "CL" Then fmed_verif(fmed_count) = "CL Client Stmt Med Trans Only"
						If the_ver = "OS" Then fmed_verif(fmed_count) = "OS Pend Out State Verification"
						If the_ver = "OT" Then fmed_verif(fmed_count) = "OT Other Document"
						If the_ver = "NO" Then fmed_verif(fmed_count) = "NO No Ver Prvd"
						If the_ver = "MX" Then fmed_verif(fmed_count) = "MX System Entered Ver By SSA"

						fmed_ref(fmed_count) = the_ref

						If the_cat = "1" Then fmed_catgry(fmed_count) = "1 HH Member"
						If the_cat = "2" Then fmed_catgry(fmed_count) = "2 Former Aged/Disa HH Mbr In NF Or Hospital"
						If the_cat = "3" Then fmed_catgry(fmed_count) = "3 Former Aged/Disa HH Decd"
						If the_cat = "4" Then fmed_catgry(fmed_count) = "4 Other Eligible"

						fmed_begin(fmed_count) = replace(the_begin, " ", "/")
						fmed_end(fmed_count) = replace(the_end, " ", "/")
						fmed_expense(fmed_count) = trim(the_amt)

						panel_row = panel_row + 1
						fmed_count = fmed_count + 1
						If panel_row = 15 Then
							pf20
							scroll_page = scroll_page + 1
							panel_row = 9
							EMReadScreen end_of_list, 9, 24, 14
							If end_of_list = "LAST PAGE" Then Exit Do
						End If
					End If
				Loop until panel_type = "__"
			End If

			Call navigate_to_MAXIS_screen("STAT", "PARE")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen pare_version, 1, 2, 73
			If pare_version = "0" Then pare_exists = FALSE
			If pare_version = "1" Then pare_exists = TRUE

			If pare_exists = TRUE Then
				pare_row = 8
				pare_array_count = 0

				Do
					EMReadScreen panel_child_ref_number, 2, pare_row, 24
					EMReadScreen panel_child_name, 25, pare_row, 27
					EMReadScreen panel_rela_type, 1, pare_row, 53
					EMReadScreen panel_rela_verif, 2, pare_row, 71

					If panel_child_ref_number <> "__" Then
						ReDim preserve pare_child_ref_nbr(pare_array_count)
						ReDim preserve pare_child_name(pare_array_count)
						ReDim preserve pare_child_member_index(pare_array_count)
						ReDim preserve pare_relationship_type(pare_array_count)
						ReDim preserve pare_verification(pare_array_count)

						pare_child_ref_nbr(pare_array_count) = panel_child_ref_number
						pare_child_name(pare_array_count) = trim(panel_child_name)

						' pare_child_member_index(pare_array_count)

						If panel_rela_type = "1" Then pare_relationship_type(pare_array_count) = "1 - Birth/Adopted Parent"
						If panel_rela_type = "2" Then pare_relationship_type(pare_array_count) = "2 - Stepchild"
						If panel_rela_type = "3" Then pare_relationship_type(pare_array_count) = "3 - Grandchild"
						If panel_rela_type = "4" Then pare_relationship_type(pare_array_count) = "4 - Relative Caregiver"
						If panel_rela_type = "5" Then pare_relationship_type(pare_array_count) = "5 - Foster Child"
						If panel_rela_type = "6" Then pare_relationship_type(pare_array_count) = "6 - Non-related Caregiver"
						If panel_rela_type = "7" Then pare_relationship_type(pare_array_count) = "7 - Legal Guardian"
						If panel_rela_type = "8" Then pare_relationship_type(pare_array_count) = "8 - Other Relative"

						If panel_rela_verif = "BC" Then pare_verification(pare_array_count) = "BC - Birth Certificate"
						If panel_rela_verif = "AR" Then pare_verification(pare_array_count) = "AR - Adoption Records"
						If panel_rela_verif = "LG" Then pare_verification(pare_array_count) = "LG - Legal Guardian"
						If panel_rela_verif = "RE" Then pare_verification(pare_array_count) = "RE - Religious Records"
						If panel_rela_verif = "HR" Then pare_verification(pare_array_count) = "HR - Hospital Records"
						If panel_rela_verif = "RP" Then pare_verification(pare_array_count) = "RP - Recognition of Parentage"
						If panel_rela_verif = "OT" Then pare_verification(pare_array_count) = "OT - Other Verification"
						If panel_rela_verif = "NO" Then pare_verification(pare_array_count) = "NO - No Verif Provided"
						If panel_rela_verif = "__" Then pare_verification(pare_array_count) = "Blank"
						If panel_rela_verif = "?_" Then pare_verification(pare_array_count) = "Delayed Verification"
					End If

					pare_row = pare_row + 1
					pare_array_count = pare_array_count + 1
					If pare_row = 18 Then
						pare_row = 8
						PF20
						EMReadScreen end_of_list, 9, 24, 14
						If end_of_list = "LAST PAGE" then Exit Do
					End If
				Loop until panel_child_ref_number = "__"
			End If

			Call navigate_to_MAXIS_screen("STAT", "PDED")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen pded_version, 1, 2, 73
			If pded_version = "0" Then pded_exists = FALSE
			If pded_version = "1" Then pded_exists = TRUE

			If pded_exists = TRUE Then
				EMReadScreen pded_guardian_fee, 8, 15, 44
				EMReadScreen pded_rep_payee_fee, 8, 15, 70
				EMReadScreen pded_shel_spec_need, 1, 18, 78

				pded_guardian_fee = replace(pded_guardian_fee, "_", "")
				pded_guardian_fee = trim(pded_guardian_fee)
				' MsgBox pded_rep_payee_fee & " 1"
				pded_rep_payee_fee = replace(pded_rep_payee_fee, "_", "")
				pded_rep_payee_fee = trim(pded_rep_payee_fee)
				' MsgBox pded_rep_payee_fee & " 2"

				If pded_shel_spec_need = "Y" Then pded_shel_spec_need = "Yes"
				If pded_shel_spec_need = "N" Then pded_shel_spec_need = "No"
				If pded_shel_spec_need = "_" Then pded_shel_spec_need = ""
			End If


			Call navigate_to_MAXIS_screen("STAT", "DIET")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen diet_version, 1, 2, 73
			If diet_version = "0" Then diet_exists = FALSE
			If diet_version = "1" Then diet_exists = TRUE

			If diet_exists = TRUE Then
				EMReadScreen diet_mf_type_one, 2, 8, 40
				EMReadScreen diet_mf_verif_one, 1, 8, 51
				EMReadScreen diet_mf_type_two, 2, 9, 40
				EMReadScreen diet_mf_verif_two, 1, 9, 51

				EMReadScreen diet_msa_type_one, 2, 11, 40
				EMReadScreen diet_msa_verif_one, 1, 11, 51
				EMReadScreen diet_msa_type_two, 2, 12, 40
				EMReadScreen diet_msa_verif_two, 1, 12, 51
				EMReadScreen diet_msa_type_three, 2, 13, 40
				EMReadScreen diet_msa_verif_three, 1, 13, 51
				EMReadScreen diet_msa_type_four, 2, 14, 40
				EMReadScreen diet_msa_verif_four, 1, 14, 51
				EMReadScreen diet_msa_type_five, 2, 15, 40
				EMReadScreen diet_msa_verif_five, 1, 15, 51
				EMReadScreen diet_msa_type_six, 2, 16, 40
				EMReadScreen diet_msa_verif_six, 1, 16, 51
				EMReadScreen diet_msa_type_seven, 2, 17, 40
				EMReadScreen diet_msa_verif_seven, 1, 17, 51
				EMReadScreen diet_msa_type_eight, 2, 18, 40
				EMReadScreen diet_msa_verif_eight, 1, 18, 51

				If diet_mf_type_one = "01" Then diet_mf_type_one = "01 - High Protein > 79 grams/day"
				If diet_mf_type_one = "02" Then diet_mf_type_one = "02 - Control Protein 40-60 grams/day"
				If diet_mf_type_one = "03" Then diet_mf_type_one = "03 - Control Protein < 40 grams/day"
				If diet_mf_type_one = "04" Then diet_mf_type_one = "04 - Lo Cholesterol"
				If diet_mf_type_one = "05" Then diet_mf_type_one = "05 - High Residue"
				If diet_mf_type_one = "06" Then diet_mf_type_one = "06 - Pregnancy and Lactation"
				If diet_mf_type_one = "07" Then diet_mf_type_one = "07 - Gluten Free"
				If diet_mf_type_one = "08" Then diet_mf_type_one = "08 - Lactose Free"
				If diet_mf_type_one = "09" Then diet_mf_type_one = "09 - Anti-Dumping"
				If diet_mf_type_one = "10" Then diet_mf_type_one = "10 - Hypoglycemic"
				If diet_mf_type_one = "11" Then diet_mf_type_one = "11 - Ketogenic"
				If diet_mf_type_one = "__" Then diet_mf_type_one = ""

				If diet_mf_type_two = "01" Then diet_mf_type_two = "01 - High Protein > 79 grams/day"
				If diet_mf_type_two = "02" Then diet_mf_type_two = "02 - Control Protein 40-60 grams/day"
				If diet_mf_type_two = "03" Then diet_mf_type_two = "03 - Control Protein < 40 grams/day"
				If diet_mf_type_two = "04" Then diet_mf_type_two = "04 - Lo Cholesterol"
				If diet_mf_type_two = "05" Then diet_mf_type_two = "05 - High Residue"
				If diet_mf_type_two = "06" Then diet_mf_type_two = "06 - Pregnancy and Lactation"
				If diet_mf_type_two = "07" Then diet_mf_type_two = "07 - Gluten Free"
				If diet_mf_type_two = "08" Then diet_mf_type_two = "08 - Lactose Free"
				If diet_mf_type_two = "09" Then diet_mf_type_two = "09 - Anti-Dumping"
				If diet_mf_type_two = "10" Then diet_mf_type_two = "10 - Hypoglycemic"
				If diet_mf_type_two = "11" Then diet_mf_type_two = "11 - Ketogenic"
				If diet_mf_type_two = "__" Then diet_mf_type_two = ""


				If diet_msa_type_one = "01" Then diet_msa_type_one = "01 - High Protein > 79 grams/day"
				If diet_msa_type_one = "02" Then diet_msa_type_one = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_one = "03" Then diet_msa_type_one = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_one = "04" Then diet_msa_type_one = "04 - Lo Cholesterol"
				If diet_msa_type_one = "05" Then diet_msa_type_one = "05 - High Residue"
				If diet_msa_type_one = "06" Then diet_msa_type_one = "06 - Pregnancy and Lactation"
				If diet_msa_type_one = "07" Then diet_msa_type_one = "07 - Gluten Free"
				If diet_msa_type_one = "08" Then diet_msa_type_one = "08 - Lactose Free"
				If diet_msa_type_one = "09" Then diet_msa_type_one = "09 - Anti-Dumping"
				If diet_msa_type_one = "10" Then diet_msa_type_one = "10 - Hypoglycemic"
				If diet_msa_type_one = "11" Then diet_msa_type_one = "11 - Ketogenic"
				If diet_msa_type_one = "__" Then diet_msa_type_one = ""

				If diet_msa_type_two = "01" Then diet_msa_type_two = "01 - High Protein > 79 grams/day"
				If diet_msa_type_two = "02" Then diet_msa_type_two = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_two = "03" Then diet_msa_type_two = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_two = "04" Then diet_msa_type_two = "04 - Lo Cholesterol"
				If diet_msa_type_two = "05" Then diet_msa_type_two = "05 - High Residue"
				If diet_msa_type_two = "06" Then diet_msa_type_two = "06 - Pregnancy and Lactation"
				If diet_msa_type_two = "07" Then diet_msa_type_two = "07 - Gluten Free"
				If diet_msa_type_two = "08" Then diet_msa_type_two = "08 - Lactose Free"
				If diet_msa_type_two = "09" Then diet_msa_type_two = "09 - Anti-Dumping"
				If diet_msa_type_two = "10" Then diet_msa_type_two = "10 - Hypoglycemic"
				If diet_msa_type_two = "11" Then diet_msa_type_two = "11 - Ketogenic"
				If diet_msa_type_two = "__" Then diet_msa_type_two = ""

				If diet_msa_type_three = "01" Then diet_msa_type_three = "01 - High Protein > 79 grams/day"
				If diet_msa_type_three = "02" Then diet_msa_type_three = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_three = "03" Then diet_msa_type_three = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_three = "04" Then diet_msa_type_three = "04 - Lo Cholesterol"
				If diet_msa_type_three = "05" Then diet_msa_type_three = "05 - High Residue"
				If diet_msa_type_three = "06" Then diet_msa_type_three = "06 - Pregnancy and Lactation"
				If diet_msa_type_three = "07" Then diet_msa_type_three = "07 - Gluten Free"
				If diet_msa_type_three = "08" Then diet_msa_type_three = "08 - Lactose Free"
				If diet_msa_type_three = "09" Then diet_msa_type_three = "09 - Anti-Dumping"
				If diet_msa_type_three = "10" Then diet_msa_type_three = "10 - Hypoglycemic"
				If diet_msa_type_three = "11" Then diet_msa_type_three = "11 - Ketogenic"
				If diet_msa_type_three = "__" Then diet_msa_type_three = ""

				If diet_msa_type_four = "01" Then diet_msa_type_four = "01 - High Protein > 79 grams/day"
				If diet_msa_type_four = "02" Then diet_msa_type_four = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_four = "03" Then diet_msa_type_four = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_four = "04" Then diet_msa_type_four = "04 - Lo Cholesterol"
				If diet_msa_type_four = "05" Then diet_msa_type_four = "05 - High Residue"
				If diet_msa_type_four = "06" Then diet_msa_type_four = "06 - Pregnancy and Lactation"
				If diet_msa_type_four = "07" Then diet_msa_type_four = "07 - Gluten Free"
				If diet_msa_type_four = "08" Then diet_msa_type_four = "08 - Lactose Free"
				If diet_msa_type_four = "09" Then diet_msa_type_four = "09 - Anti-Dumping"
				If diet_msa_type_four = "10" Then diet_msa_type_four = "10 - Hypoglycemic"
				If diet_msa_type_four = "11" Then diet_msa_type_four = "11 - Ketogenic"
				If diet_msa_type_four = "__" Then diet_msa_type_four = ""

				If diet_msa_type_five = "01" Then diet_msa_type_five = "01 - High Protein > 79 grams/day"
				If diet_msa_type_five = "02" Then diet_msa_type_five = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_five = "03" Then diet_msa_type_five = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_five = "04" Then diet_msa_type_five = "04 - Lo Cholesterol"
				If diet_msa_type_five = "05" Then diet_msa_type_five = "05 - High Residue"
				If diet_msa_type_five = "06" Then diet_msa_type_five = "06 - Pregnancy and Lactation"
				If diet_msa_type_five = "07" Then diet_msa_type_five = "07 - Gluten Free"
				If diet_msa_type_five = "08" Then diet_msa_type_five = "08 - Lactose Free"
				If diet_msa_type_five = "09" Then diet_msa_type_five = "09 - Anti-Dumping"
				If diet_msa_type_five = "10" Then diet_msa_type_five = "10 - Hypoglycemic"
				If diet_msa_type_five = "11" Then diet_msa_type_five = "11 - Ketogenic"
				If diet_msa_type_five = "__" Then diet_msa_type_five = ""

				If diet_msa_type_six = "01" Then diet_msa_type_six = "01 - High Protein > 79 grams/day"
				If diet_msa_type_six = "02" Then diet_msa_type_six = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_six = "03" Then diet_msa_type_six = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_six = "04" Then diet_msa_type_six = "04 - Lo Cholesterol"
				If diet_msa_type_six = "05" Then diet_msa_type_six = "05 - High Residue"
				If diet_msa_type_six = "06" Then diet_msa_type_six = "06 - Pregnancy and Lactation"
				If diet_msa_type_six = "07" Then diet_msa_type_six = "07 - Gluten Free"
				If diet_msa_type_six = "08" Then diet_msa_type_six = "08 - Lactose Free"
				If diet_msa_type_six = "09" Then diet_msa_type_six = "09 - Anti-Dumping"
				If diet_msa_type_six = "10" Then diet_msa_type_six = "10 - Hypoglycemic"
				If diet_msa_type_six = "11" Then diet_msa_type_six = "11 - Ketogenic"
				If diet_msa_type_six = "__" Then diet_msa_type_six = ""

				If diet_msa_type_seven = "01" Then diet_msa_type_seven = "01 - High Protein > 79 grams/day"
				If diet_msa_type_seven = "02" Then diet_msa_type_seven = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_seven = "03" Then diet_msa_type_seven = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_seven = "04" Then diet_msa_type_seven = "04 - Lo Cholesterol"
				If diet_msa_type_seven = "05" Then diet_msa_type_seven = "05 - High Residue"
				If diet_msa_type_seven = "06" Then diet_msa_type_seven = "06 - Pregnancy and Lactation"
				If diet_msa_type_seven = "07" Then diet_msa_type_seven = "07 - Gluten Free"
				If diet_msa_type_seven = "08" Then diet_msa_type_seven = "08 - Lactose Free"
				If diet_msa_type_seven = "09" Then diet_msa_type_seven = "09 - Anti-Dumping"
				If diet_msa_type_seven = "10" Then diet_msa_type_seven = "10 - Hypoglycemic"
				If diet_msa_type_seven = "11" Then diet_msa_type_seven = "11 - Ketogenic"
				If diet_msa_type_seven = "__" Then diet_msa_type_seven = ""

				If diet_msa_type_eight = "01" Then diet_msa_type_eight = "01 - High Protein > 79 grams/day"
				If diet_msa_type_eight = "02" Then diet_msa_type_eight = "02 - Control Protein 40-60 grams/day"
				If diet_msa_type_eight = "03" Then diet_msa_type_eight = "03 - Control Protein < 40 grams/day"
				If diet_msa_type_eight = "04" Then diet_msa_type_eight = "04 - Lo Cholesterol"
				If diet_msa_type_eight = "05" Then diet_msa_type_eight = "05 - High Residue"
				If diet_msa_type_eight = "06" Then diet_msa_type_eight = "06 - Pregnancy and Lactation"
				If diet_msa_type_eight = "07" Then diet_msa_type_eight = "07 - Gluten Free"
				If diet_msa_type_eight = "08" Then diet_msa_type_eight = "08 - Lactose Free"
				If diet_msa_type_eight = "09" Then diet_msa_type_eight = "09 - Anti-Dumping"
				If diet_msa_type_eight = "10" Then diet_msa_type_eight = "10 - Hypoglycemic"
				If diet_msa_type_eight = "11" Then diet_msa_type_eight = "11 - Ketogenic"
				If diet_msa_type_eight = "__" Then diet_msa_type_eight = ""

				If diet_mf_verif_one = "_" Then diet_mf_verif_one = ""
				If diet_mf_verif_two = "_" Then diet_mf_verif_two = ""
				If diet_msa_verif_one = "_" Then diet_msa_verif_one = ""
				If diet_msa_verif_two = "_" Then diet_msa_verif_two = ""
				If diet_msa_verif_three = "_" Then diet_msa_verif_three = ""
				If diet_msa_verif_four = "_" Then diet_msa_verif_four = ""
				If diet_msa_verif_five = "_" Then diet_msa_verif_five = ""
				If diet_msa_verif_six = "_" Then diet_msa_verif_six = ""
				If diet_msa_verif_seven = "_" Then diet_msa_verif_seven = ""
				If diet_msa_verif_eight	 = "_" Then diet_msa_verif_eight = ""
			End If
		End If
	end sub

	public sub collect_parent_information()

		If pare_exists = TRUE Then
			' MsgBox "PARE EXISTS for " & ref_number
			pare_row_index = 0
			Do
				For the_membs = 0 to UBound(HH_MEMB_ARRAY)
					' MsgBox "REF on PARE - " & pare_child_ref_nbr(pare_row_index) & vbCr & "REF of the HH MEMB - " & HH_MEMB_ARRAY(the_membs).ref_number
					If pare_child_ref_nbr(pare_row_index) = HH_MEMB_ARRAY(the_membs).ref_number Then
						pare_child_member_index(pare_array_count) = the_membs

						If HH_MEMB_ARRAY(the_membs).parent_one_name = "" Then

							HH_MEMB_ARRAY(the_membs).parent_one_name = full_name
							HH_MEMB_ARRAY(the_membs).parent_one_type = pare_relationship_type(pare_array_count)
							HH_MEMB_ARRAY(the_membs).parent_one_verif = pare_verification(pare_array_count)
							HH_MEMB_ARRAY(the_membs).parent_one_in_home = TRUE

						ElseIf HH_MEMB_ARRAY(the_membs).parent_two_name = "" Then
							HH_MEMB_ARRAY(the_membs).parent_two_name = full_name
							HH_MEMB_ARRAY(the_membs).parent_two_type = pare_relationship_type(pare_array_count)
							HH_MEMB_ARRAY(the_membs).parent_two_verif = pare_verification(pare_array_count)
							HH_MEMB_ARRAY(the_membs).parent_two_in_home = TRUE
						End If
						' MsgBox HH_MEMB_ARRAY(the_membs).parent_one_name

						Exit For
					End If
				Next
				pare_row_index = pare_row_index + 1
			Loop until pare_row_index > UBound(pare_child_ref_nbr)
		End If

		Call navigate_to_MAXIS_screen("STAT", "ABPS")
		Do
			abps_row = 15
			Do
				EMReadScreen abps_ref_nrb, 2, abps_row, 35
				' MsgBox "REF on ABPS - " & abps_ref_nrb & vbCr & "REF of the HH MEMB - " & ref_number
				If abps_ref_nrb = ref_number Then
					EMReadScreen abps_last_name, 24, 10, 30
					EMReadScreen abps_first_name, 12, 10, 63
					EMReadScreen abps_mid_initial, 1, 10, 80
					EMReadScreen abps_ssn, 11, 11, 30
					EMReadScreen abps_dob, 10, 11, 60
					EMReadScreen abps_gender, 1, 11, 80
					EMReadScreen abps_parental_status, 1, abps_row, 53
					EMReadScreen abps_custody, 1, abps_row, 67

					abps_last_name = replace(abps_last_name, "_", "")
					abps_first_name = replace(abps_first_name, "_", "")
					abps_mid_initial = replace(abps_mid_initial, "_", "")

					' MsgBox trim(abps_first_name) & " " & trim(abps_last_name)
					If abps_first_name = "" AND abps_last_name = "" Then abps_first_name = "Name Unknown"
					abps_ssn = replace(abps_ssn, "_", "")
					abps_ssn = trim(abps_ssn)
					abps_ssn = replace(abps_ssn, " ", "-")

					abps_dob = replace(abps_dob, "_", "")
					abps_dob = trim(abps_dob)
					abps_dob = replace(abps_dob, " ", "/")

					If parent_one_name = "" Then

						parent_one_name = trim(abps_first_name) & " " & trim(abps_last_name)
						parent_one_type = "ABSENT"
						parent_one_verif = ""
						parent_one_in_home = FALSE

					ElseIf parent_two_name = "" Then
						parent_two_name = trim(abps_first_name) & " " & trim(abps_last_name)
						parent_two_type = "ABSENT"
						parent_two_verif = ""
						parent_two_in_home = FALSE
					End If
				End If
				abps_row = abps_row + 1

				If abps_row = 18 Then
					PF20
					abps_row = 15
					EMReadScreen end_of_list, 9, 24, 14
					If end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until abps_ref_nrb = "__"
			transmit
			EMReadScreen last_abps, 7, 24, 2
		Loop until last_abps = "ENTER A"


	end sub

	Public sub choose_the_members()

	end sub

	' private sub Class_Initialize()
	' end sub
end class


class client_income

	'about the income
	public member_ref
	public member_name
	public member
	public access_denied

	public panel_name
	public panel_instance

	public unea_or_earned
	public income_type
	public income_type_code
	public income_review
	public income_verification
	public verif_explaination
	public income_start_date
	public income_end_date
	public pay_frequency
	public pay_weekday
	public hc_inc_est
	public most_recent_pay_date
	public most_recent_pay_amt
	public income_notes
	public pay_gross
	public expenses_allowed
	public expenses_not_allowed

	'JOBS
	public subsidized_income_type
	public hourly_wage
	public employer
	public prosp_pay_total
	public prosp_hours_total
	public prosp_pay_date_one
	public prosp_pay_wage_one
	public prosp_pay_date_two
	public prosp_pay_wage_two
	public prosp_pay_date_three
	public prosp_pay_wage_three
	public prosp_pay_date_four
	public prosp_pay_wage_four
	public prosp_pay_date_five
	public prosp_pay_wage_five
	public prosp_average_pay

	public retro_pay_total
	public retro_hours_total
	public retro_pay_date_one
	public retro_pay_wage_one
	public retro_pay_date_two
	public retro_pay_wage_two
	public retro_pay_date_three
	public retro_pay_wage_three
	public retro_pay_date_four
	public retro_pay_wage_four
	public retro_pay_date_five
	public retro_pay_wage_five
	public retro_average_pay

	'BUSI
	public prosp_net_cash_earnings
	public prosp_gross_cash_earnings
	public cash_earnings_verif
	public prosp_cash_expenses
	public cash_expense_verif
	public retro_net_cash_earnings
	public retro_gross_cash_earnings
	public retro_cash_expenses

	public prosp_net_ive_earnings
	public prosp_gross_ive_earnings
	public ive_earnings_verif
	public prosp_ive_expenses
	public ive_expense_verif

	public prosp_net_snap_earnings
	public prosp_gross_snap_earnings
	public snap_earnings_verif
	public prosp_snap_expenses
	public snap_expense_verif
	public retro_net_snap_earnings
	public retro_gross_snap_earnings
	public retro_snap_expenses

	public prosp_net_hc_a_earnings
	public prosp_gross_hc_a_earnings
	public hc_a_earnings_verif
	public prosp_hc_a_expenses
	public hc_a_expense_verif

	public prosp_net_hc_b_earnings
	public prosp_gross_hc_b_earnings
	public hc_b_earnings_verif
	public prosp_hc_b_expenses
	public hc_b_expense_verif

	public retro_reptd_hours
	public retro_min_wage_hours
	public prosp_reptd_hours
	public prosp_min_wage_hours

	public self_emp_method
	public self_emp_method_date

	'UNEA
	public claim_number
	public cola_month


	public sub read_member_name()
		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMWriteScreen member_ref, 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			last_name = "UNABLE TO FIND"
			first_name = "Access Denied"
			access_denied = TRUE
		Else
			access_denied = FALSE
			EMReadscreen last_name, 25, 6, 30
			EMReadscreen first_name, 12, 6, 63
		End If
		last_name = trim(replace(last_name, "_", ""))
		first_name = trim(replace(first_name, "_", ""))

		member_name = first_name & " " & last_name
		member = member_ref & " - " & member_name
		' MsgBox "~" & member & "~"
	end sub

	Public sub read_jobs_panel()
		jobs_found = FALSE
	end sub

	Public sub read_busi_panel()
		busi_found = FALSE
	end sub

	Public sub read_unea_panel()
		Call navigate_to_MAXIS_screen("STAT", "UNEA")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

		panel_name = "UNEA"
		unea_or_earned = "Unearned"

		EMReadScreen income_type, 2, 5, 37
		EMReadScreen income_verification, 1, 5, 65
		EMReadScreen income_start_date, 8, 7, 37
		EMReadScreen income_end_date, 8, 7, 68

		EmWriteScreen "X", 6, 56
		transmit
			EMReadScreen pay_frequency, 1, 10, 63
			EMReadScreen hc_inc_est, 8, 9, 65
		PF3

		EMReadScreen claim_number, 15, 6, 37
		EMReadScreen cola_month, 2, 19, 36

		EMReadScreen prosp_pay_total, 8, 18, 68
		EMReadScreen prosp_pay_date_one, 8, 13, 54
		EMReadScreen prosp_pay_wage_one, 8, 13, 68
		EMReadScreen prosp_pay_date_two, 8, 14, 54
		EMReadScreen prosp_pay_wage_two, 8, 14, 68
		EMReadScreen prosp_pay_date_three, 8, 15, 54
		EMReadScreen prosp_pay_wage_three, 8, 15, 68
		EMReadScreen prosp_pay_date_four, 8, 16, 54
		EMReadScreen prosp_pay_wage_four, 8, 16, 68
		EMReadScreen prosp_pay_date_five, 8, 17, 54
		EMReadScreen prosp_pay_wage_five, 8, 17, 68

		EMReadScreen retro_pay_total, 8, 18, 39
		EMReadScreen retro_pay_date_one, 8, 13, 25
		EMReadScreen retro_pay_wage_one, 8, 13, 39
		EMReadScreen retro_pay_date_two, 8, 14, 25
		EMReadScreen retro_pay_wage_two, 8, 14, 39
		EMReadScreen retro_pay_date_three, 8, 15, 25
		EMReadScreen retro_pay_wage_three, 8, 15, 39
		EMReadScreen retro_pay_date_four, 8, 16, 25
		EMReadScreen retro_pay_wage_four, 8, 16, 39
		EMReadScreen retro_pay_date_five, 8, 17, 25
		EMReadScreen retro_pay_wage_five, 8, 17, 39

		income_type_code = income_type
		If income_type = "01" Then income_type = "01 - RSDI, Disa"
		If income_type = "02" Then income_type = "02 - RSDI, No Disa"
		If income_type = "03" Then income_type = "03 - SSI"
		If income_type = "06" Then income_type = "06 - Non-MN PA"
		If income_type = "11" Then income_type = "11 - VA Disability Benefit"
		If income_type = "12" Then income_type = "12 - VA Pension"
		If income_type = "13" Then income_type = "13 - VA Other"
		If income_type = "38" Then income_type = "38 - VA Aid and Attendance"
		If income_type = "14" Then income_type = "14 - Unemployment Insurance"
		If income_type = "15" Then income_type = "15 - Worker's Compensation"
		If income_type = "16" Then income_type = "16 - Railroad Retirement"
		If income_type = "17" Then income_type = "17 - Other Retirement"
		If income_type = "18" Then income_type = "18 - Military Entitlement"
		If income_type = "19" Then income_type = "19 - FC Child Requesting SNAP"
		If income_type = "20" Then income_type = "20 - FC Child NOT Requesting SNAP"
		If income_type = "21" Then income_type = "21 - FC Adult Requesting SNAP"
		If income_type = "22" Then income_type = "22 - FC Adult NOT Requesting SNAP"
		If income_type = "23" Then income_type = "23 - Dividends"
		If income_type = "24" Then income_type = "24 - Interest "
		If income_type = "25" Then income_type = "25 - Counted Gifts or Prizes"
		If income_type = "26" Then income_type = "26 - Strike Benefit"
		If income_type = "27" Then income_type = "27 - Contract for Deed"
		If income_type = "28" Then income_type = "28 - Illegal Income"
		If income_type = "29" Then income_type = "29 - Other Countable"
		If income_type = "30" Then income_type = "30 - Not Counted - Infreq <30"
		If income_type = "21" Then income_type = "31 - Other SNAP Only"
		If income_type = "08" Then income_type = "08 - Direct Child Support"
		If income_type = "35" Then income_type = "35 - Direct Spousal Support"
		If income_type = "36" Then income_type = "36 - Disb Child Support"
		If income_type = "37" Then income_type = "37 - Disb Spousal Support"
		If income_type = "39" Then income_type = "39 - Disb Child Support Arrears"
		If income_type = "40" Then income_type = "40 - Disb Spousal Support Arrears"
		If income_type = "43" Then income_type = "43 - Disb Excess Child Support"
		If income_type = "44" Then income_type = "44 - MSA - Excess Income for SSI"
		If income_type = "45" Then income_type = "45 - County 88 Child Support"
		If income_type = "46" Then income_type = "46 - County 88 Gaming"
		If income_type = "47" Then income_type = "47 - Counted Tribal Income"
		If income_type = "48" Then income_type = "48 - Trust Income"
		If income_type = "49" Then income_type = "49 - Non-Recurring > $60/qtr"

		If income_verification = "1" Then income_verification = "1 - Copy of Checks"
		If income_verification = "2" Then income_verification = "2 - Award Letters"
		If income_verification = "3" Then income_verification = "3 - System Initiated"
		If income_verification = "4" Then income_verification = "4 - Colateral Statement"
		If income_verification = "5" Then income_verification = "5 - Pend Out State Verif"
		If income_verification = "6" Then income_verification = "6 - Other Document"
		If income_verification = "7" Then income_verification = "7 - Worker Initiated"
		If income_verification = "8" Then income_verification = "8 - RI Stubs"
		If income_verification = "N" Then income_verification = "N - No Verif Provided"
		' MsgBox "~" & income_verification & "~"
		income_start_date = replace(income_start_date, " ", "/")
		If income_start_date = "__/__/__" Then income_start_date = ""
		income_end_date = replace(income_end_date, " ", "/")
		If income_end_date = "__/__/__" Then income_end_date = ""

		If pay_frequency = "1" Then pay_frequency = "1 - Monthly"
		If pay_frequency = "2" Then pay_frequency = "2 - Semi-monthly"
		If pay_frequency = "3" Then pay_frequency = "3 - Biweekly"
		If pay_frequency = "4" Then pay_frequency = "4 - Weekly"
		If pay_frequency = "5" Then pay_frequency = "5 - Other"
		If pay_frequency = "_" Then pay_frequency = ""
		hc_inc_est = trim(hc_inc_est)

		'pay_weekday'

		claim_number = replace(claim_number, "_", "")

		If cola_month = "01" Then cola_month = "January"
		If cola_month = "02" Then cola_month = "February"
		If cola_month = "03" Then cola_month = "March"
		If cola_month = "04" Then cola_month = "April"
		If cola_month = "05" Then cola_month = "May"
		If cola_month = "06" Then cola_month = "June"
		If cola_month = "07" Then cola_month = "July"
		If cola_month = "08" Then cola_month = "August"
		If cola_month = "09" Then cola_month = "September"
		If cola_month = "10" Then cola_month = "October"
		If cola_month = "11" Then cola_month = "November"
		If cola_month = "12" Then cola_month = "December"
		If cola_month = "NA" Then cola_month = "Not Applicable"
		If cola_month = "__" Then cola_month = "Unspecified"

		prosp_pay_total = trim(prosp_pay_total)
		prosp_pay_date_one = replace(prosp_pay_date_one, " ", "/")
		If prosp_pay_date_one = "__/__/__" Then prosp_pay_date_one = ""
		prosp_pay_wage_one = trim(prosp_pay_wage_one)
		If prosp_pay_wage_one = "________" Then prosp_pay_wage_one = ""
		prosp_pay_date_two = replace(prosp_pay_date_two, " ", "/")
		If prosp_pay_date_two = "__/__/__" Then prosp_pay_date_two = ""
		prosp_pay_wage_two = trim(prosp_pay_wage_two)
		If prosp_pay_wage_two = "________" Then prosp_pay_wage_two = ""
		prosp_pay_date_three = replace(prosp_pay_date_three, " ", "/")
		If prosp_pay_date_three = "__/__/__" Then prosp_pay_date_three = ""
		prosp_pay_wage_three = trim(prosp_pay_wage_three)
		If prosp_pay_wage_three = "________" Then prosp_pay_wage_three = ""
		prosp_pay_date_four = replace(prosp_pay_date_four, " ", "/")
		If prosp_pay_date_four = "__/__/__" Then prosp_pay_date_four = ""
		prosp_pay_wage_four = trim(prosp_pay_wage_four)
		If prosp_pay_wage_four = "________" Then prosp_pay_wage_four = ""
		prosp_pay_date_five = replace(prosp_pay_date_five, " ", "/")
		If prosp_pay_date_five = "__/__/__" Then prosp_pay_date_five = ""
		prosp_pay_wage_five = trim(prosp_pay_wage_five)
		If prosp_pay_wage_five = "________" Then prosp_pay_wage_five = ""
		total_of_prosp_pay = 0
		number_of_checks = 0
		If prosp_pay_wage_one <> "" Then
			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_one * 1
			number_of_checks = number_of_checks + 1
		End If
		If prosp_pay_wage_two <> "" Then
			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_two * 1
			number_of_checks = number_of_checks + 1
		End If
		If prosp_pay_wage_three <> "" Then
			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_three * 1
			number_of_checks = number_of_checks + 1
		End If
		If prosp_pay_wage_four <> "" Then
			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_four * 1
			number_of_checks = number_of_checks + 1
		End If
		If prosp_pay_wage_five <> "" Then
			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_five * 1
			number_of_checks = number_of_checks + 1
		End If
		If number_of_checks <> 0 Then prosp_average_pay = total_of_prosp_pay / number_of_checks
		prosp_average_pay = prosp_average_pay & ""

		retro_pay_total = trim(retro_pay_total)
		retro_pay_date_one = replace(retro_pay_date_one, " ", "/")
		If retro_pay_date_one = "__/__/__" Then retro_pay_date_one = ""
		retro_pay_wage_one = trim(retro_pay_wage_one)
		If retro_pay_wage_one = "________" Then retro_pay_wage_one = ""
		retro_pay_date_two = replace(retro_pay_date_two, " ", "/")
		If retro_pay_date_two = "__/__/__" Then retro_pay_date_two = ""
		retro_pay_wage_two = trim(retro_pay_wage_two)
		If retro_pay_wage_two = "________" Then retro_pay_wage_two = ""
		retro_pay_date_three = replace(retro_pay_date_three, " ", "/")
		If retro_pay_date_three = "__/__/__" Then retro_pay_date_three = ""
		retro_pay_wage_three = trim(retro_pay_wage_three)
		If retro_pay_wage_three = "________" Then retro_pay_wage_three = ""
		retro_pay_date_four = replace(retro_pay_date_four, " ", "/")
		If retro_pay_date_four = "__/__/__" Then retro_pay_date_four = ""
		retro_pay_wage_four = trim(retro_pay_wage_four)
		If retro_pay_wage_four = "________" Then retro_pay_wage_four = ""
		retro_pay_date_five = replace(retro_pay_date_five, " ", "/")
		If retro_pay_date_five = "__/__/__" Then retro_pay_date_five = ""
		retro_pay_wage_five = trim(retro_pay_wage_five)
		If retro_pay_wage_five = "________" Then retro_pay_wage_five = ""
		total_of_retro_pay = 0
		number_of_checks = 0
		If retro_pay_wage_one <> "" Then
			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_one * 1
			number_of_checks = number_of_checks + 1
		End If
		If retro_pay_wage_two <> "" Then
			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_two * 1
			number_of_checks = number_of_checks + 1
		End If
		If retro_pay_wage_three <> "" Then
			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_three * 1
			number_of_checks = number_of_checks + 1
		End If
		If retro_pay_wage_four <> "" Then
			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_four * 1
			number_of_checks = number_of_checks + 1
		End If
		If retro_pay_wage_five <> "" Then
			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_five * 1
			number_of_checks = number_of_checks + 1
		End If
		If number_of_checks <> 0 Then retro_average_pay = total_of_retro_pay / number_of_checks
		retro_average_pay = retro_average_pay & ""

		If pay_frequency = "3 - Biweekly" OR pay_frequency = "4 - Weekly" Then
			If prosp_pay_date_five <> "" Then
				pay_weekday = WeekdayName(weekday(prosp_pay_date_five))
			ElseIf prosp_pay_date_four <> "" Then
				pay_weekday = WeekdayName(weekday(prosp_pay_date_four))
			ElseIf prosp_pay_date_three <> "" Then
				pay_weekday = WeekdayName(weekday(prosp_pay_date_three))
			ElseIf prosp_pay_date_two <> "" Then
				pay_weekday = WeekdayName(weekday(prosp_pay_date_two))
			ElseIf prosp_pay_date_one <> "" Then
				pay_weekday = WeekdayName(weekday(prosp_pay_date_one))
			End If

		End If

	end sub

end class

class client_assets

	public member_ref
	public member_name
	public member
	public access_denied

	public panel_name
	public panel_instance

	public asset_btn_one
	public asset_type
	public account_number
	public asset_verification
	public asset_update_date
	public withdraw_yn
	public withdraw_penalty
	public withdraw_verif
	public count_cash_yn
	public count_snap_yn
	public count_hc_yn
	public count_grh_yn
	public count_ive_yn
	public joint_owners_yn
	public share_ratio
	public next_interest_date

	public cash_value

	public acct_location
	public acct_balance
	public acct_balance_date

	public cars_year
	public cars_make
	public cars_model
	public cars_trade_in_value
	public cars_loan_value
	public cars_value_source
	public cars_amt_owed
	public cars_owed_verification
	public cars_owed_date
	public cars_use
	public cars_hc_benefit

	public secu_name
	public secu_cash_value
	public secu_cash_value_date
	public secu_face_value

	public rest_market_value
	public rest_value_verification
	public rest_amount_owed
	public rest_owed_verification
	public rest_owed_date
	public rest_property_status
	public rest_ive_repayment_agreement_date

	' function access_ACCT_panel(access_type, member_name,

	' account_type,
	' account_number,
	' account_location,
	' account_balance,
	' account_verification,
	' update_date, panel_ref_numb,
	' balance_date,
	' withdraw_penalty,
	' withdraw_yn,
	' withdraw_verif_code,
	' count_cash,
	' count_snap,
	' count_hc,
	' count_grh,
	' count_ive,
	' joint_own_yn,
	' share_ratio,
	' next_interest)

	' function access_CARS_panel(access_type, member_name,

	' cars_type,
	' cars_year,
	' cars_make,
	' cars_model,
	' cars_verif,
	' update_date, panel_ref_numb,
	' cars_trade_in,
	' cars_loan,
	' cars_source,
	' cars_owed,
	' cars_owed_verif_code,
	' cars_owed_date,
	' cars_use,
	' cars_hc_benefit,
	' cars_joint_yn,
	' cars_share)

	' function access_SECU_panel(access_type, member_name,

	' security_type,
	' security_account_number,
	' security_name,
	' security_cash_value,
	' security_verif,
	' secu_update_date,
	' panel_ref_numb,
	' security_face_value,
	' security_withdraw,
	' security_withdraw_yn,
	' security_withdraw_verif,
	' secu_cash_yn,
	' secu_snap_yn,
	' secu_hc_yn,
	' secu_grh_yn,
	' secu_ive_yn,
	' secu_joint,
	' secu_ratio,
	' security_eff_date)

	' function access_REST_panel(access_type, member_name,

	' rest_type,
	' rest_verif,
	' rest_update_date,
	' panel_ref_numb,
	' rest_market_value,
	' value_verif_code,
	' rest_amt_owed,
	' amt_owed_verif_code,
	' rest_eff_date,
	' rest_status,
	' rest_joint_yn,
	' rest_ratio,
	' repymt_agree_date)

	public sub read_member_name()
		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMWriteScreen member_ref, 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			last_name = "UNABLE TO FIND"
			first_name = "Access Denied"
			access_denied = TRUE
		Else
			access_denied = FALSE
			EMReadscreen last_name, 25, 6, 30
			EMReadscreen first_name, 12, 6, 63
		End If
		last_name = trim(replace(last_name, "_", ""))
		first_name = trim(replace(first_name, "_", ""))

		member_name = first_name & " " & last_name
		member = member_ref & " - " & member_name
		' MsgBox "~" & member & "~"
	end sub

	public sub read_cash_panel()
		Call navigate_to_MAXIS_screen("STAT", "CASH")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

		asset_type = "CASH"

		EMReadScreen cash_value, 8, 8, 39
		cash_value = trim(cash_value)

	end sub

	public sub read_acct_panel()

		Call navigate_to_MAXIS_screen("STAT", "ACCT")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

		EMReadScreen panel_type, 2, 6, 44
		EMReadScreen panel_number, 20, 7, 44
		EMReadScreen acct_location, 20, 8, 44
		EMReadScreen panel_balance, 8, 10, 46
		EMReadScreen panel_verif_code, 1, 10, 64
		EMReadScreen balance_date, 8, 11, 44
		EMReadScreen withdraw_penalty, 8, 12, 46
		EMReadScreen withdraw_yn, 1, 12, 64
		EMReadScreen withdraw_verif_code, 1, 12, 72
		EMReadScreen count_cash, 1, 14, 50
		EMReadScreen count_snap, 1, 14, 57
		EMReadScreen count_hc, 1, 14, 64
		EMReadScreen count_grh, 1, 14, 72
		EMReadScreen count_ive, 1, 14, 80
		EMReadScreen joint_own_yn, 1, 15, 44
		EMReadScreen share_ratio, 5, 15, 76
		EMReadScreen next_interest, 5, 17, 57
		EMReadScreen update_date, 8, 21, 55

		If panel_type = "SV" Then asset_type = "SV - Savings"
		If panel_type = "CK" Then asset_type = "CK - Checking"
		If panel_type = "CE" Then asset_type = "CE - Certificate of Deposit"
		If panel_type = "MM" Then asset_type = "MM - Money Market"
		If panel_type = "DC" Then asset_type = "DC - Debit Card"
		If panel_type = "KO" Then asset_type = "KO - Keogh Account"
		If panel_type = "FT" Then asset_type = "FT - Fed Thrift Savings Plan"
		If panel_type = "SL" Then asset_type = "SL - State & Local Govt"
		If panel_type = "RA" Then asset_type = "RA - Employee Ret Annuities"
		If panel_type = "NP" Then asset_type = "NP - Non-Profit Emmployee Ret"
		If panel_type = "IR" Then asset_type = "IR - Indiv Ret Acct"
		If panel_type = "RH" Then asset_type = "RH - Roth IRA"
		If panel_type = "FR" Then asset_type = "FR - Ret Plan for Employers"
		If panel_type = "CT" Then asset_type = "CT - Corp Ret Trust"
		If panel_type = "RT" Then asset_type = "RT - Other Ret Fund"
		If panel_type = "QT" Then asset_type = "QT - Qualified Tuition (529)"
		If panel_type = "CA" Then asset_type = "CA - Coverdell SV (530)"
		If panel_type = "OE" Then asset_type = "OE - Other Educational"
		If panel_type = "OT" Then asset_type = "OT - Other"

		account_number = replace(panel_number, "_", "")
		acct_location =  replace(acct_location, "_", "")
		acct_balance = trim(panel_balance)

		If panel_verif_code = "1"  Then asset_verification = "1 - Bank Statement"
		If panel_verif_code = "2"  Then asset_verification = "2 - Agcy Ver Form"
		If panel_verif_code = "3"  Then asset_verification = "3 - Coltrl Contact"
		If panel_verif_code = "5"  Then asset_verification = "5 - Other Document"
		If panel_verif_code = "6"  Then asset_verification = "6 - Personal Statement"
		If panel_verif_code = "N"  Then asset_verification = "N - No Ver Prvd"

		acct_balance_date = replace(balance_date, " ", "/")
		If acct_balance_date = "__/__/__" Then acct_balance_date = ""

		withdraw_penalty = replace(withdraw_penalty, "_", "")
		withdraw_penalty = trim(withdraw_penalty)
		withdraw_yn = replace(withdraw_yn, "_", "")
		If withdraw_verif_code = "1"  Then withdraw_verif = "1 - Bank Statement"
		If withdraw_verif_code = "2"  Then withdraw_verif = "2 - Agcy Ver Form"
		If withdraw_verif_code = "3"  Then withdraw_verif = "3 - Coltrl Contact"
		If withdraw_verif_code = "5"  Then withdraw_verif = "5 - Other Document"
		If withdraw_verif_code = "6"  Then withdraw_verif = "6 - Personal Statement"
		If withdraw_verif_code = "N"  Then withdraw_verif = "N - No Ver Prvd"

		count_cash_yn = replace(count_cash, "_", "")
		count_snap_yn = replace(count_snap, "_", "")
		count_hc_yn = replace(count_hc, "_", "")
		count_grh_yn = replace(count_grh, "_", "")
		count_ive_yn = replace(count_ive, "_", "")

		share_ratio = replace(share_ratio, " ", "")

		next_interest_date = replace(next_interest, " ", "/")
		If next_interest_date = "__/__" Then next_interest_date = ""

		asset_update_date = replace(update_date, " ", "/")

	end sub

	public sub read_secu_panel()
		Call navigate_to_MAXIS_screen("STAT", "SECU")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

        EMReadScreen panel_type, 2, 6, 50
        EMReadScreen security_account_number, 12, 7, 50
        EMReadScreen security_name, 20, 8, 50
        EMReadScreen security_cash_value, 8, 10, 52
        EMReadScreen security_eff_date, 8, 11, 35   'not output
        EMReadScreen verif_code, 1, 11, 50
        EMReadScreen security_face_value, 8, 12, 52     'not output
        EMReadScreen security_withdraw, 8, 13, 52       'not output
        EMReadScreen security_withdraw_yn, 1, 13, 72    'not output
        EMReadScreen security_withdraw_verif, 1, 13, 80 'not output

        EMReadScreen secu_cash_yn, 1, 15, 50    'not output
        EMReadScreen secu_snap_yn, 1, 15, 57    'not output
        EMReadScreen secu_hc_yn, 1, 15, 64      'not output
        EMReadScreen secu_grh_yn, 1, 15, 72     'not output
        EMReadScreen secu_ive_yn, 1, 15, 80     'not output

        EMReadScreen secu_joint, 1, 16, 44      'not output
        EMReadScreen secu_ratio, 5, 16, 76      'not output
        EMReadScreen secu_update_date, 8, 21, 55

        If panel_type = "LI" Then asset_type = "LI - Life Insurance"
        If panel_type = "ST" Then asset_type = "ST - Stocks"
        If panel_type = "BO" Then asset_type = "BO - Bonds"
        If panel_type = "CD" Then asset_type = "CD - Ctrct for Deed"
        If panel_type = "MO" Then asset_type = "MO - Mortgage Note"
        If panel_type = "AN" Then asset_type = "AN - Annuity"
        If panel_type = "OT" Then asset_type = "OT - Other"

        account_number = replace(security_account_number, "_", "")
        secu_name = replace(security_name, "_", "")

        secu_cash_value = replace(security_cash_value, "_", "")
        secu_cash_value = trim(secu_cash_value)

        secu_cash_value_date = replace(security_eff_date, " ", "/")
        If secu_cash_value_date = "__/__/__" Then secu_cash_value_date = ""

        If verif_code = "1" Then asset_verification = "1 - Agency Form"
        If verif_code = "2" Then asset_verification = "2 - Source Doc"
        If verif_code = "3" Then asset_verification = "3 - Phone Contact"
        If verif_code = "5" Then asset_verification = "5 - Other Document"
        If verif_code = "6" Then asset_verification = "6 - Personal Statement"
        If verif_code = "N" Then asset_verification = "N - No Ver Prov"

        secu_face_value = replace(security_face_value, "_", "")
        secu_face_value = trim(secu_face_value)

        withdraw_penalty = replace(security_withdraw, "_", "")
        withdraw_penalty = trim(withdraw_penalty)

        withdraw_yn = replace(security_withdraw_yn, "_", "")

        If security_withdraw_verif = "1" Then withdraw_verif = "1 - Agency Form"
        If security_withdraw_verif = "2" Then withdraw_verif = "2 - Source Doc"
        If security_withdraw_verif = "3" Then withdraw_verif = "3 - Phone Contact"
        If security_withdraw_verif = "4" Then withdraw_verif = "4 - Other Document"
        If security_withdraw_verif = "5" Then withdraw_verif = "5 - Personal Stmt"
        If security_withdraw_verif = "N" Then withdraw_verif = "N - No Ver Prov"

        count_cash_yn = replace(secu_cash_yn, "_", "")
        count_snap_yn = replace(secu_snap_yn, "_", "")
        count_hc_yn = replace(secu_hc_yn, "_", "")
        count_grh_yn = replace(secu_grh_yn, "_", "")
        count_ive_yn = replace(secu_ive_yn, "_", "")

        joint_owners_yn = replace(secu_joint, "_", "")
        share_ratio = replace(secu_ratio, " ", "")

        asset_update_date = replace(secu_update_date, " ", "/")

	end sub

	public sub read_cars_panel()
		Call navigate_to_MAXIS_screen("STAT", "CARS")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

		EMReadScreen cars_type, 1, 6, 43
		EMReadScreen cars_year, 4, 8, 31
		EMReadScreen cars_make, 15, 8, 43
		EMReadScreen cars_model, 15, 8, 66
		EMReadScreen cars_trade_in, 8, 9, 45            'not output
		EMReadScreen cars_loan, 8, 9, 62                'not output
		EMReadScreen cars_source, 1, 9, 80              'not output
		EMReadScreen cars_verif_code, 1, 10, 60
		EMReadScreen cars_owed, 8, 12, 45               'not output
		EMReadScreen cars_owed_verif_code, 1, 12, 60    'not output
		EMReadScreen cars_owed_date, 8, 13, 43          'not output
		EMReadScreen cars_use, 1, 15, 43                'not output
		EMReadScreen cars_hc_benefit, 1, 15, 76         'not output
		EMReadScreen cars_joint_yn, 1, 16, 43           'not output
		EMReadScreen cars_share, 5, 16, 76              'not output
		EMReadScreen cars_update, 8, 21, 55

		If cars_type = "1" Then asset_type = "1 - Car"
		If cars_type = "2" Then asset_type = "2 - Truck"
		If cars_type = "3" Then asset_type = "3 - Van"
		If cars_type = "4" Then asset_type = "4 - Camper"
		If cars_type = "5" Then asset_type = "5 - Motorcycle"
		If cars_type = "6" Then asset_type = "6 - Trailer"
		If cars_type = "7" Then asset_type = "7 - Other"

		cars_make = replace(cars_make, "_", "")
		cars_model = replace(cars_model, "_", "")

		cars_trade_in_value = replace(cars_trade_in, "_", "")
		cars_trade_in_value = trim(cars_trade_in_value)

		cars_loan_value = replace(cars_loan, "_", "")
		cars_loan_value = trim(cars_loan_value)

		If cars_source = "1" Then cars_value_source = "1 - NADA"
		If cars_source = "2" Then cars_value_source = "2 - Appraisal Val"
		If cars_source = "3" Then cars_value_source = "3 - Client Stmt"
		If cars_source = "4" Then cars_value_source = "4 - Other Document"

		If cars_verif_code = "1" Then asset_verification = "1 - Title"
		If cars_verif_code = "2" Then asset_verification = "2 - License Reg"
		If cars_verif_code = "3" Then asset_verification = "3 - DMV"
		If cars_verif_code = "4" Then asset_verification = "4 - Purchase Agmt"
		If cars_verif_code = "5" Then asset_verification = "5 - Other Document"
		If cars_verif_code = "N" Then asset_verification = "N - No Ver Prvd"

		cars_amt_owed = replace(cars_owed, "_", "")
		cars_amt_owed = trim(cars_amt_owed)

		If cars_owed_verif_code = "1" Then cars_owed_verification = "1 - Bank/Lending Inst Stmt"
		If cars_owed_verif_code = "2" Then cars_owed_verification = "2 - Private Lender Stmt"
		If cars_owed_verif_code = "3" Then cars_owed_verification = "3 - Other Document"
		If cars_owed_verif_code = "4" Then cars_owed_verification = "4 - Pend Out State Verif"
		If cars_owed_verif_code = "N" Then cars_owed_verification = "N - No Ver Prvd"

		cars_owed_date = replace(cars_owed_date, " ", "/")
		If cars_owed_date = "__/__/__" Then cars_owed_date = ""

		If cars_use = "1" Then cars_use = "1 - Primary Vehicle"
		If cars_use = "2" Then cars_use = "2 - Employment/Training Search"
		If cars_use = "3" Then cars_use = "3 - Disa Transportation"
		If cars_use = "4" Then cars_use = "4 - Income Producing"
		If cars_use = "5" Then cars_use = "5 - Used as Home"
		If cars_use = "7" Then cars_use = "7 - Unlicensed"
		If cars_use = "8" Then cars_use = "8 - Other Countable"
		If cars_use = "9" Then cars_use = "9 - Unavailable"
		If cars_use = "0" Then cars_use = "0 - Long Distance Employment Travel"
		If cars_use = "A" Then cars_use = "A - Carry Heating Fuel or Water"

		cars_hc_benefit = replace(cars_hc_benefit, "_", "")
		joint_owners_yn = replace(cars_joint_yn, "_", "")
		share_ratio = replace(cars_share, " ", "")

		asset_update_date = replace(cars_update, " ", "/")

	end sub

	public sub read_rest_panel()
		Call navigate_to_MAXIS_screen("STAT", "REST")
		EMWriteScreen member_ref, 20, 76
		EMWriteScreen panel_instance, 20, 79
		transmit

        EMReadScreen type_code, 1, 6, 39
        EMReadScreen type_verif_code, 2, 6, 62
        EMReadScreen rest_market_value, 10, 8, 41
        EMReadScreen value_verif_code, 2, 8, 62
        EMReadScreen rest_amt_owed, 10, 9, 41
        EMReadScreen amt_owed_verif_code, 2, 9, 62
        EMReadScreen rest_eff_date, 8, 10, 39
        EMReadScreen rest_status, 1, 12, 54
        EMReadScreen rest_joint_yn, 1, 13, 54
        EMReadScreen rest_ratio, 5, 14, 54
        EMReadScreen repymt_agree_date, 8, 16, 62
        EMReadScreen rest_update_date, 8, 21, 55

        If type_code = "1" Then asset_type = "1 - House"
        If type_code = "2" Then asset_type = "2 - Land"
        If type_code = "3" Then asset_type = "3 - Buildings"
        If type_code = "4" Then asset_type = "4 - Mobile Home"
        If type_code = "5" Then asset_type = "5 - Life Estate"
        If type_code = "6" Then asset_type = "6 - Other"

        If type_verif_code = "TX" Then asset_verification = "TX - Property Tax Statement"
        If type_verif_code = "PU" Then asset_verification = "PU - Purchase Agreement"
        If type_verif_code = "TI" Then asset_verification = "TI - Title/Deed"
        If type_verif_code = "CD" Then asset_verification = "CD - Contract for Deed"
        If type_verif_code = "CO" Then asset_verification = "CO - County Record"
        If type_verif_code = "OT" Then asset_verification = "OT - Other Document"
        If type_verif_code = "NO" Then asset_verification = "NO - No Ver Prvd"

        rest_market_value = replace(rest_market_value, "_", "")
        rest_market_value = trim(rest_market_value)

        If value_verif_code = "TX" Then rest_value_verification = "TX - Property Tax Statement"
        If value_verif_code = "PU" Then rest_value_verification = "PU - Purchase Agreement"
        If value_verif_code = "AP" Then rest_value_verification = "AP - Appraisal"
        If value_verif_code = "CO" Then rest_value_verification = "CO - County Record"
        If value_verif_code = "OT" Then rest_value_verification = "OT - Other Document"
        If value_verif_code = "NO" Then rest_value_verification = "NO - No Ver Prvd"

        rest_amount_owed = replace(rest_amt_owed, "_", "")
        rest_amount_owed = trim(rest_amount_owed)

        If amt_owed_verif_code = "MO" Then rest_owed_verification = "TI - Title/Deed"
        If amt_owed_verif_code = "LN" Then rest_owed_verification = "CD - Contract for Deed"
        If amt_owed_verif_code = "CD" Then rest_owed_verification = "CD - Contract for Deed"
        If amt_owed_verif_code = "OT" Then rest_owed_verification = "OT - Other Document"
        If amt_owed_verif_code = "NO" Then rest_owed_verification = "NO - No Ver Prvd"

        rest_owed_date = replace(rest_eff_date, " ", "/")
        If rest_owed_date = "__/__/__" Then rest_owed_date = ""

        If rest_status = "1" Then rest_property_status = "1 - Home Residence"
        If rest_status = "2" Then rest_property_status = "2 - For Sale, IV-E Rpymt Agmt"
        If rest_status = "3" Then rest_property_status = "3 - Joint Owner, Unavailable"
        If rest_status = "4" Then rest_property_status = "4 - Income Producing"
        If rest_status = "5" Then rest_property_status = "5 - Future Residence"
        If rest_status = "6" Then rest_property_status = "6 - Other"
        If rest_status = "7" Then rest_property_status = "7 - For Sale, Unavailable"

        joint_owners_yn = replace(rest_joint_yn, "_", "")
        share_ratio = replace(rest_ratio, "_", "")

        rest_ive_repayment_agreement_date = replace(repymt_agree_date, " ", "/")
        If rest_ive_repayment_agreement_date = "__/__/__" Then rest_ive_repayment_agreement_date = ""

        asset_update_date = replace(rest_update_date, " ", "/")

	end sub

end class


function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")

        EMReadScreen line_one, 22, 6, 43										'Reading all the information from the panel
        EMReadScreen line_two, 22, 7, 43
        EMReadScreen city_line, 15, 8, 43
        EMReadScreen state_line, 2, 8, 66
        EMReadScreen zip_line, 7, 9, 43
        EMReadScreen county_line, 2, 9, 66
        EMReadScreen verif_line, 2, 9, 74
        EMReadScreen homeless_line, 1, 10, 43
        EMReadScreen reservation_line, 1, 10, 74
        EMReadScreen living_sit_line, 2, 11, 43

        resi_line_one = replace(line_one, "_", "")								'This is all formatting of the information from the panel
        resi_line_two = replace(line_two, "_", "")
        resi_city = replace(city_line, "_", "")
        resi_zip = replace(zip_line, "_", "")

        If county_line = "01" Then addr_county = "01 - Aitkin"
        If county_line = "02" Then addr_county = "02 - Anoka"
        If county_line = "03" Then addr_county = "03 - Becker"
        If county_line = "04" Then addr_county = "04 - Beltrami"
        If county_line = "05" Then addr_county = "05 - Benton"
        If county_line = "06" Then addr_county = "06 - Big Stone"
        If county_line = "07" Then addr_county = "07 - Blue Earth"
        If county_line = "08" Then addr_county = "08 - Brown"
        If county_line = "09" Then addr_county = "09 - Carlton"
        If county_line = "10" Then addr_county = "10 - Carver"
        If county_line = "11" Then addr_county = "11 - Cass"
        If county_line = "12" Then addr_county = "12 - Chippewa"
        If county_line = "13" Then addr_county = "13 - Chisago"
        If county_line = "14" Then addr_county = "14 - Clay"
        If county_line = "15" Then addr_county = "15 - Clearwater"
        If county_line = "16" Then addr_county = "16 - Cook"
        If county_line = "17" Then addr_county = "17 - Cottonwood"
        If county_line = "18" Then addr_county = "18 - Crow Wing"
        If county_line = "19" Then addr_county = "19 - Dakota"
        If county_line = "20" Then addr_county = "20 - Dodge"
        If county_line = "21" Then addr_county = "21 - Douglas"
        If county_line = "22" Then addr_county = "22 - Faribault"
        If county_line = "23" Then addr_county = "23 - Fillmore"
        If county_line = "24" Then addr_county = "24 - Freeborn"
        If county_line = "25" Then addr_county = "25 - Goodhue"
        If county_line = "26" Then addr_county = "26 - Grant"
        If county_line = "27" Then addr_county = "27 - Hennepin"
        If county_line = "28" Then addr_county = "28 - Houston"
        If county_line = "29" Then addr_county = "29 - Hubbard"
        If county_line = "30" Then addr_county = "30 - Isanti"
        If county_line = "31" Then addr_county = "31 - Itasca"
        If county_line = "32" Then addr_county = "32 - Jackson"
        If county_line = "33" Then addr_county = "33 - Kanabec"
        If county_line = "34" Then addr_county = "34 - Kandiyohi"
        If county_line = "35" Then addr_county = "35 - Kittson"
        If county_line = "36" Then addr_county = "36 - Koochiching"
        If county_line = "37" Then addr_county = "37 - Lac Qui Parle"
        If county_line = "38" Then addr_county = "38 - Lake"
        If county_line = "39" Then addr_county = "39 - Lake Of Woods"
        If county_line = "40" Then addr_county = "40 - Le Sueur"
        If county_line = "41" Then addr_county = "41 - Lincoln"
        If county_line = "42" Then addr_county = "42 - Lyon"
        If county_line = "43" Then addr_county = "43 - Mcleod"
        If county_line = "44" Then addr_county = "44 - Mahnomen"
        If county_line = "45" Then addr_county = "45 - Marshall"
        If county_line = "46" Then addr_county = "46 - Martin"
        If county_line = "47" Then addr_county = "47 - Meeker"
        If county_line = "48" Then addr_county = "48 - Mille Lacs"
        If county_line = "49" Then addr_county = "49 - Morrison"
        If county_line = "50" Then addr_county = "50 - Mower"
        If county_line = "51" Then addr_county = "51 - Murray"
        If county_line = "52" Then addr_county = "52 - Nicollet"
        If county_line = "53" Then addr_county = "53 - Nobles"
        If county_line = "54" Then addr_county = "54 - Norman"
        If county_line = "55" Then addr_county = "55 - Olmsted"
        If county_line = "56" Then addr_county = "56 - Otter Tail"
        If county_line = "57" Then addr_county = "57 - Pennington"
        If county_line = "58" Then addr_county = "58 - Pine"
        If county_line = "59" Then addr_county = "59 - Pipestone"
        If county_line = "60" Then addr_county = "60 - Polk"
        If county_line = "61" Then addr_county = "61 - Pope"
        If county_line = "62" Then addr_county = "62 - Ramsey"
        If county_line = "63" Then addr_county = "63 - Red Lake"
        If county_line = "64" Then addr_county = "64 - Redwood"
        If county_line = "65" Then addr_county = "65 - Renville"
        If county_line = "66" Then addr_county = "66 - Rice"
        If county_line = "67" Then addr_county = "67 - Rock"
        If county_line = "68" Then addr_county = "68 - Roseau"
        If county_line = "69" Then addr_county = "69 - St. Louis"
        If county_line = "70" Then addr_county = "70 - Scott"
        If county_line = "71" Then addr_county = "71 - Sherburne"
        If county_line = "72" Then addr_county = "72 - Sibley"
        If county_line = "73" Then addr_county = "73 - Stearns"
        If county_line = "74" Then addr_county = "74 - Steele"
        If county_line = "75" Then addr_county = "75 - Stevens"
        If county_line = "76" Then addr_county = "76 - Swift"
        If county_line = "77" Then addr_county = "77 - Todd"
        If county_line = "78" Then addr_county = "78 - Traverse"
        If county_line = "79" Then addr_county = "79 - Wabasha"
        If county_line = "80" Then addr_county = "80 - Wadena"
        If county_line = "81" Then addr_county = "81 - Waseca"
        If county_line = "82" Then addr_county = "82 - Washington"
        If county_line = "83" Then addr_county = "83 - Watonwan"
        If county_line = "84" Then addr_county = "84 - Wilkin"
        If county_line = "85" Then addr_county = "85 - Winona"
        If county_line = "86" Then addr_county = "86 - Wright"
        If county_line = "87" Then addr_county = "87 - Yellow Medicine"
        If county_line = "89" Then addr_county = "89 - Out-of-State"
        resi_county = addr_county

		Call get_state_name_from_state_code(state_line, resi_state, TRUE)		'This function makes the state code to be the state name written out - including the code

        If homeless_line = "Y" Then addr_homeless = "Yes"
        If homeless_line = "N" Then addr_homeless = "No"
        If reservation_line = "Y" Then addr_reservation = "Yes"
        If reservation_line = "N" Then addr_reservation = "No"

        If verif_line = "SF" Then addr_verif = "SF - Shelter Form"
        If verif_line = "Co" Then addr_verif = "CO - Coltrl Stmt"
        If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
        If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
        If verif_line = "CD" Then addr_verif = "CD - Contrct for Deed"
        If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
        If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
        If verif_line = "OT" Then addr_verif = "OT - Other Document"
        If verif_line = "NO" Then addr_verif = "NO - No Ver Prvd"
        If verif_line = "?_" Then addr_verif = "? - Delayed"
        If verif_line = "__" Then addr_verif = "Blank"


        If living_sit_line = "__" Then living_situation = "Blank"
        If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roomate"
        If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
        If living_sit_line = "03" Then living_situation = "03 -  servc prvdr- foster/group home"
        If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
        If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
        If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
        If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
        If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
        If living_sit_line = "09" Then living_situation = "09 - Declined"
        If living_sit_line = "10" Then living_situation = "10 - Unknown"
        addr_living_sit = living_situation

        EMReadScreen addr_eff_date, 8, 4, 43									'reading the mail information
        EMReadScreen addr_future_date, 8, 4, 66
        EMReadScreen mail_line_one, 22, 13, 43
        EMReadScreen mail_line_two, 22, 14, 43
        EMReadScreen mail_city_line, 15, 15, 43
        EMReadScreen mail_state_line, 2, 16, 43
        EMReadScreen mail_zip_line, 7, 16, 52

        addr_eff_date = replace(addr_eff_date, " ", "/")						'cormatting the mail information
        addr_future_date = trim(addr_future_date)
        addr_future_date = replace(addr_future_date, " ", "/")
        mail_line_one = replace(mail_line_one, "_", "")
        mail_line_two = replace(mail_line_two, "_", "")
        mail_city = replace(mail_city_line, "_", "")
        mail_state = replace(mail_state_line, "_", "")
        mail_zip = replace(mail_zip_line, "_", "")

        notes_on_address = "Address effective: " & addr_eff_date & "."
        ' If mail_line_one <> "" Then
        '     If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        '     If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        ' End If
        If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."

        EMReadScreen phone_one, 14, 17, 45										'reading the phone information
        EMReadScreen phone_two, 14, 18, 45
        EMReadScreen phone_three, 14, 19, 45

        EMReadScreen type_one, 1, 17, 67
        EMReadScreen type_two, 1, 18, 67
        EMReadScreen type_three, 1, 19, 67

        phone_one = replace(phone_one, " ) ", "-")								'formatting the phone information
        phone_one = replace(phone_one, " ", "-")
        If phone_one = "___-___-____" Then phone_one = ""

        phone_two = replace(phone_two, " ) ", "-")
        phone_two = replace(phone_two, " ", "-")
        If phone_two = "___-___-____" Then phone_two = ""

        phone_three = replace(phone_three, " ) ", "-")
        phone_three = replace(phone_three, " ", "-")
        If phone_three = "___-___-____" Then phone_three = ""

        If type_one = "H" Then type_one = "Home"
        If type_one = "W" Then type_one = "Work"
        If type_one = "C" Then type_one = "Cell"
        If type_one = "M" Then type_one = "Message"
        If type_one = "T" Then type_one = "TTY/TDD"
        If type_one = "_" Then type_one = ""

        If type_two = "H" Then type_two = "Home"
        If type_two = "W" Then type_two = "Work"
        If type_two = "C" Then type_two = "Cell"
        If type_two = "M" Then type_two = "Message"
        If type_two = "T" Then type_two = "TTY/TDD"
        If type_two = "_" Then type_two = ""

        If type_three = "H" Then type_three = "Home"
        If type_three = "W" Then type_three = "Work"
        If type_three = "C" Then type_three = "Cell"
        If type_three = "M" Then type_three = "Message"
        If type_three = "T" Then type_three = "TTY/TDD"
        If type_three = "_" Then type_three = ""
    End If

end function


' show_pg_one_memb01_and_exp
' show_pg_one_address
' show_pg_memb_list
' show_q_1_7
' show_q_8_13
' show_q_14_19
' show_q_20_24
' show_qual
' show_pg_last
'
' update_addr
' update_pers

function define_main_dialog()

	BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"

	  ButtonGroup ButtonPressed
	    If page_display = show_pg_one_memb01_and_exp Then
			Text 490, 12, 60, 13, "Applicant and EXP"

			ComboBox 205, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_form_with, who_are_we_completing_the_form_with
			EditBox 290, 65, 50, 15, exp_q_1_income_this_month
			EditBox 310, 85, 50, 15, exp_q_2_assets_this_month
			EditBox 250, 105, 50, 15, exp_q_3_rent_this_month
			CheckBox 125, 125, 30, 10, "Heat", exp_pay_heat_checkbox
			CheckBox 160, 125, 65, 10, "Air Conditioning", exp_pay_ac_checkbox
			CheckBox 230, 125, 45, 10, "Electricity", exp_pay_electricity_checkbox
			CheckBox 280, 125, 35, 10, "Phone", exp_pay_phone_checkbox
			CheckBox 325, 125, 35, 10, "None", exp_pay_none_checkbox
			DropListBox 245, 140, 40, 45, "No"+chr(9)+"Yes", exp_migrant_seasonal_formworker_yn
			DropListBox 365, 155, 40, 45, "No"+chr(9)+"Yes", exp_received_previous_assistance_yn
			EditBox 80, 175, 80, 15, exp_previous_assistance_when
			EditBox 200, 175, 85, 15, exp_previous_assistance_where
			EditBox 320, 175, 85, 15, exp_previous_assistance_what
			DropListBox 160, 195, 40, 45, "No"+chr(9)+"Yes", exp_pregnant_yn
			ComboBox 255, 195, 150, 45, all_the_clients, exp_pregnant_who
			Text 70, 15, 130, 10, "Who are you completing the form with?"
			GroupBox 10, 50, 400, 165, "Expedited Questions - Do you need help right away?"
			Text 20, 70, 270, 10, "1. How much income (cash or checkes) did or will your household get this month?"
			Text 20, 90, 290, 10, "2. How much does your household (including children) have cash, checking or savings?"
			Text 20, 110, 225, 10, "3. How much does your household pay for rent/mortgage per month?"
			Text 30, 125, 90, 10, "What utilities do you pay?"
			Text 20, 145, 225, 10, "4. Is anyone in your household a migrant or seasonal farm worker?"
			Text 20, 160, 345, 10, "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
			Text 30, 180, 50, 10, "If yes, When?"
			Text 170, 180, 30, 10, "Where?"
			Text 295, 180, 25, 10, "What?"
			Text 20, 200, 135, 10, "6. Is anyone in your household pregnant?"
			Text 210, 200, 40, 10, "If yes, who?"
		ElseIf page_display = show_pg_one_address Then
			Text 495, 27, 60, 13, "CAF ADDR"
			If update_addr = FALSE Then
				Text 70, 55, 305, 15, resi_addr_street_full
				Text 70, 75, 105, 15, resi_addr_city
				Text 205, 75, 110, 45, resi_addr_state
				Text 340, 75, 35, 15, resi_addr_zip
				Text 125, 95, 45, 45, reservation_yn
				Text 245, 85, 130, 15, reservation_name
				Text 125, 115, 45, 45, homeless_yn
				Text 245, 115, 130, 45, living_situation
				Text 70, 165, 305, 15, mail_addr_street_full
				Text 70, 185, 105, 15, mail_addr_city
				Text 205, 185, 110, 45, mail_addr_state
				Text 340, 185, 35, 15, mail_addr_zip
				Text 20, 240, 90, 15, phone_one_number
				Text 125, 240, 65, 45, phone_pne_type
				Text 20, 260, 90, 15, phone_two_number
				Text 125, 260, 65, 45, phone_two_type
				Text 20, 280, 90, 15, phone_three_number
				Text 125, 280, 65, 45, phone_three_type
				Text 325, 220, 50, 15, address_change_date
				Text 255, 255, 120, 45, resi_addr_county
				PushButton 290, 300, 95, 15, "Update Information", update_information_btn
			End If
			If update_addr = TRUE Then
				EditBox 70, 50, 305, 15, resi_addr_street_full
				EditBox 70, 70, 105, 15, resi_addr_city
				DropListBox 205, 70, 110, 45, state_list, resi_addr_state
				EditBox 340, 70, 35, 15, resi_addr_zip
				DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", reservation_yn
				EditBox 245, 90, 130, 15, reservation_name
				DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", homeless_yn
				DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
				EditBox 70, 160, 305, 15, mail_addr_street_full
				EditBox 70, 180, 105, 15, mail_addr_city
				DropListBox 205, 180, 110, 45, state_list, mail_addr_state
				EditBox 340, 180, 35, 15, mail_addr_zip
				EditBox 20, 240, 90, 15, phone_one_number
				DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_pne_type
				EditBox 20, 260, 90, 15, phone_two_number
				DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_two_type
				EditBox 20, 280, 90, 15, phone_three_number
				DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_three_type
				EditBox 325, 220, 50, 15, address_change_date
				DropListBox 255, 255, 120, 45, county_list, resi_addr_county
				PushButton 290, 300, 95, 15, "Save Information", save_information_btn
			End If

			PushButton 325, 145, 50, 10, "CLEAR", clear_mail_addr_btn
			PushButton 205, 240, 35, 10, "CLEAR", clear_phone_one_btn
			PushButton 205, 260, 35, 10, "CLEAR", clear_phone_two_btn
			PushButton 205, 280, 35, 10, "CLEAR", clear_phone_three_btn
			Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
			GroupBox 10, 35, 375, 95, "Residence Address"
			Text 20, 55, 45, 10, "House/Street"
			Text 45, 75, 20, 10, "City"
			Text 185, 75, 20, 10, "State"
			Text 325, 75, 15, 10, "Zip"
			Text 20, 95, 100, 10, "Do you live on a Reservation?"
			Text 180, 95, 60, 10, "If yes, which one?"
			Text 30, 115, 90, 10, "Client Indicates Homeless:"
			Text 185, 115, 60, 10, "Living Situation?"
			GroupBox 10, 135, 375, 70, "Mailing Address"
			Text 20, 165, 45, 10, "House/Street"
			Text 45, 185, 20, 10, "City"
			Text 185, 185, 20, 10, "State"
			Text 325, 185, 15, 10, "Zip"
			GroupBox 10, 210, 235, 90, "Phone Number"
			Text 20, 225, 50, 10, "Number"
			Text 125, 225, 25, 10, "Type"
			Text 255, 225, 60, 10, "Date of Change:"
			Text 255, 245, 75, 10, "County of Residence:"
		ElseIf page_display = show_pg_memb_list Then
			Text 495, 42, 60, 13, "CAF MEMBs"
			If update_pers = FALSE Then
				Text 70, 45, 90, 15, HH_MEMB_ARRAY(selected_memb).last_name
				Text 165, 45, 75, 15, HH_MEMB_ARRAY(selected_memb).first_name
				Text 245, 45, 50, 15, HH_MEMB_ARRAY(selected_memb).mid_initial
				Text 300, 45, 175, 15, HH_MEMB_ARRAY(selected_memb).other_names
				If HH_MEMB_ARRAY(selected_memb).ssn_verif = "V - System Verified" Then
					Text 70, 75, 70, 15, HH_MEMB_ARRAY(selected_memb).ssn
				Else
					EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(selected_memb).ssn
				End If
				Text 145, 75, 70, 15, HH_MEMB_ARRAY(selected_memb).date_of_birth
				Text 220, 75, 50, 45, HH_MEMB_ARRAY(selected_memb).gender
				Text 275, 75, 90, 45, HH_MEMB_ARRAY(selected_memb).rel_to_applcnt
				Text 370, 75, 105, 45, HH_MEMB_ARRAY(selected_memb).marital_status
				Text 70, 105, 110, 15, HH_MEMB_ARRAY(selected_memb).last_grade_completed
				Text 195, 105, 70, 15, HH_MEMB_ARRAY(selected_memb).mn_entry_date
				Text 270, 105, 135, 15, HH_MEMB_ARRAY(selected_memb).former_state
				Text 400, 105, 75, 45, HH_MEMB_ARRAY(selected_memb).citizen
				Text 70, 135, 60, 45, HH_MEMB_ARRAY(selected_memb).interpreter
				Text 140, 135, 120, 15, HH_MEMB_ARRAY(selected_memb).spoken_lang
				Text 140, 165, 120, 15, HH_MEMB_ARRAY(selected_memb).written_lang
				Text 330, 145, 40, 45, HH_MEMB_ARRAY(selected_memb).ethnicity_yn
				' CheckBox 330, 165, 30, 10, "Asian", HH_MEMB_ARRAY(selected_memb).race_a_checkbox
				' CheckBox 330, 175, 30, 10, "Black", HH_MEMB_ARRAY(selected_memb).race_b_checkbox
				' CheckBox 330, 185, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(selected_memb).race_n_checkbox
				' CheckBox 330, 195, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(selected_memb).race_p_checkbox
				' CheckBox 330, 205, 130, 10, "White", HH_MEMB_ARRAY(selected_memb).race_w_checkbox
				' CheckBox 70, 200, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(selected_memb).snap_req_checkbox
				' CheckBox 125, 200, 65, 10, "Cash programs", HH_MEMB_ARRAY(selected_memb).cash_req_checkbox
				' CheckBox 195, 200, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(selected_memb).emer_req_checkbox
				' CheckBox 280, 200, 30, 10, "NONE", HH_MEMB_ARRAY(selected_memb).none_req_checkbox
				' DropListBox 15, 230, 80, 45, "Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).intend_to_reside_in_mn
				' EditBox 100, 230, 205, 15, HH_MEMB_ARRAY(selected_memb).imig_status
				' DropListBox 310, 230, 55, 45, "No"+chr(9)+"Yes", HH_MEMB_ARRAY(selected_memb).clt_has_sponsor
				' DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(selected_memb).client_verification
				' EditBox 100, 260, 435, 15, HH_MEMB_ARRAY(selected_memb).client_verification_details
				' EditBox 15, 290, 350, 15, HH_MEMB_ARRAY(selected_memb).client_notes
				PushButton 385, 330, 95, 15, "Update Information", update_information_btn
			End If
			If update_pers = TRUE Then
				EditBox 70, 45, 90, 15, HH_MEMB_ARRAY(selected_memb).last_name
				EditBox 165, 45, 75, 15, HH_MEMB_ARRAY(selected_memb).first_name
				EditBox 245, 45, 50, 15, HH_MEMB_ARRAY(selected_memb).mid_initial
				EditBox 300, 45, 175, 15, HH_MEMB_ARRAY(selected_memb).other_names
				EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(selected_memb).ssn
				EditBox 145, 75, 70, 15, HH_MEMB_ARRAY(selected_memb).date_of_birth
				DropListBox 220, 75, 50, 45, ""+chr(9)+"Male"+chr(9)+"Female", HH_MEMB_ARRAY(selected_memb).gender
				DropListBox 275, 75, 90, 45, memb_panel_relationship_list, HH_MEMB_ARRAY(selected_memb).rel_to_applcnt
				DropListBox 370, 75, 105, 45, marital_status, HH_MEMB_ARRAY(selected_memb).marital_status
				EditBox 70, 105, 110, 15, HH_MEMB_ARRAY(selected_memb).last_grade_completed
				EditBox 185, 105, 70, 15, HH_MEMB_ARRAY(selected_memb).mn_entry_date
				EditBox 260, 105, 135, 15, HH_MEMB_ARRAY(selected_memb).former_state
				DropListBox 400, 105, 75, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).citizen
				DropListBox 70, 135, 60, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).interpreter
				EditBox 140, 135, 120, 15, HH_MEMB_ARRAY(selected_memb).spoken_lang
				EditBox 140, 165, 120, 15, HH_MEMB_ARRAY(selected_memb).written_lang
				DropListBox 330, 145, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).ethnicity_yn

				PushButton 385, 330, 95, 15, "Save Information", save_information_btn
			End If
			CheckBox 330, 170, 30, 10, "Asian", HH_MEMB_ARRAY(selected_memb).race_a_checkbox
			CheckBox 330, 180, 30, 10, "Black", HH_MEMB_ARRAY(selected_memb).race_b_checkbox
			CheckBox 330, 190, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(selected_memb).race_n_checkbox
			CheckBox 330, 200, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(selected_memb).race_p_checkbox
			CheckBox 330, 210, 130, 10, "White", HH_MEMB_ARRAY(selected_memb).race_w_checkbox
			CheckBox 70, 210, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(selected_memb).snap_req_checkbox
			CheckBox 125, 210, 65, 10, "Cash programs", HH_MEMB_ARRAY(selected_memb).cash_req_checkbox
			CheckBox 195, 210, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(selected_memb).emer_req_checkbox
			CheckBox 280, 210, 30, 10, "NONE", HH_MEMB_ARRAY(selected_memb).none_req_checkbox
			DropListBox 70, 250, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).intend_to_reside_in_mn
			EditBox 155, 250, 205, 15, HH_MEMB_ARRAY(selected_memb).imig_status
			DropListBox 365, 250, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).clt_has_sponsor
			DropListBox 70, 280, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(selected_memb).client_verification
			EditBox 155, 280, 320, 15, HH_MEMB_ARRAY(selected_memb).client_verification_details
			EditBox 70, 310, 405, 15, HH_MEMB_ARRAY(selected_memb).client_notes
			If HH_MEMB_ARRAY(selected_memb).ref_number = "" Then
				GroupBox 65, 25, 415, 200, "Person " & known_membs+1
				GroupBox 65, 230, 415, 100, "Person " & known_membs+1 & " Interview Questions"
			Else
				GroupBox 65, 25, 415, 200, "Person " & known_membs+1 & " - MEMBER " & HH_MEMB_ARRAY(selected_memb).ref_number
				GroupBox 65, 230, 415, 100, "Person " & known_membs+1 & " - MEMBER " & HH_MEMB_ARRAY(selected_memb).ref_number & " Interview Questions"

			End If
			y_pos = 35
			For the_memb = 0 to UBound(HH_MEMB_ARRAY)
				If the_memb = selected_memb Then
					Text 20, y_pos + 1, 45, 10, "Person " & (the_memb + 1)
				Else
					PushButton 10, y_pos, 45, 10, "Person " & (the_memb + 1), HH_MEMB_ARRAY(the_memb).button_one
				End If
				y_pos = y_pos + 10
			Next
			y_pos = y_pos + 10
			PushButton 10, y_pos, 45, 10, "Add Person", add_person_btn
			Text 70, 35, 50, 10, "Last Name"
			Text 165, 35, 50, 10, "First Name"
			Text 245, 35, 50, 10, "Middle Name"
			Text 300, 35, 50, 10, "Other Names"
			Text 70, 65, 55, 10, "Soc Sec Number"
			Text 145, 65, 45, 10, "Date of Birth"
			Text 220, 65, 45, 10, "Gender"
			Text 275, 65, 90, 10, "Relationship to MEMB 01"
			Text 370, 65, 50, 10, "Marital Status"
			Text 70, 95, 75, 10, "Last Grade Completed"
			Text 185, 95, 55, 10, "Moved to MN on"
			Text 260, 95, 65, 10, "Moved to MN from"
			Text 400, 95, 75, 10, "US Citizen or National"
			Text 70, 125, 40, 10, "Interpreter?"
			Text 140, 125, 95, 10, "Preferred Spoken Language"
			Text 140, 155, 95, 10, "Preferred Written Language"
			GroupBox 320, 125, 155, 100, "Demographics"
			Text 330, 135, 35, 10, "Hispanic?"
			Text 330, 160, 50, 10, "Race"
			Text 70, 200, 145, 10, "Which programs is this person requesting?"
			Text 70, 240, 80, 10, "Intends to reside in MN"
			Text 155, 240, 65, 10, "Immigration Status"
			Text 365, 240, 50, 10, "Sponsor?"
			Text 70, 270, 50, 10, "Verification"
			Text 155, 270, 65, 10, "Verification Details"
			Text 70, 300, 50, 10, "Notes:"
		ElseIf page_display = show_q_1_7 Then
			Text 505, 57, 60, 13, "Q. 1 - 7"



			GroupBox 5, 15, 475, 50, "1. Does everyone in your household buy, fix or eat food with you?"
			Text 15, 30, 40, 10, "CAF Answer"
			DropListBox 55, 25, 35, 45, question_answers, question_1_yn
			Text 95, 30, 25, 10, "write-in:"
			EditBox 120, 25, 235, 15, question_1_notes
			Text 360, 30, 110, 10, "Q1 - Verification - " & question_1_verif_yn
			Text 15, 50, 60, 10, "Interview Notes:"
			EditBox 75, 45, 320, 15, question_1_interview_notes
			PushButton 400, 50, 75, 10, "ADD VERIFICATION", add_verif_1_btn

			GroupBox 5, 65, 475, 50, "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			' Text 20, 55, 115, 10, "buy or fix food due to a disability?"
			Text 15, 80, 40, 10, "CAF Answer"
			DropListBox 55, 75, 35, 45, question_answers, question_2_yn
			Text 95, 80, 25, 10, "write-in:"
			EditBox 120, 75, 235, 15, question_2_notes
			Text 360, 80, 110, 10, "Q2 - Verification - " & question_2_verif_yn
			Text 15, 100, 60, 10, "Interview Notes:"
			EditBox 75, 95, 320, 15, question_2_interview_notes
			PushButton 400, 100, 75, 10, "ADD VERIFICATION", add_verif_2_btn

			GroupBox 5, 115, 475, 50, "3. Is anyone in the household attending school?"
			Text 15, 130, 40, 10, "CAF Answer"
			DropListBox 55, 125, 35, 45, question_answers, question_3_yn
			Text 95, 130, 25, 10, "write-in:"
			EditBox 120, 125, 235, 15, question_3_notes
			Text 360, 130, 110, 10, "Q3 - Verification - " & question_3_verif_yn
			Text 15, 150, 60, 10, "Interview Notes:"
			EditBox 75, 145, 320, 15, question_3_interview_notes
			PushButton 400, 150, 75, 10, "ADD VERIFICATION", add_verif_3_btn

			GroupBox 5, 165, 475, 50, "4. Is anyone in your household temporarily not living in your home? (eg. vacation, foster care, treatment, hospital, job search)"
			' Text 20, 135, 230, 10, "(for example: vacation, foster care, treatment, hospital, job search)"
			Text 15, 180, 40, 10, "CAF Answer"
			DropListBox 55, 175, 35, 45, question_answers, question_4_yn
			Text 95, 180, 25, 10, "write-in:"
			EditBox 120, 175, 235, 15, question_4_notes
			Text 360, 180, 110, 10, "Q4 - Verification - " & question_4_verif_yn
			Text 15, 200, 60, 10, "Interview Notes:"
			EditBox 75, 195, 320, 15, question_4_interview_notes
			PushButton 400, 200, 75, 10, "ADD VERIFICATION", add_verif_4_btn

			GroupBox 5, 215, 475, 50, "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			' Text 20, 180, 185, 10, " that limits the ability to work or perform daily activities?"
			Text 15, 230, 40, 10, "CAF Answer"
			DropListBox 55, 225, 35, 45, question_answers, question_5_yn
			Text 95, 230, 25, 10, "write-in:"
			EditBox 120, 225, 235, 15, question_5_notes
			Text 360, 230, 110, 10, "Q5 - Verification - " & question_5_verif_yn
			Text 15, 250, 60, 10, "Interview Notes:"
			EditBox 75, 245, 320, 15, question_5_interview_notes
			PushButton 400, 250, 75, 10, "ADD VERIFICATION", add_verif_5_btn

			GroupBox 5, 265, 475, 50, "6. Is anyone unable to work for reasons other than illness or disability?"
			Text 15, 280, 40, 10, "CAF Answer"
			DropListBox 55, 275, 35, 45, question_answers, question_6_yn
			Text 95, 280, 25, 10, "write-in:"
			EditBox 120, 275, 235, 15, question_6_notes
			Text 360, 280, 110, 10, "Q6 - Verification - " & question_6_verif_yn
			Text 15, 300, 60, 10, "Interview Notes:"
			EditBox 75, 295, 320, 15, question_6_interview_notes
			PushButton 400, 300, 75, 10, "ADD VERIFICATION", add_verif_6_btn

			GroupBox 5, 315, 475, 50, "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
			' Text 20, 315, 350, 10, "- Stop working or quit a job?   - Refuse a job offer? - Ask to work fewer hours?   - Go on strike?"
			Text 15, 330, 40, 10, "CAF Answer"
			DropListBox 55, 325, 35, 45, question_answers, question_7_yn
			Text 95, 330, 25, 10, "write-in:"
			EditBox 120, 325, 235, 15, question_7_notes
			Text 360, 330, 110, 10, "Q7 - Verification - " & question_7_verif_yn
			Text 15, 350, 60, 10, "Interview Notes:"
			EditBox 75, 345, 320, 15, question_7_interview_notes
			PushButton 400, 350, 75, 10, "ADD VERIFICATION", add_verif_7_btn
		ElseIf page_display = show_q_8_13 Then
			Text 505, 72, 60, 13, "Q. 8 - 13"

			GroupBox 5, 10, 475, 65, "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			Text 15, 25, 40, 10, "CAF Answer"
			DropListBox 55, 20, 35, 45, question_answers, question_8_yn
			Text 95, 25, 25, 10, "write-in:"
			EditBox 120, 20, 235, 15, question_8_notes
			Text 360, 25, 110, 10, "Q8 - Verification - " & question_8_verif_yn
			Text 15, 40, 400, 10, "a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?       CAF Answer"
			DropListBox 415, 35, 35, 45, question_answers, question_8a_yn
			Text 15, 60, 60, 10, "Interview Notes:"
			EditBox 75, 55, 320, 15, question_7_interview_notes
			PushButton 400, 60, 75, 10, "ADD VERIFICATION", add_verif_8_btn

			GroupBox 5, 80, 475, 50, "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			Text 15, 95, 40, 10, "CAF Answer"
			DropListBox 55, 90, 35, 45, question_answers, question_9_yn
			Text 95, 95, 25, 10, "write-in:"
			EditBox 120, 90, 235, 15, question_9_notes
			Text 360, 95, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			PushButton 125, 100, 55, 10, "ADD JOB", add_job_btn

			' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			y_pos = 115
			' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then

					Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					PushButton 410, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					y_pos = y_pos + 10
				End If
			next
			y_pos = y_pos + 10

			y_pos = y_pos + 15
			DropListBox 10, y_pos, 60, 45, question_answers, question_10_yn
			Text 80, y_pos, 430, 10, "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			Text 540, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			y_pos = y_pos + 10
			ButtonGroup ButtonPressed
			  PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_10_btn
			Text 95, y_pos, 85, 10, "Gross Monthly Earnings:"
			Text 185, y_pos, 25, 10, "Notes:"
			y_pos = y_pos + 10
			EditBox 95, y_pos, 80, 15, question_10_monthly_earnings
			EditBox 185, y_pos, 325, 15, question_10_notes
			y_pos = y_pos + 20
			DropListBox 10, y_pos, 60, 45, question_answers, question_11_yn
			Text 80, y_pos, 255, 10, "11. Do you expect any changes in income, expenses or work hours?"
			Text 540, y_pos, 105, 10, "Q11 - Verification - " & question_11_verif_yn
			y_pos = y_pos + 10
			ButtonGroup ButtonPressed
			  PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_11_btn
			Text 95, y_pos + 5, 25, 10, "Notes:"
			EditBox 120, y_pos, 390, 15, question_11_notes
			y_pos = y_pos + 25
			Text 80, y_pos, 75, 10, "Pricipal Wage Earner"
			DropListBox 155, y_pos - 5, 175, 45, all_the_clients, pwe_selection
			y_pos = y_pos + 10
			Text 80, y_pos + 5, 370, 10, "12. Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			Text 540, y_pos, 105, 10, "Q12 - Verification - " & question_12_verif_yn
			y_pos = y_pos + 10
			ButtonGroup ButtonPressed
			  PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_12_btn
			y_pos = y_pos + 10
			DropListBox 80, y_pos, 60, 45, question_answers, question_12_rsdi_yn
			Text 150, y_pos + 5, 70, 10, "RSDI                      $"
			EditBox 220, y_pos, 35, 15, question_12_rsdi_amt
			DropListBox 305, y_pos, 60, 45, question_answers, question_12_ssi_yn
			Text 375, y_pos + 5, 85, 10, "SSI                                 $"
			EditBox 460, y_pos, 35, 15, question_12_ssi_amt
			y_pos = y_pos + 15
			DropListBox 80, y_pos, 60, 45, question_answers, question_12_va_yn
			Text 150, y_pos + 5, 70, 10, "VA                          $"
			EditBox 220, y_pos, 35, 15, question_12_va_amt
			DropListBox 305, y_pos, 60, 45, question_answers, question_12_ui_yn
			Text 375, y_pos + 5, 85, 10, "Unemployment Ins          $"
			EditBox 460, y_pos, 35, 15, question_12_ui_amt
			y_pos = y_pos + 15
			DropListBox 80, y_pos, 60, 45, question_answers, question_12_wc_yn
			Text 150, y_pos + 5, 70, 10, "Workers Comp       $"
			EditBox 220, y_pos, 35, 15, question_12_wc_amt
			DropListBox 305, y_pos, 60, 45, question_answers, question_12_ret_yn
			Text 375, y_pos + 5, 85, 10, "Retirement Ben.              $"
			EditBox 460, y_pos, 35, 15, question_12_ret_amt
			y_pos = y_pos + 15
			DropListBox 80, y_pos, 60, 45, question_answers, question_12_trib_yn
			Text 150, y_pos + 5, 70, 10, "Tribal Payments      $"
			EditBox 220, y_pos, 35, 15, question_12_trib_amt
			DropListBox 305, y_pos, 60, 45, question_answers, question_12_cs_yn
			Text 375, y_pos + 5, 85, 10, "Child/Spousal Support    $"
			EditBox 460, y_pos, 35, 15, question_12_cs_amt
			y_pos = y_pos + 15
			DropListBox 80, y_pos, 60, 45, question_answers, question_12_other_yn
			Text 150, y_pos + 5, 110, 10, "Other unearned income          $"
			EditBox 250, y_pos, 35, 15, question_12_other_amt
			y_pos = y_pos + 20
			Text 95, y_pos + 5, 25, 10, "Notes:"
			EditBox 120, y_pos, 390, 15, question_12_notes
			y_pos = y_pos + 25
			DropListBox 10, y_pos, 60, 45, question_answers, question_13_yn
			Text 0, 0, 0, 0, ""
			Text 0, 0, 0, 0, ""
			Text 0, 0, 0, 0, ""
			Text 0, 0, 0, 0, ""
			Text 0, 0, 0, 0, ""
			Text 80, y_pos, 400, 10, "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			Text 540, y_pos, 105, 10, "Q13 - Verification - " & question_13_verif_yn
			y_pos = y_pos + 10
			ButtonGroup ButtonPressed
			  PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_13_btn
			Text 95, y_pos + 5, 25, 10, "Notes:"
			EditBox 120, y_pos, 390, 15, question_13_notes
			y_pos = y_pos + 20
		ElseIf page_display = show_q_14_19 Then
			Text 505, 87, 60, 13, "Q. 14 - 19"

			DropListBox 95, 20, 60, 45, question_answers, question_14_rent_yn
			DropListBox 300, 20, 60, 45, question_answers, question_14_subsidy_yn
			DropListBox 95, 35, 60, 45, question_answers, question_14_mortgage_yn
			DropListBox 300, 35, 60, 45, question_answers, question_14_association_yn
			DropListBox 95, 50, 60, 45, question_answers, question_14_insurance_yn
			DropListBox 300, 50, 60, 45, question_answers, question_14_room_yn
			DropListBox 95, 65, 60, 45, question_answers, question_14_taxes_yn
			EditBox 135, 85, 390, 15, question_14_notes
			DropListBox 95, 120, 60, 45, question_answers, question_15_heat_ac_yn
			DropListBox 265, 120, 60, 45, question_answers, question_15_electricity_yn
			DropListBox 415, 120, 60, 45, question_answers, question_15_cooking_fuel_yn
			DropListBox 95, 135, 60, 45, question_answers, question_15_water_and_sewer_yn
			DropListBox 265, 135, 60, 45, question_answers, question_15_garbage_yn
			DropListBox 415, 135, 60, 45, question_answers, question_15_phone_yn
			DropListBox 95, 150, 60, 45, question_answers, question_15_liheap_yn
			EditBox 120, 165, 390, 15, question_15_notes
			DropListBox 10, 190, 60, 45, question_answers, question_16_yn
			EditBox 120, 210, 390, 15, question_16_notes
			DropListBox 10, 235, 60, 45, question_answers, question_17_yn
			EditBox 120, 255, 390, 15, question_17_notes
			DropListBox 10, 280, 60, 45, question_answers, question_18_yn
			EditBox 120, 300, 390, 15, question_18_notes
			DropListBox 10, 325, 60, 45, question_answers, question_19_yn
			EditBox 120, 335, 390, 15, question_19_notes

			PushButton 580, 20, 75, 10, "ADD VERIFICATION", add_verif_14_btn
			PushButton 580, 120, 75, 10, "ADD VERIFICATION", add_verif_15_btn
			PushButton 580, 200, 75, 10, "ADD VERIFICATION", add_verif_16_btn
			PushButton 580, 245, 75, 10, "ADD VERIFICATION", add_verif_17_btn
			PushButton 580, 290, 75, 10, "ADD VERIFICATION", add_verif_18_btn
			PushButton 580, 335, 75, 10, "ADD VERIFICATION", add_verif_19_btn

			Text 80, 10, 220, 10, "14. Does your household have the following housing expenses?"
			Text 165, 25, 70, 10, "Rent"
			Text 370, 25, 100, 10, "Rent or Section 8 Subsidy"
			Text 165, 40, 125, 10, "Mortgage/contract for deed payment"
			Text 370, 40, 70, 10, "Association fees"
			Text 165, 55, 85, 10, "Homeowner's insurance"
			Text 370, 55, 70, 10, "Room and/or board"
			Text 165, 70, 100, 10, "Real estate taxes"
			Text 110, 90, 25, 10, "Notes:"
			Text 560, 10, 105, 10, "Q14 - Verification - " & question_14_verif_yn
			Text 80, 110, 290, 10, "15. Does your household have the following utility expenses any time during the year? "
			Text 165, 125, 85, 10, "Heating/air conditioning"
			Text 335, 125, 70, 10, "Electricity"
			Text 485, 125, 70, 10, "Cooking fuel"
			Text 165, 140, 75, 10, "Water and sewer"
			Text 335, 140, 60, 10, "Garbage removal"
			Text 485, 140, 70, 10, "Phone/cell phone"
			Text 165, 155, 375, 10, "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"
			Text 95, 170, 25, 10, "Notes:"
			Text 560, 110, 105, 10, "Q15 - Verification - " & question_15_verif_yn
			Text 80, 190, 345, 10, "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working,"
			Text 95, 200, 125, 10, "looking for work or going to school?"
			Text 95, 215, 25, 10, "Notes:"
			Text 560, 190, 105, 10, "Q16 - Verification - " & question_16_verif_yn
			Text 80, 235, 380, 10, "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working,"
			Text 95, 245, 125, 10, "looking for work or going to school?"
			Text 95, 260, 25, 10, "Notes:"
			Text 560, 235, 105, 10, "Q17 - Verification - " & question_17_verif_yn
			Text 80, 280, 430, 10, "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support"
			Text 95, 290, 215, 10, "or contribute to a tax dependent who does not live in your home?"
			Text 95, 305, 25, 10, "Notes:"
			Text 0, 0, 0, 0, ""
			Text 0, 0, 0, 0, ""
			Text 560, 280, 105, 10, "Q18 - Verification - " & question_18_verif_yn
			Text 80, 325, 255, 10, "19. For SNAP only: Does anyone in the household have medical expenses? "
			Text 95, 340, 25, 10, "Notes:"
			Text 560, 325, 105, 10, "Q19 - Verification - " & question_19_verif_yn
		ElseIf page_display = show_q_20_24 Then
			Text 505, 102, 60, 13, "Q. 20 - 24"

			DropListBox 80, 25, 60, 45, question_answers, question_20_cash_yn
			DropListBox 285, 25, 60, 45, question_answers, question_20_acct_yn
			DropListBox 80, 40, 60, 45, question_answers, question_20_secu_yn
			DropListBox 285, 40, 60, 45, question_answers, question_20_cars_yn
			EditBox 120, 60, 390, 15, question_20_notes
			DropListBox 10, 85, 60, 45, question_answers, question_21_yn
			EditBox 120, 95, 390, 15, question_21_notes
			DropListBox 10, 120, 60, 45, question_answers, question_22_yn
			EditBox 120, 130, 390, 15, question_22_notes
			DropListBox 10, 155, 60, 45, question_answers, question_23_yn
			EditBox 120, 165, 390, 15, question_23_notes
			DropListBox 80, 205, 60, 45, question_answers, question_24_rep_payee_yn
			DropListBox 285, 205, 60, 45, question_answers, question_24_guardian_fees_yn
			DropListBox 80, 220, 60, 45, question_answers, question_24_special_diet_yn
			DropListBox 285, 220, 60, 45, question_answers, question_24_high_housing_yn
			EditBox 120, 240, 390, 15, question_24_notes

			PushButton 560, 20, 75, 10, "ADD VERIFICATION", add_verif_20_btn
			PushButton 560, 95, 75, 10, "ADD VERIFICATION", add_verif_21_btn
			PushButton 560, 130, 75, 10, "ADD VERIFICATION", add_verif_22_btn
			PushButton 560, 165, 75, 10, "ADD VERIFICATION", add_verif_23_btn
			PushButton 560, 200, 75, 10, "ADD VERIFICATION", add_verif_24_btn

			Text 80, 10, 280, 10, "20. Does anyone in the household own, or is anyone buying, any of the following?"
			Text 150, 30, 70, 10, "Cash"
			Text 355, 30, 175, 10, "Bank accounts (savings, checking, debit card, etc.)"
			Text 150, 45, 125, 10, "Stocks, bonds, annuities, 401k, etc."
			Text 355, 45, 180, 10, "Vehicles (cars, trucks, motorcycles, campers, trailers)"
			Text 95, 65, 25, 10, "Notes:"
			Text 540, 10, 105, 10, "Q20 - Verification - " & question_20_verif_yn
			Text 80, 85, 420, 10, "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? "
			Text 95, 100, 25, 10, "Notes:"
			Text 540, 85, 105, 10, "Q21 - Verification - " & question_21_verif_yn
			Text 80, 120, 305, 10, "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			Text 95, 135, 25, 10, "Notes:"
			Text 540, 120, 105, 10, "Q22 - Verification - " & question_22_verif_yn
			Text 80, 155, 250, 10, "23. For children under the age of 19, are both parents living in the home?"
			Text 95, 170, 25, 10, "Notes:"
			Text 540, 155, 105, 10, "Q23 - Verification - " & question_23_verif_yn
			Text 80, 190, 325, 10, "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			Text 150, 210, 95, 10, "Representative Payee fees"
			Text 355, 210, 105, 10, "Guardian Conservator fees"
			Text 150, 225, 125, 10, "Physician-perscribed special diet"
			Text 355, 225, 105, 10, "High housing costs"
			Text 95, 245, 25, 10, "Notes:"
			Text 540, 190, 105, 10, "Q24 - Verification - " & question_24_verif_yn
		ElseIf page_display = show_qual Then
			Text 492, 117, 60, 13, "CAF QUAL Q"

			DropListBox 220, 40, 30, 45, "No"+chr(9)+"Yes", qual_question_one
			ComboBox 340, 40, 105, 45, all_the_clients, qual_memb_one
			DropListBox 220, 80, 30, 45, "No"+chr(9)+"Yes", qual_question_two
			ComboBox 340, 80, 105, 45, all_the_clients, qual_memb_two
			DropListBox 220, 110, 30, 45, "No"+chr(9)+"Yes", qual_question_three
			ComboBox 340, 110, 105, 45, all_the_clients, qual_memb_there
			DropListBox 220, 140, 30, 45, "No"+chr(9)+"Yes", qual_question_four
			ComboBox 340, 140, 105, 45, all_the_clients, qual_memb_four
			DropListBox 220, 160, 30, 45, "No"+chr(9)+"Yes", qual_question_five
			ComboBox 340, 160, 105, 45, all_the_clients, qual_memb_five

			PushButton 340, 185, 50, 15, "Next", next_btn
			PushButton 285, 190, 50, 10, "Back", back_btn

			Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
			Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			Text 10, 110, 195, 30, "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
			Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			Text 260, 40, 70, 10, "Household Member:"
			Text 260, 80, 70, 10, "Household Member:"
			Text 260, 110, 70, 10, "Household Member:"
			Text 260, 140, 70, 10, "Household Member:"
			Text 260, 160, 70, 10, "Household Member:"
		ElseIf page_display = show_pg_last Then
			Text 490, 132, 60, 13, "CAF Last Page"

			EditBox 135, 50, 60, 15, caf_form_date
			DropListBox 135, 70, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_yn

			PushButton 35, 90, 105, 15, "Complete CAF Form Detail", complete_caf_questions
			PushButton 10, 35, 145, 10, "Open RIGHTS AND RESPONSIBLITIES ", open_r_and_r_button

			Text 10, 10, 160, 20, "Confirm the client is signing this form and attesting to the information provided verbally."
			Text 70, 55, 55, 10, "CAF Form Date:"
			Text 10, 75, 120, 10, "Cient signature accepted verbally?"
	    ' ElseIf page_display =  Then
		End If


		' show_pg_one_memb01_and_exp
		' show_pg_one_address
		' show_pg_memb_list
		' show_q_1_7
		' show_q_8_13
		' show_q_14_19
		' show_q_20_24
		' show_qual
		' show_pg_last
		'
		' update_addr
		' update_pers


		If page_display <> show_pg_one_memb01_and_exp 	Then PushButton 485, 10, 60, 13, "Applicant & EXP", caf_page_one_btn
		If page_display <> show_pg_one_address 			Then PushButton 485, 25, 60, 13, "CAF ADDR", caf_addr_btn
		' If page_display <> show_pg_memb_list AND page_display <> show_pg_memb_info AND  page_display <> show_pg_imig Then PushButton 485, 25, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_pg_memb_list 			Then PushButton 485, 40, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_q_1_7 					Then PushButton 485, 55, 60, 13, "Q. 1 - 7", caf_q_1_7_btn
		If page_display <> show_q_8_13 					Then PushButton 485, 70, 60, 13, "Q. 8 - 13", caf_q_8_13_btn
		If page_display <> show_q_14_19 				Then PushButton 485, 85, 60, 13, "Q. 14 - 19", caf_q_14_19_btn
		If page_display <> show_q_20_24 				Then PushButton 485, 100, 60, 13, "Q. 20 - 24", caf_q_20_24_btn
		' If page_display <> show_q_6 Then PushButton 485, 100, 60, 13, "Q. 6", caf_q_6_btn
		' If page_display <> show_q_7 Then PushButton 485, 115, 60, 13, "Q. 7", caf_q_7_btn
		' If page_display <> show_q_8 Then PushButton 485, 130, 60, 13, "Q. 8", caf_q_8_btn
		' If page_display <> show_q_9 Then PushButton 485, 145, 60, 13, "Q. 9", caf_q_9_btn
		' If page_display <> show_q_10 Then PushButton 485, 160, 60, 13, "Q. 10", caf_q_10_btn
		' If page_display <> show_q_11 Then PushButton 485, 175, 60, 13, "Q. 11", caf_q_11_btn
		' If page_display <> show_q_12 Then PushButton 485, 190, 60, 13, "Q. 12", caf_q_12_btn
		' If page_display <> show_q_13 Then PushButton 485, 205, 60, 13, "Q. 13", caf_q_13_btn
		' If page_display <> show_q_14_15 Then PushButton 485, 220, 60, 13, "Q. 14 and 15", caf_q_14_15_btn
		' If page_display <> show_q_16_18 Then PushButton 485, 235, 60, 13, "Q. 16, 17, and 18", caf_q_16_17_18_btn
		' If page_display <> show_q_19 Then PushButton 485, 250, 60, 13, "Q. 19", caf_q_19_btn
		' If page_display <> show_q_20_21 Then PushButton 485, 265, 60, 13, "Q. 20 and 21", caf_q_20_21_btn
		' If page_display <> show_q_22 Then PushButton 485, 280, 60, 13, "Q. 22", caf_q_22_btn
		' If page_display <> show_q_23 Then PushButton 485, 295, 60, 13, "Q. 23", caf_q_23_btn
		' If page_display <> show_q_24 Then PushButton 485, 310, 60, 13, "Q. 24", caf_q_24_btn
		If page_display <> show_qual 					Then PushButton 485, 115, 60, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last 				Then PushButton 485, 130, 60, 13, "CAF Last Page", caf_last_page_btn
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn

	EndDialog

end function

function dialog_movement()
	' case_has_imig = FALSE
	' MsgBox ButtonPressed
	For i = 0 to Ubound(HH_MEMB_ARRAY)
		' If HH_MEMB_ARRAY(i).imig_exists = TRUE Then case_has_imig = TRUE
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = HH_MEMB_ARRAY(i).button_one Then
			If page_display = show_pg_memb_list Then selected_memb = i
			' If page_display = show_pg_imig Then selected_memb = i
			' If page_display = show_q_12 Then memb_to_match = HH_MEMB_ARRAY(i).ref_number
			' If page_display = show_q_7 Then stwk_selected = i
			' If page_display = show_q_19 Then fmed_selected = i
			' If second_page_display = ssi_unea Then memb_to_match = HH_MEMB_ARRAY(i).ref_number
			' MsgBox "Selected Memb - " & selected_memb
		End If
	Next
	If ButtonPressed = add_verif_1_btn Then Call verif_details_dlg(1)
	If ButtonPressed = add_verif_2_btn Then Call verif_details_dlg(2)
	If ButtonPressed = add_verif_3_btn Then Call verif_details_dlg(3)
	If ButtonPressed = add_verif_4_btn Then Call verif_details_dlg(4)
	If ButtonPressed = add_verif_5_btn Then Call verif_details_dlg(5)
	If ButtonPressed = add_verif_6_btn Then Call verif_details_dlg(6)
	If ButtonPressed = add_verif_7_btn Then Call verif_details_dlg(7)
	If ButtonPressed = add_verif_8_btn Then Call verif_details_dlg(8)
	If ButtonPressed = add_verif_9_btn Then Call verif_details_dlg(9)
	If ButtonPressed = add_verif_10_btn Then Call verif_details_dlg(10)
	If ButtonPressed = add_verif_11_btn Then Call verif_details_dlg(11)
	If ButtonPressed = add_verif_12_btn Then Call verif_details_dlg(12)
	If ButtonPressed = add_verif_13_btn Then Call verif_details_dlg(13)
	If ButtonPressed = add_verif_14_btn Then Call verif_details_dlg(14)
	If ButtonPressed = add_verif_15_btn Then Call verif_details_dlg(15)
	If ButtonPressed = add_verif_16_btn Then Call verif_details_dlg(16)
	If ButtonPressed = add_verif_17_btn Then Call verif_details_dlg(17)
	If ButtonPressed = add_verif_18_btn Then Call verif_details_dlg(18)
	If ButtonPressed = add_verif_19_btn Then Call verif_details_dlg(19)
	If ButtonPressed = add_verif_20_btn Then Call verif_details_dlg(20)
	If ButtonPressed = add_verif_21_btn Then Call verif_details_dlg(21)
	If ButtonPressed = add_verif_22_btn Then Call verif_details_dlg(22)
	If ButtonPressed = add_verif_23_btn Then Call verif_details_dlg(23)
	If ButtonPressed = add_verif_24_btn Then Call verif_details_dlg(24)

	If ButtonPressed = update_information_btn Then
		If page_display = show_pg_one_address Then update_addr = TRUE
		If page_display = show_pg_memb_list Then update_pers = TRUE
	End If
	If ButtonPressed = save_information_btn Then
		If page_display = show_pg_one_address Then update_addr = FALSE
		If page_display = show_pg_memb_list Then update_pers = FALSE
	End If
	If ButtonPressed = clear_mail_addr_btn Then
		' phone_one_number = ""
		' phone_pne_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_one_btn Then
		phone_one_number = ""
		phone_pne_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_two_btn Then
		phone_two_number = ""
		phone_two_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_three_btn Then
		phone_three_number = ""
		phone_three_type = "Select One..."
	End If
	' For i = 0 to Ubound(ASSET_ARRAY, 1)
	' 	' MsgBox HH_MEMB_ARRAY(i).button_one
	' 	If ButtonPressed = ASSET_ARRAY(i).asset_btn_one Then
	' 		memb_to_match = ASSET_ARRAY(i).member_ref
	' 		inst_to_match = ASSET_ARRAY(i).panel_instance
	' 	End If
	' Next
	' MsgBox ButtonPressed
	If page_display = show_pg_memb_info AND ButtonPressed = -1 Then ButtonPressed = next_memb_btn
	' If page_display = show_pg_imig AND ButtonPressed = -1 Then ButtonPressed = next_imig_btn

	' If ButtonPressed = next_imig_btn Then
	' 	imig_questions_step = imig_questions_step + 1
	'
	' 	If imig_questions_step = 2 AND HH_MEMB_ARRAY(memb_selected).imig_q_2_required = FALSE Then imig_questions_step = 3
	' 	If imig_questions_step = 4 AND HH_MEMB_ARRAY(memb_selected).imig_q_4_required = FALSE Then imig_questions_step = 5
	' 	If imig_questions_step = 5 AND HH_MEMB_ARRAY(memb_selected).imig_q_5_required = FALSE Then imig_questions_step = 6
	' 	If imig_questions_step > 5 Then
	' 		ButtonPressed = next_memb_btn
	' 		imig_questions_step = 1
	' 	End If
	' End If
	' If ButtonPressed = prev_imig_btn Then
	' 	imig_questions_step = imig_questions_step - 1
	' 	If imig_questions_step = 2 AND HH_MEMB_ARRAY(memb_selected).imig_q_2_required = FALSE Then imig_questions_step = 1
	' 	If imig_questions_step = 5 AND HH_MEMB_ARRAY(memb_selected).imig_q_5_required = FALSE Then imig_questions_step = 4
	' 	If imig_questions_step = 4 AND HH_MEMB_ARRAY(memb_selected).imig_q_4_required = FALSE Then imig_questions_step = 3
	'
	' 	If imig_questions_step < 1 Then imig_questions_step = 1
	' End If
	If ButtonPressed = next_memb_btn Then
		memb_selected = memb_selected + 1
		If memb_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	End If
	' If ButtonPressed = next_stwk_btn Then
	' 	stwk_selected = stwk_selected + 1
	' 	Do
	' 		If HH_MEMB_ARRAY(stwk_selected).stwk_exists = FALSE Then stwk_selected = stwk_selected + 1
	' 		If stwk_selected > UBound(HH_MEMB_ARRAY, 1) Then Exit Do
	' 	Loop Until  HH_MEMB_ARRAY(stwk_selected).stwk_exists = TRUE
	' 	If stwk_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	' End If
	' If ButtonPressed = next_fmed_btn Then
	' 	fmed_selected = fmed_selected + 1
	' 	Do
	' 		If HH_MEMB_ARRAY(fmed_selected).fmed_exists = FALSE Then fmed_selected = fmed_selected + 1
	' 		If fmed_selected > UBound(HH_MEMB_ARRAY, 1) Then Exit Do
	' 	Loop Until  HH_MEMB_ARRAY(fmed_selected).fmed_exists = TRUE
	' 	If fmed_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	' End If
	If ButtonPressed = -1 Then ButtonPressed = next_btn
	If ButtonPressed = next_btn Then
		If page_display = show_pg_one_memb01_and_exp 	Then ButtonPressed = caf_addr_btn
		If page_display = show_pg_one_address 			Then ButtonPressed = caf_membs_btn
		If page_display = show_pg_memb_list 			Then ButtonPressed = caf_q_1_7_btn
		If page_display = show_q_1_7 					Then ButtonPressed = caf_q_8_13_btn
		If page_display = show_q_8_13 					Then ButtonPressed = caf_q_14_19_btn
		If page_display = show_q_14_19 					Then ButtonPressed = caf_q_20_24_btn
		If page_display = show_q_20_24 					Then ButtonPressed = caf_qual_q_btn
		If page_display = show_qual 					Then ButtonPressed = caf_last_page_btn
		If page_display = show_pg_last 					Then ButtonPressed = finish_interview_btn

		' If page_display = show_pg_one Then ButtonPressed = caf_membs_btn
		' If page_display = show_pg_memb_list Then ButtonPressed = HH_memb_detail_review
		' If page_display = show_pg_memb_info AND case_has_imig = FALSE Then ButtonPressed = show_pg_imig
		' If page_display = show_pg_memb_info AND case_has_imig = TRUE Then ButtonPressed = caf_q_1_2_btn
		' If page_display = show_pg_imig Then ButtonPressed = caf_q_1_2_btn
		' If page_display = show_q_1_2 Then ButtonPressed = caf_q_3_btn
		' If page_display = show_q_3 Then ButtonPressed = caf_q_4_btn
		' If page_display = show_q_4 Then ButtonPressed = caf_q_5_btn
		' If page_display = show_q_5 Then ButtonPressed = caf_q_6_btn
		' If page_display = show_q_6 Then ButtonPressed = caf_q_7_btn
		' If page_display = show_q_7 Then ButtonPressed = caf_q_8_btn
		' If page_display = show_q_8 Then ButtonPressed = caf_q_9_btn
		' If page_display = show_q_9 Then ButtonPressed = caf_q_10_btn
		' If page_display = show_q_10 Then ButtonPressed = caf_q_11_btn
		' If page_display = show_q_11 Then ButtonPressed = caf_q_12_btn
		' If page_display = show_q_12 Then
		' 	If second_page_display = main_unea Then ButtonPressed = rsdi_btn
		' 	If second_page_display = rsdi_unea Then ButtonPressed = ssi_btn
		' 	If second_page_display = ssi_unea Then ButtonPressed = va_btn
		' 	If second_page_display = va_unea Then ButtonPressed = ui_btn
		' 	If second_page_display = ui_unea Then ButtonPressed = wc_btn
		' 	If second_page_display = wc_unea Then ButtonPressed = ret_btn
		' 	If second_page_display = ret_unea Then ButtonPressed = tribal_btn
		' 	If second_page_display = tribal_unea Then ButtonPressed = cs_btn
		' 	If second_page_display = cs_unea Then ButtonPressed = ss_btn
		' 	If second_page_display = ss_unea Then ButtonPressed = other_btn
		' 	If second_page_display = other_unea Then ButtonPressed = caf_q_13_btn
		' End If
		' If page_display = show_q_13 Then ButtonPressed = caf_q_14_15_btn
		' If page_display = show_q_14_15 Then ButtonPressed = caf_q_16_17_18_btn
		' If page_display = show_q_16_18 Then ButtonPressed = caf_q_19_btn
		' If page_display = show_q_19 Then ButtonPressed = caf_q_20_21_btn
		' If page_display = show_q_20_21 Then
		' 	If second_page_display = main_asset Then ButtonPressed = cash_btn
		' 	If second_page_display = cash_asset Then ButtonPressed = acct_btn
		' 	If second_page_display = acct_asset Then ButtonPressed = secu_btn
		' 	If second_page_display = secu_asset Then ButtonPressed = cars_btn
		' 	If second_page_display = cars_asset Then ButtonPressed = rest_btn
		' 	If second_page_display = rest_asset Then ButtonPressed = caf_q_22_btn
		' End If
		' If page_display = show_q_22 Then ButtonPressed = caf_q_23_btn
		' If page_display = show_q_23 Then ButtonPressed = caf_q_24_btn
		' If page_display = show_q_24 Then ButtonPressed = caf_qual_q_btn
		' If page_display = show_qual Then ButtonPressed = caf_last_page_btn
		' If page_display = show_pg_last Then ButtonPressed =
	End If

	If ButtonPressed = caf_page_one_btn Then
		page_display = show_pg_one_memb01_and_exp
	End If
	If ButtonPressed = caf_addr_btn Then
		page_display = show_pg_one_address
	End If
	If ButtonPressed = caf_membs_btn Then
		page_display = show_pg_memb_list
	End If
	If ButtonPressed = caf_q_1_7_btn Then
		page_display = show_q_1_7
	End If
	If ButtonPressed = caf_q_8_13_btn Then
		page_display = show_q_8_13
	End If
	If ButtonPressed = caf_q_14_19_btn Then
		page_display = show_q_14_19
	End If
	If ButtonPressed = caf_q_20_24_btn Then
		page_display = show_q_20_24
	End If
	If ButtonPressed = caf_qual_q_btn Then
		page_display = show_qual
	End If
	If ButtonPressed = caf_last_page_btn Then
		page_display = show_pg_last
	End If


	' If ButtonPressed = caf_page_one_btn Then
	' 	page_display = show_pg_one
	' End If
	' If ButtonPressed = caf_membs_btn Then
	' 	page_display = show_pg_memb_list
	' End If
	' If ButtonPressed = hh_list_btn Then
	' 	page_display = show_pg_memb_list
	' End If
	' If ButtonPressed = HH_memb_detail_review Then
	' 	page_display = show_pg_memb_info
	' End If
	' If ButtonPressed = hh_imig_btn Then
	' 	page_display = show_pg_imig
	' End If
	' If ButtonPressed = caf_q_1_2_btn Then
	' 	page_display = show_q_1_2
	' End If
	' If ButtonPressed = caf_q_3_btn Then
	' 	page_display = show_q_3
	' End If
	' If ButtonPressed = caf_q_4_btn Then
	' 	page_display = show_q_4
	' End If
	' If ButtonPressed = caf_q_5_btn Then
	' 	page_display = show_q_5
	' End If
	' If ButtonPressed = caf_q_6_btn Then
	' 	page_display = show_q_6
	' End If
	' If ButtonPressed = caf_q_7_btn Then
	' 	page_display = show_q_7
	' End If
	' If ButtonPressed = caf_q_8_btn Then
	' 	page_display = show_q_8
	' End If
	' If ButtonPressed = caf_q_9_btn Then
	' 	page_display = show_q_9
	' End If
	' If ButtonPressed = caf_q_10_btn Then
	' 	page_display = show_q_10
	' End If
	' If ButtonPressed = caf_q_11_btn Then
	' 	page_display = show_q_11
	' End If
	' If ButtonPressed = caf_q_12_btn Then
	' 	page_display = show_q_12
	' 	second_page_display = main_unea
	' End If
	' If ButtonPressed = rsdi_btn 	Then second_page_display = rsdi_unea
	' If ButtonPressed = ssi_btn		Then second_page_display = ssi_unea
	' If ButtonPressed = va_btn		Then second_page_display = va_unea
	' If ButtonPressed = ui_btn		Then second_page_display = ui_unea
	' If ButtonPressed = wc_btn		Then second_page_display = wc_unea
	' If ButtonPressed = ret_btn		Then second_page_display = ret_unea
	' If ButtonPressed = tribal_btn	Then second_page_display = tribal_unea
	' If ButtonPressed = cs_btn		Then second_page_display = cs_unea
	' If ButtonPressed = ss_btn		Then second_page_display = ss_unea
	' If ButtonPressed = other_btn	Then second_page_display = other_unea
	' If ButtonPressed = main_btn		Then second_page_display = main_unea
	'
	' If ButtonPressed = caf_q_13_btn Then
	' 	page_display = show_q_13
	' End If
	' If ButtonPressed = caf_q_14_15_btn Then
	' 	page_display = show_q_14_15
	' End If
	' If ButtonPressed = caf_q_16_17_18_btn Then
	' 	page_display = show_q_16_18
	' End If
	' If ButtonPressed = caf_q_19_btn Then
	' 	page_display = show_q_19
	' End If
	' If ButtonPressed = caf_q_20_21_btn Then
	' 	page_display = show_q_20_21
	' 	second_page_display = main_asset
	' End If
	' If ButtonPressed = cash_btn		Then second_page_display = cash_asset
	' If ButtonPressed = acct_btn		Then second_page_display = acct_asset
	' If ButtonPressed = secu_btn		Then second_page_display = secu_asset
	' If ButtonPressed = cars_btn		Then second_page_display = cars_asset
	' If ButtonPressed = rest_btn		Then second_page_display = rest_asset
	' If ButtonPressed = main_asset_btn		Then second_page_display = main_asset
	' If ButtonPressed = caf_q_22_btn Then
	' 	page_display = show_q_22
	' End If
	' If ButtonPressed = caf_q_23_btn Then
	' 	page_display = show_q_23
	' End If
	' If ButtonPressed = caf_q_24_btn Then
	' 	page_display = show_q_24
	' End If
	' If ButtonPressed = caf_qual_q_btn Then
	' 	page_display = show_qual
	' End If
	' If ButtonPressed = caf_last_page_btn Then
	' 	page_display = show_pg_last
	' End If
	' If ButtonPressed = finish_interview_btn then leave_loop = TRUE
	'
	' If page_display <> show_pg_memb_info AND page_display <> show_pg_imig Then memb_selected = ""
	' If page_display <> show_q_7 Then stwk_selected = ""
	' If page_display <> show_q_12 AND page_display <> show_q_20_21 Then memb_to_match = ""
	' If page_display <> show_q_19 Then fmed_selected = ""
	'
	' If page_display <> show_q_20_21 Then inst_to_match = ""
	If ButtonPressed = finish_interview_btn Then leave_loop = TRUE
	If ButtonPressed > 10000 Then
		save_button = ButtonPressed
		If ButtonPressed = page_1_step_1_btn Then call explain_dialog_actions("PAGE 1", "STEP 1")
		If ButtonPressed = page_1_step_2_btn Then call explain_dialog_actions("PAGE 1", "STEP 2")

		ButtonPressed = save_button
	End If

end function

function get_state_name_from_state_code(state_code, state_name, include_state_code)
    If state_code = "NB" Then state_name = "MN Newborn"							'This is the list of all the states connected to the code.
    If state_code = "FC" Then state_name = "Foreign Country"
    If state_code = "UN" Then state_name = "Unknown"
    If state_code = "AL" Then state_name = "Alabama"
    If state_code = "AK" Then state_name = "Alaska"
    If state_code = "AZ" Then state_name = "Arizona"
    If state_code = "AR" Then state_name = "Arkansas"
    If state_code = "CA" Then state_name = "California"
    If state_code = "CO" Then state_name = "Colorado"
    If state_code = "CT" Then state_name = "Connecticut"
    If state_code = "DE" Then state_name = "Delaware"
    If state_code = "DC" Then state_name = "District Of Columbia"
    If state_code = "FL" Then state_name = "Florida"
    If state_code = "GA" Then state_name = "Georgia"
    If state_code = "HI" Then state_name = "Hawaii"
    If state_code = "ID" Then state_name = "Idaho"
    If state_code = "IL" Then state_name = "Illnois"
    If state_code = "IN" Then state_name = "Indiana"
    If state_code = "IA" Then state_name = "Iowa"
    If state_code = "KS" Then state_name = "Kansas"
    If state_code = "KY" Then state_name = "Kentucky"
    If state_code = "LA" Then state_name = "Louisiana"
    If state_code = "ME" Then state_name = "Maine"
    If state_code = "MD" Then state_name = "Maryland"
    If state_code = "MA" Then state_name = "Massachusetts"
    If state_code = "MI" Then state_name = "Michigan"
	If state_code = "MN" Then state_name = "Minnesota"
    If state_code = "MS" Then state_name = "Mississippi"
    If state_code = "MO" Then state_name = "Missouri"
    If state_code = "MT" Then state_name = "Montana"
    If state_code = "NE" Then state_name = "Nebraska"
    If state_code = "NV" Then state_name = "Nevada"
    If state_code = "NH" Then state_name = "New Hampshire"
    If state_code = "NJ" Then state_name = "New Jersey"
    If state_code = "NM" Then state_name = "New Mexico"
    If state_code = "NY" Then state_name = "New York"
    If state_code = "NC" Then state_name = "North Carolina"
    If state_code = "ND" Then state_name = "North Dakota"
    If state_code = "OH" Then state_name = "Ohio"
    If state_code = "OK" Then state_name = "Oklahoma"
    If state_code = "OR" Then state_name = "Oregon"
    If state_code = "PA" Then state_name = "Pennsylvania"
    If state_code = "RI" Then state_name = "Rhode Island"
    If state_code = "SC" Then state_name = "South Carolina"
    If state_code = "SD" Then state_name = "South Dakota"
    If state_code = "TN" Then state_name = "Tennessee"
    If state_code = "TX" Then state_name = "Texas"
    If state_code = "UT" Then state_name = "Utah"
    If state_code = "VT" Then state_name = "Vermont"
    If state_code = "VA" Then state_name = "Virginia"
    If state_code = "WA" Then state_name = "Washington"
    If state_code = "WV" Then state_name = "West Virginia"
    If state_code = "WI" Then state_name = "Wisconsin"
    If state_code = "WY" Then state_name = "Wyoming"
    If state_code = "PR" Then state_name = "Puerto Rico"
    If state_code = "VI" Then state_name = "Virgin Islands"

    If include_state_code = TRUE Then state_name = state_code & " " & state_name	'This adds the code to the state name if seelected
end function

function split_phone_number_into_parts(phone_variable, phone_left, phone_mid, phone_right)
'This function is to take the information provided as a phone number and split it up into the 3 parts
    phone_variable = trim(phone_variable)
    If phone_variable <> "" Then
        phone_variable = replace(phone_variable, "(", "")						'formatting the phone variable to get rid of symbols and spaces
        phone_variable = replace(phone_variable, ")", "")
        phone_variable = replace(phone_variable, "-", "")
        phone_variable = replace(phone_variable, " ", "")
        phone_variable = trim(phone_variable)
        phone_left = left(phone_variable, 3)									'reading the certain sections of the variable for each part.
        phone_mid = mid(phone_variable, 4, 3)
        phone_right = right(phone_variable, 4)
        phone_variable = "(" & phone_left & ")" & phone_mid & "-" & phone_right
    End If
end function

function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
'This function will asses the variables provided as the footer month and year to be sure it is correct.
    If IsNumeric(footer_month) = FALSE Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be a number, review and reenter the footer month information."
    Else
        footer_month = footer_month * 1
        If footer_month > 12 OR footer_month < 1 Then err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be between 1 and 12. Review and reenter the footer month information."
        footer_month = right("00" & footer_month, 2)
    End If

    If len(footer_year) < 2 Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be at least 2 characters long, review and reenter the footer year information."
    Else
        If IsNumeric(footer_year) = FALSE Then
            err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be a number, review and reenter the footer year information."
        Else
            footer_year = right("00" & footer_year, 2)
        End If
    End If
end function

function save_your_work()
'This function records the variables into a txt file so that it can be retrieved by the script if run later.

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	If MAXIS_case_number <> "" Then
		local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	Else
		local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"
	End If
	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then
			.DeleteFile(local_changelog_path)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

		If .FileExists(local_changelog_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)

			'Write the contents of the text file
			objTextStream.WriteLine "PRE - ATC - " & all_the_clients
			objTextStream.WriteLine "PRE - WHO - " & who_are_we_completing_the_form_with

			objTextStream.WriteLine "EXP - 1 - " & exp_q_1_income_this_month
			objTextStream.WriteLine "EXP - 2 - " & exp_q_2_assets_this_month
			objTextStream.WriteLine "EXP - 3 - RENT - " & exp_q_3_rent_this_month
			objTextStream.WriteLine "EXP - 3 - HEAT - " & exp_pay_heat_checkbox
			objTextStream.WriteLine "EXP - 3 - ACON - " & exp_pay_ac_checkbox
			objTextStream.WriteLine "EXP - 3 - ELEC - " & exp_pay_electricity_checkbox
			objTextStream.WriteLine "EXP - 3 - PHON - " & exp_pay_phone_checkbox
			objTextStream.WriteLine "EXP - 3 - NONE - " & exp_pay_none_checkbox
			objTextStream.WriteLine "EXP - 4 - " & exp_migrant_seasonal_formworker_yn
			objTextStream.WriteLine "EXP - 5 - PREV - " & exp_received_previous_assistance_yn
			objTextStream.WriteLine "EXP - 5 - WHEN - " & exp_previous_assistance_when
			objTextStream.WriteLine "EXP - 5 - WHER - " & exp_previous_assistance_where
			objTextStream.WriteLine "EXP - 5 - WHAT - " & exp_previous_assistance_what
			objTextStream.WriteLine "EXP - 6 - PREG - " & exp_pregnant_yn
			objTextStream.WriteLine "EXP - 6 - WHO? - " & exp_pregnant_who

			objTextStream.WriteLine "ADR - RESI - STR - " & resi_addr_street_full
			objTextStream.WriteLine "ADR - RESI - CIT - " & resi_addr_city
			objTextStream.WriteLine "ADR - RESI - STA - " & resi_addr_state
			objTextStream.WriteLine "ADR - RESI - ZIP - " & resi_addr_zip

			objTextStream.WriteLine "ADR - RESI - RES - " & reservation_yn
			objTextStream.WriteLine "ADR - RESI - NAM - " & reservation_name

			objTextStream.WriteLine "ADR - RESI - HML - " & homeless_yn

			objTextStream.WriteLine "ADR - RESI - LIV - " & living_situation

			objTextStream.WriteLine "ADR - MAIL - STR - " & mail_addr_street_full
			objTextStream.WriteLine "ADR - MAIL - CIT - " & mail_addr_city
			objTextStream.WriteLine "ADR - MAIL - STA - " & mail_addr_state
			objTextStream.WriteLine "ADR - MAIL - ZIP - " & mail_addr_zip

			objTextStream.WriteLine "ADR - PHON - NON - " & phone_one_number
			objTextStream.WriteLine "ADR - PHON - TON - " & phone_pne_type
			objTextStream.WriteLine "ADR - PHON - NTW - " & phone_two_number
			objTextStream.WriteLine "ADR - PHON - TTW - " & phone_two_type
			objTextStream.WriteLine "ADR - PHON - NTH - " & phone_three_number
			objTextStream.WriteLine "ADR - PHON - TTH - " & phone_three_type

			objTextStream.WriteLine "ADR - DATE - " & address_change_date
			objTextStream.WriteLine "ADR - CNTY - " & resi_addr_county

			objTextStream.WriteLine "01A - " & question_1_yn
			objTextStream.WriteLine "01N - " & question_1_notes
			objTextStream.WriteLine "01V - " & question_1_verif_yn
			objTextStream.WriteLine "01D - " & question_1_verif_details

			objTextStream.WriteLine "02A - " & question_2_yn
			objTextStream.WriteLine "02N - " & question_2_notes
			objTextStream.WriteLine "02V - " & question_2_verif_yn
			objTextStream.WriteLine "02D - " & question_2_verif_details

			objTextStream.WriteLine "03A - " & question_3_yn
			objTextStream.WriteLine "03N - " & question_3_notes
			objTextStream.WriteLine "03V - " & question_3_verif_yn
			objTextStream.WriteLine "03D - " & question_3_verif_details

			objTextStream.WriteLine "04A - " & question_4_yn
			objTextStream.WriteLine "04N - " & question_4_notes
			objTextStream.WriteLine "04V - " & question_4_verif_yn
			objTextStream.WriteLine "04D - " & question_4_verif_details

			objTextStream.WriteLine "05A - " & question_5_yn
			objTextStream.WriteLine "05N - " & question_5_notes
			objTextStream.WriteLine "05V - " & question_5_verif_yn
			objTextStream.WriteLine "05D - " & question_5_verif_details

			objTextStream.WriteLine "06A - " & question_6_yn
			objTextStream.WriteLine "06N - " & question_6_notes
			objTextStream.WriteLine "06V - " & question_6_verif_yn
			objTextStream.WriteLine "06D - " & question_6_verif_details

			objTextStream.WriteLine "07A - " & question_7_yn
			objTextStream.WriteLine "07N - " & question_7_notes
			objTextStream.WriteLine "07V - " & question_7_verif_yn
			objTextStream.WriteLine "07D - " & question_7_verif_details

			objTextStream.WriteLine "08A - " & question_8_yn
			objTextStream.WriteLine "08N - " & question_8_notes
			objTextStream.WriteLine "08V - " & question_8_verif_yn
			objTextStream.WriteLine "08D - " & question_8_verif_details

			objTextStream.WriteLine "09A - " & question_9_yn
			objTextStream.WriteLine "09N - " & question_9_notes
			objTextStream.WriteLine "09V - " & question_9_verif_yn
			objTextStream.WriteLine "09D - " & question_9_verif_details

			objTextStream.WriteLine "10A - " & question_10_yn
			objTextStream.WriteLine "10N - " & question_10_notes
			objTextStream.WriteLine "10V - " & question_10_verif_yn
			objTextStream.WriteLine "10D - " & question_10_verif_details
			objTextStream.WriteLine "10G - " & question_10_monthly_earnings

			objTextStream.WriteLine "11A - " & question_11_yn
			objTextStream.WriteLine "11N - " & question_11_notes
			objTextStream.WriteLine "11V - " & question_11_verif_yn
			objTextStream.WriteLine "11D - " & question_11_verif_details

			objTextStream.WriteLine "PWE - " & pwe_selection

			objTextStream.WriteLine "12A - RS - " & question_12_yn
			objTextStream.WriteLine "12A - SS - " & question_12_yn
			objTextStream.WriteLine "12A - VA - " & question_12_yn
			objTextStream.WriteLine "12A - UI - " & question_12_yn
			objTextStream.WriteLine "12A - WC - " & question_12_yn
			objTextStream.WriteLine "12A - RT - " & question_12_yn
			objTextStream.WriteLine "12A - TP - " & question_12_yn
			objTextStream.WriteLine "12A - CS - " & question_12_yn
			objTextStream.WriteLine "12A - OT - " & question_12_yn
			objTextStream.WriteLine "12N - " & question_12_notes
			objTextStream.WriteLine "12V - " & question_12_verif_yn
			objTextStream.WriteLine "12D - " & question_12_verif_details

			objTextStream.WriteLine "13A - " & question_13_yn
			objTextStream.WriteLine "13N - " & question_13_notes
			objTextStream.WriteLine "13V - " & question_13_verif_yn
			objTextStream.WriteLine "13D - " & question_13_verif_details

			objTextStream.WriteLine "14A - RT - " &  question_14_rent_yn
			objTextStream.WriteLine "14A - SB - " &  question_14_subsidy_yn
			objTextStream.WriteLine "14A - MT - " &  question_14_mortgage_yn
			objTextStream.WriteLine "14A - AS - " &  question_14_association_yn
			objTextStream.WriteLine "14A - IN - " &  question_14_insurance_yn
			objTextStream.WriteLine "14A - RM - " &  question_14_room_yn
			objTextStream.WriteLine "14A - TX - " &  question_14_taxes_yn
			objTextStream.WriteLine "14N - " & question_14_notes
			objTextStream.WriteLine "14V - " & question_14_verif_yn
			objTextStream.WriteLine "14D - " & question_14_verif_details

			objTextStream.WriteLine "15A - HA - " & question_15_heat_ac_yn
			objTextStream.WriteLine "15A - EL - " & question_15_electricity_yn
			objTextStream.WriteLine "15A - CF - " & question_15_cooking_fuel_yn
			objTextStream.WriteLine "15A - WS - " & question_15_water_and_sewer_yn
			objTextStream.WriteLine "15A - GR - " & question_15_garbage_yn
			objTextStream.WriteLine "15A - PN - " & question_15_phone_yn
			objTextStream.WriteLine "15A - LP - " & question_15_liheap_yn
			objTextStream.WriteLine "15N - " & question_15_notes
			objTextStream.WriteLine "15V - " & question_15_verif_yn
			objTextStream.WriteLine "15D - " & question_15_verif_details

			objTextStream.WriteLine "16A - " & question_16_yn
			objTextStream.WriteLine "16N - " & question_16_notes
			objTextStream.WriteLine "16V - " & question_16_verif_yn
			objTextStream.WriteLine "16D - " & question_16_verif_details

			objTextStream.WriteLine "17A - " & question_17_yn
			objTextStream.WriteLine "17N - " & question_17_notes
			objTextStream.WriteLine "17V - " & question_17_verif_yn
			objTextStream.WriteLine "17D - " & question_17_verif_details

			objTextStream.WriteLine "18A - " & question_18_yn
			objTextStream.WriteLine "18N - " & question_18_notes
			objTextStream.WriteLine "18V - " & question_18_verif_yn
			objTextStream.WriteLine "18D - " & question_18_verif_details

			objTextStream.WriteLine "19A - " & question_19_yn
			objTextStream.WriteLine "19N - " & question_19_notes
			objTextStream.WriteLine "19V - " & question_19_verif_yn
			objTextStream.WriteLine "19D - " & question_19_verif_details

			objTextStream.WriteLine "20A - CA - " & question_20_cash_yn
			objTextStream.WriteLine "20A - AC - " & question_20_acct_yn
			objTextStream.WriteLine "20A - SE - " & question_20_secu_yn
			objTextStream.WriteLine "20A - CR - " & question_20_cars_yn
			objTextStream.WriteLine "20N - " & question_20_notes
			objTextStream.WriteLine "20V - " & question_20_verif_yn
			objTextStream.WriteLine "20D - " & question_20_verif_details

			objTextStream.WriteLine "21A - " & question_21_yn
			objTextStream.WriteLine "21N - " & question_21_notes
			objTextStream.WriteLine "21V - " & question_21_verif_yn
			objTextStream.WriteLine "21D - " & question_21_verif_details

			objTextStream.WriteLine "22A - " & question_22_yn
			objTextStream.WriteLine "22N - " & question_22_notes
			objTextStream.WriteLine "22V - " & question_22_verif_yn
			objTextStream.WriteLine "22D - " & question_22_verif_details

			objTextStream.WriteLine "23A - " & question_23_yn
			objTextStream.WriteLine "23N - " & question_23_notes
			objTextStream.WriteLine "23V - " & question_23_verif_yn
			objTextStream.WriteLine "23D - " & question_23_verif_details

			objTextStream.WriteLine "24A - RP - " & question_24_rep_payee_yn
			objTextStream.WriteLine "24A - GF - " & question_24_guardian_fees_yn
			objTextStream.WriteLine "24A - SD - " & question_24_special_diet_yn
			objTextStream.WriteLine "24A - HH - " & question_24_high_housing_yn
			objTextStream.WriteLine "24N - " & question_24_notes
			objTextStream.WriteLine "24V - " & question_24_verif_yn
			objTextStream.WriteLine "24D - " & question_24_verif_details

			objTextStream.WriteLine "QQ1A - " & qual_question_one
			objTextStream.WriteLine "QQ1M - " & qual_memb_one
			objTextStream.WriteLine "QQ2A - " & qual_question_two
			objTextStream.WriteLine "QQ2M - " & qual_memb_two
			objTextStream.WriteLine "QQ3A - " & qual_question_three
			objTextStream.WriteLine "QQ3M - " & qual_memb_there
			objTextStream.WriteLine "QQ4A - " & qual_question_four
			objTextStream.WriteLine "QQ4M - " & qual_memb_four
			objTextStream.WriteLine "QQ5A - " & qual_question_five
			objTextStream.WriteLine "QQ5M - " & qual_memb_five

			For known_membs = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				objTextStream.WriteLine "ARR - ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_last_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_first_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_other_names, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_dob, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_gender, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_former_state, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_citizen, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_written_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_notes, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
			Next

			for this_jobs = 0 to UBOUND(JOBS_ARRAY, 2)
				objTextStream.WriteLine "ARR - JOBS_ARRAY - " & JOBS_ARRAY(jobs_employee_name, this_jobs)&"~"&JOBS_ARRAY(jobs_hourly_wage, this_jobs)&"~"&JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)&"~"&JOBS_ARRAY(jobs_employer_name, this_jobs)&"~"&JOBS_ARRAY(jobs_notes, this_jobs)
			Next

			'Close the object so it can be opened again shortly
			objTextStream.Close

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

function restore_your_work(vars_filled)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run
'
	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	If MAXIS_case_number <> "" Then local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	If no_case_number_checkbox = checked Then local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			If pull_variables = vbYes Then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_caf_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE

				array_counters = 0
				known_membs = 0
				known_jobs = 0
				For Each text_line in saved_caf_details
					' MsgBox "~" & left(text_line, 9) & "~"
					' MsgBox text_line
					If left(text_line, 9) = "PRE - WHO" Then who_are_we_completing_the_form_with = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - ATC" Then all_the_clients = Mid(text_line, 13)
					If left(text_line, 7) = "EXP - 1" Then exp_q_1_income_this_month = Mid(text_line, 11)
					If left(text_line, 7) = "EXP - 2" Then exp_q_2_assets_this_month = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 3 - RENT" Then exp_q_3_rent_this_month = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - HEAT" Then exp_pay_heat_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - ACON" Then exp_pay_ac_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - ELEC" Then exp_pay_electricity_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - PHON" Then exp_pay_phone_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - NONE" Then exp_pay_none_checkbox = Mid(text_line, 18)
					If left(text_line, 7) = "EXP - 4" Then exp_migrant_seasonal_formworker_yn = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 5 - PREV" Then exp_received_previous_assistance_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHEN" Then exp_previous_assistance_when = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHER" Then exp_previous_assistance_where = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHAT" Then exp_previous_assistance_what = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - PREG" Then exp_pregnant_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - WHO?" Then exp_pregnant_who = Mid(text_line, 18)
					If left(text_line, 3) = "ADR" Then
						' MsgBox "~" & mid(text_line, 7, 10) & "~"
						If mid(text_line, 7, 10) = "RESI - STR" Then resi_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - CIT" Then resi_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - STA" Then resi_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - ZIP" Then resi_addr_zip = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - RES" Then reservation_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - NAM" Then reservation_name = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - HML" Then homeless_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - LIV" Then living_situation = MID(text_line, 20)

						If mid(text_line, 7, 10) = "MAIL - STR" Then mail_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - CIT" Then mail_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - STA" Then mail_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - ZIP" Then mail_addr_zip = MID(text_line, 20)

						If mid(text_line, 7, 10) = "PHON - NON" Then phone_one_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TON" Then phone_one_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTW" Then phone_two_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTW" Then phone_two_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTH" Then phone_three_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTH" Then phone_three_type = MID(text_line, 20)

						If mid(text_line, 7, 4) = "DATE" Then address_change_date = MID(text_line, 13)
						If mid(text_line, 7, 4) = "CNTY" Then resi_addr_county = MID(text_line, 13)

					End If
					' If left(text_line, 3) = "" Then  = Mid(text_line, 7)
					If left(text_line, 3) = "01A" Then question_1_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01N" Then question_1_notes = Mid(text_line, 7)
					If left(text_line, 3) = "01V" Then question_1_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01D" Then question_1_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "02A" Then question_2_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02N" Then question_2_notes = Mid(text_line, 7)
					If left(text_line, 3) = "02V" Then question_2_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02D" Then question_2_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "03A" Then question_3_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03N" Then question_3_notes = Mid(text_line, 7)
					If left(text_line, 3) = "03V" Then question_3_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03D" Then question_3_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "04A" Then question_4_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04N" Then question_4_notes = Mid(text_line, 7)
					If left(text_line, 3) = "04V" Then question_4_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04D" Then question_4_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "05A" Then question_5_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05N" Then question_5_notes = Mid(text_line, 7)
					If left(text_line, 3) = "05V" Then question_5_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05D" Then question_5_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "06A" Then question_6_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06N" Then question_6_notes = Mid(text_line, 7)
					If left(text_line, 3) = "06V" Then question_6_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06D" Then question_6_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "07A" Then question_7_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07N" Then question_7_notes = Mid(text_line, 7)
					If left(text_line, 3) = "07V" Then question_7_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07D" Then question_7_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "08A" Then question_8_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08N" Then question_8_notes = Mid(text_line, 7)
					If left(text_line, 3) = "08V" Then question_8_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08D" Then question_8_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "09A" Then question_9_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09N" Then question_9_notes = Mid(text_line, 7)
					If left(text_line, 3) = "09V" Then question_9_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09D" Then question_9_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "10A" Then question_10_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10N" Then question_10_notes = Mid(text_line, 7)
					If left(text_line, 3) = "10V" Then question_10_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10D" Then question_10_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "10G" Then question_10_monthly_earnings = Mid(text_line, 7)

					If left(text_line, 3) = "11A" Then question_11_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11N" Then question_11_notes = Mid(text_line, 7)
					If left(text_line, 3) = "11V" Then question_11_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11D" Then question_11_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "PWE" Then pwe_selection = Mid(text_line, 7)

					If left(text_line, 8) = "12A - RS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - SS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - VA" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - UI" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - WC" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - RT" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - TP" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - CS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - OT" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 3) = "12N" Then question_12_notes = Mid(text_line, 7)
					If left(text_line, 3) = "12V" Then question_12_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "12D" Then question_12_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "13A" Then question_13_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13N" Then question_13_notes = Mid(text_line, 7)
					If left(text_line, 3) = "13V" Then question_13_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13D" Then question_13_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "14A - RT" Then  question_14_rent_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - SB" Then  question_14_subsidy_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - MT" Then  question_14_mortgage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - AS" Then  question_14_association_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - IN" Then  question_14_insurance_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - RM" Then  question_14_room_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - TX" Then  question_14_taxes_yn = Mid(text_line, 12)
					If left(text_line, 3) = "14N" Then question_14_notes = Mid(text_line, 7)
					If left(text_line, 3) = "14V" Then question_14_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "14D" Then question_14_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "15A - HA" Then question_15_heat_ac_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - EL" Then question_15_electricity_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - CF" Then question_15_cooking_fuel_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - WS" Then question_15_water_and_sewer_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - GR" Then question_15_garbage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - PN" Then question_15_phone_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - LP" Then question_15_liheap_yn = Mid(text_line, 12)
					If left(text_line, 3) = "15N" Then question_15_notes = Mid(text_line, 7)
					If left(text_line, 3) = "15V" Then question_15_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "15D" Then question_15_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "16A" Then question_16_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16N" Then question_16_notes = Mid(text_line, 7)
					If left(text_line, 3) = "16V" Then question_16_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16D" Then question_16_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "17A" Then question_17_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17N" Then question_17_notes = Mid(text_line, 7)
					If left(text_line, 3) = "17V" Then question_17_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17D" Then question_17_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "18A" Then question_18_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18N" Then question_18_notes = Mid(text_line, 7)
					If left(text_line, 3) = "18V" Then question_18_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18D" Then question_18_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "19A" Then question_19_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19N" Then question_19_notes = Mid(text_line, 7)
					If left(text_line, 3) = "19V" Then question_19_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19D" Then question_19_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "20A - CA" Then question_20_cash_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - AC" Then question_20_acct_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - SE" Then question_20_secu_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - CR" Then question_20_cars_yn = Mid(text_line, 12)
					If left(text_line, 3) = "20N" Then question_20_notes = Mid(text_line, 7)
					If left(text_line, 3) = "20V" Then question_20_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "20D" Then question_20_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "21A" Then question_21_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21N" Then question_21_notes = Mid(text_line, 7)
					If left(text_line, 3) = "21V" Then question_21_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21D" Then question_21_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "22A" Then question_22_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22N" Then question_22_notes = Mid(text_line, 7)
					If left(text_line, 3) = "22V" Then question_22_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22D" Then question_22_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "23A" Then question_23_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23N" Then question_23_notes = Mid(text_line, 7)
					If left(text_line, 3) = "23V" Then question_23_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23D" Then question_23_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "24A - RP" Then question_24_rep_payee_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - GF" Then question_24_guardian_fees_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - SD" Then question_24_special_diet_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - HH" Then question_24_high_housing_yn = Mid(text_line, 12)
					If left(text_line, 3) = "24N" Then question_24_notes = Mid(text_line, 7)
					If left(text_line, 3) = "24V" Then question_24_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "24D" Then question_24_verif_details = Mid(text_line, 7)

					If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ1M" Then qual_memb_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2A" Then qual_question_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2M" Then qual_memb_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3A" Then qual_question_three = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3M" Then qual_memb_there = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4A" Then qual_question_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4M" Then qual_memb_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5A" Then qual_question_five = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5M" Then qual_memb_five = Mid(text_line, 8)

					' If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)

					If left(text_line, 3) = "ARR" Then
						If MID(text_line, 7, 17) = "ALL_CLIENTS_ARRAY" Then
							array_info = Mid(text_line, 27)
							array_info = split(array_info, "~")
							ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, known_membs)
							ALL_CLIENTS_ARRAY(memb_last_name, known_membs) 				= array_info(0)
							ALL_CLIENTS_ARRAY(memb_first_name, known_membs) 			= array_info(1)
							ALL_CLIENTS_ARRAY(memb_mid_name, known_membs) 				= array_info(2)
							ALL_CLIENTS_ARRAY(memb_other_names, known_membs) 			= array_info(3)
							ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs) 				= array_info(4)
							ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs) 			= array_info(5)
							ALL_CLIENTS_ARRAY(memb_dob, known_membs) 					= array_info(6)
							ALL_CLIENTS_ARRAY(memb_gender, known_membs) 				= array_info(7)
							ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs) 			= array_info(8)
							ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs) 		= array_info(9)
							ALL_CLIENTS_ARRAY(memi_last_grade, known_membs) 			= array_info(10)
							ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs) 			= array_info(11)
							ALL_CLIENTS_ARRAY(memi_former_state, known_membs) 			= array_info(12)
							ALL_CLIENTS_ARRAY(memi_citizen, known_membs) 				= array_info(13)
							ALL_CLIENTS_ARRAY(memb_interpreter, known_membs) 			= array_info(14)
							ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs) 		= array_info(15)
							ALL_CLIENTS_ARRAY(memb_written_language, known_membs) 		= array_info(16)
							ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs) 				= array_info(17)
							ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs) 		= array_info(18)
							ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs) 		= array_info(19)
							ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs) 		= array_info(20)
							ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs) 		= array_info(21)
							ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs) 		= array_info(22)
							ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs) 			= array_info(23)
							ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs) 			= array_info(24)
							ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs) 			= array_info(25)
							ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs) 			= array_info(26)
							ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs) 	= array_info(27)
							ALL_CLIENTS_ARRAY(clt_imig_status, known_membs) 			= array_info(28)
							ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs) 				= array_info(29)
							ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs) 				= array_info(30)
							ALL_CLIENTS_ARRAY(clt_verif_details, known_membs) 			= array_info(31)
							ALL_CLIENTS_ARRAY(memb_notes, known_membs) 					= array_info(32)
							ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs) 				= array_info(33)
							known_membs = known_membs + 1
						End If

						If MID(text_line, 7, 10) = "JOBS_ARRAY" Then
							array_info = Mid(text_line, 20)
							array_info = split(array_info, "~")
							ReDim Preserve JOBS_ARRAY(jobs_notes, known_jobs)
							JOBS_ARRAY(jobs_employee_name, known_jobs) 			= array_info(0)
							JOBS_ARRAY(jobs_hourly_wage, known_jobs) 			= array_info(1)
							JOBS_ARRAY(jobs_gross_monthly_earnings, known_jobs) = array_info(2)
							JOBS_ARRAY(jobs_employer_name, known_jobs) 			= array_info(3)
							JOBS_ARRAY(jobs_notes, known_jobs) 					= array_info(4)
							known_jobs = known_jobs + 1
						End If
					End If
				Next
			End If
		End If
	End With
end function

'THESE FUNCTIONS ARE ALL THE INDIVIDUAL DIALOGS WITHIN THE MAIN DIALOG LOOP
function dlg_page_one_pers_and_exp()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 416, 240, "CAF Person and Expedited"
			  ComboBox 205, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_form_with, who_are_we_completing_the_form_with
			  EditBox 290, 65, 50, 15, exp_q_1_income_this_month
			  EditBox 310, 85, 50, 15, exp_q_2_assets_this_month
			  EditBox 250, 105, 50, 15, exp_q_3_rent_this_month
			  CheckBox 125, 125, 30, 10, "Heat", exp_pay_heat_checkbox
			  CheckBox 160, 125, 65, 10, "Air Conditioning", exp_pay_ac_checkbox
			  CheckBox 230, 125, 45, 10, "Electricity", exp_pay_electricity_checkbox
			  CheckBox 280, 125, 35, 10, "Phone", exp_pay_phone_checkbox
			  CheckBox 325, 125, 35, 10, "None", exp_pay_none_checkbox
			  DropListBox 245, 140, 40, 45, "No"+chr(9)+"Yes", exp_migrant_seasonal_formworker_yn
			  DropListBox 365, 155, 40, 45, "No"+chr(9)+"Yes", exp_received_previous_assistance_yn
			  EditBox 80, 175, 80, 15, exp_previous_assistance_when
			  EditBox 200, 175, 85, 15, exp_previous_assistance_where
			  EditBox 320, 175, 85, 15, exp_previous_assistance_what
			  DropListBox 160, 195, 40, 45, "No"+chr(9)+"Yes", exp_pregnant_yn
			  ComboBox 255, 195, 150, 45, all_the_clients, exp_pregnant_who
			  ButtonGroup ButtonPressed
				PushButton 305, 220, 50, 15, "Next", next_btn
			    CancelButton 360, 220, 50, 15
			  Text 70, 15, 130, 10, "Who are you completing the form with?"
			  GroupBox 10, 50, 400, 165, "Expedited Questions - Do you need help right away?"
			  Text 20, 70, 270, 10, "1. How much income (cash or checkes) did or will your household get this month?"
			  Text 20, 90, 290, 10, "2. How much does your household (including children) have cash, checking or savings?"
			  Text 20, 110, 225, 10, "3. How much does your household pay for rent/mortgage per month?"
			  Text 30, 125, 90, 10, "What utilities do you pay?"
			  Text 20, 145, 225, 10, "4. Is anyone in your household a migrant or seasonal farm worker?"
			  Text 20, 160, 345, 10, "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
			  Text 30, 180, 50, 10, "If yes, When?"
			  Text 170, 180, 30, 10, "Where?"
			  Text 295, 180, 25, 10, "What?"
			  Text 20, 200, 135, 10, "6. Is anyone in your household pregnant?"
			  Text 210, 200, 40, 10, "If yes, who?"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

		Loop until ButtonPressed = next_btn
	Loop until err_msg = ""
	If exp_pregnant_who = "Select or Type" Then exp_pregnant_who = ""

	show_caf_pg_1_pers_dlg = FALSE
	caf_pg_1_pers_dlg_cleared = TRUE
end function


function dlg_page_one_address()

	If resi_addr_street_full = blank Then show_known_addr = FALSE
	If resi_addr_county = "" Then resi_addr_county = "27 Hennepin"
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""

			If show_known_addr = TRUE Then
				BeginDialog Dialog1, 0, 0, 391, 285, "CAF Address"
				  Text 70, 55, 305, 15, resi_addr_street_full
				  Text 70, 75, 105, 15, resi_addr_city
				  Text 205, 75, 110, 45, resi_addr_state
				  Text 340, 75, 35, 15, resi_addr_zip
				  Text 125, 95, 45, 45, reservation_yn
				  Text 245, 85, 130, 15, reservation_name
				  Text 125, 115, 45, 45, homeless_yn
				  Text 245, 115, 130, 45, living_situation
				  Text 70, 155, 305, 15, mail_addr_street_full
				  Text 70, 175, 105, 15, mail_addr_city
				  Text 205, 175, 110, 45, mail_addr_state
				  Text 340, 175, 35, 15, mail_addr_zip
				  Text 20, 225, 90, 15, phone_one_number
				  Text 125, 225, 65, 45, phone_pne_type
				  Text 20, 245, 90, 15, phone_two_number
				  Text 125, 245, 65, 45, phone_two_type
				  Text 20, 265, 90, 15, phone_three_number
				  Text 125, 265, 65, 45, phone_three_type
				  Text 325, 205, 50, 15, address_change_date
				  Text 255, 240, 120, 45, resi_addr_county
				  ButtonGroup ButtonPressed
					PushButton 280, 265, 50, 15, "Next", next_btn
					PushButton 280, 253, 50, 10, "Back", back_btn
				    CancelButton 335, 265, 50, 15
				    PushButton 290, 20, 95, 15, "Update Information", update_information_btn
					' PushButton 290, 20, 95, 15, "Save Information", save_information_btn
				    PushButton 325, 135, 50, 10, "CLEAR", clear_mail_addr_btn
					PushButton 205, 220, 35, 10, "CLEAR", clear_phone_one_btn
				    PushButton 205, 240, 35, 10, "CLEAR", clear_phone_two_btn
				    PushButton 205, 260, 35, 10, "CLEAR", clear_phone_three_btn
				  Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
				  GroupBox 10, 190, 235, 90, "Phone Number"
				  Text 20, 55, 45, 10, "House/Street"
				  Text 45, 75, 20, 10, "City"
				  Text 185, 75, 20, 10, "State"
				  Text 325, 75, 15, 10, "Zip"
				  Text 20, 95, 100, 10, "Do you live on a Reservation?"
				  Text 180, 95, 60, 10, "If yes, which one?"
				  Text 30, 115, 90, 10, "Client Indicates Homeless:"
				  Text 185, 115, 60, 10, "Living Situation?"
				  GroupBox 10, 35, 375, 95, "Residence Address"
				  Text 20, 155, 45, 10, "House/Street"
				  Text 45, 175, 20, 10, "City"
				  Text 185, 175, 20, 10, "State"
				  Text 325, 175, 15, 10, "Zip"
				  GroupBox 10, 125, 375, 70, "Mailing Address"
				  Text 20, 205, 50, 10, "Number"
				  Text 125, 205, 25, 10, "Type"
				  Text 255, 205, 60, 10, "Date of Change:"
				  Text 255, 225, 75, 10, "County of Residence:"
				EndDialog
			End If

			If show_known_addr = FALSE Then
				BeginDialog Dialog1, 0, 0, 391, 285, "CAF Address"
				  EditBox 70, 50, 305, 15, resi_addr_street_full
				  EditBox 70, 70, 105, 15, resi_addr_city
				  DropListBox 205, 70, 110, 45, state_list, resi_addr_state
				  EditBox 340, 70, 35, 15, resi_addr_zip
				  DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", reservation_yn
				  EditBox 245, 90, 130, 15, reservation_name
				  DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", homeless_yn
				  DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
				  EditBox 70, 150, 305, 15, mail_addr_street_full
				  EditBox 70, 170, 105, 15, mail_addr_city
				  DropListBox 205, 170, 110, 45, state_list, mail_addr_state
				  EditBox 340, 170, 35, 15, mail_addr_zip
				  EditBox 20, 220, 90, 15, phone_one_number
				  DropListBox 125, 220, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_pne_type
				  EditBox 20, 240, 90, 15, phone_two_number
				  DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_two_type
				  EditBox 20, 260, 90, 15, phone_three_number
				  DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_three_type
				  EditBox 325, 200, 50, 15, address_change_date
				  DropListBox 255, 235, 120, 45, county_list, resi_addr_county
				  ButtonGroup ButtonPressed
					PushButton 280, 265, 50, 15, "Next", next_btn
					PushButton 280, 253, 50, 10, "Back", back_btn
				    CancelButton 335, 265, 50, 15
				    ' PushButton 290, 20, 95, 15, "Update Information", update_information_btn
					PushButton 290, 20, 95, 15, "Save Information", save_information_btn
				    PushButton 325, 135, 50, 10, "CLEAR", clear_mail_addr_btn
				    PushButton 205, 220, 35, 10, "CLEAR", clear_phone_one_btn
				    PushButton 205, 240, 35, 10, "CLEAR", clear_phone_two_btn
				    PushButton 205, 260, 35, 10, "CLEAR", clear_phone_three_btn
				  Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
				  GroupBox 10, 190, 235, 90, "Phone Number"
				  Text 20, 55, 45, 10, "House/Street"
				  Text 45, 75, 20, 10, "City"
				  Text 185, 75, 20, 10, "State"
				  Text 325, 75, 15, 10, "Zip"
				  Text 20, 95, 100, 10, "Do you live on a Reservation?"
				  Text 180, 95, 60, 10, "If yes, which one?"
				  Text 30, 115, 90, 10, "Client Indicates Homeless:"
				  Text 185, 115, 60, 10, "Living Situation?"
				  GroupBox 10, 35, 375, 95, "Residence Address"
				  Text 20, 155, 45, 10, "House/Street"
				  Text 45, 175, 20, 10, "City"
				  Text 185, 175, 20, 10, "State"
				  Text 325, 175, 15, 10, "Zip"
				  GroupBox 10, 125, 375, 70, "Mailing Address"
				  Text 20, 205, 50, 10, "Number"
				  Text 125, 205, 25, 10, "Type"
				  Text 255, 205, 60, 10, "Date of Change:"
				  Text 255, 225, 75, 10, "County of Residence:"
				EndDialog
			End If

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE
			Call validate_phone_number(err_msg, "*", phone_one_number, TRUE)
			Call validate_phone_number(err_msg, "*", phone_two_number, TRUE)
			Call validate_phone_number(err_msg, "*", phone_three_number, TRUE)

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = update_information_btn Then show_known_addr = FALSE
			If ButtonPressed = save_information_btn Then show_known_addr = TRUE
			If ButtonPressed = clear_mail_addr_btn Then
				mail_addr_street_full = ""
				mail_addr_city = ""
				mail_addr_state = "Select One..."
				mail_addr_zip = ""
			End If
			If ButtonPressed = clear_phone_one_btn Then
				phone_one_number = ""
				phone_pne_type = "Select One..."
			End If
			If ButtonPressed = clear_phone_two_btn Then
				phone_two_number = ""
				phone_two_type = "Select One..."
			End If
			If ButtonPressed = clear_phone_three_btn Then
				phone_three_number = ""
				phone_three_type = "Select One..."
			End If
			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_1_pers_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_1_addr_dlg = FALSE
		caf_pg_1_addr_dlg_cleared = TRUE
	End If
end function

function dlg_page_two_household_comp()

	known_membs = 0
	shown_known_pers_detail = TRUE
	If ALL_CLIENTS_ARRAY(memb_last_name, known_membs) = "" Then shown_known_pers_detail = FALSE
	go_back = FALSE
	Do
		Do
			btn_placeholder = 3001
			For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb) = btn_placeholder
				btn_placeholder = btn_placeholder + 1
			Next

			err_msg = ""
			Dialog1 = ""

			If shown_known_pers_detail = TRUE Then
				BeginDialog Dialog1, 0, 0, 550, 385, "Household Member Information"
				  Text 20, 45, 105, 15, ALL_CLIENTS_ARRAY(memb_last_name, known_membs)
				  Text 130, 45, 90, 15, ALL_CLIENTS_ARRAY(memb_first_name, known_membs)
				  Text 225, 45, 70, 15, ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)
				  Text 300, 45, 175, 15, ALL_CLIENTS_ARRAY(memb_other_names, known_membs)
				  If ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs) = "V - System Verified" Then
					  Text 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  Else
					  EditBox 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  End If
				  Text 95, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_dob, known_membs)
				  Text 170, 75, 50, 45, ALL_CLIENTS_ARRAY(memb_gender, known_membs)
				  Text 225, 75, 140, 45, ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)
				  Text 370, 75, 105, 45, ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)
				  Text 20, 105, 130, 15, ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)
				  Text 155, 105, 70, 15, ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)
				  Text 230, 105, 165, 15, ALL_CLIENTS_ARRAY(memi_former_state, known_membs)
				  Text 400, 105, 75, 45, ALL_CLIENTS_ARRAY(memi_citizen, known_membs)
				  Text 20, 135, 60, 45, ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)
				  Text 90, 135, 170, 15, ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)
				  Text 90, 165, 170, 15, ALL_CLIENTS_ARRAY(memb_written_language, known_membs)
				  Text 280, 155, 40, 45, ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)
				  CheckBox 330, 155, 30, 10, "Asian", ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)
				  CheckBox 330, 165, 30, 10, "Black", ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)
				  CheckBox 330, 175, 120, 10, "American Indian or Alaska Native", ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)
				  CheckBox 330, 185, 130, 10, "Pacific Islander and Native Hawaiian", ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)
				  CheckBox 330, 195, 130, 10, "White", ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)
				  CheckBox 20, 195, 50, 10, "SNAP (food)", ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)
				  CheckBox 75, 195, 65, 10, "Cash programs", ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)
				  CheckBox 145, 195, 85, 10, "Emergency Assistance", ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)
				  CheckBox 235, 195, 30, 10, "NONE", ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)
				  DropListBox 15, 230, 80, 45, "Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)
				  EditBox 100, 230, 205, 15, ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)
				  DropListBox 310, 230, 55, 45, "No"+chr(9)+"Yes", ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)
				  DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
				  EditBox 100, 260, 435, 15, ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)
				  EditBox 15, 290, 350, 15, ALL_CLIENTS_ARRAY(memb_notes, known_membs)
				  ButtonGroup ButtonPressed
					PushButton 430, 290, 50, 15, "Next", next_btn
					PushButton 375, 295, 50, 10, "Back", back_btn
					CancelButton 485, 290, 50, 15
					PushButton 330, 5, 95, 15, "Update Information", update_information_btn
					' PushButton 330, 5, 95, 15, "Save Information", save_information_btn
					y_pos = 35
					For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
						If the_memb = known_membs Then
							Text 498, y_pos + 1, 45, 10, "Person " & (the_memb + 1)
						Else
							PushButton 490, y_pos, 45, 10, "Person " & (the_memb + 1), ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb)
						End If
						y_pos = y_pos + 10
					Next
					PushButton 490, 230, 45, 10, "Add Person", add_person_btn
				  Text 10, 10, 315, 10, "Known Member in MAXIS - current information. If there are changes, press this button to update:"
				  If ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs) = "" Then
					  GroupBox 10, 25, 475, 190, "Person " & known_membs+1
				  Else
					  GroupBox 10, 25, 475, 190, "Person " & known_membs+1 & " - MEMBER " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				  End If
				  Text 20, 35, 50, 10, "Last Name"
				  Text 130, 35, 50, 10, "First Name"
				  Text 225, 35, 50, 10, "Middle Name"
				  Text 300, 35, 50, 10, "Other Names"
				  Text 20, 65, 55, 10, "Soc Sec Number"
				  Text 95, 65, 45, 10, "Date of Birth"
				  Text 170, 65, 45, 10, "Gender"
				  Text 225, 65, 90, 10, "Relationship to MEMB 01"
				  Text 370, 65, 50, 10, "Marital Status"
				  Text 20, 95, 75, 10, "Last Grade Completed"
				  Text 155, 95, 55, 10, "Moved to MN on"
				  Text 230, 95, 65, 10, "Moved to MN from"
				  Text 400, 95, 75, 10, "US Citizen or National"
				  Text 20, 125, 40, 10, "Interpreter?"
				  Text 90, 125, 95, 10, "Preferred Spoken Language"
				  Text 90, 155, 95, 10, "Preferred Written Language"
				  GroupBox 270, 135, 205, 75, "Demographics - OPTIONAL"
				  Text 280, 145, 35, 10, "Hispanic?"
				  Text 330, 145, 50, 10, "Race"
				  Text 20, 185, 145, 10, "Which programs is this person requesting?"
				  Text 15, 220, 80, 10, "Intends to reside in MN"
				  Text 100, 220, 65, 10, "Immigration Status"
				  Text 310, 220, 50, 10, "Sponsor?"
				  Text 15, 250, 50, 10, "Verification"
				  Text 100, 250, 65, 10, "Verification Details"
				  Text 15, 280, 50, 10, "Notes:"
				EndDialog

			End If

			If shown_known_pers_detail = FALSE Then

				BeginDialog Dialog1, 0, 0, 550, 385, "Household Member Information"
				  EditBox 20, 45, 105, 15, ALL_CLIENTS_ARRAY(memb_last_name, known_membs)
				  EditBox 130, 45, 90, 15, ALL_CLIENTS_ARRAY(memb_first_name, known_membs)
				  EditBox 225, 45, 70, 15, ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)
				  EditBox 300, 45, 175, 15, ALL_CLIENTS_ARRAY(memb_other_names, known_membs)
				  EditBox 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  EditBox 95, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_dob, known_membs)
				  DropListBox 170, 75, 50, 45, ""+chr(9)+"Male"+chr(9)+"Female", ALL_CLIENTS_ARRAY(memb_gender, known_membs)
				  DropListBox 225, 75, 140, 45, memb_panel_relationship_list, ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)
				  DropListBox 370, 75, 105, 45, marital_status, ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)
				  EditBox 20, 105, 130, 15, ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)
				  EditBox 155, 105, 70, 15, ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)
				  EditBox 230, 105, 165, 15, ALL_CLIENTS_ARRAY(memi_former_state, known_membs)
				  DropListBox 400, 105, 75, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memi_citizen, known_membs)
				  DropListBox 20, 135, 60, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)
				  EditBox 90, 135, 170, 15, ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)
				  EditBox 90, 165, 170, 15, ALL_CLIENTS_ARRAY(memb_written_language, known_membs)
				  DropListBox 280, 155, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)
				  CheckBox 330, 155, 30, 10, "Asian", ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)
				  CheckBox 330, 165, 30, 10, "Black", ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)
				  CheckBox 330, 175, 120, 10, "American Indian or Alaska Native", ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)
				  CheckBox 330, 185, 130, 10, "Pacific Islander and Native Hawaiian", ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)
				  CheckBox 330, 195, 130, 10, "White", ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)
				  CheckBox 20, 195, 50, 10, "SNAP (food)", ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)
				  CheckBox 75, 195, 65, 10, "Cash programs", ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)
				  CheckBox 145, 195, 85, 10, "Emergency Assistance", ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)
				  CheckBox 235, 195, 30, 10, "NONE", ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)
				  DropListBox 15, 230, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)
				  EditBox 100, 230, 205, 15, ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)
				  DropListBox 310, 230, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)
				  DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
				  EditBox 100, 260, 435, 15, ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)
				  EditBox 15, 290, 350, 15, ALL_CLIENTS_ARRAY(memb_notes, known_membs)
				  ButtonGroup ButtonPressed
				    PushButton 430, 290, 50, 15, "Next", next_btn
					PushButton 375, 295, 50, 10, "Back", back_btn
				    CancelButton 485, 290, 50, 15
				    ' PushButton 330, 5, 95, 15, "Update Information", update_information_btn
					PushButton 330, 5, 95, 15, "Save Information", save_information_btn
					y_pos = 35
					For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
						PushButton 490, y_pos, 45, 10, "Person " & (the_memb + 1), ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb)
						y_pos = y_pos + 10
					Next
				    PushButton 490, 230, 45, 10, "Add Person", add_person_btn
				  Text 10, 10, 315, 10, "Known Member in MAXIS - current information. If there are changes, press this button to update:"
				  GroupBox 10, 25, 475, 190, "MEMBER " &  ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				  Text 20, 35, 50, 10, "Last Name"
				  Text 130, 35, 50, 10, "First Name"
				  Text 225, 35, 50, 10, "Middle Name"
				  Text 300, 35, 50, 10, "Other Names"
				  Text 20, 65, 55, 10, "Soc Sec Number"
				  Text 95, 65, 45, 10, "Date of Birth"
				  Text 170, 65, 45, 10, "Gender"
				  Text 225, 65, 90, 10, "Relationship to MEMB 01"
				  Text 370, 65, 50, 10, "Marital Status"
				  Text 20, 95, 75, 10, "Last Grade Completed"
				  Text 155, 95, 55, 10, "Moved to MN on"
				  Text 230, 95, 65, 10, "Moved to MN from"
				  Text 400, 95, 75, 10, "US Citizen or National"
				  Text 20, 125, 40, 10, "Interpreter?"
				  Text 90, 125, 95, 10, "Preferred Spoken Language"
				  Text 90, 155, 95, 10, "Preferred Written Language"
				  GroupBox 270, 135, 205, 75, "Demographics - OPTIONAL"
				  Text 280, 145, 35, 10, "Hispanic?"
				  Text 330, 145, 50, 10, "Race"
				  Text 20, 185, 145, 10, "Which programs is this person requesting?"
				  Text 15, 220, 80, 10, "Intends to reside in MN"
				  Text 100, 220, 65, 10, "Immigration Status"
				  Text 310, 220, 50, 10, "Sponsor?"
				  Text 15, 250, 50, 10, "Verification"
				  Text 100, 250, 65, 10, "Verification Details"
				  Text 15, 280, 50, 10, "Notes:"
				EndDialog

			End If

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = next_btn Then
				known_membs = known_membs + 1
				If known_membs =< UBound(ALL_CLIENTS_ARRAY, 2) Then ButtonPressed = ""
			End If
			If ButtonPressed = update_information_btn Then shown_known_pers_detail = FALSE
			If ButtonPressed = save_information_btn Then shown_known_pers_detail = TRUE
			For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				If ButtonPressed = ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb) Then known_membs = the_memb
			Next
			If ButtonPressed = add_person_btn Then
				last_clt = UBound(ALL_CLIENTS_ARRAY, 2)
				new_clt = last_clt + 1
				ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, new_clt)
				known_membs = new_clt
			End If
			If ButtonPressed = back_btn Then
				If known_membs = 0 Then
					go_back = TRUE
					ButtonPressed = next_btn
					err_msg = ""
					show_caf_pg_1_addr_dlg = TRUE
				Else
					known_membs = known_membs - 1
				End If
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_2_hhcomp_dlg = FALSE
		caf_pg_2_hhcomp_dlg_cleared = TRUE
	End If

end function

function dlg_page_three_household_info()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "Tell Us About Your Household"
			  DropListBox 10, 10, 60, 45, question_answers, question_1_yn
			  EditBox 120, 20, 235, 15, question_1_notes
			  DropListBox 10, 45, 60, 45, question_answers, question_2_yn
			  EditBox 120, 65, 235, 15, question_2_notes
			  DropListBox 10, 90, 60, 45, question_answers, question_3_yn
			  EditBox 120, 100, 235, 15, question_3_notes
			  DropListBox 10, 125, 60, 45, question_answers, question_4_yn
			  EditBox 120, 145, 235, 15, question_4_notes
			  DropListBox 10, 170, 60, 45, question_answers, question_5_yn
			  EditBox 120, 190, 235, 15, question_5_notes
			  DropListBox 10, 215, 60, 45, question_answers, question_6_yn
			  EditBox 120, 225, 235, 15, question_6_notes
			  DropListBox 10, 250, 60, 45, question_answers, question_7_yn
			  EditBox 120, 280, 235, 15, question_7_notes
			  ButtonGroup ButtonPressed
			    PushButton 360, 285, 50, 15, "Next", next_btn
			    PushButton 360, 275, 50, 10, "Back", back_btn
			    CancelButton 415, 285, 50, 15
			    PushButton 380, 20, 75, 10, "ADD VERIFICATION", add_verif_1_btn
			    PushButton 380, 55, 75, 10, "ADD VERIFICATION", add_verif_2_btn
			    PushButton 380, 100, 75, 10, "ADD VERIFICATION", add_verif_3_btn
			    PushButton 380, 135, 75, 10, "ADD VERIFICATION", add_verif_4_btn
			    PushButton 380, 180, 75, 10, "ADD VERIFICATION", add_verif_5_btn
			    PushButton 380, 225, 75, 10, "ADD VERIFICATION", add_verif_6_btn
			    PushButton 380, 260, 75, 10, "ADD VERIFICATION", add_verif_7_btn
			  Text 80, 10, 230, 10, "1. Does everyone in your household buy, fix or eat food with you?"
			  Text 95, 25, 25, 10, "Notes:"
			  Text 360, 10, 100, 10, "Q1 - Verification - " & question_1_verif_yn
			  Text 80, 45, 245, 10, "2. Is anyone in the household, who is age 60 or over or disabled, unable to "
			  Text 90, 55, 115, 10, "buy or fix food due to a disability?"
			  Text 95, 70, 25, 10, "Notes:"
			  Text 360, 45, 100, 10, "Q2 - Verification - " & question_2_verif_yn
			  Text 80, 90, 165, 10, "3. Is anyone in the household attending school?"
			  Text 95, 105, 25, 10, "Notes:"
			  Text 360, 90, 100, 10, "Q3 - Verification - " & question_3_verif_yn
			  Text 80, 125, 230, 10, "4. Is anyone in your household temporarily not living in your home?"
			  Text 90, 135, 230, 10, "(for example: vacation, foster care, treatment, hospital, job search)"
			  Text 95, 150, 25, 10, "Notes:"
			  Text 360, 125, 100, 10, "Q4 - Verification - " & question_4_verif_yn
			  Text 80, 170, 255, 10, "5. Is anyone blind, or does anyone have a physical or mental health condition"
			  Text 90, 180, 185, 10, " that limits the ability to work or perform daily activities?"
			  Text 95, 195, 25, 10, "Notes:"
			  Text 360, 170, 100, 10, "Q5 - Verification - " & question_5_verif_yn
			  Text 80, 215, 245, 10, "6. Is anyone unable to work for reasons other than illness or disability?"
			  Text 95, 230, 25, 10, "Notes:"
			  Text 360, 215, 100, 10, "Q6 - Verification - " & question_6_verif_yn
			  Text 80, 250, 170, 10, "7. In the last 60 days did anyone in the household: "
			  Text 90, 260, 165, 20, "- Stop working or quit a job?   - Refuse a job offer? - Ask to work fewer hours?   - Go on strike?"
			  Text 95, 285, 25, 10, "Notes:"
			  Text 360, 250, 100, 10, "Q7 - Verification - " & question_7_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_1_btn Then Call verif_details_dlg(1)
			If ButtonPressed = add_verif_2_btn Then Call verif_details_dlg(2)
			If ButtonPressed = add_verif_3_btn Then Call verif_details_dlg(3)
			If ButtonPressed = add_verif_4_btn Then Call verif_details_dlg(4)
			If ButtonPressed = add_verif_5_btn Then Call verif_details_dlg(5)
			If ButtonPressed = add_verif_6_btn Then Call verif_details_dlg(6)
			If ButtonPressed = add_verif_7_btn Then Call verif_details_dlg(7)

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_2_hhcomp_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_3_hhinfo_dlg = FALSE
		caf_pg_3_hhinfo_dlg_cleared = TRUE
	End If

end function

function dlg_page_four_income()
	go_back = FALSE
	Do
		Do
			err_msg = ""

			btn_placeholder = 4000
			dlg_len = 350
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				JOBS_ARRAY(jobs_edit_btn, each_job) = btn_placeholder
				btn_placeholder = btn_placeholder + 1
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then dlg_len = dlg_len + 10
			next

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 650, dlg_len, "What kinds of income do you have?"
			  DropListBox 10, 10, 60, 45, question_answers, question_8_yn
			  Text 80, 10, 290, 10, "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			  Text 540, 10, 105, 10, "Q8 - Verification - " & question_8_verif_yn
			  ButtonGroup ButtonPressed
			    PushButton 560, 20, 75, 10, "ADD VERIFICATION", add_verif_8_btn
			  DropListBox 10, 25, 60, 45, question_answers, question_8a_yn
			  Text 90, 25, 350, 10, "a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			  Text 95, 40, 25, 10, "Notes:"
			  EditBox 120, 35, 390, 15, question_8_notes
			  DropListBox 10, 55, 60, 45, question_answers, question_9_yn
			  Text 80, 55, 350, 10, "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			  ButtonGroup ButtonPressed
			    PushButton 430, 55, 55, 10, "ADD JOB", add_job_btn
			  Text 540, 55, 105, 10, "Q9 - Verification - " & question_9_verif_yn
			  ButtonGroup ButtonPressed
			    PushButton 560, 65, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			  y_pos = 65
			  ' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			  for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				  ' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				  If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then

					  Text 95, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					  ButtonGroup ButtonPressed
					    PushButton 495, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					  y_pos = y_pos + 10
				  End If
			  next
			  y_pos = y_pos + 10
			  Text 95, y_pos, 25, 10, "Notes:"
			  EditBox 120, y_pos - 5, 390, 15, question_9_notes
			  y_pos = y_pos + 15
			  DropListBox 10, y_pos, 60, 45, question_answers, question_10_yn
			  Text 80, y_pos, 430, 10, "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			  Text 540, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_10_btn
			  Text 95, y_pos, 85, 10, "Gross Monthly Earnings:"
			  Text 185, y_pos, 25, 10, "Notes:"
			  y_pos = y_pos + 10
			  EditBox 95, y_pos, 80, 15, question_10_monthly_earnings
			  EditBox 185, y_pos, 325, 15, question_10_notes
			  y_pos = y_pos + 20
			  DropListBox 10, y_pos, 60, 45, question_answers, question_11_yn
			  Text 80, y_pos, 255, 10, "11. Do you expect any changes in income, expenses or work hours?"
			  Text 540, y_pos, 105, 10, "Q11 - Verification - " & question_11_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_11_btn
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_11_notes
			  y_pos = y_pos + 25
			  Text 80, y_pos, 75, 10, "Pricipal Wage Earner"
			  DropListBox 155, y_pos - 5, 175, 45, all_the_clients, pwe_selection
			  y_pos = y_pos + 10
			  Text 80, y_pos + 5, 370, 10, "12. Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			  Text 540, y_pos, 105, 10, "Q12 - Verification - " & question_12_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_12_btn
			  y_pos = y_pos + 10
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_rsdi_yn
			  Text 150, y_pos + 5, 70, 10, "RSDI                      $"
			  EditBox 220, y_pos, 35, 15, question_12_rsdi_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ssi_yn
			  Text 375, y_pos + 5, 85, 10, "SSI                                 $"
			  EditBox 460, y_pos, 35, 15, question_12_ssi_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_va_yn
			  Text 150, y_pos + 5, 70, 10, "VA                          $"
			  EditBox 220, y_pos, 35, 15, question_12_va_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ui_yn
			  Text 375, y_pos + 5, 85, 10, "Unemployment Ins          $"
			  EditBox 460, y_pos, 35, 15, question_12_ui_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_wc_yn
			  Text 150, y_pos + 5, 70, 10, "Workers Comp       $"
			  EditBox 220, y_pos, 35, 15, question_12_wc_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ret_yn
			  Text 375, y_pos + 5, 85, 10, "Retirement Ben.              $"
			  EditBox 460, y_pos, 35, 15, question_12_ret_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_trib_yn
			  Text 150, y_pos + 5, 70, 10, "Tribal Payments      $"
			  EditBox 220, y_pos, 35, 15, question_12_trib_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_cs_yn
			  Text 375, y_pos + 5, 85, 10, "Child/Spousal Support    $"
			  EditBox 460, y_pos, 35, 15, question_12_cs_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_other_yn
			  Text 150, y_pos + 5, 110, 10, "Other unearned income          $"
			  EditBox 250, y_pos, 35, 15, question_12_other_amt
			  y_pos = y_pos + 20
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_12_notes
			  y_pos = y_pos + 25
			  DropListBox 10, y_pos, 60, 45, question_answers, question_13_yn
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 80, y_pos, 400, 10, "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			  Text 540, y_pos, 105, 10, "Q13 - Verification - " & question_13_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
				PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_13_btn
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_13_notes
			  y_pos = y_pos + 20
			  ButtonGroup ButtonPressed
			    PushButton 540, y_pos, 50, 15, "Next", next_btn
			    PushButton 485, y_pos + 5, 50, 10, "Back", back_btn
			    CancelButton 595, y_pos, 50, 15
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_8_btn Then Call verif_details_dlg(8)
			If ButtonPressed = add_verif_9_btn Then Call verif_details_dlg(9)
			If ButtonPressed = add_verif_10_btn Then Call verif_details_dlg(10)
			If ButtonPressed = add_verif_11_btn Then Call verif_details_dlg(11)
			If ButtonPressed = add_verif_12_btn Then Call verif_details_dlg(12)
			If ButtonPressed = add_verif_13_btn Then Call verif_details_dlg(13)

			If ButtonPressed = add_job_btn Then
				another_job = ""
				count = 0
				for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
					count = count + 1
					If JOBS_ARRAY(jobs_employer_name, each_job) = "" AND JOBS_ARRAY(jobs_employee_name, each_job) = "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) = "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) = "" Then
						another_job = each_job
					End If
				Next
				If another_job = "" Then
					another_job = count
					ReDim Preserve JOBS_ARRAY(jobs_notes, another_job)
				End If
				Call jobs_details_dlg(another_job)
			End If

			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				If ButtonPressed = JOBS_ARRAY(jobs_edit_btn, each_job) Then
					Call jobs_details_dlg(each_job)
				End If
			next

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_3_hhinfo_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_4_income_dlg = FALSE
		caf_pg_4_income_dlg_cleared = TRUE
	End If

end function

function dlg_page_five_expenses()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""

			BeginDialog Dialog1, 0, 0, 550, 385, "What kinds of expenses do you have?"
			  DropListBox 95, 20, 60, 45, question_answers, question_14_rent_yn
			  DropListBox 300, 20, 60, 45, question_answers, question_14_subsidy_yn
			  DropListBox 95, 35, 60, 45, question_answers, question_14_mortgage_yn
			  DropListBox 300, 35, 60, 45, question_answers, question_14_association_yn
			  DropListBox 95, 50, 60, 45, question_answers, question_14_insurance_yn
			  DropListBox 300, 50, 60, 45, question_answers, question_14_room_yn
			  DropListBox 95, 65, 60, 45, question_answers, question_14_taxes_yn
			  EditBox 135, 85, 390, 15, question_14_notes
			  DropListBox 95, 120, 60, 45, question_answers, question_15_heat_ac_yn
			  DropListBox 265, 120, 60, 45, question_answers, question_15_electricity_yn
			  DropListBox 415, 120, 60, 45, question_answers, question_15_cooking_fuel_yn
			  DropListBox 95, 135, 60, 45, question_answers, question_15_water_and_sewer_yn
			  DropListBox 265, 135, 60, 45, question_answers, question_15_garbage_yn
			  DropListBox 415, 135, 60, 45, question_answers, question_15_phone_yn
			  DropListBox 95, 150, 60, 45, question_answers, question_15_liheap_yn
			  EditBox 120, 165, 390, 15, question_15_notes
			  DropListBox 10, 190, 60, 45, question_answers, question_16_yn
			  EditBox 120, 210, 390, 15, question_16_notes
			  DropListBox 10, 235, 60, 45, question_answers, question_17_yn
			  EditBox 120, 255, 390, 15, question_17_notes
			  DropListBox 10, 280, 60, 45, question_answers, question_18_yn
			  EditBox 120, 300, 390, 15, question_18_notes
			  DropListBox 10, 325, 60, 45, question_answers, question_19_yn
			  EditBox 120, 335, 390, 15, question_19_notes
			  ButtonGroup ButtonPressed
			    PushButton 560, 355, 50, 15, "Next", next_btn
			    PushButton 505, 360, 50, 10, "Back", back_btn
			    CancelButton 615, 355, 50, 15
			    PushButton 580, 20, 75, 10, "ADD VERIFICATION", add_verif_14_btn
			    PushButton 580, 120, 75, 10, "ADD VERIFICATION", add_verif_15_btn
			    PushButton 580, 200, 75, 10, "ADD VERIFICATION", add_verif_16_btn
			    PushButton 580, 245, 75, 10, "ADD VERIFICATION", add_verif_17_btn
			    PushButton 580, 290, 75, 10, "ADD VERIFICATION", add_verif_18_btn
			    PushButton 580, 335, 75, 10, "ADD VERIFICATION", add_verif_19_btn
			  Text 80, 10, 220, 10, "14. Does your household have the following housing expenses?"
			  Text 165, 25, 70, 10, "Rent"
			  Text 370, 25, 100, 10, "Rent or Section 8 Subsidy"
			  Text 165, 40, 125, 10, "Mortgage/contract for deed payment"
			  Text 370, 40, 70, 10, "Association fees"
			  Text 165, 55, 85, 10, "Homeowner's insurance"
			  Text 370, 55, 70, 10, "Room and/or board"
			  Text 165, 70, 100, 10, "Real estate taxes"
			  Text 110, 90, 25, 10, "Notes:"
			  Text 560, 10, 105, 10, "Q14 - Verification - " & question_14_verif_yn
			  Text 80, 110, 290, 10, "15. Does your household have the following utility expenses any time during the year? "
			  Text 165, 125, 85, 10, "Heating/air conditioning"
			  Text 335, 125, 70, 10, "Electricity"
			  Text 485, 125, 70, 10, "Cooking fuel"
			  Text 165, 140, 75, 10, "Water and sewer"
			  Text 335, 140, 60, 10, "Garbage removal"
			  Text 485, 140, 70, 10, "Phone/cell phone"
			  Text 165, 155, 375, 10, "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"
			  Text 95, 170, 25, 10, "Notes:"
			  Text 560, 110, 105, 10, "Q15 - Verification - " & question_15_verif_yn
			  Text 80, 190, 345, 10, "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working,"
			  Text 95, 200, 125, 10, "looking for work or going to school?"
			  Text 95, 215, 25, 10, "Notes:"
			  Text 560, 190, 105, 10, "Q16 - Verification - " & question_16_verif_yn
			  Text 80, 235, 380, 10, "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working,"
			  Text 95, 245, 125, 10, "looking for work or going to school?"
			  Text 95, 260, 25, 10, "Notes:"
			  Text 560, 235, 105, 10, "Q17 - Verification - " & question_17_verif_yn
			  Text 80, 280, 430, 10, "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support"
			  Text 95, 290, 215, 10, "or contribute to a tax dependent who does not live in your home?"
			  Text 95, 305, 25, 10, "Notes:"
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 560, 280, 105, 10, "Q18 - Verification - " & question_18_verif_yn
			  Text 80, 325, 255, 10, "19. For SNAP only: Does anyone in the household have medical expenses? "
			  Text 95, 340, 25, 10, "Notes:"
			  Text 560, 325, 105, 10, "Q19 - Verification - " & question_19_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_14_btn Then Call verif_details_dlg(14)
			If ButtonPressed = add_verif_15_btn Then Call verif_details_dlg(15)
			If ButtonPressed = add_verif_16_btn Then Call verif_details_dlg(16)
			If ButtonPressed = add_verif_17_btn Then Call verif_details_dlg(17)
			If ButtonPressed = add_verif_18_btn Then Call verif_details_dlg(18)
			If ButtonPressed = add_verif_19_btn Then Call verif_details_dlg(19)


			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_4_income_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_5_expenses_dlg = FALSE
		caf_pg_5_expenses_dlg_cleared = TRUE
	End If

end function

function dlg_page_six_assets_and_other()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "What do you own? Other Information"
			  DropListBox 80, 25, 60, 45, question_answers, question_20_cash_yn
			  DropListBox 285, 25, 60, 45, question_answers, question_20_acct_yn
			  DropListBox 80, 40, 60, 45, question_answers, question_20_secu_yn
			  DropListBox 285, 40, 60, 45, question_answers, question_20_cars_yn
			  EditBox 120, 60, 390, 15, question_20_notes
			  DropListBox 10, 85, 60, 45, question_answers, question_21_yn
			  EditBox 120, 95, 390, 15, question_21_notes
			  DropListBox 10, 120, 60, 45, question_answers, question_22_yn
			  EditBox 120, 130, 390, 15, question_22_notes
			  DropListBox 10, 155, 60, 45, question_answers, question_23_yn
			  EditBox 120, 165, 390, 15, question_23_notes
			  DropListBox 80, 205, 60, 45, question_answers, question_24_rep_payee_yn
			  DropListBox 285, 205, 60, 45, question_answers, question_24_guardian_fees_yn
			  DropListBox 80, 220, 60, 45, question_answers, question_24_special_diet_yn
			  DropListBox 285, 220, 60, 45, question_answers, question_24_high_housing_yn
			  EditBox 120, 240, 390, 15, question_24_notes
			  ButtonGroup ButtonPressed
			    PushButton 540, 240, 50, 15, "Next", next_btn
			    PushButton 540, 230, 50, 10, "Back", back_btn
			    CancelButton 595, 240, 50, 15
			    PushButton 560, 20, 75, 10, "ADD VERIFICATION", add_verif_20_btn
			    PushButton 560, 95, 75, 10, "ADD VERIFICATION", add_verif_21_btn
			    PushButton 560, 130, 75, 10, "ADD VERIFICATION", add_verif_22_btn
			    PushButton 560, 165, 75, 10, "ADD VERIFICATION", add_verif_23_btn
			    PushButton 560, 200, 75, 10, "ADD VERIFICATION", add_verif_24_btn
			  Text 80, 10, 280, 10, "20. Does anyone in the household own, or is anyone buying, any of the following?"
			  Text 150, 30, 70, 10, "Cash"
			  Text 355, 30, 175, 10, "Bank accounts (savings, checking, debit card, etc.)"
			  Text 150, 45, 125, 10, "Stocks, bonds, annuities, 401k, etc."
			  Text 355, 45, 180, 10, "Vehicles (cars, trucks, motorcycles, campers, trailers)"
			  Text 95, 65, 25, 10, "Notes:"
			  Text 540, 10, 105, 10, "Q20 - Verification - " & question_20_verif_yn
			  Text 80, 85, 420, 10, "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? "
			  Text 95, 100, 25, 10, "Notes:"
			  Text 540, 85, 105, 10, "Q21 - Verification - " & question_21_verif_yn
			  Text 80, 120, 305, 10, "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			  Text 95, 135, 25, 10, "Notes:"
			  Text 540, 120, 105, 10, "Q22 - Verification - " & question_22_verif_yn
			  Text 80, 155, 250, 10, "23. For children under the age of 19, are both parents living in the home?"
			  Text 95, 170, 25, 10, "Notes:"
			  Text 540, 155, 105, 10, "Q23 - Verification - " & question_23_verif_yn
			  Text 80, 190, 325, 10, "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			  Text 150, 210, 95, 10, "Representative Payee fees"
			  Text 355, 210, 105, 10, "Guardian Conservator fees"
			  Text 150, 225, 125, 10, "Physician-perscribed special diet"
			  Text 355, 225, 105, 10, "High housing costs"
			  Text 95, 245, 25, 10, "Notes:"
			  Text 540, 190, 105, 10, "Q24 - Verification - " & question_24_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_20_btn Then Call verif_details_dlg(20)
			If ButtonPressed = add_verif_21_btn Then Call verif_details_dlg(21)
			If ButtonPressed = add_verif_22_btn Then Call verif_details_dlg(22)
			If ButtonPressed = add_verif_23_btn Then Call verif_details_dlg(23)
			If ButtonPressed = add_verif_24_btn Then Call verif_details_dlg(24)


			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_5_expenses_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_6_other_dlg = FALSE
		caf_pg_6_other_dlg_cleared = TRUE
	End If
end function


function dlg_qualifying_questions()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "CAF Qualifying Questions"
			  DropListBox 220, 40, 30, 45, "No"+chr(9)+"Yes", qual_question_one
			  ComboBox 340, 40, 105, 45, all_the_clients, qual_memb_one
			  DropListBox 220, 80, 30, 45, "No"+chr(9)+"Yes", qual_question_two
			  ComboBox 340, 80, 105, 45, all_the_clients, qual_memb_two
			  DropListBox 220, 110, 30, 45, "No"+chr(9)+"Yes", qual_question_three
			  ComboBox 340, 110, 105, 45, all_the_clients, qual_memb_there
			  DropListBox 220, 140, 30, 45, "No"+chr(9)+"Yes", qual_question_four
			  ComboBox 340, 140, 105, 45, all_the_clients, qual_memb_four
			  DropListBox 220, 160, 30, 45, "No"+chr(9)+"Yes", qual_question_five
			  ComboBox 340, 160, 105, 45, all_the_clients, qual_memb_five
			  ButtonGroup ButtonPressed
			    CancelButton 395, 185, 50, 15
			    PushButton 340, 185, 50, 15, "Next", next_btn
			    PushButton 285, 190, 50, 10, "Back", back_btn
			  Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
			  Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			  Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			  Text 10, 110, 195, 30, "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			  Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
			  Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			  Text 260, 40, 70, 10, "Household Member:"
			  Text 260, 80, 70, 10, "Household Member:"
			  Text 260, 110, 70, 10, "Household Member:"
			  Text 260, 140, 70, 10, "Household Member:"
			  Text 260, 160, 70, 10, "Household Member:"
			EndDialog

			dialog Dialog1

			cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_5_expenses_dlg = TRUE
			End If
		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_qual_questions_dlg = FALSE
		caf_caf_qual_questions_dlg_cleared = TRUE
	End If

end function


function dlg_signature()
	go_back = FALSE
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "Form dates and signatures"
			  EditBox 135, 50, 60, 15, caf_form_date
			  DropListBox 135, 70, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_yn
			  ButtonGroup ButtonPressed
			    PushButton 35, 90, 105, 15, "Complete CAF Form Detail", complete_caf_questions
			    PushButton 5, 95, 25, 10, "BACK", back_btn
			    PushButton 10, 35, 145, 10, "Open RIGHTS AND RESPONSIBLITIES ", open_r_and_r_button
			    CancelButton 145, 90, 50, 15
			  Text 10, 10, 160, 20, "Confirm the client is signing this form and attesting to the information provided verbally."
			  Text 70, 55, 55, 10, "CAF Form Date:"
			  Text 10, 75, 120, 10, "Cient signature accepted verbally?"
			EndDialog

			dialog Dialog1

			cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = complete_caf_questions
			If ButtonPressed = open_r_and_r_button Then open_URL_in_browser("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG")

			If IsDate(caf_form_date) = FALSE Then
				err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was received."
			Else
				If DateDiff("d", date, caf_form_date) > 0 Then err_msg = err_msg & vbNewLine & "* The date of the CAF form is listed as a future date, a form cannot be listed as received inthe future, please review the form date."
			End If
			If client_signed_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the client has signed the form correctly by selecting 'yes' or 'no'."

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = complete_caf_questions
				err_msg = ""
				show_caf_qual_questions_dlg = TRUE
			End If
		Loop until ButtonPressed = complete_caf_questions
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_sig_dlg = FALSE
		caf_sig_dlg_cleared = TRUE
	End If
end function

function read_all_the_MEMBs()
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

	clt_count = 0

	For each hh_clt in client_array

		ReDim Preserve HH_MEMB_ARRAY(clt_count)
		Set HH_MEMB_ARRAY(clt_count) = new mx_hh_member
		HH_MEMB_ARRAY(clt_count).ref_number = hh_clt
		HH_MEMB_ARRAY(clt_count).define_the_member
		HH_MEMB_ARRAY(clt_count).button_one = 500 + clt_count
		HH_MEMB_ARRAY(clt_count).button_two = 600 + clt_count
		memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(clt_count).ref_number & " - " & HH_MEMB_ARRAY(clt_count).full_name
		If HH_MEMB_ARRAY(clt_count).fs_pwe = "Yes" Then the_pwe_for_this_case = HH_MEMB_ARRAY(clt_count).ref_number & " - " & HH_MEMB_ARRAY(clt_count).full_name

		ReDim Preserve ALL_ANSWERS_ARRAY(ans_notes, clt_count)
		clt_count = clt_count + 1
	Next

	For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
		HH_MEMB_ARRAY(i).collect_parent_information

		If HH_MEMB_ARRAY(i).parent_one_in_home = TRUE AND HH_MEMB_ARRAY(i).parent_two_in_home = TRUE Then
			HH_MEMB_ARRAY(i).parents_in_home = "Both parents in the home"
		ElseIf HH_MEMB_ARRAY(i).parent_one_in_home = FALSE AND HH_MEMB_ARRAY(i).parent_two_in_home = FALSE Then
			HH_MEMB_ARRAY(i).parents_in_home = "Neither parent in the home"
		ElseIf HH_MEMB_ARRAY(i).parent_one_in_home = TRUE AND HH_MEMB_ARRAY(i).parent_two_in_home = FALSE Then
			HH_MEMB_ARRAY(i).parents_in_home = "1 parent in the home"
		ElseIf HH_MEMB_ARRAY(i).parent_one_in_home = FALSE AND HH_MEMB_ARRAY(i).parent_two_in_home = TRUE Then
			HH_MEMB_ARRAY(i).parents_in_home = "1 parent in the home"
		End If
	Next

	rela_counter = 0
	For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
		If HH_MEMB_ARRAY(i).rel_to_applcnt <> "Self" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Not Related" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Live-in Attendant" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Unknown" Then
			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(0).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = HH_MEMB_ARRAY(i).rel_to_applcnt

			rela_counter = rela_counter + 1

			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

		 	' MsgBox "Member Count - " & i & vbNewLine & "Relationship - " & HH_MEMB_ARRAY(i).rel_to_applcnt
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(0).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Spouse" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Spouse"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Child" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Parent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Parent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Child"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Sibling" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Sibling"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Sibling" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Sibling"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Child" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Parent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Parent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Child"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Aunt" OR HH_MEMB_ARRAY(i).rel_to_applcnt = "Uncle" Then
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Neice"
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Nephew"
			End If
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Nephew" OR HH_MEMB_ARRAY(i).rel_to_applcnt = "Neice" Then
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Aunt"
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Uncle"
			End If
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Cousin" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Cousin"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Grandparent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandchild"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Grandchild" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandparent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Other Relative" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

			rela_counter = rela_counter + 1
		End If
	Next

	Call navigate_to_MAXIS_screen("STAT", "SIBL")

	sibl_row = 7
	Do
		EMReadScreen sibl_group_nbr, 2, sibl_row, 28
		If sibl_group_nbr <> "__" Then
			sibl_col = 39
			Do
				EMReadScreen sibl_ref_nrb, 2, sibl_row, sibl_col
				If sibl_ref_nrb <> "__" Then
					list_of_siblings = list_of_siblings & "~" & sibl_ref_nrb
					sibl_col = sibl_col + 4
				End If
			Loop until sibl_ref_nrb = "__"
			' MsgBox "here"
			list_of_siblings = right(list_of_siblings, len(list_of_siblings) - 1)
			sibl_array = split(list_of_siblings, "~")
			For each memb_sibling in sibl_array
				For each other_sibling in sibl_array
					If memb_sibling <> other_sibling Then
						ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

						ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = memb_sibling
						ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = other_sibling
						ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Sibling"

						For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
							If HH_MEMB_ARRAY(i).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
							If HH_MEMB_ARRAY(i).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
						Next

						rela_counter = rela_counter + 1
					End If
				Next
			Next
		End If
		sibl_row = sibl_row + 1
	Loop until sibl_group_nbr = "__"

	Call navigate_to_MAXIS_screen("STAT", "PARE")

	'Need to add a way to find relationship verifications for MEMB 01
	EMWriteScreen HH_MEMB_ARRAY(0).ref_number, 20, 76
	transmit



	For i = 0 to UBound(HH_MEMB_ARRAY, 1)						'we start with 1 because 0 is MEMB 01 and that parental relationshipare all known because of MEMB
		EMWriteScreen HH_MEMB_ARRAY(i).ref_number, 20, 76
		transmit

		pare_row = 8
		Do
			EMReadScreen child_ref_nbr, 2, pare_row, 24
			If child_ref_nbr <> "__" Then
				If i = 0 Then
					For known_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
						If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, known_rela) = "01" AND child_ref_nbr = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, known_rela) THen
							EMReadScreen pare_verif, 2, pare_row, 71
							If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "BC - Birth Certificate"
							If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "AR - Adoption Records"
							If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "LG = Legal Guardian"
							If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "RE - Religious Records"
							If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "HR - Hospital Records"
							If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "RP - Recognition of Parentage"
							If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "OT - Other Verifciation"
							If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "NO - No Verif Provided"
						End If
					Next
				Else
					EMReadScreen pare_type, 1, pare_row, 53
					EMReadScreen pare_verif, 2, pare_row, 71

					ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = child_ref_nbr

					If pare_type = "1" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Parent"
					If pare_type = "2" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Parent"
					If pare_type = "3" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandparent"
					If pare_type = "4" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Relative Caregiver"
					If pare_type = "5" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Foster Parent"
					If pare_type = "6" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Not Related"
					If pare_type = "7" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Legal Guardian"
					If pare_type = "8" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

					If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "BC - Birth Certificate"
					If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "AR - Adoption Records"
					If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "LG = Legal Guardian"
					If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RE - Religious Records"
					If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "HR - Hospital Records"
					If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RP - Recognition of Parentage"
					If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "OT - Other Verifciation"
					If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "NO - No Verif Provided"

					for x = 0 to UBound(HH_MEMB_ARRAY, 1)
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
					Next

					rela_counter = rela_counter + 1

					ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = child_ref_nbr
					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number

					If pare_type = "1" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Child"
					If pare_type = "2" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Child"
					If pare_type = "3" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandchild"
					If pare_type = "4" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Relative Caregiver"
					If pare_type = "5" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Foster Child"
					If pare_type = "6" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Not Related"
					If pare_type = "7" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Legal Guardian"
					If pare_type = "8" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

					If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "BC - Birth Certificate"
					If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "AR - Adoption Records"
					If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "LG = Legal Guardian"
					If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RE - Religious Records"
					If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "HR - Hospital Records"
					If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RP - Recognition of Parentage"
					If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "OT - Other Verifciation"
					If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "NO - No Verif Provided"

					for x = 0 to UBound(HH_MEMB_ARRAY, 1)
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
					Next

					rela_counter = rela_counter + 1

				End If
			End If
			pare_row = pare_row + 1

		Loop until child_ref_nbr = "__"
	Next

	' For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
	' 	ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, the_rela)
	' 	ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, the_rela)
	'
	' 	' MsgBox "Relationship detail:" & vbNewLine & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) & " is the " & ALL_HH_RELATIONSHIPS_ARRAY(rela_type, the_rela) & " of " & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela)
	' Next
end function

function verif_details_dlg(question_number)
	Select Case question_number
		Case 1
			verif_selection = question_1_verif_yn
			verif_detials = question_1_verif_details
			question_words = "1. Does everyone in your household buy, fix or eat food with you?"
		Case 2
			verif_selection = question_2_verif_yn
			verif_detials = question_2_verif_details
			question_words = "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		Case 3
			verif_selection = question_3_verif_yn
			verif_detials = question_3_verif_details
			question_words = "3. Is anyone in the household attending school?"
		Case 4
			verif_selection = question_4_verif_yn
			verif_detials = question_4_verif_details
			question_words = "4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)"
		Case 5
			verif_selection = question_5_verif_yn
			verif_detials = question_5_verif_details
			question_words = "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
		Case 6
			verif_selection = question_6_verif_yn
			verif_detials = question_6_verif_details
			question_words = "6. Is anyone unable to work for reasons other than illness or disability?"
		Case 7
			verif_selection = question_7_verif_yn
			verif_detials = question_7_verif_details
			question_words = "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
		Case 8
			verif_selection = question_8_verif_yn
			verif_detials = question_8_verif_details
			question_words = "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
		Case 9
			verif_selection = question_9_verif_yn
			verif_detials = question_9_verif_details
			question_words = "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
		Case 10
			verif_selection = question_10_verif_yn
			verif_detials = question_10_verif_details
			question_words = "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
		Case 11
			verif_selection = question_11_verif_yn
			verif_detials = question_11_verif_details
			question_words = "11. Do you expect any changes in income, expenses or work hours?"
		Case 12
			verif_selection = question_12_verif_yn
			verif_detials = question_12_verif_details
			question_words = "12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
		Case 13
			verif_selection = question_13_verif_yn
			verif_detials = question_13_verif_details
			question_words = "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
		Case 14
			verif_selection = question_14_verif_yn
			verif_detials = question_14_verif_details
			question_words = "14. Does your household have the following housing expenses?"
		Case 15
			verif_selection = question_15_verif_yn
			verif_detials = question_15_verif_details
			question_words = "15. Does your household have the following utility expenses any time during the year?"
		Case 16
			verif_selection = question_16_verif_yn
			verif_detials = question_16_verif_details
			question_words = "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
		Case 17
			verif_selection = question_17_verif_yn
			verif_detials = question_17_verif_details
			question_words = "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
		Case 18
			verif_selection = question_18_verif_yn
			verif_detials = question_18_verif_details
			question_words = "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
		Case 19
			verif_selection = question_19_verif_yn
			verif_detials = question_19_verif_details
			question_words = "19. For SNAP only: Does anyone in the household have medical expenses? "
		Case 20
			verif_selection = question_20_verif_yn
			verif_detials = question_20_verif_details
			question_words = "20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
		Case 21
			verif_selection = question_21_verif_yn
			verif_detials = question_21_verif_details
			question_words = "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)"
		Case 22
			verif_selection = question_22_verif_yn
			verif_detials = question_22_verif_details
			question_words = "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
		Case 23
			verif_selection = question_23_verif_yn
			verif_detials = question_23_verif_details
			question_words = "23. For children under the age of 19, are both parents living in the home?"
		Case 24
			verif_selection = question_24_verif_yn
			verif_detials = question_24_verif_details
			question_words = "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
	End Select


	BeginDialog Dialog1, 0, 0, 396, 95, "Add Verification"
	  DropListBox 60, 35, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", verif_selection
	  EditBox 60, 55, 330, 15, verif_detials
	  ButtonGroup ButtonPressed
	    PushButton 340, 75, 50, 15, "Return", return_btn
		PushButton 145, 35, 50, 10, "CLEAR", clear_btn
	  Text 10, 10, 380, 20, question_words
	  Text 10, 40, 45, 10, "Verification: "
	  Text 20, 60, 30, 10, "Details:"
	EndDialog

	Do
		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = clear_btn Then
			verif_selection = "Not Needed"
			verif_detials = ""
		End If
	Loop until ButtonPressed = return_btn

	Select Case question_number
		Case 1
			question_1_verif_yn = verif_selection
			question_1_verif_details = verif_detials
		Case 2
			question_2_verif_yn = verif_selection
			question_2_verif_details = verif_detials
		Case 3
			question_3_verif_yn = verif_selection
			question_3_verif_details = verif_detials
		Case 4
			question_4_verif_yn = verif_selection
			question_4_verif_details = verif_detials
		Case 5
			question_5_verif_yn = verif_selection
			question_5_verif_details = verif_detials
		Case 6
			question_6_verif_yn = verif_selection
			question_6_verif_details = verif_detials
		Case 7
			question_7_verif_yn = verif_selection
			question_7_verif_details = verif_detials
		Case 8
			question_8_verif_yn = verif_selection
			question_8_verif_details = verif_detials
		Case 9
			question_9_verif_yn = verif_selection
			question_9_verif_details = verif_detials
		Case 10
			question_10_verif_yn = verif_selection
			question_10_verif_details = verif_detials
		Case 11
			question_11_verif_yn = verif_selection
			question_11_verif_details = verif_detials
		Case 12
			question_12_verif_yn = verif_selection
			question_12_verif_details = verif_detials
		Case 13
			question_13_verif_yn = verif_selection
			question_13_verif_details = verif_detials
		Case 14
			question_14_verif_yn = verif_selection
			question_14_verif_details = verif_detials
		Case 15
			question_15_verif_yn = verif_selection
			question_15_verif_details = verif_detials
		Case 16
			question_16_verif_yn = verif_selection
			question_16_verif_details = verif_detials
		Case 17
			question_17_verif_yn = verif_selection
			question_17_verif_details = verif_detials
		Case 18
			question_18_verif_yn = verif_selection
			question_18_verif_details = verif_detials
		Case 19
			question_19_verif_yn = verif_selection
			question_19_verif_details = verif_detials
		Case 20
			question_20_verif_yn = verif_selection
			question_20_verif_details = verif_detials
		Case 21
			question_21_verif_yn = verif_selection
			question_21_verif_details = verif_detials
		Case 22
			question_22_verif_yn = verif_selection
			question_22_verif_details = verif_detials
		Case 23
			question_23_verif_yn = verif_selection
			question_23_verif_details = verif_detials
		Case 24
			question_24_verif_yn = verif_selection
			question_24_verif_details = verif_detials
	End Select

end function

function jobs_details_dlg(this_jobs)
	Do
		pick_a_client = replace(all_the_clients, "Select or Type", "Select One...")
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 321, 130, "Add Job"
		  DropListBox 10, 35, 135, 45, pick_a_client+chr(9)+"", JOBS_ARRAY(jobs_employee_name, this_jobs)
		  EditBox 150, 35, 60, 15, JOBS_ARRAY(jobs_hourly_wage, this_jobs)
		  EditBox 215, 35, 100, 15, JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)
		  EditBox 10, 65, 305, 15, JOBS_ARRAY(jobs_employer_name, this_jobs)
		  EditBox 35, 90, 280, 15, JOBS_ARRAY(jobs_notes, this_jobs)
		  ButtonGroup ButtonPressed
		    PushButton 265, 110, 50, 15, "Return", return_btn
		    PushButton 265, 10, 50, 10, "CLEAR", clear_job_btn
		  Text 10, 10, 100, 10, "Enter Job Details/Information"
		  Text 10, 25, 70, 10, "EMPLOYEE NAME:"
		  Text 150, 25, 60, 10, "HOURLY WAGE:"
		  Text 215, 25, 105, 10, "GROSS MONTHLY EARNINGS:"
		  Text 10, 55, 110, 10, "EMPLOYER/BUSINESS NAME:"
		  Text 10, 95, 25, 10, "Notes:"
		EndDialog


		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = clear_job_btn Then
			JOBS_ARRAY(jobs_employee_name, this_jobs) = ""
			JOBS_ARRAY(jobs_hourly_wage, this_jobs) = ""
			JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs) = ""
			JOBS_ARRAY(jobs_employer_name, this_jobs) = ""
			JOBS_ARRAY(jobs_notes, this_jobs) = ""
		End If
	Loop until ButtonPressed = return_btn
	If JOBS_ARRAY(jobs_employee_name, this_jobs) = "Select One..." Then JOBS_ARRAY(jobs_employee_name, this_jobs) = ""

end function

function format_phone_number(phone_variable, format_type)
'This function formats phone numbers to match the specificed format.
	' format_type_options:
	'  (xxx)xxx-xxxx
	'  xxx-xxx-xxxx
	'  xxx xxx xxxx
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) = 10 Then
		left_phone = left(phone_variable, 3)
		mid_phone = mid(phone_variable, 4, 3)
		right_phone = right(phone_variable, 4)
		format_type = lcase(format_type)
		If format_type = "(xxx)xxx-xxxx" Then
			phone_variable = "(" & left_phone & ")" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx-xxx-xxxx" Then
			phone_variable = left_phone & "-" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx xxx xxxx" Then
			phone_variable = left_phone & " " & mid_phone & " " & right_phone
		End If
	Else
		phone_variable = original_phone_var
	End If
end function

function validate_phone_number(err_msg_variable, list_delimiter, phone_variable, allow_to_be_blank)
'This isn't working yet
'This function will review to ensure a variale appears to be a phone number.
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) <> 10 Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " Phone numbers should be entered as a 10 digit number. Please incldue the area code or check the number to ensure the correct information is entered."
	If len(phone_variable) = 0 then
		If allow_to_be_blank = TRUE then err_msg_variable = ""
	End If
	phone_variable = original_phone_var
end function

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

const memb_ref_numb                 = 00
const memb_last_name                = 01
const memb_first_name               = 02
const memb_mid_name					= 03
const memb_other_names				= 04
const memb_age                      = 05
const memb_remo_checkbox            = 06
const memb_new_checkbox             = 07
const clt_grh_status                = 08
const clt_hc_status                 = 09
const clt_snap_status               = 10
const memb_id_verif                 = 11
const memb_soc_sec_numb             = 12
const memb_ssn_verif                = 13
const memb_dob                      = 14
const memb_dob_verif                = 15
const memb_gender                   = 16
const memb_rel_to_applct            = 17
const memb_spoken_language          = 18
const memb_written_language         = 19
const memb_interpreter              = 20
const memb_alias                    = 21
const memb_ethnicity                = 22
const memb_race                     = 23
const memb_race_a_checkbox			= 24
const memb_race_b_checkbox			= 25
const memb_race_n_checkbox			= 26
const memb_race_p_checkbox			= 27
const memb_race_w_checkbox			= 28
const memi_marriage_status          = 29
const memi_spouse_ref               = 30
const memi_spouse_name              = 31
const memi_designated_spouse        = 32
const memi_marriage_date            = 33
const memi_marriage_verif           = 34
const memi_citizen                  = 35
const memi_citizen_verif            = 36
const memi_last_grade               = 37
const memi_in_MN_less_12_mo         = 38
const memi_resi_verif               = 39
const memi_MN_entry_date            = 40
const memi_former_state             = 41
const memi_other_FS_end             = 42
const clt_snap_checkbox				= 43
const clt_cash_checkbox				= 44
const clt_emer_checkbox				= 45
const clt_none_checkbox 			= 46
const clt_nav_btn					= 47
const clt_intend_to_reside_mn		= 48
const clt_imig_status				= 49
const clt_sponsor_yn 				= 50
const clt_verif_yn					= 51
const clt_verif_details				= 52
const memb_notes                    = 53

const jobs_employee_name 			= 0
const jobs_hourly_wage 				= 1
const jobs_gross_monthly_earnings	= 2
const jobs_employer_name 			= 3
const jobs_edit_btn					= 4
const jobs_notes 					= 5

Const end_of_doc = 6			'This is for word document ennumeration

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
Dim JOBS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)
ReDim JOBS_ARRAY(jobs_notes, 0)

'These are all the definitions for droplists
state_list = "Select One..."
state_list = state_list+chr(9)+"AL Alabama"
state_list = state_list+chr(9)+"AK Alaska"
state_list = state_list+chr(9)+"AZ Arizona"
state_list = state_list+chr(9)+"AR Arkansas"
state_list = state_list+chr(9)+"CA California"
state_list = state_list+chr(9)+"CO Colorado"
state_list = state_list+chr(9)+"CT Connecticut"
state_list = state_list+chr(9)+"DE Delaware"
state_list = state_list+chr(9)+"DC District Of Columbia"
state_list = state_list+chr(9)+"FL Florida"
state_list = state_list+chr(9)+"GA Georgia"
state_list = state_list+chr(9)+"HI Hawaii"
state_list = state_list+chr(9)+"ID Idaho"
state_list = state_list+chr(9)+"IL Illnois"
state_list = state_list+chr(9)+"IN Indiana"
state_list = state_list+chr(9)+"IA Iowa"
state_list = state_list+chr(9)+"KS Kansas"
state_list = state_list+chr(9)+"KY Kentucky"
state_list = state_list+chr(9)+"LA Louisiana"
state_list = state_list+chr(9)+"ME Maine"
state_list = state_list+chr(9)+"MD Maryland"
state_list = state_list+chr(9)+"MA Massachusetts"
state_list = state_list+chr(9)+"MI Michigan"
state_list = state_list+chr(9)+"MN Minnesota"
state_list = state_list+chr(9)+"MS Mississippi"
state_list = state_list+chr(9)+"MO Missouri"
state_list = state_list+chr(9)+"MT Montana"
state_list = state_list+chr(9)+"NE Nebraska"
state_list = state_list+chr(9)+"NV Nevada"
state_list = state_list+chr(9)+"NH New Hampshire"
state_list = state_list+chr(9)+"NJ New Jersey"
state_list = state_list+chr(9)+"NM New Mexico"
state_list = state_list+chr(9)+"NY New York"
state_list = state_list+chr(9)+"NC North Carolina"
state_list = state_list+chr(9)+"ND North Dakota"
state_list = state_list+chr(9)+"OH Ohio"
state_list = state_list+chr(9)+"OK Oklahoma"
state_list = state_list+chr(9)+"OR Oregon"
state_list = state_list+chr(9)+"PA Pennsylvania"
state_list = state_list+chr(9)+"RI Rhode Island"
state_list = state_list+chr(9)+"SC South Carolina"
state_list = state_list+chr(9)+"SD South Dakota"
state_list = state_list+chr(9)+"TN Tennessee"
state_list = state_list+chr(9)+"TX Texas"
state_list = state_list+chr(9)+"UT Utah"
state_list = state_list+chr(9)+"VT Vermont"
state_list = state_list+chr(9)+"VA Virginia"
state_list = state_list+chr(9)+"WA Washington"
state_list = state_list+chr(9)+"WV West Virginia"
state_list = state_list+chr(9)+"WI Wisconsin"
state_list = state_list+chr(9)+"WY Wyoming"
state_list = state_list+chr(9)+"PR Puerto Rico"
state_list = state_list+chr(9)+"VI Virgin Islands"

memb_panel_relationship_list = "Select One..."
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"01 Applicant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"02 Spouse"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"03 Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"04 Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"05 Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"06 Step Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"08 Step Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"09 Step Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"10 Aunt"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"11 Uncle"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"12 Niece"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"13 Nephew"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"14 Cousin"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"15 Grandparent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"16 Grandchild"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"17 Other Relative"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"18 Legal Guardian"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"24 Not Related"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"25 Live-In Attendant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"27 Unknown"

marital_status = "Select One..."
marital_status = marital_status+chr(9)+"N  Never Married"
marital_status = marital_status+chr(9)+"M  Married Living With Spouse"
marital_status = marital_status+chr(9)+"S  Married Living Apart (Sep)"
marital_status = marital_status+chr(9)+"L  Legally Sep"
marital_status = marital_status+chr(9)+"D  Divorced"
marital_status = marital_status+chr(9)+"W  Widowed"

question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"

'Dimming all the variables because they are defined and set within functions
Dim who_are_we_completing_the_form_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, exp_pay_heat_checkbox, exp_pay_ac_checkbox, exp_pay_electricity_checkbox, exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_pne_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county, caf_form_date, all_the_clients, err_msg

Dim question_1_yn, question_1_notes, question_1_verif_yn, question_1_verif_details
Dim question_2_yn, question_2_notes, question_2_verif_yn, question_2_verif_details
Dim question_3_yn, question_3_notes, question_3_verif_yn, question_3_verif_details
Dim question_4_yn, question_4_notes, question_4_verif_yn, question_4_verif_details
Dim question_5_yn, question_5_notes, question_5_verif_yn, question_5_verif_details
Dim question_6_yn, question_6_notes, question_6_verif_yn, question_6_verif_details
Dim question_7_yn, question_7_notes, question_7_verif_yn, question_7_verif_details
Dim question_8_yn, question_8a_yn, question_8_notes, question_8_verif_yn, question_8_verif_details
Dim question_9_yn, question_9_notes, question_9_verif_yn, question_9_verif_details
Dim question_10_yn, question_10_notes, question_10_verif_yn, question_10_verif_details, question_10_monthly_earnings
Dim question_11_yn, question_11_notes, question_11_verif_yn, question_11_verif_details
Dim pwe_selection
Dim question_12_yn, question_12_notes, question_12_verif_yn, question_12_verif_details
Dim question_12_rsdi_yn, question_12_rsdi_amt, question_12_ssi_yn, question_12_ssi_amt,  question_12_va_yn, question_12_va_amt, question_12_ui_yn, question_12_ui_amt, question_12_wc_yn, question_12_wc_amt, question_12_ret_yn, question_12_ret_amt, question_12_trib_yn, question_12_trib_amt, question_12_cs_yn, question_12_cs_amt, question_12_other_yn, question_12_other_amt
Dim question_13_yn, question_13_notes, question_13_verif_yn, question_13_verif_details
Dim question_14_yn, question_14_notes, question_14_verif_yn, question_14_verif_details
Dim question_14_rent_yn, question_14_subsidy_yn, question_14_mortgage_yn, question_14_association_yn, question_14_insurance_yn, question_14_room_yn, question_14_taxes_yn
Dim question_15_yn, question_15_notes, question_15_verif_yn, question_15_verif_details
Dim question_15_heat_ac_yn, question_15_electricity_yn, question_15_cooking_fuel_yn, question_15_water_and_sewer_yn, question_15_garbage_yn, question_15_phone_yn, question_15_liheap_yn
Dim question_16_yn, question_16_notes, question_16_verif_yn, question_16_verif_details
Dim question_17_yn, question_17_notes, question_17_verif_yn, question_17_verif_details
Dim question_18_yn, question_18_notes, question_18_verif_yn, question_18_verif_details
Dim question_19_yn, question_19_notes, question_19_verif_yn, question_19_verif_details
Dim question_20_yn, question_20_notes, question_20_verif_yn, question_20_verif_details
Dim question_20_cash_yn, question_20_acct_yn, question_20_secu_yn, question_20_cars_yn
Dim question_21_yn, question_21_notes, question_21_verif_yn, question_21_verif_details
Dim question_22_yn, question_22_notes, question_22_verif_yn, question_22_verif_details
Dim question_23_yn, question_23_notes, question_23_verif_yn, question_23_verif_details
Dim question_24_yn, question_24_notes, question_24_verif_yn, question_24_verif_details
Dim question_24_rep_payee_yn, question_24_guardian_fees_yn, question_24_special_diet_yn, question_24_high_housing_yn
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_there, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five

Dim show_pg_one_memb01_and_exp, show_pg_one_address, show_pg_memb_list, show_q_1_7
Dim show_q_8_13, show_q_14_19, show_q_20_24, show_qual, show_pg_last

show_pg_one_memb01_and_exp	= 1
show_pg_one_address			= 2
show_pg_memb_list			= 3
show_q_1_7					= 4
show_q_8_13					= 5
show_q_14_19				= 6
show_q_20_24				= 7
show_qual					= 8
show_pg_last				= 9

update_addr = FALSE
update_pers = FALSE
page_display = 1
'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call check_for_MAXIS(true)
Call MAXIS_case_number_finder(MAXIS_case_number)
CAF_datestamp = date & ""
interview_date = date & ""

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 235, "Interview Script Case number dialog"
  EditBox 105, 90, 60, 15, MAXIS_case_number
  EditBox 105, 110, 50, 15, CAF_datestamp
  DropListBox 105, 130, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  CheckBox 110, 165, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 150, 165, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 190, 165, 35, 10, "EMER", EMER_on_CAF_checkbox
  ButtonGroup ButtonPressed
    OkButton 260, 215, 50, 15
    CancelButton 315, 215, 50, 15
    PushButton 125, 215, 15, 15, "!", tips_and_tricks_button
  Text 10, 10, 360, 10, "Start this script at the beginning of the interview and keep it running during the entire course of the interview."
  Text 10, 20, 60, 10, "This script will:"
  Text 20, 30, 170, 10, "- Guide you through all of the interview questions."
  Text 20, 40, 170, 10, "- Capture client answers for CASE:NOTE"
  Text 20, 50, 260, 10, "- Create a document of the interview answers to be saved in the ECF Case File."
  Text 20, 60, 245, 10, "- Provide verbiage guidance for consistent resident interview experience."
  Text 20, 70, 260, 10, "- Store the interview date, time, and legth in a database (an FNS requirement)."
  Text 50, 95, 50, 10, "Case number:"
  Text 10, 115, 90, 10, "Date Application Received:"
  Text 40, 135, 60, 10, "Actual CAF Form:"
  GroupBox 105, 150, 125, 30, "Programs marked on CAF"
  Text 145, 220, 105, 10, "Look for me for Tips and Tricks!"
  Text 15, 185, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
  DropListBox 25, 195, 295, 45, "Alert at the time you attempt to leave the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
EndDialog
'Showing the case number dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
		If no_case_number_checkbox = checked Then err_msg = ""
        ' Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

show_known_addr = FALSE
vars_filled = FALSE

Call back_to_SELF
Call restore_your_work(vars_filled)			'looking for a 'restart' run
Call convert_date_into_MAXIS_footer_month(CAF_datestamp, MAXIS_footer_month, MAXIS_footer_year)
If vars_filled = TRUE Then show_known_addr = TRUE		'This is a setting for the address dialog to see the view

'If we already know the variables because we used 'restore your work' OR if there is no case number, we don't need to read the information from MAXIS
If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then

	Call back_to_SELF

	Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)			'Then we create an array of the the full hh list for looping purpoases
	full_hh_list = Split(list_for_array, chr(9))

	Call read_all_the_MEMBs

	'Now we gather the address information that exists in MAXIS
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, mail_line_one, mail_line_two, mail_addr_city, mail_addr_state, mail_addr_zip, addr_eff_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_pne_type, phone_two_type, phone_four_type)
	resi_addr_street_full = resi_line_one & " " & resi_line_two
	resi_addr_street_full = trim(resi_addr_street_full)
	mail_addr_street_full = mail_line_one & " " & mail_line_two
	mail_addr_street_full = trim(mail_addr_street_full)

	show_known_addr = TRUE
End If


'Giving the buttons specific unumerations so they don't think they are eachother
next_btn					= 1000
back_btn					= 1010
update_information_btn		= 1020
save_information_btn		= 1030
clear_mail_addr_btn			= 1040
clear_phone_one_btn			= 1041
clear_phone_two_btn			= 1042
clear_phone_three_btn		= 1043
add_person_btn				= 1050
add_verif_1_btn				= 1060
add_verif_2_btn				= 1061
add_verif_3_btn				= 1062
add_verif_4_btn				= 1063
add_verif_5_btn				= 1064
add_verif_6_btn				= 1065
add_verif_7_btn				= 1066
add_verif_8_btn				= 1070
add_verif_9_btn				= 1071
add_verif_10_btn			= 1072
add_verif_11_btn			= 1073
add_verif_12_btn			= 1074
add_verif_12_btn			= 1075
add_job_btn					= 1076
add_verif_14_btn			= 1080
add_verif_15_btn			= 1081
add_verif_16_btn			= 1082
add_verif_17_btn			= 1083
add_verif_18_btn			= 1084
add_verif_19_btn			= 1085
add_verif_20_btn			= 1090
add_verif_21_btn			= 1091
add_verif_22_btn			= 1092
add_verif_23_btn			= 1093
add_verif_24_btn			= 1094
clear_job_btn				= 1100
open_r_and_r_button 		= 1200
caf_page_one_btn			= 1300
caf_addr_btn				= 1400
caf_membs_btn				= 1500
caf_q_1_7_btn				= 1600
caf_q_8_13_btn				= 1700
caf_q_14_19_btn				= 1800
caf_q_20_24_btn				= 1900
caf_qual_q_btn				= 2000
caf_last_page_btn			= 2100
finish_interview_btn		= 2200

selected_memb = 0
' 'Presetting booleans for the dialog looping
' show_caf_pg_1_pers_dlg = TRUE
' show_caf_pg_1_addr_dlg = TRUE
' show_caf_pg_2_hhcomp_dlg = TRUE
' show_caf_pg_3_hhinfo_dlg = TRUE
' show_caf_pg_4_income_dlg = TRUE
' show_caf_pg_5_expenses_dlg = TRUE
' show_caf_pg_6_other_dlg = TRUE
' show_caf_qual_questions_dlg = TRUE
' show_caf_sig_dlg = TRUE
'
' caf_pg_1_pers_dlg_cleared = FALSE
' caf_pg_1_addr_dlg_cleared = FALSE
' caf_pg_2_hhcomp_dlg_cleared = FALSE
' caf_pg_3_hhinfo_dlg_cleared = FALSE
' caf_pg_4_income_dlg_cleared = FALSE
' caf_pg_5_expenses_dlg_cleared = FALSE
' caf_pg_6_other_dlg_cleared = FALSE
' caf_caf_qual_questions_dlg_cleared = FALSE
' caf_sig_dlg_cleared = FALSE
'
' 'This is where all of the main dialogs are called.
' 'They loop together so that you can move between all of the different dialogs.
' Do
' 	Do
' 		Do
' 			Do
' 				Do
' 					Do
' 						Do
' 							Do
' 								Do
' 									show_confirmation = TRUE
' 									If caf_pg_1_pers_dlg_cleared = FALSE Then show_caf_pg_1_pers_dlg = TRUE
' 									If caf_pg_1_addr_dlg_cleared = FALSE Then show_caf_pg_1_addr_dlg = TRUE
' 									If caf_pg_2_hhcomp_dlg_cleared = FALSE Then show_caf_pg_2_hhcomp_dlg = TRUE
' 									If caf_pg_3_hhinfo_dlg_cleared = FALSE Then show_caf_pg_3_hhinfo_dlg = TRUE
' 									If caf_pg_4_income_dlg_cleared = FALSE Then show_caf_pg_4_income_dlg = TRUE
' 									If caf_pg_5_expenses_dlg_cleared = FALSE Then show_caf_pg_5_expenses_dlg = TRUE
' 									If caf_pg_6_other_dlg_cleared = FALSE Then show_caf_pg_6_other_dlg = TRUE
' 									If caf_caf_qual_questions_dlg_cleared = FALSE Then show_caf_qual_questions_dlg = TRUE
'
' 									If caf_sig_dlg_cleared = FALSE Then show_caf_sig_dlg = TRUE
'
' 									If show_caf_pg_1_pers_dlg = TRUE Then Call dlg_page_one_pers_and_exp
'
' 								Loop until show_caf_pg_1_pers_dlg = FALSE
' 								save_your_work
' 								If show_caf_pg_1_addr_dlg = TRUE Then Call dlg_page_one_address
' 							Loop until show_caf_pg_1_addr_dlg = FALSE
' 							save_your_work
' 							If show_caf_pg_2_hhcomp_dlg = TRUE Then Call dlg_page_two_household_comp
' 						Loop until show_caf_pg_2_hhcomp_dlg = FALSE
' 						save_your_work
' 						If show_caf_pg_3_hhinfo_dlg = TRUE Then Call dlg_page_three_household_info
' 					Loop until show_caf_pg_3_hhinfo_dlg = FALSE
' 					save_your_work
' 					If show_caf_pg_4_income_dlg = TRUE Then Call dlg_page_four_income
' 				Loop until show_caf_pg_4_income_dlg = FALSE
' 				save_your_work
' 				If show_caf_pg_5_expenses_dlg = TRUE Then Call dlg_page_five_expenses
' 			Loop until show_caf_pg_5_expenses_dlg = FALSE
' 			save_your_work
' 			If show_caf_pg_6_other_dlg = TRUE Then Call dlg_page_six_assets_and_other
' 		Loop until show_caf_pg_6_other_dlg = FALSE
' 		save_your_work
' 		If show_caf_qual_questions_dlg = TRUE Then Call dlg_qualifying_questions
' 	Loop until show_caf_qual_questions_dlg = FALSE
' 	save_your_work
' 	If show_caf_sig_dlg = TRUE Then Call dlg_signature
' Loop until show_caf_sig_dlg = FALSE
' save_your_work

Do
	Do

		' MsgBox page_display
		Dialog1 = Empty
		call define_main_dialog

		err_msg = ""

		prev_page = page_display


		dialog Dialog1
		cancel_confirmation
		' MsgBox  HH_MEMB_ARRAY(0).ans_imig_status
		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

		' Call assess_imig_questions
		' call save_entered_information
		' For i = 0 to UBound(HH_MEMB_ARRAY)
		' 	If page_display = show_pg_imig Then
		' 		' HH_MEMB_ARRAY(i).save_entered_information
		' 		HH_MEMB_ARRAY(i).assess_imig_questions
		' 	End If
		' next

		If page_display <> prev_page Then
			'ADD FUNCTIONS HERE TO EVALUATE THE COMPLETION OF EACH PAGE
		End If

		' MsgBox "ButtonPressed - " & ButtonPressed

		call dialog_movement


	Loop until leave_loop = TRUE
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "CAF Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "CAF Form Details - CASE #" & MAXIS_case_number			'Title of the document
objWord.Visible = True														'Let the worker see the document

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "NOTES on INTERVIEW"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

If MAXIS_case_number <> "" Then objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information
' If no_case_number_checkbox = checked Then objSelection.TypeText "New Case - no case number" & vbCr
objSelection.TypeText "Interview Date: " & date & vbCR
objSelection.TypeText "DATE OF APPLICATION: " & application_date & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR
objSelection.TypeText "Completed over the phone with: " & who_are_we_completing_the_form_with & vbCR

'Program CAF Information
caf_progs = ""
for each_memb = 0 to UBOUND(HH_MEMB_ARRAY)
	If HH_MEMB_ARRAY(each_memb).snap_req_checkbox = checked AND InStr(caf_progs, "SNAP") = 0 Then caf_progs = caf_progs & ", SNAP"
	If HH_MEMB_ARRAY(each_memb).cash_req_checkbox = checked AND InStr(caf_progs, "Cash") = 0 Then caf_progs = caf_progs & ", Cash"
	If HH_MEMB_ARRAY(each_memb).emer_req_checkbox = checked AND InStr(caf_progs, "EMER") = 0 Then caf_progs = caf_progs & ", EMER"
Next
If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
objSelection.TypeText "CAF requesting: " & caf_progs & vbCr
objSelection.Font.Size = "11"


'Ennumeration for SetHeight and SetWidth
'wdAdjustFirstColumn	2	Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
	' wdAdjustNone			0	Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
	' wdAdjustProportional	1	Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
	' wdAdjustSameWidth		3	Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.


objSelection.TypeText "PERSON 1"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 16, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objPers1Table = objDoc.Tables(1)		'Creates the table with the specific index'
'This table will be formatted to look similar to the structure of CAF Page 1

objPers1Table.AutoFormat(16)							'This adds the borders to the table and formats it
objPers1Table.Columns(1).Width = 500					'This sets the width of the table.

for row = 1 to 15 Step 2
	objPers1Table.Cell(row, 1).SetHeight 10, 2			'setting the heights of the rows
Next
for row = 2 to 16 Step 2
	objPers1Table.Cell(row, 1).SetHeight 15, 2
Next

'Now we are going to look at the the first and second rows. These have 4 cells to add details in and we will split the row into those 4 then resize them
For row = 1 to 2
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 140, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 85, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
'Now going to each cell and setting teh font size
For col = 1 to 4
	objPers1Table.Cell(1, col).Range.Font.Size = 6
	objPers1Table.Cell(2, col).Range.Font.Size = 12
Next

'Adding the headers
objPers1Table.Cell(1, 1).Range.Text = "APPLICANT'S LEGAL NAME - LAST"
objPers1Table.Cell(1, 2).Range.Text = "FIRST NAME"
objPers1Table.Cell(1, 3).Range.Text = "MIDDLE NAME"
objPers1Table.Cell(1, 4).Range.Text = "OTHER NAMES YOU USE"

'Adding the detail from the dialog
objPers1Table.Cell(2, 1).Range.Text = HH_MEMB_ARRAY(0).last_name
objPers1Table.Cell(2, 2).Range.Text = HH_MEMB_ARRAY(0).first_name
objPers1Table.Cell(2, 3).Range.Text = HH_MEMB_ARRAY(0).mid_initial
objPers1Table.Cell(2, 4).Range.Text = HH_MEMB_ARRAY(0).other_names

' objPers1Table.Cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleNone			'commented out code dealing with borders
' objPers1Table.Cell(1, 3).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 4).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Borders(wdBorderBottom) = wdLineStyleNone

'Now formatting rows 3 and 4 - 3 is the header and 4 is the actual information
For row = 3 to 4
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 110, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 115, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
For col = 1 to 4
	objPers1Table.Cell(3, col).Range.Font.Size = 6
	objPers1Table.Cell(4, col).Range.Font.Size = 12
Next
'Adding the words to rows 3 and 4
objPers1Table.Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
objPers1Table.Cell(3, 2).Range.Text = "DATE OF BIRTH"
objPers1Table.Cell(3, 3).Range.Text = "GENDER"
objPers1Table.Cell(3, 4).Range.Text = "MARITAL STATUS"

objPers1Table.Cell(4, 1).Range.Text = HH_MEMB_ARRAY(0).ssn
objPers1Table.Cell(4, 2).Range.Text = HH_MEMB_ARRAY(0).date_of_birth
objPers1Table.Cell(4, 3).Range.Text = HH_MEMB_ARRAY(0).gender
objPers1Table.Cell(4, 4).Range.Text = HH_MEMB_ARRAY(0).marital_status

'Now formatting rows 5 and 6 - 5 is the header and 6 is the actual information
For row = 5 to 6
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(5, col).Range.Font.Size = 6
	objPers1Table.Cell(6, col).Range.Font.Size = 12
Next
'Adding the words to rows 5 and 6
objPers1Table.Cell(5, 1).Range.Text = "ADDRESS WHERE YOU LIVE"
' objPers1Table.Cell(5, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(5, 2).Range.Text = "CITY"
objPers1Table.Cell(5, 3).Range.Text = "STATE"
objPers1Table.Cell(5, 4).Range.Text = "ZIP CODE"

objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full
' objPers1Table.Cell(6, 2).Range.Text = ""
objPers1Table.Cell(6, 2).Range.Text = resi_addr_city
objPers1Table.Cell(6, 3).Range.Text = LEFT(resi_addr_state, 2)
objPers1Table.Cell(6, 4).Range.Text = resi_addr_zip

'Now formatting rows 7 and 8 - 7 is the header and 8 is the actual information
For row = 7 to 8
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(7, col).Range.Font.Size = 6
	objPers1Table.Cell(8, col).Range.Font.Size = 12
Next
'Adding the words to rows 7 and 8
objPers1Table.Cell(7, 1).Range.Text = "MAILING ADDRESS"
' objPers1Table.Cell(7, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(7, 2).Range.Text = "CITY"
objPers1Table.Cell(7, 3).Range.Text = "STATE"
objPers1Table.Cell(7, 4).Range.Text = "ZIP CODE"

objPers1Table.Cell(8, 1).Range.Text = mail_addr_street_full
' objPers1Table.Cell(8, 2).Range.Text = ""
objPers1Table.Cell(8, 2).Range.Text = mail_addr_city
objPers1Table.Cell(8, 3).Range.Text = LEFT(mail_addr_state, 2)
objPers1Table.Cell(8, 4).Range.Text = mail_addr_zip

'Now formatting rows 9 and 10 - 9 is the header and 10 is the actual information
For row = 9 to 10
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 105, 2
	objPers1Table.Cell(row, 2).SetWidth 105, 2
	objPers1Table.Cell(row, 3).SetWidth 105, 2
	objPers1Table.Cell(row, 4).SetWidth 185, 2
Next
For col = 1 to 4
	objPers1Table.Cell(9, col).Range.Font.Size = 6
	objPers1Table.Cell(10, col).Range.Font.Size = 11
Next
'Adding the words to rows 9 and 10
objPers1Table.Cell(9, 1).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 2).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 3).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 4).Range.Text = "DO YOU LIVE ON A RESERVATION?"

'formatting the phone numbers so they all match and fit
Call format_phone_number(phone_one_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_two_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_three_number, "xxx-xxx-xxxx")
If phone_pne_type = "" OR phone_pne_type = "Select One..." Then
	phone_one_info = phone_one_number
Else
	phone_one_info = phone_one_number & " (" & left(phone_pne_type, 1) & ")"
End If

If phone_two_type = "" OR phone_two_type = "Select One..." Then
	phone_two_info = phone_two_number
Else
	phone_two_info = phone_two_number & " (" & left(phone_two_type, 1) & ")"
End If
If phone_three_type = "" OR phone_three_type = "Select One..." Then
	phone_three_info = phone_three_number
Else
	phone_three_info = phone_three_number & " (" & left(phone_three_type, 1) & ")"
End If
objPers1Table.Cell(10, 1).Range.Text = phone_one_info
objPers1Table.Cell(10, 2).Range.Text = phone_two_info
objPers1Table.Cell(10, 3).Range.Text = phone_three_info
objPers1Table.Cell(10, 4).Range.Text = reservation_yn & " - " & reservation_name

'Now formatting rows 11 and 12 - 11 is the header and 12 is the actual information
For row = 11 to 12
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 190, 2
	objPers1Table.Cell(row, 3).SetWidth 190, 2
Next
For col = 1 to 3
	objPers1Table.Cell(11, col).Range.Font.Size = 6
	objPers1Table.Cell(12, col).Range.Font.Size = 12
Next
'Adding the words to rows 11 and 12
objPers1Table.Cell(11, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
objPers1Table.Cell(11, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
objPers1Table.Cell(11, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

objPers1Table.Cell(12, 1).Range.Text = HH_MEMB_ARRAY(0).interpreter
objPers1Table.Cell(12, 2).Range.Text = HH_MEMB_ARRAY(0).spoken_lang
objPers1Table.Cell(12, 3).Range.Text = HH_MEMB_ARRAY(0).written_lang

'Now formatting rows 13 and 14 - 13 is the header and 14 is the actual information
For row = 13 to 14
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 270, 2
	objPers1Table.Cell(row, 3).SetWidth 110, 2
Next
For col = 1 to 3
	objPers1Table.Cell(13, col).Range.Font.Size = 6
	objPers1Table.Cell(14, col).Range.Font.Size = 12
Next
'Adding the words to rows 13 and 14
objPers1Table.Cell(13, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
objPers1Table.Cell(13, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
objPers1Table.Cell(13, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

objPers1Table.Cell(14, 1).Range.Text = HH_MEMB_ARRAY(0).last_grade_completed
objPers1Table.Cell(14, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(0).mn_entry_date & "   From: " & HH_MEMB_ARRAY(0).former_state
objPers1Table.Cell(14, 3).Range.Text = HH_MEMB_ARRAY(0).citizen

'Now formatting rows 15 and 16 - 15 is the header and 16 is the actual information
For row = 15 to 16
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 275, 2
	objPers1Table.Cell(row, 2).SetWidth 95, 2
	objPers1Table.Cell(row, 3).SetWidth 130, 2
Next
For col = 1 to 3
	objPers1Table.Cell(15, col).Range.Font.Size = 6
	objPers1Table.Cell(16, col).Range.Font.Size = 12
Next
'Adding the words to rows 15 and 16
objPers1Table.Cell(15, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
objPers1Table.Cell(15, 2).Range.Text = "ETHNICITY"
objPers1Table.Cell(15, 3).Range.Text = "RACE"

'defining a string that lists the programs based on the checkboxes of programs from the dialog'
If HH_MEMB_ARRAY(0).none_req_checkbox = checked then progs_applying_for = "NONE"
If HH_MEMB_ARRAY(0).snap_req_checkbox = checked then progs_applying_for = progs_applying_for & ", SNAP"
If HH_MEMB_ARRAY(0).cash_req_checkbox = checked then progs_applying_for = progs_applying_for & ", Cash"
If HH_MEMB_ARRAY(0).emer_req_checkbox = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

'defining a string of the races that were selected from checkboxes in the dialog.
If HH_MEMB_ARRAY(0).race_a_checkbox = checked then race_to_enter = race_to_enter & ", Asian"
If HH_MEMB_ARRAY(0).race_b_checkbox = checked then race_to_enter = race_to_enter & ", Black"
If HH_MEMB_ARRAY(0).race_n_checkbox = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
If HH_MEMB_ARRAY(0).race_p_checkbox = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
If HH_MEMB_ARRAY(0).race_w_checkbox = checked then race_to_enter = race_to_enter & ", White"
If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

objPers1Table.Cell(16, 1).Range.Text = progs_applying_for
objPers1Table.Cell(16, 2).Range.Text = HH_MEMB_ARRAY(0).ethnicity_yn
objPers1Table.Cell(16, 3).Range.Text = race_to_enter

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

objSelection.TypeText "NOTES: " & HH_MEMB_ARRAY(0).client_notes & vbCR

objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF 1 - EXPEDITED QUESTIONS"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
set objEXPTable = objDoc.Tables(2)		'Creates the table with the specific index'

objEXPTable.AutoFormat(16)							'This adds the borders to the table and formats it
objEXPTable.Columns(1).Width = 375					'Setting the widths of the columns
objEXPTable.Columns(2).Width = 120
for col = 1 to 2
	for row = 1 to 8
		objEXPTable.Cell(row, col).Range.Font.Bold = TRUE	'Making the cell text bold.
	next
next

'Adding the Expedited text to the table for Expedited
objEXPTable.Cell(1, 1).Range.Text = "1. How much income (cash or checks) did or will your household get this month?"
objEXPTable.Cell(1, 2).Range.Text = exp_q_1_income_this_month

objEXPTable.Cell(2, 1).Range.Text = "2. How much does your household (including children) have cash, checking or savings?"
objEXPTable.Cell(2, 2).Range.Text = exp_q_2_assets_this_month

objEXPTable.Cell(3, 1).Range.Text = "3. How much does your household pay for rent/mortgage per month?"
objEXPTable.Cell(3, 2).Range.Text = exp_q_3_rent_this_month

objEXPTable.Cell(4, 1).Range.Text = "   What utilities do you pay?"
If exp_pay_heat_checkbox = checked Then util_pay = util_pay & "Heat, "
If exp_pay_ac_checkbox = checked Then util_pay = util_pay & "Air Conditioning, "
If exp_pay_electricity_checkbox = checked Then util_pay = util_pay & "Electricity, "
If exp_pay_phone_checkbox = checked Then util_pay = util_pay & "Phone, "
If exp_pay_none_checkbox = checked Then util_pay = util_pay & "NONE"
If right(util_pay, 2) = ", " Then util_pay = left(util_pay, len(util_pay) - 2)
objEXPTable.Cell(4, 2).Range.Text = util_pay

objEXPTable.Cell(5, 1).Range.Text = "4. Is anyone in your household a migrant or seasonal farm worker?"
objEXPTable.Cell(5, 2).Range.Text = exp_migrant_seasonal_formworker_yn

objEXPTable.Cell(6, 1).Range.Text = "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
objEXPTable.Cell(6, 2).Range.Text = exp_received_previous_assistance_yn

objEXPTable.Rows(7).Cells.Split 1, 6, TRUE										'Splitting the cells to add more detail for the three questions here
objEXPTable.Cell(7, 1).Range.Text = "When?"
objEXPTable.Cell(7, 2).Range.Text = exp_previous_assistance_when
objEXPTable.Cell(7, 3).Range.Text = "Where?"
objEXPTable.Cell(7, 4).Range.Text = exp_previous_assistance_where
objEXPTable.Cell(7, 5).Range.Text = "What?"
objEXPTable.Cell(7, 6).Range.Text = exp_previous_assistance_what

objEXPTable.Cell(8, 1).Range.Text = "6. Is anyone in your household pregnant?"
If exp_pregnant_who <> "" Then
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn & ", " &  exp_pregnant_who
Else
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn
End If

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.Font.Bold = TRUE
objSelection.TypeText "AGENCY USE:" & vbCr
objSelection.Font.Bold = FALSE
objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(0).intend_to_reside_in_mn & vbCr
objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(0).clt_has_sponsor & vbCr
objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(0).imig_status & vbCr
objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(0).client_verification & vbCr
If HH_MEMB_ARRAY(0).client_verification_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(0).client_verification_details & vbCr

'Now we have a dynamic number of tables
'each table has to be defined with its index so we need to have a variable to increment
table_count = 3			'table index variable
If UBound(HH_MEMB_ARRAY) <> 0 Then
	ReDim TABLE_ARRAY(UBound(HH_MEMB_ARRAY)-1)		'defining the table array for as many persons aas are in the household - each person gets their own table
	array_counters = 0		'the incrementer for the table array'

	For each_member = 1 to UBound(HH_MEMB_ARRAY)
		objSelection.TypeText "PERSON " & each_member + 1
		Set objRange = objSelection.Range										'range is needed to create tables
		objDoc.Tables.Add objRange, 10, 1										'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)			'Creates the table with the specific index - using the vairable index
		table_count = table_count + 1											'incrementing the table index'

		'This table is now formatted to match how the CAF looks with person information.
		'This formatting uses 'spliting' and resizing to make theym look like the CAF
		TABLE_ARRAY(array_counters).AutoFormat(16)								'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 500

		for row = 1 to 9 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
		Next
		for row = 2 to 10 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
		Next

		For row = 1 to 2
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 140, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 85, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 85, 2
			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
		Next
		For col = 1 to 4
			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
		Next

		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "LEGAL NAME - LAST"
		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "FIRST NAME"
		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "MIDDLE NAME"
		TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "OTHER NAMES"

		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = HH_MEMB_ARRAY(each_member).last_name
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = HH_MEMB_ARRAY(each_member).first_name
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = HH_MEMB_ARRAY(each_member).mid_initial
		TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = HH_MEMB_ARRAY(each_member).other_names

		For row = 3 to 4
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 5, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 95, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 80, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 65, 2
			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
			TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 70, 2
		Next
		For col = 1 to 5
			TABLE_ARRAY(array_counters).Cell(3, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(4, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
		TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "DATE OF BIRTH"
		TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "GENDER"
		TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "RELATIONSHIP TO YOU"
		TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "MARITAL STATUS"

		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = HH_MEMB_ARRAY(each_member).ssn
		TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = HH_MEMB_ARRAY(each_member).date_of_birth
		TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = HH_MEMB_ARRAY(each_member).gender
		TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = HH_MEMB_ARRAY(each_member).rel_to_applcnt
		TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = Left(HH_MEMB_ARRAY(each_member).marital_status, 1)

		For row = 5 to 6
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 190, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 190, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(5, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(6, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
		TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
		TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = HH_MEMB_ARRAY(each_member).interpreter
		TABLE_ARRAY(array_counters).Cell(6, 2).Range.Text = HH_MEMB_ARRAY(each_member).spoken_lang
		TABLE_ARRAY(array_counters).Cell(6, 3).Range.Text = HH_MEMB_ARRAY(each_member).written_lang

		For row = 7 to 8
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 270, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(7, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(8, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
		TABLE_ARRAY(array_counters).Cell(7, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
		TABLE_ARRAY(array_counters).Cell(7, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = HH_MEMB_ARRAY(each_member).last_grade_completed
		TABLE_ARRAY(array_counters).Cell(8, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(each_member).mn_entry_date & "   From: " & HH_MEMB_ARRAY(each_member).former_state
		TABLE_ARRAY(array_counters).Cell(8, 3).Range.Text = HH_MEMB_ARRAY(each_member).citizen

		For row = 9 to 10
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 275, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 95, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 130, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(9, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(10, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(9, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
		TABLE_ARRAY(array_counters).Cell(9, 2).Range.Text = "ETHNICITY"
		TABLE_ARRAY(array_counters).Cell(9, 3).Range.Text = "RACE"

		progs_applying_for = ""
		If HH_MEMB_ARRAY(each_member).none_req_checkbox = checked then progs_applying_for = "NONE"
		If HH_MEMB_ARRAY(each_member).snap_req_checkbox = checked then progs_applying_for = progs_applying_for & ", SNAP"
		If HH_MEMB_ARRAY(each_member).cash_req_checkbox = checked then progs_applying_for = progs_applying_for & ", Cash"
		If HH_MEMB_ARRAY(each_member).emer_req_checkbox = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
		If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

		race_to_enter = ""
		If HH_MEMB_ARRAY(each_member).race_a_checkbox = checked then race_to_enter = race_to_enter & ", Asian"
		If HH_MEMB_ARRAY(each_member).race_b_checkbox = checked then race_to_enter = race_to_enter & ", Black"
		If HH_MEMB_ARRAY(each_member).race_n_checkbox = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
		If HH_MEMB_ARRAY(each_member).race_p_checkbox = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
		If HH_MEMB_ARRAY(each_member).race_w_checkbox = checked then race_to_enter = race_to_enter & ", White"
		If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

		TABLE_ARRAY(array_counters).Cell(10, 1).Range.Text = progs_applying_for
		TABLE_ARRAY(array_counters).Cell(10, 2).Range.Text = HH_MEMB_ARRAY(each_member).ethnicity_yn
		TABLE_ARRAY(array_counters).Cell(10, 3).Range.Text = race_to_enter


		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

		objSelection.TypeText "NOTES: " & HH_MEMB_ARRAY(each_member).client_notes & vbCR
		objSelection.Font.Bold = TRUE
		objSelection.TypeText "AGENCY USE:" & vbCr
		objSelection.Font.Bold = FALSE
		objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(each_member).intend_to_reside_in_mn & vbCr
		objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(each_member).clt_has_sponsor & vbCr
		objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(each_member).imig_status & vbCr
		objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(each_member).client_verification & vbCr
		If HH_MEMB_ARRAY(each_member).client_verification_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(each_member).client_verification_details & vbCr

		array_counters = array_counters + 1
	Next
Else
	objSelection.TypeText "THERE ARE NO OTHER PEOPLE TO BE LISTED ON THIS APPLICATION" & vbCr
	ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
End If

'This is the rest of the verbiage from the CAF. It is not kept in tables - for the most part
objSelection.TypeText "Q 1. Does everyone in your household buy, fix or eat food with you?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_1_yn & vbCr
If question_1_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_1_notes & vbCr
If question_1_verif_yn <> "Mot Needed" AND question_1_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_1_verif_yn & vbCr
If question_1_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_1_verif_details & vbCr

objSelection.TypeText "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_2_yn & vbCr
If question_2_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_2_notes & vbCr
If question_2_verif_yn <> "Mot Needed" AND question_2_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_2_verif_yn & vbCr
If question_2_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_2_verif_details & vbCr

objSelection.TypeText "Q 3. Is anyone in the household attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_3_yn & vbCr
If question_3_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_3_notes & vbCr
If question_3_verif_yn <> "Mot Needed" AND question_3_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_3_verif_yn & vbCr
If question_3_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_3_verif_details & vbCr

objSelection.TypeText "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_4_yn & vbCr
If question_4_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_4_notes & vbCr
If question_4_verif_yn <> "Mot Needed" AND question_4_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_4_verif_yn & vbCr
If question_4_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_4_verif_details & vbCr

objSelection.TypeText "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_5_yn & vbCr
If question_5_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_5_notes & vbCr
If question_5_verif_yn <> "Mot Needed" AND question_5_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_5_verif_yn & vbCr
If question_5_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_5_verif_details & vbCr

objSelection.TypeText "Q 6. Is anyone unable to work for reasons other than illness or disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_6_yn & vbCr
If question_6_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_6_notes & vbCr
If question_6_verif_yn <> "Mot Needed" AND question_6_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_6_verif_yn & vbCr
If question_6_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_6_verif_details & vbCr

objSelection.TypeText "Q 7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_7_yn & vbCr
If question_7_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_7_notes & vbCr
If question_7_verif_yn <> "Mot Needed" AND question_7_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_7_verif_yn & vbCr
If question_7_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_7_verif_details & vbCr

objSelection.TypeText "Q 8. Has anyone in the household had a job or been self-employed in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_8_yn & vbCr
objSelection.TypeText "Q 8.a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?" & vbCr
objSelection.TypeText chr(9) & question_8a_yn & vbCr
If question_8_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_8_notes & vbCr
If question_8_verif_yn <> "Mot Needed" AND question_8_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_8_verif_yn & vbCr
If question_8_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_8_verif_details & vbCr

objSelection.TypeText "Q 9. Does anyone in the household have a job or expect to get income from a job this month or next month?" & vbCr

job_added = FALSE
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
		job_added = TRUE

		all_the_tables = UBound(TABLE_ARRAY) + 1
		ReDim Preserve TABLE_ARRAY(all_the_tables)
		Set objRange = objSelection.Range					'range is needed to create tables
		objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		table_count = table_count + 1

		TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 400

		TABLE_ARRAY(array_counters).Cell(1, 1).SetHeight 10, 2
		TABLE_ARRAY(array_counters).Cell(3, 1).SetHeight 10, 2

		TABLE_ARRAY(array_counters).Cell(2, 1).SetHeight 15, 2
		TABLE_ARRAY(array_counters).Cell(4, 1).SetHeight 15, 2

		For row = 1 to 2
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 200, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 90, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Font.Size = 6
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Font.Size = 12

		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "EMPLOYEE NAME"
		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "HOURLY WAGE"
		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "GROSS MONTHLY EARNINGS"
		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = JOBS_ARRAY(jobs_employee_name, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = JOBS_ARRAY(jobs_hourly_wage, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)

		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "EMPLOYER/BUSINESS NAME"
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = JOBS_ARRAY(jobs_employer_name, each_job)

		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		' objSelection.TypeParagraph()						'adds a line between the table and the next information

		array_counters = array_counters + 1

		objSelection.TypeText "NOTES: " & JOBS_ARRAY(jobs_notes, each_job) & vbCR
	End If
next

If job_added = FALSE Then objSelection.TypeText chr(9) & "THERE ARE NO JOBS ENTERED." & vbCr

If question_9_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_9_notes & vbCr
If question_9_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_9_verif_yn & vbCr
If question_9_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_9_verif_details & vbCr

objSelection.TypeText "Q 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_10_yn & vbCr
If question_10_monthly_earnings <> "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: " & question_10_monthly_earnings & vbCr
If question_10_monthly_earnings = "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: NONE LISTED" & vbCr
If question_10_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_10_notes & vbCr
If question_10_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_10_verif_yn & vbCr
If question_10_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_10_verif_details & vbCr

objSelection.TypeText "Q 11. Do you expect any changes in income, expenses or work hours?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_11_yn & vbCr
If question_11_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_11_notes & vbCr
If question_11_verif_yn <> "Mot Needed" AND question_11_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_11_verif_yn & vbCr
If question_11_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_11_verif_details & vbCr

objSelection.Font.Bold = TRUE
objSelection.TypeText "Principal Wage Earner (PWE)" & vbCr
objSelection.Font.Bold = FALSE

all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 2					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
TABLE_ARRAY(array_counters).Columns(1).Width = 200
TABLE_ARRAY(array_counters).Columns(2).Width = 200

TABLE_ARRAY(array_counters).Cell(1, 1).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(1, 2).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(2, 1).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(2, 2).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(1, 1).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Font.Size = 12
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Font.Size = 12

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text ="DESIGNATED PWE"
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text =pwe_selection
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text ="SIGNATURE OF APPLICANT"
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text ="VERBAL SIGNATURE"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

objSelection.TypeText "Q 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 5, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 500

For row = 1 to 4
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 75, 2
Next
TABLE_ARRAY(array_counters).Rows(5).Cells.Split 1, 3, TRUE

TABLE_ARRAY(array_counters).Cell(5, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(5, 2).SetWidth 175, 2
TABLE_ARRAY(array_counters).Cell(5, 3).SetWidth 75, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_12_rsdi_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "RSDI"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "$ " & question_12_rsdi_amt
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = question_12_ssi_yn
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = "SSI"
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "$ " & question_12_ssi_amt

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_12_va_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Veteran Benefits (VA)"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = "$ " & question_12_va_amt
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = question_12_ui_yn
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = "Unemployment Insurance"
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "$ " & question_12_ui_amt

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_12_wc_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Workers' Compensation"
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "$ " & question_12_wc_amt
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = question_12_ret_yn
TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "Retirement Benefits"
TABLE_ARRAY(array_counters).Cell(3, 6).Range.Text = "$ " & question_12_ret_amt

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_12_trib_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Tribal payments"
TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = "$ " & question_12_trib_amt
TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = question_12_cs_yn
TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = "Child or Spousal support"
TABLE_ARRAY(array_counters).Cell(4, 6).Range.Text = "$ " & question_12_cs_amt

TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = question_12_other_yn
TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "Other unearned income"
TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "$ " & question_12_other_amt

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_12_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_12_notes & vbCr
If question_12_verif_yn <> "Mot Needed" AND question_12_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_12_verif_yn & vbCr
If question_12_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_12_verif_details & vbCr

objSelection.TypeText "Q 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_13_yn & vbCr
If question_13_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_13_notes & vbCr
If question_13_verif_yn <> "Mot Needed" AND question_13_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_13_verif_yn & vbCr
If question_13_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_13_verif_details & vbCr

objSelection.TypeText "Q 14. Does your household have the following housing expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 3
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next
TABLE_ARRAY(array_counters).Rows(4).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(4, 1).SetWidth 90, 2
TABLE_ARRAY(array_counters).Cell(4, 2).SetWidth 430, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_14_rent_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Rent (include mobile home lot rental)"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_14_subsidy_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Rent or Section 8 subsidy"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_14_mortgage_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Mortgage/contract for deed payment"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_14_association_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Association fees"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_14_insurance_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Homeowner's insurance (if not included in mortgage) "
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = question_14_room_yn
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "Room and/or board"

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_14_taxes_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Real estate taxes (if not included in mortgage)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_14_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_14_notes & vbCr
If question_14_verif_yn <> "Mot Needed" AND question_14_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_14_verif_yn & vbCr
If question_14_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_14_verif_details & vbCr

objSelection.TypeText "Q 15. Does your household have the following utility expenses any time during the year?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 3, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 525

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 100, 2
Next
TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(3, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(3, 2).SetWidth 450, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_15_heat_ac_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Heating/air conditioning"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_15_electricity_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Electricity"
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = question_15_cooking_fuel_yn
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "Cooking fuel"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_15_water_and_sewer_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Water and sewer"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_15_garbage_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Garbage removal"
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = question_15_phone_yn
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "Phone/cell phone"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_15_liheap_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_15_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_15_notes & vbCr
If question_15_verif_yn <> "Mot Needed" AND question_15_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_15_verif_yn & vbCr
If question_15_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_15_verif_details & vbCr

objSelection.TypeText "Q 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_13_yn & vbCr
If question_16_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_16_notes & vbCr
If question_16_verif_yn <> "Mot Needed" AND question_16_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_16_verif_yn & vbCr
If question_16_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_16_verif_details & vbCr

objSelection.TypeText "Q 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_13_yn & vbCr
If question_17_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_17_notes & vbCr
If question_17_verif_yn <> "Mot Needed" AND question_17_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_17_verif_yn & vbCr
If question_17_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_17_verif_details & vbCr

objSelection.TypeText "Q 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_18_yn & vbCr
If question_18_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_18_notes & vbCr
If question_18_verif_yn <> "Mot Needed" AND question_18_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_18_verif_yn & vbCr
If question_18_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_18_verif_details & vbCr

objSelection.TypeText "Q 19. For SNAP only: Does anyone in the household have medical expenses? " & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_19_yn & vbCr
If question_19_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_19_notes & vbCr
If question_19_verif_yn <> "Mot Needed" AND question_19_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_19_verif_yn & vbCr
If question_19_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_19_verif_details & vbCr

objSelection.TypeText "Q 20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. " & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_20_cash_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Cash"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_20_acct_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Bank accounts (savings, checking, debit card, etc.)"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_20_secu_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Stocks, bonds, annuities, 401K, etc."
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_20_cars_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Vehicles (cars, trucks, motorcycles, campers, trailers)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_20_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_20_notes & vbCr
If question_20_verif_yn <> "Mot Needed" AND question_20_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_20_verif_yn & vbCr
If question_20_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_20_verif_details & vbCr


objSelection.TypeText "Q 21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_21_yn & vbCr
If question_21_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_21_notes & vbCr
If question_21_verif_yn <> "Mot Needed" AND question_21_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_21_verif_yn & vbCr
If question_21_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_21_verif_details & vbCr

objSelection.TypeText "Q 22. For recertifications only: Did anyone move in or out of your home in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_22_yn & vbCr
If question_22_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_22_notes & vbCr
If question_22_verif_yn <> "Mot Needed" AND question_22_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_22_verif_yn & vbCr
If question_22_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_22_verif_details & vbCr

objSelection.TypeText "Q 23. For children under the age of 19, are both parents living in the home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_23_yn & vbCr
If question_23_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_23_notes & vbCr
If question_23_verif_yn <> "Mot Needed" AND question_23_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_23_verif_yn & vbCr
If question_23_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_23_verif_details & vbCr

objSelection.TypeText "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_24_rep_payee_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Representative Payee fees"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_24_guardian_fees_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Guardian or Conservator fees"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_24_special_diet_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Physician-prescribed special diet "
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_24_high_housing_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "High housing costs"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_24_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_24_notes & vbCr
If question_24_verif_yn <> "Mot Needed" AND question_24_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_24_verif_yn & vbCr
If question_24_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_24_verif_details & vbCr

objSelection.TypeText "CAF QUALIFYING QUESTIONS" & vbCr

objSelection.TypeText "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?" & vbCr
objSelection.TypeText chr(9) & qual_question_one & vbCr
If trim(qual_memb_one) <> "" AND qual_memb_one <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_one & vbCr
objSelection.TypeText "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?" & vbCr
objSelection.TypeText chr(9) & qual_question_two & vbCr
If trim(qual_memb_two) <> "" AND qual_memb_two <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_two & vbCr
objSelection.TypeText "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?" & vbCr
objSelection.TypeText chr(9) & qual_question_three & vbCr
If trim(qual_memb_there) <> "" AND qual_memb_there <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_there & vbCr
objSelection.TypeText "Has anyone in your household been convicted of a drug felony in the past 10 years?" & vbCr
objSelection.TypeText chr(9) & qual_question_four & vbCr
If trim(qual_memb_four) <> "" AND qual_memb_four <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_four & vbCr
objSelection.TypeText "Is anyone in your household currently violating a condition of parole, probation or supervised release?" & vbCr
objSelection.TypeText chr(9) & qual_question_five & vbCr
If trim(qual_memb_five) <> "" AND qual_memb_five <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_five & vbCr

objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE
objSelection.TypeText "Verbal Signature accepted on " & caf_form_date

' MsgBox "DOC IS CREATED"			'This can be used for testing so we don't add fake documents to the assignment folder.

'Here we are creating the file path and saving the file
file_safe_date = replace(date, "/", "-")		'dates cannot have / for a file name so we change it to a -
'We set the file path and name based on case number and date. We can add other criteria if important.
'This MUST have the 'pdf' file extension to work

' If MAXIS_case_number <> "" Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CAF Forms for ECF\CAF - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
' If no_case_number_checkbox = checked Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CAF Forms for ECF\CAF - NEW CASE " & Left(ALL_CLIENTS_ARRAY(memb_first_name, 0), 1) & ". " & ALL_CLIENTS_ARRAY(memb_last_name, 0) & " on " & file_safe_date & ".pdf"
pdf_doc_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\TEMP - Interview Notes PDF Folder\CAF - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"

'Now we save the document.
'MS Word allows us to save directly as a PDF instead of a DOC.
'the file path must be PDF
'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
objDoc.SaveAs pdf_doc_path, 17

'Now we interact with the system again
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'This looks to see if the PDF file has been correctly saved. If it has the file will exists in the pdf file path
If objFSO.FileExists(pdf_doc_path) = TRUE Then
	'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
	objDoc.Close wdDoNotSaveChanges
	objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"
	'this is the file for the 'save your work' functionality.
	If MAXIS_case_number <> "" Then
		local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	Else
		local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"
	End If

	'we are checking the save your work text file. If it exists we need to delete it because we don't want to save that information locally.
	If objFSO.FileExists(local_changelog_path) = True then
		objFSO.DeleteFile(local_changelog_path)			'DELETE
	End If

	' 'Now we case note!
	' Call start_a_blank_case_note
	' Call write_variable_in_CASE_NOTE("CAF Form completed via Phone")
	' Call write_variable_in_CASE_NOTE("Form information taken verbally per COVID Waiver Allowance.")
	' Call write_variable_in_CASE_NOTE("Form information taken on " & caf_form_date)
	' Call write_variable_in_CASE_NOTE("CAF for application date: " & application_date)
	' Call write_variable_in_CASE_NOTE("CAF information saved and will be added to ECF within a few days. Detail can be viewed in 'Assignments Folder'.")
	' Call write_variable_in_CASE_NOTE("---")
	' Call write_variable_in_CASE_NOTE(worker_signature)

	Call write_variable_in_CASE_NOTE("THIS IS WHERE THE CASE NOTE GOES")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)


	'setting the end message
	end_msg = "Success! The information you have provided for the CAF form has been saved to the Assignments forlder so the CAF Form can be updated and added to ECF. The case can be processed using the information saved in the PDF. Additional notes and information are needed or case processing. This script has NOT updated MAXIS or added CAF processing notes."

	'Now we ask if the worker would like the PDF to be opened by the script before the script closes
	'This is helpful because they may not be familiar with where these are saved and they could work from the PDF to process the reVw
	reopen_pdf_doc_msg = MsgBox("The information about the CAF has been saved to a PDF on the LAN to be added to the DHS form and added to ECF." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbSystemModal + vbYesNo, "Open PDF Doc?")
	If reopen_pdf_doc_msg = vbYes Then
		run_path = chr(34) & pdf_doc_path & chr(34)
		wshshell.Run run_path
		end_msg = end_msg & vbCr & vbCr & "The PDF has been opened for you to view the information that has been saved."
	End If
Else
	end_msg = "Something has gone wrong - the CAF information has NOT been saved correctly to be processed." & vbCr & vbCr & "You can either save the Word Document that has opened as a PDF in the Assignment folder OR Close that document without saving and RERUN the script. Your details have been saved and the script can reopen them and attampt to create the files again. When the script is running, it is best to not interrupt the process."
End If

Call script_end_procedure_with_error_report(end_msg)
